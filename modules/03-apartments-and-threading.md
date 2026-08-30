# Module 3 — Threading and apartments

**Time: 1.5 weeks. This is the hardest module, and the source of the worst production bugs.**

Modules 1 and 2 assumed one thread. Reality has many. This module explains COM's threading model — why it exists, what it guarantees, and the precise ways it fails.

**Contents**

- [3.1 The problem apartments solve](#31-the-problem-apartments-solve)
- [3.2 The three apartment types](#32-the-three-apartment-types)
- [3.3 `ThreadingModel` — where an object actually lands](#33-threadingmodel--where-an-object-actually-lands)
- [3.4 `CoInitializeEx` in detail](#34-coinitializeex-in-detail)
- [3.5 The cardinal rule: never pass a raw interface pointer across apartments](#35-the-cardinal-rule-never-pass-a-raw-interface-pointer-across-apartments)
- [3.6 Reentrancy — the subtlest part of COM](#36-reentrancy--the-subtlest-part-of-com)
- [3.7 `IMessageFilter` — controlling reentrancy](#37-imessagefilter--controlling-reentrancy)
- [3.8 Deadlocks](#38-deadlocks)
- [3.9 Apartment-related error codes](#39-apartment-related-error-codes)
- [3.10 LAB 3.1 — Prove that marshaling is required](#310-lab-31--prove-that-marshaling-is-required)
- [3.11 LAB 3.2 — Build a deadlock, then diagnose it in WinDbg](#311-lab-32--build-a-deadlock-then-diagnose-it-in-windbg)
- [3.12 LAB 3.3 — The ThreadingModel matrix](#312-lab-33--the-threadingmodel-matrix)
- [3.13 .NET and apartments](#313-net-and-apartments)
- [3.14 Checkpoint](#314-checkpoint)
- [3.15 Rules to carry forward](#315-rules-to-carry-forward)

---

## 3.1 The problem apartments solve

COM objects come from arbitrary vendors. Some are thread-safe; most, historically, were not. A client has no way to know.

> An object is **thread-safe** if several threads can use it **at the same time**, with no coordination between the callers, and it still behaves correctly. To manage that, the object has to protect its own internal state — with a lock, or with interlocked operations like the `InterlockedIncrement` you used on `m_cRef` in Module 1.
>
> "Not thread-safe" does **not** mean broken. It means *"I am correct only if one thread uses me at a time."* Most 1990s components were written that way deliberately: locking costs time on every call, and single-threaded was the normal case.
>
> Here is the part that matters for this module: **thread-safety is a property of the implementation, not of the interface.** Two objects can expose a byte-identical vtable, one safe and one not. No header, GUID, or type library records the difference.

Consider the simplest possible unsafe object — a counter incremented with `m_value++`.

**There is only one object.** Both threads hold a pointer to the *same* instance, so there is exactly one `m_value`, at one address in memory. What each thread owns privately is a **CPU register**, because `m_value++` is not one indivisible action — it is three:

```cpp
reg     = m_value;      // 1. read the field into this thread's register
reg     = reg + 1;      // 2. add one, in the register
m_value = reg;          // 3. write the register back to the field
```

Both threads run those three steps against that one shared field. The three right-hand columns below are the three storage locations involved — **two private registers and one shared field** — showing what each holds after every step:

```
  time │ what happens                     │ A's reg │ B's reg │ m_value
  ─────┼──────────────────────────────────┼─────────┼─────────┼─────────
    1  │ A: reg = m_value                 │    5    │    -    │    5
    2  │ B: reg = m_value                 │    5    │    5    │    5
    3  │ A: reg = reg + 1                 │    6    │    5    │    5
    4  │ B: reg = reg + 1                 │    6    │    6    │    5
    5  │ A: m_value = reg                 │    6    │    6    │    6
    6  │ B: m_value = reg                 │    6    │    6    │    6   ← should be 7
```

Look at **time 2**: thread B reads `m_value` and gets **5**, because thread A has read it but not yet written anything back. From that moment both threads are working from the same starting value. Each adds one, each arrives at 6, and each stores 6. Two increments were executed, but the shared field only advanced by one. **The correct answer was 7.**

Nothing crashes and no error is returned. The object is simply, quietly wrong. `InterlockedIncrement` fixes exactly this, by making all three steps a single indivisible operation that no other thread can interleave with.

Now do it to a **reference count** and you get Module 1's two failure modes directly: a lost `AddRef` leaves the count too low, so `delete this` runs while someone is still holding the object — a use-after-free. A lost `Release` leaves it too high, and the object never dies — a leak.

Two bad options:

1. **Assume everything is thread-safe.** Every legacy single-threaded component corrupts instantly.
2. **Lock everything globally.** Correct but catastrophically slow, and deadlock-prone.

COM's answer is a third option:

> **Let the component declare its thread-safety contract at registration time, and let the runtime enforce it automatically.**

That declared contract is the **apartment**. An apartment is not a thread and not a process — it is a **set of rules about which threads may call an object directly**.

> **The one-sentence definition:** an apartment is a group of threads within which interface pointers may be passed freely; across apartments, they may not.

---

## 3.2 The three apartment types

### STA — Single-Threaded Apartment

```cpp
CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
```

- **Exactly one thread** lives in the apartment. A process may have many STAs, each with one thread.
- **Any number of objects** may live in that one STA. An apartment is a *scope*, not a wrapper around a single object — the thread holds ordinary raw pointers to all of them and calls them directly, with no proxies and no marshaling, because they are all on its side of the boundary.
- Objects in an STA are **only ever called on that one thread**.
- Therefore the object needs **no locking at all** for its own state.
- Because they share the single thread, those objects also **serialize against one another**: a slow method on one object delays every call to every other object in the same STA, related or not.

#### How a call from another apartment gets in

**The problem.** An STA object may only ever be touched by the apartment's one thread. But a caller may be on some *other* thread — and if that thread invoked the method directly, it would break the exact guarantee the apartment exists to provide.

So the call has to **change threads** on the way in:

1. the calling thread stops and waits;
2. the **STA's own thread** runs the method;
3. the return value travels back, and the calling thread carries on.

Everything else in this section exists to make step 2 happen.

Windows already had a mechanism for handing work to another thread — the one the UI has used since 1985 — and COM reuses it wholesale: **the thread message queue**.

> Every Win32 thread can own a **message queue**: an inbox that **other** threads are allowed to drop items into, and that **the owning thread** empties whenever it chooses. It is the standard Windows way to make one thread perform work on behalf of another — which is precisely the shape of the problem above.

Reusing the queue drags in two practical details, and these are the two things people then confuse *with* the queue:

1. **You need an address to send to.** A Windows message is addressed to a **window handle**, not to a thread. So when a thread enters an STA, COM creates a window for it: hidden, message-only, class `OleMainThreadWndClass`. Its entire job is to be a valid destination whose owning thread happens to be the STA thread.
2. **A queue is only an inbox.** Something must take items out of it and act on them. That is the **message pump** — `GetMessage` + `DispatchMessage` in a loop — and it is **your** code, not COM's. This is the origin of the rule that an STA thread *must pump*.

So there are three pieces:

| Piece | Belongs to | Role |
|---|---|---|
| **Message queue** | the **thread** | the *inbox*. Messages sit here until someone collects them. |
| **Hidden window** (`OleMainThreadWndClass`) | the **apartment** | the *address*. Callers need an `HWND` to send to; this is it. |
| **Message pump** (`GetMessage` + `DispatchMessage`) | **your code** | the *engine*. Takes messages out of the queue and hands them to the target window's procedure. |

**The window is not the queue.** The window is an address that Windows resolves to a queue — specifically, the queue of the thread that created that window. Sending to the window is simply *how you get a message into that particular thread's inbox*.

End to end, a cross-apartment call makes a **round trip**:

```
   CALLER (another apartment)              ║  STA THREAD
   ════════════════════════════════════════╬══════════════════════════════════
                                           ║
   pCalc->Add(2, 3, &r)                    ║
     on a proxy                            ║
        │                                  ║
        │ 1. proxy marshals 2 and 3        ║
        ▼                                  ║
   RPC channel ─ 2. send to hidden window ─╫───►  OleMainThreadWndClass
        │                                  ║      (the STA's hidden HWND)
        │                                  ║      │
        │                                  ║      │  Windows routes it to the
        │                                  ║      │  queue of the thread that
        │                                  ║      │  owns that window
        │                                  ║      ▼
        │                                  ║      message queue
        │  the caller's thread now         ║      │
        │  BLOCKS here, waiting            ║      │  3. GetMessage +
        │  for a reply                     ║      │     DispatchMessage
        │                                  ║      │
        │                                  ║      │  No pump? It stops
        │                                  ║      │  here. Forever.
        │                                  ║      ▼
        │                                  ║      the window procedure runs
        │                                  ║      │
        │                                  ║      │  4. the stub calls
        │                                  ║      ▼     the real method
        │                                  ║      Calculator::Add(2, 3, &rStub)
        │                                  ║      │     rStub = 5
        │                                  ║      │
        │ ◄─── 5. reply: HRESULT + [out] ──╫──────┘
        ▼                                  ║
   6. proxy unmarshals, writes 5 into      ║
      the caller's own r, returns S_OK     ║
                                           ║
```

Three things about the return leg (step 5) that are easy to get wrong:

- **The reply does not travel through a message queue.** It comes back on the **same RPC channel** the request went out on. The caller is blocked *inside that channel* — which is why a hung caller shows `combase!...SendReceive` on its stack, not a wait on a window.
- **There are two `r` variables.** The server never writes to the caller's memory. The stub passes its *own* local (`rStub` above), and step 6 copies the value across. This is precisely why `[out]` must be declared in IDL: COM can only copy back what it was told about (Module 4).
- **If the caller is itself an STA, it pumps messages while blocked**, so calls can still arrive *into* it while it waits. That is where reentrancy comes from (§3.6). The reply itself still arrives on the channel, not the queue.

So "the STA isn't pumping" and "the call was never delivered" are two descriptions of the same failure: the message reached the inbox, and nobody ever collected it.

You can see that window yourself: `spy++`, or `!handle` in WinDbg, shows a window of class `OleMainThreadWndClass` on every STA thread in the process.

> **The console-app trap.** Creating a window also creates the thread's queue — so a console STA ends up with a perfectly working inbox and no message loop anywhere to empty it, and nothing warns you. An **MTA** thread gets neither: no hidden window, no queue from COM, and no obligation to pump.

**Not every call goes through the queue.** A call from *inside* the same apartment is an ordinary vtable call — direct, no marshaling, no message, no proxy, and the same cost as any C++ virtual call. The queue is only involved when a call **crosses an apartment boundary**, which is the whole reason §3.5's cardinal rule exists.

An STA thread that blocks in `WaitForSingleObject` is therefore a **stopped mail truck**: calls pile up in the inbox and nothing is delivered. That is the origin of most COM hangs.

#### The "main STA"

One STA per process is special: **the first thread to call `CoInitializeEx` with `COINIT_APARTMENTTHREADED`** creates the **main STA** (also called the main apartment).

It matters for one reason — components registered with **no `ThreadingModel` value at all** (the pre-1996 default) are always placed there, regardless of which thread creates them. So every such component in the process funnels through that one thread, and they serialize against each other even if they're completely unrelated.

Two consequences worth remembering:

- If the main STA thread exits, or calls `CoUninitialize` while those objects are alive, they're destroyed and every holder gets `RPC_E_DISCONNECTED`.
- In a GUI app the main STA is normally the UI thread — so a legacy component's work lands on the thread that's also trying to redraw your window.

`CoGetApartmentType` (§3.4) reports `APTTYPE_MAINSTA` for it.

### MTA — Multi-Threaded Apartment

```cpp
CoInitializeEx(nullptr, COINIT_MULTITHREADED);
```

- **One MTA per process**, containing **many threads**.
- Objects in the MTA can be called on **any** MTA thread, **concurrently**.
- Therefore the object **must be fully thread-safe** — it locks its own state.
- No message pumping. Calls are delivered on RPC worker threads.
- Much faster for server-side work; no window, no queue, no serialization.

### NA / TNA — Neutral Apartment

```
ThreadingModel = "Neutral"
```

- One per process. Has **no threads of its own**.
- A call from *any* apartment enters the NA on the **caller's own thread**, without a thread switch.
- The object must be thread-safe (like MTA), but callers avoid the cost of marshaling/thread-switching.
- Best of both worlds for stateless, thread-safe, high-call-volume components. Used heavily by COM+ and by system components.

### Summary

| | STA | MTA | NA |
|---|---|---|---|
| Threads per apartment | exactly 1 | many | 0 (borrows caller's) |
| **Objects per apartment** | **many** | **many** | **many** |
| Instances per process | many | 1 | 1 |
| Object must be thread-safe | **no** | **yes** | **yes** |
| Requires message pump | **yes** | no | no |
| Calls serialized by COM | yes | no | no |
| Reentrancy during outbound calls | **yes** | no | no |
| Typical use | UI, legacy components | servers, worker pools | high-perf stateless |

### Which one should *your thread* join?

This is a decision you make per thread, in the `CoInitializeEx` call. A single process routinely does both: an STA UI thread plus a pool of MTA workers.

**Join the MTA when:**

- The thread has **no UI and no message loop** — a service, a daemon, a background worker, a request handler.
- You want **real concurrency**: several threads calling the same object at once, rather than queued behind one another.
- The thread needs to **block** on kernel objects (`WaitForSingleObject`, a semaphore, I/O). An MTA thread may block freely; an STA thread that blocks stops delivering calls.
- You want to **avoid reentrancy**. An MTA call simply blocks until it returns; it does not dispatch other work in the middle of your method (§3.6). This alone removes an entire class of bug.
- The objects you use are registered `Free`, `Both`, or `Neutral` — then no thread switch is needed and calls run at full vtable speed.

**Join an STA when:**

- **The thread owns windows.** UI has thread affinity; a thread with windows must be an STA. This is not negotiable.
- You need components registered **`ThreadingModel = Apartment`** — which is most legacy, Office, and shell components. Call one of those from an MTA and COM creates a host STA for it, so *every single call* is marshaled across a thread switch. That is a common and easily-missed performance ticket.
- You depend on APIs that require an STA anyway: shell dialogs and drag-drop, MAPI, most Office automation, WinForms and WPF.
- Your own object is **not thread-safe** and you would rather have COM serialize calls than write the locking yourself.

**The rule of thumb:** MTA is the better default for server-side and background work; STA is required for UI and for the large body of legacy components that assume it. If you can choose freely and none of the STA constraints apply, choose the MTA — it is faster and it cannot surprise you with reentrancy.

> **In .NET:** thread-pool threads and `Task` continuations are **MTA**. WinForms/WPF entry points are marked `[STAThread]`, and `Thread.SetApartmentState` must be called before the thread starts. A great many "works in a console app, hangs in my service" reports are exactly this difference.

---

## 3.3 `ThreadingModel` — where an object actually lands

This registry value on `InprocServer32` (Module 2) is the object's own declaration of what it can tolerate. **It does not say where the object runs; it says where COM is allowed to put it.**

### The one rule

> COM places the object in an apartment its declared model allows — **creating one if none suitable exists**. If that apartment is not the caller's, then every call between them is marshaled.

Everything in this section follows from that single sentence. The tables below are just it, worked out.

### What each value declares

| Value | What the object is saying |
|---|---|
| *(absent)* | "I am not thread-safe, and I predate this setting existing." |
| `Apartment` | "Call me on one thread only. COM, serialize for me." |
| `Free` | "I am thread-safe. Put me in the MTA." |
| `Both` | "I am thread-safe, and I'll live in whichever apartment my caller is in." |
| `Neutral` | "I am thread-safe, and I never want a thread switch." |

### Where it lands, and what that costs

| Value | Called from an STA | Called from the MTA |
|---|---|---|
| *(absent)* | the process's **main STA** — direct *only* if you are that thread | the **main STA** — marshaled |
| `Apartment` | the **caller's own STA** — direct | a **host STA** that COM creates — marshaled |
| `Free` | the **MTA** — marshaled | the **caller's MTA** — direct |
| `Both` | the **caller's own STA** — direct | the **MTA** — direct |
| `Neutral` | the **NA** — direct | the **NA** — direct |

Read one row aloud to check it has landed: *"`Free`, called from an STA — COM puts the object in the MTA, that is not where I am, so every call gets marshaled."*

### The three rows that cause real tickets

- **`Apartment`, called from the MTA.** COM creates a hidden host STA, and every single call becomes a marshaled call with a thread switch. This is the classic *"why is my component 100× slower in the service than in my test app"*.
- **No value at all.** Worse: every such object in the process funnels into the *one* main STA and serializes there, however unrelated they are. This is why legacy components don't scale.
- **`Both`.** What well-written components use — thread-safe, yet willing to live in an STA, so COM never has to marshal in either direction.

> **Support reflex:** for any "slow / serialized / deadlocks under load" ticket, read `ThreadingModel` first, then work out the client's apartment. That two-value lookup resolves a surprising share of cases.

---

## 3.4 `CoInitializeEx` in detail

```cpp
HRESULT CoInitializeEx(LPVOID pvReserved, DWORD dwCoInit);
```

Rules that matter:

- **Per thread.** Every thread that touches COM must call it. `CO_E_NOTINITIALIZED` (`0x800401F0`) is the penalty for forgetting.
- **Returns `S_FALSE`** if the thread is already initialized *with the same model*. That is **success** — and you still owe a `CoUninitialize`.
- **Returns `RPC_E_CHANGED_MODE` (`0x80010106`)** if the thread is already initialized with a *different* model. This is a **failure**, and you do **not** owe a `CoUninitialize`. Getting this wrong causes an unbalanced uninit, which tears COM down under other code's feet.

The correct RAII wrapper:

```cpp
class ComInit
{
    HRESULT m_hr;
public:
    explicit ComInit(DWORD model = COINIT_APARTMENTTHREADED)
        : m_hr(CoInitializeEx(nullptr, model)) {}
    ~ComInit() { if (SUCCEEDED(m_hr)) CoUninitialize(); }   // S_FALSE counts as success
    HRESULT hr() const { return m_hr; }
    ComInit(const ComInit&) = delete;
    ComInit& operator=(const ComInit&) = delete;
};
```

Note `SUCCEEDED(m_hr)`, not `m_hr == S_OK`. Both `S_OK` and `S_FALSE` require a matching `CoUninitialize`; `RPC_E_CHANGED_MODE` does not. WIL's `wil::unique_couninitialize_call` / `wil::CoInitializeEx` implement exactly this.

- `CoInitialize(nullptr)` is `CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED)`. Prefer the explicit form.
- **`CoUninitialize` on the last thread of an STA destroys all objects in it.** Calling it too early, or forgetting a matching pair, produces `RPC_E_DISCONNECTED` for anyone still holding pointers.

### Which apartment am I actually in?

`CoGetApartmentType` answers this at runtime. It is one of the most useful diagnostic calls in COM and almost nobody knows it exists.

```cpp
APTTYPE          type;
APTTYPEQUALIFIER qualifier;
HRESULT hr = CoGetApartmentType(&type, &qualifier);
```

| `APTTYPE` | Meaning |
|---|---|
| `APTTYPE_STA` | An ordinary single-threaded apartment |
| `APTTYPE_MAINSTA` | The **main** STA (§3.2) |
| `APTTYPE_MTA` | The multi-threaded apartment |
| `APTTYPE_NA` | The neutral apartment |

| `APTTYPEQUALIFIER` | Meaning |
|---|---|
| `APTTYPEQUALIFIER_NONE` | Explicitly initialized, as you'd expect |
| `APTTYPEQUALIFIER_IMPLICIT_MTA` | **Never called `CoInitializeEx`** — in the MTA by default. See below. |
| `APTTYPEQUALIFIER_NA_ON_STA` / `_MTA` / `_MAINSTA` / `_IMPLICIT_MTA` | Currently executing inside the NA, having entered from that apartment type |

A drop-in probe worth keeping in your toolkit:

```cpp
void DumpApartment(const char* tag)
{
    APTTYPE t; APTTYPEQUALIFIER q;
    HRESULT hr = CoGetApartmentType(&t, &q);
    printf("[%s] thread %lu  hr=0x%08X  type=%d  qualifier=%d\n",
           tag, GetCurrentThreadId(), hr, (int)t, (int)q);
}
```

It returns `CO_E_NOTINITIALIZED` on a thread with no apartment and no implicit MTA — which is itself a useful answer.

> **Use it constantly.** It settles in one call the question Lab 3.3's whole matrix is about, and it's the fastest way to check a customer's *claim* about which apartment their code runs on against reality. Claims and reality differ more often than you'd expect.

### The implicit MTA — a genuine trap

A thread that has **never** called `CoInitializeEx` is not necessarily apartment-less. If the MTA already exists in the process — because some *other* thread joined it — then such a thread is treated as being in the MTA **implicitly**.

Why that hurts:

- COM calls may "just work" on a thread you never initialized — right up until the last explicit MTA thread leaves, at which point they stop working.
- The behaviour therefore depends on **which unrelated thread happened to start first**. That's the source of the maddening class of bug that appears only under specific timing, or only on certain machines.
- The tell is `CoGetApartmentType` reporting `APTTYPE_MTA` with qualifier `APTTYPEQUALIFIER_IMPLICIT_MTA`. If you see that in a diagnostic, someone forgot a `CoInitializeEx`.

### Keeping the MTA alive without a dedicated thread

Historically, keeping the MTA alive meant spawning a thread that called `CoInitializeEx(nullptr, COINIT_MULTITHREADED)` and then slept forever. Windows 8 added the proper API:

```cpp
CO_MTA_USAGE_COOKIE cookie = nullptr;
CoIncrementMTAUsage(&cookie);      // the MTA now exists; no thread wasted
// ...
CoDecrementMTAUsage(cookie);
```

Use this in a DLL that needs the MTA to be available but has no business dictating the host process's threading model.

---

## 3.5 The cardinal rule: never pass a raw interface pointer across apartments

```cpp
// On the main thread:
//     CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);       // main thread joins an STA
//     CoCreateInstance(CLSID_Calculator, ..., (void**)&g_pCalc);
// The object is registered ThreadingModel = Apartment, so it now lives
// in THAT thread's STA and may only ever be called on THAT thread.

// THIS IS A BUG. Always.
ICalculator* g_pCalc = nullptr;

DWORD WINAPI WorkerThread(LPVOID)
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);   // a DIFFERENT apartment
    long r;
    g_pCalc->Add(1, 2, &r);      // <-- raw pointer from another apartment
    CoUninitialize();
    return 0;
}
```

### Why this fails

`g_pCalc` is a **direct pointer to the object's vtable**. So `g_pCalc->Add(1, 2, &r)` is an ordinary C++ virtual call: load a function address out of the vtable, jump to it. **No COM code runs at all.**

That is the entire problem, and it is worth stating precisely: COM is not being *tricked* here, COM is simply **never invoked**. Nothing checks the thread, because nothing is running that could check.

Put that next to the round-trip diagram in §3.2. Every step in it — the proxy, the message, the window, the queue, the thread switch — happens only because a **proxy** stood between the caller and the object. Copy the raw pointer instead and all of it is skipped: `Calculator::Add` simply executes **on the worker thread**.

#### What is broken

`ThreadingModel = Apartment` is the object saying *"I am not thread-safe."* COM's half of that deal is *"then I will make sure only one thread ever calls you."*

COM keeps its promise **using the proxy**. Go around the proxy and there is nothing left: the object has **no locks of its own**, because it was told it would never need any.

#### What it costs you

| Corrupted | Symptom |
|---|---|
| The object's fields | Wrong answers; bad data written to files or a database |
| Its reference count (often plain `++`, not `InterlockedIncrement`) | Random `0xC0000005`, or a leak — Module 1's two failure modes |
| Its internal lists and buffers | Access violations deep inside the component |
| A window, or **thread-local storage**, it set up on the STA thread | Silently stops working — the other thread sees a different, empty slot |

> **Thread-local storage (TLS)** is memory that is private to each thread: one variable name, a separate value per thread (`thread_local` in C++, `TlsAlloc`/`TlsGetValue` in Win32). Components use it to cache per-thread state. Arrive on a different thread and you read *that* thread's slot instead — usually empty, occasionally someone else's data.

And the call still returns `S_OK`. Every time.

### Why it is so hard to catch

What happens next depends on something invisible at the call site — **what that pointer actually pointed at**:

| What `g_pCalc` really held | Result |
|---|---|
| A **raw pointer** to the real object (in-proc server, same process) | Nothing detects anything. The method runs on the wrong thread and corrupts state **silently, intermittently, and only under load**. |
| A **proxy** (the object was already in another apartment or process) | Proxies *do* enforce thread affinity — you get `RPC_E_WRONG_THREAD` (`0x8001010E`). This is the **lucky** case, because it is loud and immediate. |

And in testing it will usually appear to work: a single-threaded test never puts two threads inside the object at once, so the corruption never materializes. It surfaces in production, under concurrency, as data corruption with no stack trace pointing anywhere near this line.

**That is why the rule is absolute rather than conditional.** "It worked when I tried it" carries no information here.

### The fix, in one sentence

The pointer must be **marshaled**: converted into a form that can travel between apartments, then converted back into a **proxy** that is valid in the destination apartment.

A proxy looks identical to the caller — same interface, same method names, same calling code. The only difference is what happens *inside* the call: instead of jumping straight into the object, it runs the round trip from §3.2. So the fix does not change how you use the pointer. It changes **how you obtain one on the other thread.**

```cpp
// WRONG                                    // RIGHT
g_pCalc->Add(1, 2, &r);                     ICalculator* pLocal = /* marshal it here */;
//  ^ a pointer that belongs                pLocal->Add(1, 2, &r);
//    to another apartment                  pLocal->Release();
//                                          //  ^ a proxy that belongs to THIS apartment
```

### Four ways to get that proxy — and which one you actually need

Do not read all four with equal attention on a first pass:

| | Approach | Use it when | On a first read |
|---|---|---|---|
| **A** | `CoMarshalInterThreadInterfaceInStream` | a one-shot handoff to one specific thread | skim — it is the primitive the others are built on |
| **B** | **The Global Interface Table (GIT)** | a pointer used repeatedly, from many threads | **read properly** — the everyday answer in C++ |
| **C** | Let the framework do it | you are in .NET or C++/WinRT | **read properly** — the everyday answer everywhere else |
| **D** | Agility (FTM / `IAgileObject`) | you are *writing* the component | come back when you author one |

> **If you take one thing from this section:** share the **GIT cookie**, never the pointer. In C++/WinRT share a `winrt::agile_ref`. In .NET the runtime already does it for you.

### Option A — `CoMarshalInterThreadInterfaceInStream`

For a one-shot handoff to a specific thread.

> **`IStream`, briefly.** COM's standard byte-stream interface (`Read`, `Write`, `Seek`, `Commit`, …). Here it's just a transport: `CoMarshalInterThreadInterfaceInStream` serializes the marshaling data into an in-memory stream, and the destination apartment deserializes a proxy back out of it. You never read or write it yourself. It's a plain COM object, so you `Release` it like anything else — except that `CoGetInterfaceAndReleaseStream` does that for you.

**What is actually in that stream?** Not the object, and not a copy of it. The call does two things: it creates a **stub** in the *source* apartment — the piece that will receive marshaled calls and invoke the real method — and it writes a small descriptor of that stub (an **OBJREF**: which apartment, which stub, which interface) into the stream. The object itself never moves.

```cpp
DWORD WINAPI WorkerThread(LPVOID pv);

// ─── main thread — the SOURCE apartment (an STA) ──────────────────────
int main()
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    ICalculator* pCalc = nullptr;
    CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_INPROC_SERVER,
                     IID_ICalculator, (void**)&pCalc);
    // pCalc is valid ONLY on this thread.

    // 1. Package it. Creates a stub here, and writes a descriptor of that
    //    stub into a fresh in-memory stream.
    IStream* pStream = nullptr;
    HRESULT hr = CoMarshalInterThreadInterfaceInStream(IID_ICalculator, pCalc, &pStream);
    if (FAILED(hr)) { /* ... */ }

    // 2. The STREAM pointer is what may cross threads — never pCalc itself.
    HANDLE hThread = CreateThread(nullptr, 0, WorkerThread, pStream, 0, nullptr);

    // 3. This thread must keep pumping or the worker's calls never arrive (§3.2),
    //    and must stay alive, because the object lives in THIS apartment.
    ULONG index = 0;
    CoWaitForMultipleHandles(COWAIT_DEFAULT, INFINITE, 1, &hThread, &index);

    CloseHandle(hThread);
    pCalc->Release();
    CoUninitialize();
    return 0;
}

// ─── worker thread — the DESTINATION apartment (the MTA) ──────────────
DWORD WINAPI WorkerThread(LPVOID pv)
{
    IStream* pStream = static_cast<IStream*>(pv);

    CoInitializeEx(nullptr, COINIT_MULTITHREADED);      // MUST happen first

    // 4. Unpack. Builds a PROXY valid in THIS apartment, and releases the
    //    stream for you — even on failure.
    ICalculator* pProxy = nullptr;
    HRESULT hr = CoGetInterfaceAndReleaseStream(pStream, IID_ICalculator, (void**)&pProxy);
    if (FAILED(hr)) { CoUninitialize(); return 1; }

    // 5. Looks like an ordinary call. Actually performs the §3.2 round trip:
    //    marshal → main thread's queue → main thread runs Add → reply.
    long r = 0;
    pProxy->Add(1, 2, &r);        // r == 3, and Add() ran on the MAIN thread

    pProxy->Release();
    CoUninitialize();
    return 0;
}
```

**Which thread may touch what:**

| | Created on | Usable from |
|---|---|---|
| `pCalc` (the real pointer) | main thread | **main thread only** |
| `pStream` | main thread | **either thread** — this is the entire point |
| `pProxy` | worker thread | **worker thread only** |
| `Calculator::Add` | — | always executes on the **main** thread |

Critical details:

- The stream is **single-use**. Unmarshal exactly once.
- If you never unmarshal, you leak the stub and the object — use `CoReleaseMarshalData(pStream)` to clean up an unused stream.
- `CoGetInterfaceAndReleaseStream` releases the stream **even on failure**. Don't release it yourself.
- The resulting proxy belongs to the destination apartment. Passing *it* to a third apartment is the same bug again.

### Option B — the Global Interface Table (GIT)

For a pointer used by many threads, repeatedly. This is the right tool most of the time.

```cpp
// Once per process
IGlobalInterfaceTable* g_pGIT = nullptr;
CoCreateInstance(CLSID_StdGlobalInterfaceTable, nullptr, CLSCTX_INPROC_SERVER,
                 IID_IGlobalInterfaceTable, (void**)&g_pGIT);

// --- Source apartment: register once, get a process-wide cookie ---
DWORD cookie = 0;
g_pGIT->RegisterInterfaceInGlobal(pCalc, IID_ICalculator, &cookie);

// --- ANY apartment, ANY number of times: get a proxy valid HERE ---
ICalculator* pLocal = nullptr;
g_pGIT->GetInterfaceFromGlobal(cookie, IID_ICalculator, (void**)&pLocal);
// ... use pLocal ...
pLocal->Release();

// --- When completely done ---
g_pGIT->RevokeInterfaceFromGlobal(cookie);
```

Rules:

- The **cookie** (a `DWORD`) is the thing you're allowed to share between threads — not the pointer.
- `GetInterfaceFromGlobal` returns a **new reference each time**; `Release` it.
- The GIT holds a reference until you `RevokeInterfaceFromGlobal`. **Forgetting to revoke is a permanent leak** — a very common one.
- Works even for in-proc objects, where it may hand back the raw pointer if source and destination are the same apartment. Safe either way.

### Option C — let the framework do it

- **.NET**: when managed code uses a COM object, the CLR wraps it in a **Runtime Callable Wrapper (RCW)** — an ordinary-looking .NET object that holds the real COM interface pointer, marshals your calls to it, and `Release`s it when collected. RCWs are apartment-aware, so handing one to another thread simply works: the CLR marshals for you. You mostly stop thinking about this — until you hit `InvalidCastException` on an interface with no marshaling support registered. (Module 6 covers RCWs, and their mirror image the CCW, in full.)
- **C++/WinRT**: `winrt::agile_ref<T>` wraps a GIT registration with RAII.

```cpp
winrt::agile_ref<ICalculator> agile{ calc };   // register
// on any other thread:
auto local = agile.get();                       // proxy for this apartment
```

### Option D — make marshaling unnecessary: agility

> **Component authors only.** If you are a *client* — someone handed you a pointer and you need to use it on another thread — Options A–C are your entire toolkit, and you can skip ahead to §3.6 on a first read. Come back here when you write a component of your own, or when you need to understand why someone else's component behaves the way it does.

Options A–C all *marshal*: they produce a proxy so the call can cross safely. There is a fourth approach — build an object that **needs no marshaling at all**, because it is genuinely safe to call from any apartment.

Such an object is called **agile**.

#### The Free-Threaded Marshaler (FTM)

The classic mechanism. Your object aggregates a system-supplied helper whose `IMarshal` implementation says, in effect: *"don't build a proxy — just pass the raw pointer across."*

```cpp
// Raw form, simplified
HRESULT Calculator::Init()
{
    return CoCreateFreeThreadedMarshaler(
        static_cast<IUnknown*>(static_cast<ICalculator*>(this)), &m_pFTM);
}

HRESULT Calculator::QueryInterface(REFIID riid, void** ppv)
{
    if (riid == IID_IMarshal)                     // hand marshaling to the FTM
        return m_pFTM->QueryInterface(riid, ppv);
    /* ... normal handling ... */
}
```

In ATL it's three lines:

```cpp
BEGIN_COM_MAP(CCalculator)
    COM_INTERFACE_ENTRY(ICalculator)
    COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()

DECLARE_GET_CONTROLLING_UNKNOWN()
HRESULT FinalConstruct()
{ return CoCreateFreeThreadedMarshaler(GetControllingUnknown(), &m_pUnkMarshaler); }

CComPtr<IUnknown> m_pUnkMarshaler;
```

**Three requirements. Violate any and you get a bug that is extremely hard to find:**

| Requirement | Why |
|---|---|
| The object is **genuinely thread-safe** | It will be called concurrently from many threads with zero serialization |
| Registered `ThreadingModel = Both` (or `Neutral`) | Otherwise COM's placement rules contradict the FTM's promise |
| The object holds **no apartment-affine state** | See below — this is the one people get wrong |

That last requirement is the killer. If your "agile" object stores an interface pointer to a **non-agile** object, you have smuggled apartment affinity inside it. Thread A stores a pointer that's only valid in apartment A; thread B calls your object and it uses that pointer from apartment B — the exact violation from §3.5, now concealed inside a component advertising itself as safe.

> **Rule:** an FTM object may hold pointers only to other *agile* objects. Anything else must go through the GIT or an agile reference.

#### `IAgileObject` — the modern marker

`IAgileObject` is a **marker interface** — it declares no methods. Answering `QueryInterface` for `IID_IAgileObject` with yourself asserts: *"I am safe in any apartment; do not marshal me."*

```cpp
// Client side: is this safe to hand to another thread?
CComPtr<IAgileObject> spAgile;
bool isAgile = SUCCEEDED(pObj->QueryInterface(IID_IAgileObject, (void**)&spAgile));
```

Declarative rather than mechanical, and COM/WinRT infrastructure checks for it. All WinRT objects are agile by default unless they explicitly opt out (Appendix B).

#### `IAgileReference` — the modern GIT

Windows 8.1 added a cleaner replacement for the Global Interface Table:

```cpp
CComPtr<IAgileReference> spRef;
HRESULT hr = RoGetAgileReference(AGILEREFERENCE_DEFAULT,
                                 IID_ICalculator, pCalc, &spRef);

// On ANY other thread, any number of times:
CComPtr<ICalculator> spLocal;
hr = spRef->Resolve(IID_ICalculator, reinterpret_cast<void**>(&spLocal));
```

Why it beats the GIT (Option B):

- **RAII lifetime.** Release the `IAgileReference` and you're done — no cookie to track, no `RevokeInterfaceFromGlobal` to forget, so the classic GIT leak simply cannot happen.
- No need to create and hold a GIT instance yourself.
- If the target is already agile, `Resolve` hands back the raw pointer with no proxy overhead at all.

`winrt::agile_ref` and `wil::com_agile_ref` wrap this. **For new code, prefer `IAgileReference` to the GIT.**

### Choosing between the four

| Approach | Use when | Cost / risk |
|---|---|---|
| **A** — stream marshal | One-shot handoff to one specific thread | Manual, single-use, easy to leak if never unmarshaled |
| **B** — GIT | Repeated use; older codebases | Must remember `Revoke` |
| **C** — `IAgileReference` | Repeated use; **new code** | RAII — the default choice |
| **D** — agile object / FTM | You **own** the object and it's genuinely thread-safe | No marshaling at all, so fastest — but the three requirements are strict and violations are near-invisible |

Note that A–C are things a **client** does to a pointer it was given. D is something an **author** builds into the object. If you're supporting someone else's component, you get A–C; only the vendor can choose D.

---

## 3.6 Reentrancy — the subtlest part of COM

This is the concept that separates people who "know COM" from people who can debug it.

### What reentrancy means here

> **Reentrancy** is a single thread beginning a **second** COM call before the first one has returned — so two calls are live on that one thread at the same time, one nested inside the other.

Be careful what that does *not* mean. It is **not** two threads running at once. An STA still has exactly one thread, and only one thing executes at any instant. The problem is that the two calls **interleave**: your first call is suspended part-way through, arbitrary other code runs on the same thread, and then your call resumes on top of whatever that code changed.

### Why STAs have it and MTAs do not

It comes straight out of the message pump from §3.2.

**When an STA thread makes an outbound COM call that crosses an apartment boundary, it does not simply block. It pumps messages while it waits.** And pumping means dispatching whatever is sitting in the queue — including *incoming* COM calls.

Why would it do something so dangerous? Because the alternative is worse. If it blocked, and the server called back into it, you would deadlock instantly: the callback needs the STA to pump, and the STA is busy waiting for the call that produced the callback. Pumping keeps the STA reachable.

An MTA thread has no queue and no pump at all. When it makes an outbound call it blocks in the channel and dispatches nothing; incoming calls to MTA objects are delivered on *other* RPC worker threads instead.

| | STA | MTA |
|---|---|---|
| While waiting on an outbound call | **pumps** — runs other work | **blocks** — runs nothing |
| Can a second call begin on the same thread, mid-call? | **yes** | no |
| Incoming calls arrive on | this same thread | other RPC threads |
| So your hazard is | **reentrancy** — state changes under you | **concurrency** — two threads at once |
| And your defence is | re-check state after every outbound call | locks |

**Neither model is free.** The STA spares you from writing locks and charges you reentrancy instead. The MTA removes reentrancy and charges you thread-safety. That trade-off is the centre of this whole module.

### Reentrancy in action — the user clicks Cancel

Here is a Submit handler on a UI thread — so, an STA — and the Cancel button next to it. **There is nothing wrong with either of these.** No missing lock, no misuse of COM, no clever trick. Written in an ordinary single-threaded program they would both be correct:

```cpp
void Order::Submit()                 // runs on the UI thread
{
    m_pPayments->Charge(100);        // out-of-proc call: the STA pumps while it waits
    m_submitted = true;              // record that it worked
}

void OrderDialog::OnCancel()         // the Cancel button, same UI thread
{
    DestroyWindow(m_hwnd);
    m_pOrder->Release();             // the dialog held the last reference to the Order
    m_pOrder = nullptr;
}
```

`Charge` talks to another process and takes a couple of seconds. While the user waits, they change their mind and click **Cancel**:

```
   ONE thread (the UI STA). Time runs downward.

   ┌─ Submit()  ................................. call #1
   │    Charge(100) ──────►  out-of-proc; the STA PUMPS while it waits
   │                             │
   │                             │   the user clicks Cancel → WM_COMMAND lands
   │                             │   in the queue → the pump dispatches it,
   │                             │   on THIS thread, right now
   │                             ▼
   │   ┌─ OnCancel()  ........................... runs NESTED inside Submit()
   │   │    DestroyWindow(m_hwnd)
   │   │    m_pOrder->Release()  →  ref count hits 0  →  the Order is DELETED
   │   └─ returns
   │                             │
   │    ◄────────────────────────┘   Charge finally returns
   │    m_submitted = true       ←   writes into freed memory; `this` is gone
   └─ returns
```

The object was destroyed **in the middle of its own method**, and the last line is a use-after-free — Module 1's crash family, produced without a second thread existing anywhere.

**Nothing was introduced or planted here.** The defect is not inside `Submit` at all — it is that an STA runs *other work* in the middle of an outbound call. Move the identical function to an MTA thread, where the call simply blocks and no `OnCancel` can run on it, and the code is correct.

That is the whole lesson of this section: **on an STA, every outbound cross-apartment call is a hole in your function through which arbitrary other code can run.** You cannot see the hole by reading the function.

**In a debugger you recognize it instantly:**

```
   Order::OnCancel               <-- the nested call
   DispatchMessage
   ...the pump...
   combase!CoWaitForMultipleHandles
   Payments::Charge              <-- still inside Submit
   Order::Submit                 <-- the original call
   main
```

Your code, a message pump, and then **your code again, higher up the same stack**. That sandwich is the signature of reentrancy in every dump you will ever read.

**And it is not only Cancel buttons.** While that outbound call is pumping, anything in the queue can run on this thread:

- another client's call into your object;
- `WM_PAINT`, `WM_TIMER`, and any user input;
- the user closing the window your object depends on;
- a second entry into the very method you are already inside.

This is not a corner case; it is normal STA behaviour. It is why Office automation code sometimes behaves bizarrely, and why UI code that calls out-of-proc COM has to be written defensively.

### Defending against reentrancy

You cannot switch reentrancy off. An STA that stopped pumping would deadlock instead (§3.8), so the pump is not optional — which means **every STA method that makes an outbound call has to be written to survive being interrupted part-way through**.

Four habits cover almost all of it:

```cpp
void CMyObject::DoWork()
{
    if (m_inDoWork) return;              // 1. Reentrancy guard: refuse a second entry
    m_inDoWork = true;
    auto guard = wil::scope_exit([&]{ m_inDoWork = false; });

    EnableWindow(m_hwnd, FALSE);         // 2. Disable the UI that could re-enter you

    auto self = ComPtr<IUnknown>(this);  // 3. Keep YOURSELF alive across the call
                                         //    - this is the Cancel bug above

    m_pRemoteServer->LongOperation();    //    <-- the hole: anything may run here

    if (!IsWindow(m_hwnd)) return;       // 4. Re-validate everything afterwards:
    m_state = ReadStateAgain();          //    any cached value may now be stale
}
```

Point 3 deserves emphasis: **you can be destroyed during your own method call.** `AddRef`ing `this` for the duration of any call that might pump is standard practice in UI-adjacent COM code — and it is precisely what would have saved `Order::Submit`.

> **Why `IMessageFilter` gets its own section.** Everything above is defence you write *inside your own methods*: it assumes the reentrant call has already arrived and you must cope. `IMessageFilter` (§3.7) is the other lever — it lets you tell COM **which calls to dispatch at all** while you are busy, so some never arrive. Your code defending itself, versus COM policing the door. It is separate because that interface also does a second, unrelated job: deciding what happens when your *outbound* calls are rejected by a busy server.

---

## 3.7 `IMessageFilter` — controlling reentrancy

Where §3.6's habits protect your code *after* a reentrant call has arrived, `IMessageFilter` decides **whether it arrives at all**. It lets an STA say what to do with incoming calls while it is busy — and, separately, what to do when its own outbound calls are rejected. Register it with `CoRegisterMessageFilter`.

```cpp
class MessageFilter : public IMessageFilter
{
public:
    // Called when a call ARRIVES while this STA is busy.
    DWORD STDMETHODCALLTYPE HandleInComingCall(
        DWORD dwCallType, HTASK htaskCaller, DWORD dwTickCount,
        LPINTERFACEINFO lpInterfaceInfo) override
    {
        // SERVERCALL_ISHANDLED  - process it now (allows reentrancy)
        // SERVERCALL_RETRYLATER - tell the caller to retry (we're busy)
        // SERVERCALL_REJECTED   - refuse outright
        if (m_inCriticalSection) return SERVERCALL_RETRYLATER;
        return SERVERCALL_ISHANDLED;
    }

    // Called when OUR outgoing call was rejected by the callee.
    DWORD STDMETHODCALLTYPE RetryRejectedCall(
        HTASK htaskCallee, DWORD dwTickCount, DWORD dwRejectType) override
    {
        if (dwRejectType == SERVERCALL_RETRYLATER && dwTickCount < 30000)
            return 100;      // retry in 100 ms
        return (DWORD)-1;    // give up -> caller gets RPC_E_CALL_REJECTED
    }

    // Called while we're waiting for an outgoing call and a message arrives.
    DWORD STDMETHODCALLTYPE MessagePending(
        HTASK htaskCallee, DWORD dwTickCount, DWORD dwPendingType) override
    {
        // PENDINGMSG_CANCELCALL      - abandon the outgoing call
        // PENDINGMSG_WAITNOPROCESS   - keep waiting, DON'T dispatch (recommended)
        // PENDINGMSG_WAITDEFPROCESS  - keep waiting, dispatch some messages
        return PENDINGMSG_WAITNOPROCESS;
    }
    // ... IUnknown ...
};
```

**You have seen this UI:** the Office "This action cannot be completed because the other program is busy — Switch To / Retry" dialog is OLE's *default* message filter doing `RetryRejectedCall`. When a customer reports that dialog, they are reporting a `SERVERCALL_RETRYLATER` storm.

`MessagePending` returning `PENDINGMSG_WAITNOPROCESS` is the safe default: it prevents user input from re-entering your code mid-call, which eliminates a whole class of reentrancy bugs at the cost of a briefly unresponsive window.

---

## 3.8 Deadlocks

> A **deadlock** is two or more threads each waiting for something that the other will only give up *after* it stops waiting. Neither can move, neither times out, and the process simply stops. No error is returned, because nothing failed — everything is still patiently waiting.

In ordinary multithreaded code, the thing being waited for is a lock. **In COM it usually is not.** The scarce resource is a *thread's attention*: a call into an STA can only ever be executed by that apartment's one thread (§3.2). If that thread is busy doing something else, the call cannot be delivered — and the caller waits forever.

This is what makes COM deadlocks awkward to read in a dump. You are not looking for a thread holding a mutex; you are looking for a thread that is **not pumping**, and the thing it is effectively withholding is its own message queue.

Every pattern below is the same circle in different clothes:

> **A waits for B → B needs A to pump or return → A won't, because it is waiting.**

### Pattern 1 — STA thread blocks without pumping

```cpp
// UI (STA) thread
EnterCriticalSection(&g_cs);
WaitForSingleObject(hWorkerDone, INFINITE);   // <-- NOT pumping
LeaveCriticalSection(&g_cs);

// Worker (MTA) thread
pStaObject->Notify();     // marshaled -> posted to the STA's hidden window
                          // never delivered, because the STA isn't pumping
SetEvent(hWorkerDone);    // never reached
```

Classic mutual deadlock. **Fix:** use `CoWaitForMultipleHandles`, which pumps COM messages while waiting:

```cpp
DWORD idx = 0;
HRESULT hr = CoWaitForMultipleHandles(
    COWAIT_DEFAULT,      // pump appropriately for this apartment
    INFINITE,
    1, &hWorkerDone,
    &idx);
```

On an STA it pumps COM (and, with `COWAIT_DISPATCH_WINDOW_MESSAGES`, window messages). On an MTA it degrades to a plain wait. **This function should be your default wait primitive in any code that might run on an STA.**

### Pattern 2 — lock inversion across a callback

```
UI thread:      holds g_lock  ──► calls server.LongOp()  ──► blocked in RPC
Server thread:  calls back into client ──► client's handler wants g_lock ──► blocked
```

The callback is delivered on the UI thread (it's an STA, so it pumps and dispatches), and the handler tries to take `g_lock`, which the UI thread already holds — if the lock isn't recursive, instant self-deadlock; if it is recursive, you get *unexpected reentrancy* instead, which is worse because it corrupts state silently.

**Fixes:** never hold a lock across an outbound COM call; make callbacks lock-free (post to a queue and return immediately); or use `IMessageFilter` to reject incoming calls while the lock is held.

### Pattern 3 — `DllMain` / loader lock

Never call COM from `DllMain`. `CoInitializeEx`, `CoCreateInstance`, and `Release` on a proxy can all load DLLs or wait, while you hold the loader lock. Result: a hang that shows `ntdll!LdrpDrainWorkQueue` or `LdrpLoadDll` in the stack.

### Pattern 4 — `CoUninitialize` with live objects

Tearing down an STA while other apartments hold proxies to its objects yields `RPC_E_DISCONNECTED` (`0x80010108`) for them, and can hang the uninitializing thread while COM tries to notify stubs.

---

## 3.9 Apartment-related error codes

| HRESULT | Symbol | Meaning / usual cause |
|---|---|---|
| `0x800401F0` | `CO_E_NOTINITIALIZED` | `CoInitializeEx` not called on this thread |
| `0x80010106` | `RPC_E_CHANGED_MODE` | Thread already initialized with a different apartment model |
| `0x8001010E` | `RPC_E_WRONG_THREAD` | Raw pointer used from the wrong apartment |
| `0x8001010D` | `RPC_E_CANTCALLOUT_ININPUTSYNCCALL` | Outbound COM call during a `SendMessage`-style inbound call (typically inside a window proc handling `WM_*` sent from another thread, or during drag-drop) |
| `0x8001010A` | `RPC_E_SERVERCALL_RETRYLATER` | Callee's message filter said "busy, retry" |
| `0x80010001` | `RPC_E_CALL_REJECTED` | Callee refused; retry gave up |
| `0x80010108` | `RPC_E_DISCONNECTED` | Server object/apartment is gone; proxy is stale |
| `0x8001011F` | `RPC_E_THREAD_NOT_INIT` | Thread not COM-initialized when a callback arrived |
| `0x80010005` | `RPC_E_CALL_REENTERED` | Reentrancy blocked by the message filter |

---

## 3.10 LAB 3.1 — Prove that marshaling is required

> **Requirements**
> - **Tools:** Visual Studio C++ with **ATL** (`atlbase.h` for `CComPtr`); WinDbg for the `dps` vtable comparison in the exercises.
> - **Elevation:** required once, to register the DLL.
> - **Bitness:** x64 throughout.
> - **Depends on:** the Module 1 `Calculator`, packaged as the Module 2 in-proc server and registered with **`ThreadingModel = Apartment`** — without that value COM will not enforce the STA rules this lab depends on.
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/). Change `ThreadingModel` from `"Both"` to `"Apartment"` in `DllRegisterServer`, rebuild, and re-register.
> - **Time:** ~90 min.

§3.5 claims a raw interface pointer cannot legally cross an apartment boundary. This lab proves it: the same object, called three different ways from a second thread, with only the marshaled route behaving correctly.

The payoff is in the exercises, where you print both pointers and *see* the proxy — a different address for the same object.

Build a console app. Use the `Calculator` object from Module 1 (in-proc, `ThreadingModel = Apartment` so COM enforces STA rules).

```cpp
#include <windows.h>
#include <objbase.h>
#include <atlbase.h>
#include <cstdio>

CLSID CLSID_Calculator = /* your CLSID */;

// ---------------------------------------------------------------- attempt 1
// WRONG: raw pointer shared across apartments.
struct RawArgs { ICalculator* p; };

DWORD WINAPI WorkerRaw(LPVOID pv)
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    auto* a = static_cast<RawArgs*>(pv);
    long r = 0;
    HRESULT hr = a->p->Add(1, 2, &r);
    printf("[raw]  hr=0x%08X r=%ld   <-- undefined behaviour even if it 'works'\n", hr, r);
    CoUninitialize();
    return 0;
}

// ---------------------------------------------------------------- attempt 2
// RIGHT: marshal via a stream.
DWORD WINAPI WorkerStream(LPVOID pv)
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    IStream* pStm = static_cast<IStream*>(pv);

    CComPtr<ICalculator> spProxy;
    HRESULT hr = CoGetInterfaceAndReleaseStream(
        pStm, IID_ICalculator, reinterpret_cast<void**>(&spProxy));
    if (SUCCEEDED(hr))
    {
        long r = 0;
        hr = spProxy->Add(1, 2, &r);
        printf("[strm] hr=0x%08X r=%ld  (proxy at %p)\n", hr, r, spProxy.p);
    }
    else printf("[strm] unmarshal failed 0x%08X\n", hr);
    CoUninitialize();
    return 0;
}

// ---------------------------------------------------------------- attempt 3
// RIGHT: GIT, reusable from any thread any number of times.
struct GitArgs { IGlobalInterfaceTable* pGit; DWORD cookie; };

DWORD WINAPI WorkerGit(LPVOID pv)
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    auto* a = static_cast<GitArgs*>(pv);

    for (int i = 0; i < 3; ++i)
    {
        CComPtr<ICalculator> spLocal;
        HRESULT hr = a->pGit->GetInterfaceFromGlobal(
            a->cookie, IID_ICalculator, reinterpret_cast<void**>(&spLocal));
        long r = 0;
        if (SUCCEEDED(hr)) hr = spLocal->Add(i, 100, &r);
        printf("[git]  iter %d hr=0x%08X r=%ld\n", i, hr, r);
    }
    CoUninitialize();
    return 0;
}

int main()
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);   // main thread = STA
    printf("main thread %lu is an STA\n", GetCurrentThreadId());

    CComPtr<ICalculator> spCalc;
    HRESULT hr = spCalc.CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_INPROC_SERVER);
    if (FAILED(hr)) { printf("create failed 0x%08X\n", hr); return 1; }
    printf("object/proxy in main STA = %p\n", spCalc.p);

    // 1 - the wrong way
    RawArgs raw{ spCalc.p };
    HANDLE h = CreateThread(nullptr, 0, WorkerRaw, &raw, 0, nullptr);
    WaitForSingleObject(h, INFINITE); CloseHandle(h);

    // 2 - stream marshaling
    IStream* pStm = nullptr;
    hr = CoMarshalInterThreadInterfaceInStream(IID_ICalculator, spCalc, &pStm);
    if (SUCCEEDED(hr))
    {
        h = CreateThread(nullptr, 0, WorkerStream, pStm, 0, nullptr);
        // NOTE: main thread must PUMP while waiting, or the STA can't serve the proxy.
        DWORD idx;
        CoWaitForMultipleHandles(COWAIT_DEFAULT, INFINITE, 1, &h, &idx);
        CloseHandle(h);
    }

    // 3 - GIT
    CComPtr<IGlobalInterfaceTable> spGit;
    hr = spGit.CoCreateInstance(CLSID_StdGlobalInterfaceTable, nullptr, CLSCTX_INPROC_SERVER);
    DWORD cookie = 0;
    spGit->RegisterInterfaceInGlobal(spCalc, IID_ICalculator, &cookie);

    GitArgs ga{ spGit.p, cookie };
    h = CreateThread(nullptr, 0, WorkerGit, &ga, 0, nullptr);
    DWORD idx;
    CoWaitForMultipleHandles(COWAIT_DEFAULT, INFINITE, 1, &h, &idx);
    CloseHandle(h);

    spGit->RevokeInterfaceFromGlobal(cookie);   // REQUIRED or you leak
    CoUninitialize();
    return 0;
}
```

### What to observe

1. **The proxy address differs from the object address.** Print both. `spCalc.p` in the STA vs `spProxy.p` in the MTA — different pointers to the same logical object. Use `dps` on each in WinDbg: one vtable is your `Calculator`, the other is `combase!CStdProxy`-ish. **This is location transparency made visible.**
2. **Replace `CoWaitForMultipleHandles` with `WaitForSingleObject`** in the stream case. It hangs. That's Pattern 1 from §3.8, reproduced in ten lines. Restore it and it works.
3. **Forget `RevokeInterfaceFromGlobal`** and add ref-count tracing (Module 1) — the object never destructs.
4. The raw-pointer case may print a plausible answer. **That's the trap.** Note in your journal: "appears to work" is not evidence of correctness in COM threading.

---

## 3.11 LAB 3.2 — Build a deadlock, then diagnose it in WinDbg

> **Requirements**
> - **Tools:** **WinDbg** (Microsoft Store or the SDK's *Debugging Tools for Windows*) with symbols configured — set `_NT_SYMBOL_PATH=srv*C:\Symbols*https://msdl.microsoft.com/download/symbols` **before** launching, or `~*kb` will show no `combase` frames and the lab teaches you nothing. ATL for `CComPtr`.
> - **Elevation:** not required to debug a process you launched yourself.
> - **Bitness:** x64 — and the debugger must match the target.
> - **Depends on:** Lab 3.1 (the registered `Calculator` and the GIT plumbing).
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/) registered as `Apartment`, exactly as for Lab 3.1. You do not need Lab 3.1's client — this lab's listing is self-contained.
> - **Time:** ~2 h.

This lab is the single best preparation for real hang tickets.

```cpp
CRITICAL_SECTION g_cs;
HANDLE g_hWorkerDone;
DWORD  g_git_cookie;
IGlobalInterfaceTable* g_pGit;

DWORD WINAPI Worker(LPVOID)
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    CComPtr<ICalculator> sp;
    g_pGit->GetInterfaceFromGlobal(g_git_cookie, IID_ICalculator, (void**)&sp);

    long r = 0;
    sp->Add(1, 1, &r);          // marshaled call INTO the STA -> needs the STA to pump

    SetEvent(g_hWorkerDone);
    CoUninitialize();
    return 0;
}

int main()
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    InitializeCriticalSection(&g_cs);
    g_hWorkerDone = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    // ... create object, create GIT, register cookie ...

    HANDLE h = CreateThread(nullptr, 0, Worker, nullptr, 0, nullptr);

    EnterCriticalSection(&g_cs);
    WaitForSingleObject(g_hWorkerDone, INFINITE);   // <-- STA stops pumping. HANG.
    LeaveCriticalSection(&g_cs);
    // never reached
}
```

### Diagnose it

1. Reproduce the hang. Capture a dump:
   ```powershell
   procdump -ma -o Lab32.exe C:\dumps\hang.dmp
   # or Task Manager -> right-click process -> Create dump file
   ```
2. Open in WinDbg, load symbols, then:
   ```
   0:000> ~*kb
   ```
   Look for:
   - **Thread 0 (STA):** `ntdll!NtWaitForSingleObject` → `KERNELBASE!WaitForSingleObject` → your `main`. **No `CoWaitForMultipleHandles`, no message pump.** That's the smoking gun.
   - **Worker thread:** `ntdll!NtWaitForSingleObject` → `combase!CSyncClientCall::SwitchAptAndDispatchCall` → `combase!CRpcChannelBuffer::SendReceive` → `RPCRT4!NdrpClientCall` → your proxy call.
3. Confirm the shape:
   ```
   0:000> !locks              ; who owns which critical section
   0:000> !runaway            ; CPU time per thread - all near zero = a hang, not a spin
   0:000> !uniqstack          ; deduplicated stacks, faster to scan
   ```
4. **Learn this signature by heart:**

   > *A thread inside `combase!...SendReceive`/`SwitchAptAndDispatchCall`, plus an STA thread in a non-pumping wait, equals a classic COM STA deadlock.*

5. Fix it three ways and confirm each:
   - Replace `WaitForSingleObject` with `CoWaitForMultipleHandles`.
   - Make the main thread an MTA (`COINIT_MULTITHREADED`) — no pumping needed, and the call is delivered on an RPC thread.
   - Restructure so the STA never blocks: have the worker post a message instead.

### Useful WinDbg commands for the COM part of a dump

```
!error <hr>                     ; decode an HRESULT
dps <interface-ptr> L8          ; dump the vtable -> identifies the real implementation
lmvm combase                    ; confirm symbols are loaded
~*e !clrstack                   ; if managed code is involved (SOS)
!cs -l                          ; locked critical sections and owners
!handle 0 f Event               ; event handles and their state
```

---

## 3.12 LAB 3.3 — The ThreadingModel matrix

> **Requirements**
> - **Tools:** Visual Studio C++; `regsvr32`; **five fresh GUIDs** (`New-Guid` in PowerShell, or `guidgen.exe`).
> - **Elevation:** required — five CLSID registrations.
> - **Bitness:** x64.
> - **Depends on:** the Module 2 in-proc server, with `DllGetClassObject` extended to accept all five CLSIDs and return the same factory.
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/) — extend `DllGetClassObject` and `DllRegisterServer` to cover five CLSIDs, one per `ThreadingModel`.
> - **Time:** ~2 h.

**Do this lab.** It converts §3.3's table from something you read into something you know.

1. Take the Module 2 in-proc server. Register the **same DLL** under **four different CLSIDs**, one for each `ThreadingModel`: `Apartment`, `Free`, `Both`, `Neutral`. (Add a fifth with the value absent if you want the full picture.)
   - Have `DllGetClassObject` accept all five CLSIDs and return the same factory.
2. Add this to the object's methods:

```cpp
HRESULT STDMETHODCALLTYPE Add(long a, long b, long* r) override
{
    printf("  Add() executing on thread %lu\n", GetCurrentThreadId());
    *r = a + b;
    return S_OK;
}
```

3. Write a client that, for each CLSID, does the following from **an STA thread** and from **an MTA thread**:

```cpp
printf("client thread = %lu\n", GetCurrentThreadId());
CComPtr<ICalculator> sp;
sp.CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER);
printf("  pointer = %p  (is this the object or a proxy?)\n", sp.p);
long r; sp->Add(1, 2, &r);
```

4. Fill in this matrix from what you actually observe:

| ThreadingModel | STA client: same thread? | STA client: proxy? | MTA client: same thread? | MTA client: proxy? |
|---|---|---|---|---|
| (absent) | | | | |
| Apartment | | | | |
| Free | | | | |
| Both | | | | |
| Neutral | | | | |

**How to tell a proxy from a direct pointer:** compare `Add()`'s reported thread ID with the client's thread ID. Different → marshaled. Also `dps sp.p L4` in the debugger: if the vtable symbols are in `combase`/`rpcrt4` rather than your DLL, it's a proxy.

**Better still, use `CoGetApartmentType`** (§3.4). Call `DumpApartment()` from the client *and* from inside `Add()`, and print both. That tells you the apartment type of each end directly, rather than inferring it from thread IDs — and it distinguishes `APTTYPE_MAINSTA` from an ordinary `APTTYPE_STA`, which the thread ID alone can't. Add the `(absent)` ThreadingModel row and you'll see the main-STA behaviour from §3.2 with your own eyes.

5. Measure the cost. Time 100,000 calls in each cell:

```cpp
LARGE_INTEGER f, t0, t1; QueryPerformanceFrequency(&f);
QueryPerformanceCounter(&t0);
for (int i = 0; i < 100000; ++i) { long r; sp->Add(i, 1, &r); }
QueryPerformanceCounter(&t1);
printf("  %.2f us/call\n", 1e6 * double(t1.QuadPart - t0.QuadPart) / f.QuadPart / 100000);
```

Expect roughly: direct call ≈ single-digit **nanoseconds**; cross-apartment in-proc ≈ **tens of microseconds**. That's a factor of ~1000–10000. **This number is your argument** the next time someone asks why a component is slow in a service.

Keep the completed matrix in your notes. It is the highest-value single page in this course.

---

## 3.13 .NET and apartments

| Concept | Detail |
|---|---|
| `[STAThread]` | On `Main`. Required for WinForms/WPF. Sets the main thread's apartment before COM initializes. |
| `[MTAThread]` | Default for console apps in .NET Core/5+. |
| `Thread.SetApartmentState(ApartmentState.STA)` | Must be called **before** the thread starts. |
| **Thread pool threads are MTA.** Always. | So `Task.Run(() => officeApp.DoThing())` marshals every call and can deadlock against the UI. |
| `async`/`await` | Continuations resume on the captured `SynchronizationContext`. In WPF/WinForms that's the UI (STA) thread — usually what you want with COM. `ConfigureAwait(false)` sends the continuation to the **thread pool (MTA)**, which silently changes the apartment your COM calls run in. **This is a real and common bug source.** |
| RCW | The CLR creates a per-apartment RCW; it handles marshaling for you, using the GIT internally. |
| `Marshal.ReleaseComObject` | Decrements the RCW's internal count, *not* the COM ref count directly. Module 6. |

Rule of thumb for .NET + COM: **do not use `ConfigureAwait(false)` on code paths that touch STA COM objects.**

---

## 3.14 Checkpoint

1. What exactly does an apartment guarantee, and what does it *not* guarantee?
2. A component is registered `ThreadingModel = Apartment`. A .NET service creates it from a thread-pool thread. Describe precisely what COM does, and predict the performance characteristic.
3. Why does an STA thread pump messages while waiting for an outbound COM call, and name two problems that behaviour creates.
4. `CoInitializeEx` returns `S_FALSE`. Do you owe a `CoUninitialize`? It returns `RPC_E_CHANGED_MODE`. Do you owe one then?
5. You're handed a dump. Thread 7 is in `combase!CSyncClientCall::SwitchAptAndDispatchCall`; thread 0 is in `KERNELBASE!WaitForSingleObject` under `main`. Diagnosis, and what would you ask the developer to change?
6. Why is passing a *proxy* from apartment B to apartment C just as wrong as passing the original raw pointer from A to B?
7. A developer says "I fixed the deadlock by making the wait pump messages with a `PeekMessage`/`DispatchMessage` loop." What's the risk they've introduced?
8. Name the one API you should use to register an interface pointer for repeated use from many apartments, and the one call people forget.
9. What does it mean for an object to be **agile**, and what are the three requirements for using the free-threaded marshaler safely?
10. A vendor's component uses the FTM and is registered `ThreadingModel = Both`. It stores an `IStream*` it was handed at initialization. What's wrong, and how would the failure present?
11. `CoGetApartmentType` returns `APTTYPE_MTA` with `APTTYPEQUALIFIER_IMPLICIT_MTA`. What does that tell you, and why is it fragile?

<details>
<summary>Answers</summary>

1. **Guarantees:** that an object will only be called by threads in its own apartment, so an STA object needs no locking; and that COM will interpose a proxy/stub when a call crosses a boundary. **Does not guarantee:** that the object's own outbound calls are safe, that reentrancy won't occur (STAs pump!), that the object is safe if *you* violate the rules by sharing raw pointers, or anything at all about non-COM state the object touches.

2. Thread-pool threads are MTA. The object declares STA-only, so COM creates (or reuses) a **host STA** in the process and puts the object there. Every call is then marshaled with a thread switch: ~tens of microseconds per call, and **all calls from all threads serialize through that one host STA thread**. Characteristic: fine at low load, falls off a cliff under concurrency, with worker threads piling up in `SendReceive`.

3. It pumps so that callbacks and incoming calls can still be delivered — otherwise a server calling back into the client would deadlock. Problems: (a) **reentrancy** — arbitrary code, including a second entry into the same method, can run mid-call; (b) **UI events are processed**, so windows can be destroyed and state invalidated under you; the object can even be released to zero during its own method.

4. `S_FALSE` is success — **yes**, you owe a matching `CoUninitialize`. `RPC_E_CHANGED_MODE` is a failure — **no**, and calling one would unbalance whoever legitimately initialized the thread.

5. Classic STA deadlock: thread 7 is blocked in a cross-apartment COM call waiting for the STA to service it; thread 0 is the STA, blocked in a non-pumping wait so it never will. Ask the developer to replace the blocking wait with `CoWaitForMultipleHandles`, or to restructure so the STA never blocks (post a message / use a completion callback). Also check whether a lock is held across the outbound call.

6. Because a proxy is itself an object with apartment affinity — it belongs to apartment B. Using it from C is exactly the same violation. Marshaling must happen for *every* boundary crossing; there's no transitivity.

7. A raw `PeekMessage`/`DispatchMessage` loop dispatches **all** messages, including user input. So the user can click buttons and re-enter application code at an arbitrary point mid-operation. `CoWaitForMultipleHandles` with `COWAIT_DEFAULT` pumps only what COM needs; `IMessageFilter::MessagePending` returning `PENDINGMSG_WAITNOPROCESS` is the controlled way to suppress input.

8. `IGlobalInterfaceTable::RegisterInterfaceInGlobal` (get a cookie, share the cookie). The forgotten call is `RevokeInterfaceFromGlobal` — omitting it is a permanent leak of the object. For new code, `RoGetAgileReference` is better precisely because RAII makes that mistake impossible.

9. **Agile** means the object is safe to call directly from any apartment, so COM never needs to build a proxy for it. The three FTM requirements: (a) the object must be **genuinely thread-safe**, since it will be called concurrently with no serialization; (b) it must be registered `ThreadingModel = Both` or `Neutral`; (c) it must hold **no apartment-affine state** — no pointers to non-agile objects.

10. It violates requirement (c). `IStream` is not agile, so the stored pointer is only valid in whichever apartment supplied it. Once the component is called from a second apartment — which the FTM actively encourages, since no proxy is created — it uses that pointer across an apartment boundary. Presentation: works perfectly in single-threaded testing; under concurrency produces `RPC_E_WRONG_THREAD`, or silent corruption, or intermittent `E_NOINTERFACE`. It's especially nasty because the component *advertises* itself as safe, so investigators look elsewhere first.

11. That the thread **never called `CoInitializeEx`** and is in the MTA only because some *other* thread created the MTA. It's fragile because the behaviour depends on unrelated startup ordering: COM calls work while an explicit MTA thread exists and stop when the last one leaves. Symptom shape: works on one machine or one run, fails on another, with no code difference. The fix is to initialize the thread explicitly.

</details>

---

## 3.15 Rules to carry forward

1. Every thread that touches COM calls `CoInitializeEx`, and balances it — treating `S_FALSE` as success and `RPC_E_CHANGED_MODE` as failure.
2. Never share a raw interface pointer across apartments. Marshal via stream, GIT, or — preferably in new code — `IAgileReference`.
3. `CoGetApartmentType` answers "which apartment am I really in?" in one call. Use it instead of guessing; watch for `APTTYPEQUALIFIER_IMPLICIT_MTA`.
4. An object is only agile if it is genuinely thread-safe **and** holds no apartment-affine state. The FTM does not make an object safe; it asserts that it already is.
5. `ThreadingModel` + client apartment determines whether you get a direct call or a ~1000× slower marshaled one. Check both, first.
6. `Both` is the right default for new thread-safe in-proc components.
7. An STA thread must never block without pumping — use `CoWaitForMultipleHandles`.
8. Never hold a lock across an outbound COM call.
9. Assume reentrancy on any STA call that crosses a boundary: `AddRef` yourself, re-validate all state afterwards.
10. Never call COM from `DllMain`.
11. In .NET, remember thread-pool threads are MTA, and `ConfigureAwait(false)` moves you there.
12. `SendReceive` + a non-pumping STA in the same dump = STA deadlock. Recognize it in five seconds.

---

**Next: [Module 4 — Interfaces, IDL, MIDL, and marshaling](04-idl-and-marshaling.md)**
