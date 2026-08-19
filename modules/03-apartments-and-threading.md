# Module 3 — Threading and apartments

**Time: 1.5 weeks. This is the hardest module, and the source of the worst production bugs.**

Modules 1 and 2 assumed one thread. Reality has many. This module explains COM's threading model — why it exists, what it guarantees, and the precise ways it fails.

---

## 3.1 The problem apartments solve

COM objects come from arbitrary vendors. Some are thread-safe; most, historically, were not. A client has no way to know.

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
- Objects in an STA are **only ever called on that one thread**.
- Therefore the object needs **no locking at all** for its own state.
- Calls arriving from other apartments are **serialized through a hidden window**. COM creates a message-only window (class `OleMainThreadWndClass`) for each STA thread and posts/sends messages to it.
- **Consequence:** the STA thread *must pump messages*, or incoming calls never get delivered.

The hidden window is the whole mechanism. You can see it:

```powershell
# From a running STA process, in WinDbg:
#   !handle / spy++ will show windows of class OleMainThreadWndClass
```

Because delivery rides on the message queue, an STA thread that blocks in `WaitForSingleObject` is a **stopped mail truck**: calls queue up and nothing is delivered. That is the origin of most COM hangs.

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
| Instances per process | many | 1 | 1 |
| Object must be thread-safe | **no** | **yes** | **yes** |
| Requires message pump | **yes** | no | no |
| Calls serialized by COM | yes | no | no |
| Reentrancy during outbound calls | **yes** | no | no |
| Typical use | UI, legacy components | servers, worker pools | high-perf stateless |

---

## 3.3 `ThreadingModel` — where an object actually lands

This registry value on `InprocServer32` (Module 2) declares the object's contract. **It does not say where the object runs; it says where COM is allowed to put it.**

| `ThreadingModel` | Meaning | STA client creates it in… | MTA client creates it in… |
|---|---|---|---|
| *(absent)* | Legacy "single-threaded" | **The main STA** (first STA initialized in the process) | The main STA — **every call is marshaled** |
| `Apartment` | Object is STA-only | The **caller's own STA** — direct calls, fast | A **host STA** created by COM — every call marshaled |
| `Free` | Object is MTA-safe | The **MTA** — every call from the STA is marshaled | The caller's MTA — direct calls, fast |
| `Both` | Object works either way | The **caller's own STA** — direct | The **MTA** — direct |
| `Neutral` | Object is NA | The NA — direct call, no thread switch | The NA — direct call |

**Read that table twice.** It explains an enormous amount of real-world behaviour:

- A component marked `Apartment` used from a thread pool (MTA) gets a **hidden host STA** created for it, and *every single call is a cross-apartment marshaled call with a thread switch*. This is the classic "why is my component 100× slower in the service than in the test app" case.
- A component with **no** `ThreadingModel` value is worse still: everything funnels into the process's *main* STA, serializing all clients through one thread. This is why legacy components don't scale.
- `Both` is what well-written components use: the object is thread-safe *and* willing to live in an STA, so COM never needs to marshal.

> **Support reflex:** for any "component is slow / serialized / deadlocks under load" ticket, read `ThreadingModel` first, then determine the client's apartment. That 2-value lookup solves a surprising fraction of cases.

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
// THIS IS A BUG. Always.
ICalculator* g_pCalc = nullptr;

DWORD WINAPI WorkerThread(LPVOID)
{
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    long r;
    g_pCalc->Add(1, 2, &r);      // <-- raw pointer from another apartment
    CoUninitialize();
    return 0;
}
```

What actually happens depends on the object, and **that is the problem**:

- If the object is in an STA and you call it from the MTA, you bypass the serialization the object was promised. Its unsynchronized state gets corrupted — **silently, intermittently, under load only**.
- If a proxy was involved, you may get `RPC_E_WRONG_THREAD` (`0x8001010E`) — the lucky case, because it's loud.
- Frequently it *appears to work* in testing and fails in production. This is why the rule is absolute rather than conditional.

The pointer must be **marshaled**: converted into a form that can travel, then converted back into a proxy valid in the destination apartment.

### Option A — `CoMarshalInterThreadInterfaceInStream`

For a one-shot handoff to a specific thread.

> **`IStream`, briefly.** COM's standard byte-stream interface (`Read`, `Write`, `Seek`, `Commit`, …). Here it's just a transport: `CoMarshalInterThreadInterfaceInStream` serializes the marshaling data into an in-memory stream, and the destination apartment deserializes a proxy back out of it. You never read or write it yourself. It's a plain COM object, so you `Release` it like anything else — except that `CoGetInterfaceAndReleaseStream` does that for you.

```cpp
// --- Source apartment ---
IStream* pStream = nullptr;
HRESULT hr = CoMarshalInterThreadInterfaceInStream(
    IID_ICalculator, pCalc, &pStream);
// Hand pStream to the other thread (e.g. as the thread parameter).

// --- Destination apartment (after its own CoInitializeEx) ---
ICalculator* pCalcProxy = nullptr;
hr = CoGetInterfaceAndReleaseStream(
    pStream, IID_ICalculator, reinterpret_cast<void**>(&pCalcProxy));
// pStream is consumed and released by the call, success or failure.
// pCalcProxy is a PROXY valid ONLY in this apartment.
```

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

- **.NET**: RCWs are apartment-aware; the CLR marshals automatically. You mostly stop thinking about this — until you hit `InvalidCastException` on an interface with no marshaling support.
- **C++/WinRT**: `winrt::agile_ref<T>` wraps a GIT registration with RAII.

```cpp
winrt::agile_ref<ICalculator> agile{ calc };   // register
// on any other thread:
auto local = agile.get();                       // proxy for this apartment
```

### Option D — make marshaling unnecessary: agility

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

**When an STA thread makes an outbound COM call that crosses an apartment boundary, it does not simply block. It pumps messages while it waits.**

Why? Because if it blocked, and the server called back into it, you'd deadlock instantly. Pumping keeps the STA responsive so callbacks can be delivered.

The consequence is severe:

```cpp
void CMyObject::DoWork()
{
    m_state = Busy;
    m_pRemoteServer->LongOperation();   // <-- pumps messages here!
                                        //     ANY incoming call can run NOW,
                                        //     on THIS thread, RE-ENTERING this object.
    assert(m_state == Busy);            // may fail: something else ran and changed it
}
```

During `LongOperation()`, on the *same thread*:

- Another client's call into your object can execute.
- `WM_PAINT`, `WM_TIMER`, and user input can be processed.
- The user can click "Close" and your window can be destroyed **while you're inside `DoWork`**.
- A second `DoWork` can start before the first finishes.

This is not a corner case; it is normal STA behaviour. It's why Office automation code sometimes behaves bizarrely, and why UI code that calls out-of-proc COM must be written defensively.

### Defensive patterns

```cpp
void CMyObject::DoWork()
{
    if (m_inDoWork) return;              // 1. Reentrancy guard
    m_inDoWork = true;
    auto guard = wil::scope_exit([&]{ m_inDoWork = false; });

    EnableWindow(m_hwnd, FALSE);         // 2. Disable UI that could re-enter

    auto self = ComPtr<IUnknown>(this);  // 3. Keep yourself alive across the call;
                                         //    a reentrant call could Release you to zero

    m_pRemoteServer->LongOperation();

    if (!IsWindow(m_hwnd)) return;       // 4. Re-validate EVERYTHING after the call
    m_state = ReadStateAgain();          //    Cached values may be stale.
}
```

Point 3 deserves emphasis: **you can be destroyed during your own method call.** `AddRef`ing `this` for the duration of any call that might pump is standard practice in UI-adjacent COM code.

---

## 3.7 `IMessageFilter` — controlling reentrancy

`IMessageFilter` lets an STA decide what to do with incoming calls while it's busy, and what to do when its own calls are rejected. Register with `CoRegisterMessageFilter`.

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
