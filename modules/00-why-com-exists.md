# Module 0 — Why COM Exists

Before any API, one question: why does COM exist at all? The answer is a specific, concrete problem with C++ — one you can watch happen — and everything else in this course is a consequence of solving it.

**What this module covers**

Four demonstrations of why a C++ class cannot safely cross a binary boundary, followed by COM's answer to it — the three pillars of immutable interfaces, reference counting, and location transparency. Along the way you pick up the vocabulary the rest of the course assumes, take apart a working COM call three lines long, and see where COM still runs in Windows today. There is no code to write here.

**Contents**

- [0.0 Start here: COM in plain language](#00-start-here-com-in-plain-language)
- [0.1 The problem: C++ has no stable binary contract](#01-the-problem-c-has-no-stable-binary-contract)
- [0.2 The insight](#02-the-insight)
- [0.3 The three pillars](#03-the-three-pillars)
- [0.4 Vocabulary — learn these now](#04-vocabulary--learn-these-now)
- [0.5 The canonical "hello world"](#05-the-canonical-hello-world)
- [0.6 Where COM lives today (why this matters for your job)](#06-where-com-lives-today-why-this-matters-for-your-job)
- [0.7 Checkpoint](#07-checkpoint)
- [0.8 Tooling setup (do this now)](#08-tooling-setup-do-this-now)

---

## 0.0 Start here: COM in plain language

*New to COM? Read this section slowly and everything after it will land. Already comfortable with C++ binary-compatibility problems? Skim to §0.1.*

### The one-sentence version

> **COM is a standard that lets your program use code written by someone else — in a different language, compiled years earlier, possibly running in a different process or on a different machine — where neither side knows anything about the other except an agreed list of functions.**

### The analogy

Think about the electrical socket in your wall.

- The socket manufacturer and the kettle manufacturer never spoke to each other.
- Neither knows how the other's product works inside.
- Either can completely redesign their internals, and everything still works.
- A kettle from 1995 still works in a socket installed yesterday.

All of that rests on one thing: **the shape of the plug is frozen.** It's a published contract nobody is allowed to change. Everything else is free to vary.

COM is that idea applied to software. The "plug shape" is called an **interface**: a fixed, published, permanently frozen list of functions. Any component offering that interface can be used by any program that knows it — regardless of language, compiler, vintage, or location.

### What using COM actually looks like

Stripped of all syntax, every COM program does this:

```
1.  "Give me an object of type «shortcut file»,
     and let me talk to it through the «shortcut» interface."

2.  Call some functions on it.

3.  "Does this same object also support the «save to a file» interface?"
        → if yes, call those functions too.

4.  "I'm finished with it."
```

That is genuinely all of it. The whole course is the detail behind those four steps:

| Step | What COM calls it | Module |
|---|---|---|
| 1. Give me an object | **Activation** | 2 |
| 2. Call functions | **Calling through the vtable** | 1, 4 |
| 3. What else can you do? | **`QueryInterface`** | 1 |
| 4. I'm finished | **Reference counting** (`Release`) | 1 |

### See it working right now

Before any theory — open PowerShell and run these three lines:

```powershell
$fso = New-Object -ComObject Scripting.FileSystemObject
$fso.GetTempName()
$fso.GetSpecialFolder(2)
```

You just did all four steps. You asked the system for an object by name, called two functions on it, and PowerShell will release it for you. You didn't install anything, didn't reference a library, and have no idea what language `Scripting.FileSystemObject` was written in (C++, as it happens, in the 1990s).

That's COM. Everything else is detail.

### Two habits to unlearn

Newcomers import assumptions from ordinary object-oriented programming. Drop these two now:

- **As a client, you never create a COM object with `new`, and never `delete` it.** You *ask the system* for one, and you *tell it* when you're done. The object destroys itself when the last user lets go. (Someone does eventually call `new` — the component's own author, inside their own binary, where it's perfectly safe. You just never see it. Module 1 §1.4 and Module 2 show both sides.)
- **You never see the object's data, or even its class.** You only ever hold a pointer to a **list of functions**. What sits behind that list is deliberately hidden from you — and that hiding is the entire point, not an inconvenience.

### You already use COM every day

- Right-clicking a file in Explorer — every context-menu entry is a COM object
- Opening a File → Open dialog
- Anything playing video or rendering 3D (DirectX, Media Foundation)
- `Get-CimInstance` / `Get-WmiObject` in PowerShell
- Excel, Word and Outlook automation
- Windows Update, Task Scheduler, Volume Shadow Copy
- Every WinRT / Windows App SDK app — WinRT is built on COM

That last point is why this matters for support work: **a large share of "class not registered," "access denied," "the app hangs," and "works on my machine" tickets are COM tickets.**

### How much C++ do you need?

Less than you might fear.

| To do this | You need |
|---|---|
| Follow the concepts (§0.0, 0.3, 0.4, 0.6 and most of the course) | No C++ at all |
| Do the labs | Ordinary C++ — classes, inheritance, pointers. Not expert level. |
| Read §0.1 below in full detail | Some knowledge of how compilers lay out objects |

**§0.1 is the deep justification for why COM exists.** It uses advanced vocabulary in a few places. If it gets heavy, **read the bold conclusion of each problem and move on** — you'll find it easy after Module 1. There's a four-sentence summary at the end of §0.2 that captures everything §0.1 proves.

### The one idea to carry out of §0.1

> **Shipping a C++ class across a binary boundary is unsafe. COM is the fix.**

That sentence is doing a lot of work. Here it is a word at a time.

#### "A binary"

A compiled file — a `.exe` or a `.dll`. It holds machine code, not source. It was produced at one particular moment, by one particular compiler, with one particular set of settings.

#### "A binary boundary"

The line where execution leaves one binary and enters another.

**Inside** a single binary, everything agrees automatically. One compiler compiled all of it in one pass, so it made the same decision everywhere about how big each object is, where its fields sit, how arguments are passed, and which heap `new` uses — and it *verified* those decisions as it went.

**Across** binaries, every one of those guarantees evaporates. The other file may have been built last year, by a different compiler, by a different company, with different flags. Nothing checks that the two agree — and nothing *can*, because by then there is no source left to compare, only machine code.

```
   ┌───────────── YourApp.exe ──────────────┐      ┌──── MathLib.dll ────┐
   │                                        │      │                     │
   │   your code ──calls──► your other code │      │   vendor's code     │
   │            ▲                           │      │          ▲          │
   │            └─ one compiler, one pass,  │      │          │          │
   │               fully verified           │      │          │          │
   └────────────────────┬───────────────────┘      └──────────┼──────────┘
                        │                                     │
                        └───────────── calls ─────────────────┘
                                       ▲
                            THE BINARY BOUNDARY
                       different compiler, different day,
                        nothing verifies the two agree
```

The technical name for "the set of rules two binaries must agree on" is the **ABI** — Application Binary Interface. **C++ has no standard ABI**, so two compilers can disagree about all of it and both still be correct. COM's contribution is to define one tiny ABI that everybody can implement.

#### "Shipping a class"

Handing someone a **header file** describing your C++ class, plus a **DLL** containing it, and inviting them to compile against it.

The header tells *their* compiler how big your objects are, where each field sits, and which slot each method occupies. Their compiled code then bakes those numbers in **permanently**. If you later change any of them, their binary is still using last year's numbers — and neither of you finds out.

#### "Unsafe"

This does **not** mean "occasionally crashes." It means the mismatch **cannot be detected**:

- It compiles cleanly.
- It links cleanly.
- It usually appears to run correctly at first.
- When it does fail, it fails **at runtime, silently, far from the cause** — a wrong answer, or memory corruption that crashes something unrelated ten minutes later.

A loud, immediate error would be a minor inconvenience. Silent, delayed, undetectable failure is the worst possible outcome — which is why this problem was worth inventing an entire technology to solve.

#### "COM is the fix"

COM shrinks what crosses the boundary down to the one thing both sides can provably agree on: **an array of function pointers, called by position.**

No object sizes. No field layouts. No heaps. No name mangling. No exceptions. Nothing left that can quietly disagree.

Everything you'll read in §0.1 is the evidence for this claim; everything in §0.2 is the consequence of it.

---

## 0.1 The problem: C++ has no stable binary contract

### The scenario

You work at a company that sells `MathLib`, a C++ library. A few thousand customer applications are built on it.

```cpp
// MathLib.h — you ship this header to every customer
class MathLib
{
public:
    virtual double Add(double a, double b);
    virtual double Divide(double a, double b);
private:
    int m_precision;
};
```

Customers `#include` your header, compile against it, and ship their products alongside `MathLib.dll`.

**Everything works.** Hold on to that. Nothing warns anyone that this arrangement is fragile.

Eighteen months later you release v2. Customers do exactly what DLLs exist to allow — they drop in the new `MathLib.dll` **without recompiling** — and their applications start misbehaving.

To understand why, you first have to know what their compiler did with your header.

### What your customer's compiler permanently baked in

When their compiler read `class MathLib`, it hardcoded three facts into their binary. Permanently. Nothing can change them short of recompiling:

| Baked into the customer's binary | Derived from | Needed for |
|---|---|---|
| **The object's size** | the list of data members | `new MathLib()` |
| **The offset of each field** | the order of data members | reading `m_precision` |
| **The slot number of each method** | the order of `virtual` declarations | calling `Divide()` |

The first two are intuitive. The third is the one that surprises people, so:

#### How a virtual call actually works

When a C++ class has `virtual` functions, the compiler builds a **vtable** ("virtual function table") for that class: a plain array of function pointers, one per virtual method, in **declaration order**. There is exactly **one vtable per class** — built at compile time, stored in the binary's read-only data, shared by every instance.

Each *object* then carries a hidden pointer — the **vptr** — to its class's vtable. Under **MSVC** (Microsoft Visual C++, the compiler that ships with Visual Studio) it sits at offset 0, the first bytes of the object.

> Note the hedge in that sentence. *Where* the vptr sits is a **compiler's choice**, not a language rule — MSVC puts it first, other compilers may not. That's a small preview of Problem 3.

```
  a MathLib object                the MathLib vtable (one per class, shared)
  ┌──────────────┐               ┌─────────────────────────┐
  │ vptr         │──────────────►│ &MathLib::Add           │  slot 0
  ├──────────────┤               ├─────────────────────────┤
  │ m_precision  │               │ &MathLib::Divide        │  slot 1
  └──────────────┘               └─────────────────────────┘
```

So `obj->Divide(10, 2)` does **not** call a fixed address. It means:

> **"Read the vptr. Go to slot 1. Call whatever is there."**

The **slot number** is fixed in the caller at compile time. The **address** is looked up at run time.

That split is genuinely useful — it's what lets v2 rewrite `Divide`'s implementation freely, and the customer picks up the new code automatically. But it also means the caller and the callee must agree, forever, on *which slot is which*. Nothing enforces that agreement.

### The four ways this breaks

Each of the next four problems invalidates one of those baked-in facts. Read them as a progression — they get steadily further outside your control.

| # | What changes | Who controls it | What it invalidates |
|---|---|---|---|
| **1** | You add a data member | you | the object's size |
| **2** | You add a virtual method | you | the slot numbers |
| **3** | Your customer changes compiler | **your customer** | all of it at once |
| **4** | Your customer isn't using C++ | **nobody** | it never worked at all |

As you read, watch for the one thing all four have in common. That common factor is what COM exists to eliminate.

### Problem 1: You add a private data member

```cpp
private:
    int m_precision;
    int m_rounding;   // new in v2
```

Harmless-looking. It's `private` — no customer can even name it.

But the object's **size changed**. The customer's `new MathLib()` still allocates the v1 size, because that number was baked in eighteen months ago. Your v2 code then writes to `m_rounding`, which sits past the end of that allocation.

**Result:** heap corruption. Not at the moment of the bad write — later, somewhere else, in code that did nothing wrong. The customer must recompile to fix it, but they may no longer have the source, or may have already shipped.

### Problem 2: You add a virtual method in the middle

```cpp
    virtual double Add(double a, double b);
    virtual double Subtract(double a, double b);   // inserted in v2
    virtual double Divide(double a, double b);
```

You put `Subtract` next to `Add` because that reads better. Nothing was removed, nothing was renamed, and no existing signature changed.

But slots are assigned in declaration order, so in v2 the vtable is now:

```
  v1 (what the customer compiled against)   v2 (what they're actually running)
  slot 0  Add                               slot 0  Add
  slot 1  Divide          ◄── they call     slot 1  Subtract     ◄── they get
                                            slot 2  Divide
```

The customer's binary still says "go to slot 1." So `Divide(10, 2)` runs `Subtract` and returns **8**.

**Result:** no crash. No error. No warning. Just wrong numbers, in production, for as long as it takes someone to notice. This is the worst failure mode in the entire module.

### Problem 3: Your customer changes compiler

So far you broke things by changing your own code. Now you don't have to do anything at all — the customer simply rebuilds *their* product with a different compiler, a different *version* of the same compiler, or just different compiler flags.

Five independent things can now disagree:

| What differs | In plain terms | What breaks |
|---|---|---|
| **Name mangling** | Compilers encode a function's name, class and parameter types into one symbol, each in their own private format | The customer's linker literally cannot find `MathLib::Add` — it's looking for a different spelling |
| **Class layout** | Where the compiler puts the hidden pointer and the members, especially with multiple inheritance | Reads land on the wrong field |
| **Calling convention** | Who places arguments where, and who cleans up the stack afterwards | Stack corruption on every single call |
| **Exception handling** | How a `throw` unwinds the stack | An exception thrown inside your DLL cannot be caught in their EXE |
| **C runtime heap** | Each copy of the C runtime has its own private heap | Memory allocated by your `new` cannot be freed by their `delete` — two heaps, immediate corruption |

**None of these is a bug.** Every compiler is free to choose, and every choice is defensible. There is simply no rule requiring them to agree — and no way for either side to detect the disagreement at build time.

**Result:** anything from a link error (the merciful case — at least it's loud) to silent stack corruption on every call.

### Problem 4: Your customer isn't using C++ at all

Your customer writes Visual Basic, Python, or C#. There is no `#include`, no header, no way to express a C++ class in their language.

**Result:** the conversation ends before it starts. This isn't a version-compatibility problem — it's a demonstration that the *entire approach* only ever worked for one language.

### What all four have in common

Go back and look at the four results together:

| | The customer's binary assumed… | …but reality was |
|---|---|---|
| 1 | this object is 8 bytes | it's 12 |
| 2 | `Divide` is at slot 1 | slot 1 is `Subtract` |
| 3 | arguments are passed *this* way | they're passed *that* way |
| 4 | — | there's no shared vocabulary at all |

Every row is the same shape: **the customer's compiler was forced to make assumptions about your implementation, and nothing in the system keeps those assumptions true.**

That's the real problem. Not any single one of the four — the *category*. And notice what makes it so damaging: in cases 1, 2 and 3 everything compiles, links, and usually appears to run. The failure arrives later, silently, somewhere else.

**The conclusion:** a C++ class is an excellent way to share code *within* one binary compiled at one moment by one compiler, where the compiler verifies every one of those assumptions. It is a terrible way to share code *across* a binary boundary, because far too much invisible detail becomes part of the contract — and nobody is checking.

The question COM answers is therefore: **what is the smallest possible thing two binaries can agree on, that's still enough to call each other's code?**

---

## 0.2 The insight

Answering that question means looking for something that survives all four problems.

Go back to Problem 2. Notice what *didn't* fail there. The mechanism worked perfectly: the customer's code read a vptr, indexed into an array, and called what it found. No corruption, no compiler disagreement, no ambiguity. The **machinery** was flawless.

The only thing that went wrong was that the *contents* of slot 1 changed.

And that's the whole insight. "Read a pointer, index an array, call the entry, pass `this` as the first argument" is a pure machine-level operation. There is nothing compiler-specific about it, nothing language-specific, nothing that two vendors could reasonably implement differently. Any compiler can do it. Any language that can make an indirect call can do it. It survives Problem 3 untouched, and it's simple enough that Problem 4 dissolves too.

It has exactly one weakness: someone has to promise not to renumber the slots.

So COM's founding move is to extract that promise:

> **Freeze the vtable. Make it the contract.**

A COM **interface** is a vtable layout that is declared immutable, forever, once published. From that single decision, everything else follows:

| Rule | Because |
|---|---|
| Interfaces have no data members | Data members imply size, and size implies fragility |
| Interfaces are pure virtual | Any implementation would be compiler-specific |
| You never modify a published interface | The vtable is frozen |
| Methods return `HRESULT`, not exceptions | Exception ABIs are compiler-specific |
| Methods use `__stdcall` | An explicitly chosen, stable calling convention |
| The object frees its own memory | Because the caller's heap is a different heap |
| Clients never call `new`/`delete` on an object | The object's size is the object's business |
| Objects are reference counted | The client can't know when others are done |

Every "weird" thing about COM is a direct consequence of "the vtable is the contract."

### The whole argument in four sentences

If §0.1 was heavy, this is all you actually need to carry forward:

1. Sharing a C++ class across a binary boundary is unsafe, because too much invisible detail — sizes, layouts, conventions, heaps — silently becomes part of the contract.
2. But **one** thing *is* safe to share: an array of function pointers, called by position.
3. So COM declares that array — the vtable — **to be** the contract, and forbids changing it once published.
4. Every rule in the table above exists to protect that one promise.

When something in COM later looks arbitrary or over-complicated, come back to point 3. It is almost always the answer.

---

## 0.3 The three pillars

### Pillar 1: Interfaces are immutable contracts

An interface is identified by a **GUID** (a 128-bit number), called an **IID** (Interface ID). Once you ship `IFoo` with IID `{...}`, that IID means *exactly that vtable layout* forever, on every machine on Earth.

Need to add a method? You publish a **new** interface, `IFoo2`, with a **new** IID. Old clients keep asking for `IFoo` and keep working. New clients ask for `IFoo2`. Both can be served by the same object.

This is why Windows has `IShellFolder` and `IShellFolder2`, `IPropertyStore` and `IPropertyStoreCapabilities`, `IClassFactory` and `IClassFactory2`. It looks ugly. It is the reason 30-year-old shell extensions still load.

### Pillar 2: Objects are reference counted

A COM object may be used by code you've never heard of. Nobody has global knowledge of who holds a pointer. So the object counts its own references and destroys itself at zero. Every interface carries `AddRef` and `Release` for this purpose — they're part of `IUnknown`, which every interface inherits.

### Pillar 3: Location transparency

This is the payoff for the first two pillars. The name is unhelpful, so here is what it actually means:

> **You write exactly the same code whether the object is inside your own process, in a different process on the same machine, or on a computer in another country.**

One set of source lines covers all three cases. No `#ifdef`, no separate "remote" version of your function, no recompilation, nothing in your calling code that mentions location at all.

Being precise about the claim, since it's easy to over- or under-read:

- It does **not** mean an object relocates while it's running. Objects don't move around.
- It means the object's location is decided **once, at creation time** (Module 2) — by a registry setting or a single flag in one line — and **every line after that is unaffected**.

So the difference between "local" and "on a server" is one argument in one call, in one place. The other ten thousand lines of your program are identical either way.

#### Why that's even possible

Look again at what a client actually does with an object: it holds a pointer, reads the vptr, goes to a slot, and calls the address it finds there.

Now notice everything the client **never** does:

- It never reads the object's fields.
- It never knows the object's size.
- It never takes the address of anything inside the object.

So the client has **no way to check** that the thing it's calling is the real object. Every observation available to it would look identical if it were talking to a convincing stand-in.

COM exploits exactly that.

#### The proxy

When the real object lives somewhere the client can't reach directly, COM quietly hands the client a **proxy** instead — a small local object that:

- exposes **the identical vtable**, so the client's already-compiled code works unchanged, and
- implements every method by packing up the arguments, shipping them to wherever the real object lives, waiting, and returning the result.

On the far side a **stub** does the mirror image: unpack the arguments, call the real object, pack the result, send it back.

Think of an interpreter on a phone call. You speak normally. The interpreter deals with the fact that the other party is elsewhere and doesn't share your language. **You don't change how you talk** — that's the transparency.

```
     THE SAME CLIENT LINE IN BOTH CASES:   pCalc->Add(2, 3, &r);

  In-process                          Cross-process or cross-machine
 ┌─────────────────────┐          ┌────────────────────┐   ┌─────────────────────┐
 │ client              │          │ client             │   │ server process      │
 │   │                 │          │   │                │   │                     │
 │   ▼                 │          │   ▼                │   │                     │
 │ ┌─────────────┐     │          │ ┌────────┐         │   │  ┌────────┐         │
 │ │ real object │     │          │ │ PROXY  │──net/IPC┼───┼─►│  STUB  │         │
 │ └─────────────┘     │          │ └────────┘         │   │  └───┬────┘         │
 │                     │          │  identical vtable  │   │      ▼              │
 │  direct call        │          │                    │   │  ┌─────────────┐    │
 │  ~nanoseconds       │          │                    │   │  │ real object │    │
 └─────────────────────┘          └────────────────────┘   │  └─────────────┘    │
                                                            └─────────────────────┘
```

The client can't tell these apart, because no observation it is capable of making would differ.

#### This works *only* because of Pillar 1

Substituting a stand-in is possible only because the contract is **exactly a vtable layout and nothing else**.

If interfaces had data members, or the client could take the address of a field, or objects were created with `new`, then a proxy would have to reproduce the object's *memory* as well as its functions — and the illusion would collapse immediately.

**This is why COM interfaces look so restrictive.** Every restriction in §0.2's table is what buys you this pillar.

#### What it costs

Transparency is about **correctness, not performance**. Your source doesn't change — but the call absolutely does:

| Where the object is | Cost of one call |
|---|---|
| Same process, no boundary crossed | a few **nanoseconds** |
| Another process, same machine | tens of **microseconds** |
| Another machine | **milliseconds** |

That's up to a millionfold difference for an identical line of source. Which is exactly why *"why did this component suddenly get slow?"* is such a common support ticket — usually, somebody's object quietly moved.

Module 3 has you measure this yourself; Module 4 covers the packing and unpacking (called **marshaling**); Module 7 covers the cross-machine case.

#### Seeing it in one line

PowerShell's CIM/WMI cmdlets are COM clients. The only difference between these two commands is *where the object gets created* — everything after that is the same machinery:

```powershell
Get-CimInstance Win32_OperatingSystem                                    # object is local
Get-CimInstance Win32_OperatingSystem -ComputerName FS01 -Protocol DCOM  # object is on FS01
```

Same cmdlet, same properties, same code. Different continent, if you like. That's location transparency.

---

## 0.4 Vocabulary — learn these now

| Term | Meaning |
|---|---|
| **GUID** | 128-bit unique ID. Written `{6B29FC40-CA47-1067-B31D-00DD010662DA}`. Generated by `guidgen.exe` or `New-Guid`. |
| **Interface / IID** | A frozen vtable layout, identified by a GUID called an IID. Named `IFoo` by convention. |
| **Coclass / CLSID** | A concrete *implementation* — a class that implements one or more interfaces. Identified by a GUID called a CLSID. |
| **ProgID** | A human-readable alias for a CLSID, e.g. `Excel.Application`, `MSXML2.DOMDocument.6.0`. Not guaranteed unique; a convenience for scripting. |
| **Component / Server** | The DLL or EXE containing the implementation. |
| **Client** | Anything that calls a COM object. |
| **Type library / LIBID** | A compiled description of interfaces and coclasses, used for late binding and interop. |
| **Apartment** | A thread-safety boundary. Module 3. |
| **Marshaling** | Packaging a call so it can cross an apartment/process/machine boundary. |
| **Proxy / Stub** | Client-side and server-side halves of the marshaling machinery. |
| **SCM** | COM's **Service Control Manager** — the broker that turns a CLSID into a running server. Confusingly, this is **not** the Windows Service Control Manager that starts and stops services; they merely share a name. COM's SCM is implemented by the `RpcSs` / `DcomLaunch` services. |
| **DCOM** | **Distributed COM** — the same COM, carried over a network transport so a client can use an object in another process or on another machine. Module 7. |
| **AppID** | Registry entry holding *process-wide* settings for out-of-proc servers: identity, permissions, surrogate. |

**Crucial distinction that beginners always get wrong:**

- **CLSID** = "which implementation do I want?" — passed to `CoCreateInstance`.
- **IID** = "which contract do I want to talk through?" — passed to `CoCreateInstance` and `QueryInterface`.

One CLSID implements many IIDs. One IID is implemented by many CLSIDs.

```cpp
CoCreateInstance(
    CLSID_ShellLink,          // WHICH object  (the implementation)
    nullptr,
    CLSCTX_INPROC_SERVER,
    IID_IShellLinkW,          // WHICH contract (the interface)
    (void**)&pLink);
```

---

## 0.5 The canonical "hello world"

### What this program does, in plain English

It asks Windows for a **shell link** object — the thing behind a `.lnk` shortcut file. It points that object at Notepad and gives it a description. Then it asks the *same object* whether it also knows how to **save itself to a file**; it does, so we save it. Finally we let go of everything we were holding.

Before you read the code, notice what's missing from that description: we never name a C++ class, never include a header describing the object's contents, and never link against a shell library. We know only **two GUIDs** and a list of functions.

Read it for shape, not detail. You'll write something like it in Module 1.

```cpp
#include <windows.h>
#include <shobjidl.h>   // IShellLink
#include <shlguid.h>
#include <objbase.h>
#include <stdio.h>

int wmain()
{
    // 1. Initialize COM on THIS THREAD. Every thread that uses COM must do this.
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return 1;

    IShellLinkW* pLink = nullptr;

    // 2. Ask the SCM to create an object of class CLSID_ShellLink,
    //    and hand me back its IShellLinkW interface.
    hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
                          IID_IShellLinkW, reinterpret_cast<void**>(&pLink));
    if (SUCCEEDED(hr))
    {
        // 3. Use it through the vtable.
        pLink->SetPath(L"C:\\Windows\\notepad.exe");
        pLink->SetDescription(L"My first COM object");

        // 4. Ask the SAME OBJECT for a DIFFERENT interface.
        IPersistFile* pFile = nullptr;
        hr = pLink->QueryInterface(IID_IPersistFile, reinterpret_cast<void**>(&pFile));
        if (SUCCEEDED(hr))
        {
            pFile->Save(L"C:\\Temp\\test.lnk", TRUE);
            pFile->Release();          // 5. Release EVERY interface you obtain.
        }

        pLink->Release();
    }
    else
    {
        wprintf(L"CoCreateInstance failed: 0x%08X\n", hr);
    }

    // 6. Tear down COM on this thread.
    CoUninitialize();
    return 0;
}
```

Six things happened, and each maps to a module in this course:

1. `CoInitializeEx` → **apartments** (Module 3)
2. `CoCreateInstance` → **activation and the registry** (Module 2)
3. Calling through the vtable → **interfaces and IDL** (Module 4)
4. `QueryInterface` → **`IUnknown`** (Module 1)
5. `Release` → **lifetime** (Module 1)
6. Failure handling → **HRESULTs and diagnostics** (Module 8)

> **What is `IPersistFile`?** One of COM's standard *persistence* interfaces — "save yourself to this path" / "load yourself from it." Note that `CShellLink` implements `IShellLink` for its **behaviour** and `IPersistFile` for its **storage**: two orthogonal contracts on one object, which is precisely what `QueryInterface` is for. The family is covered in [Appendix A §A.2](appendix-a-monikers-and-persistence.md#a2-persistence--the-ipersist-family).

Notice what is *absent*: no `new`, no `delete`, no header describing `CShellLink`'s members, no link against a shell library. The client knows only two GUIDs and one vtable layout.

---

## 0.6 Where COM lives today (why this matters for your job)

COM is not legacy. It's load-bearing infrastructure:

- **Windows Shell** — every context menu handler, property sheet, thumbnail provider, preview handler, `IFileDialog`, jump lists.
- **WinRT / Windows App SDK** — WinRT is *built on* COM. `IInspectable` derives from `IUnknown`. Every WinRT object is a COM object with extra rules.
- **DirectX / Direct2D / Direct3D / DXGI / Media Foundation / WIC** — all COM.
- **WMI** — `IWbemServices`, `IWbemClassObject`. Every WMI support case is a COM case.
- **Office automation** — `Excel.Application` etc.
- **MSXML, ADO, OLE DB, ADSI**.
- **Task Scheduler, BITS, Windows Update Agent, Volume Shadow Copy (VSS), Windows Search** — all exposed as COM.
- **PowerShell's `New-Object -ComObject`**, and .NET's entire interop layer.
- **Internet Explorer / WebView2 hosting**, ActiveX (still present in embedded/industrial software).

For a support engineer, this means: **a large fraction of "the app hangs," "access denied," "class not registered," and "works on my machine" cases are COM cases.** Recognizing the signature is most of the job.

---

## 0.7 Checkpoint

Answer in writing before moving on. Don't look them up — reason from first principles.

### Part A — concepts (no C++ needed)

If you only do one part, do this one. These are the ideas the rest of the course builds on.

1. In your own words, what is a COM **interface**, and how is it different from a C++ class?
2. Why do you never call `new` or `delete` on a COM object? Who decides when it's destroyed?
3. An object supports interfaces `IFoo` and `IBar`. You hold an `IFoo` pointer. How do you find out whether it also does `IBar` — and what have you got if it does?
4. Explain "location transparency" to a colleague in two sentences, without using the word "marshaling."
5. Name three Windows features you personally used this week that are COM underneath.

### Part B — the deeper argument

These test §0.1 and §0.2. If §0.1 was heavy, come back to these after Module 1.

6. A vendor ships `Widget.dll` with a C++ class. Their v2 adds one private `int`. The customer doesn't recompile. Describe precisely what goes wrong, and when.
7. Why does COM forbid adding a method to a published interface, but permit shipping a *new* interface with an extra method — even though both change what the object can do?
8. Why do COM methods return `HRESULT` rather than throwing? Give two independent reasons.
9. Why must a COM object hand back memory through a shared allocator, rather than letting the caller use `free()`?

<details>
<summary>Answers (open only after writing yours)</summary>

**Part A**

1. An interface is a frozen, published **list of functions** and nothing else — no data, no implementation, and it may never change once released. A C++ class bundles functions *with* data and implementation, so its size and layout leak into the contract; that's exactly what makes it unsafe to share across a binary boundary. One more difference worth noting: a single object can offer many interfaces, and a single interface can be offered by many unrelated objects.

2. Because `new` requires knowing the object's **size**, and `delete` requires knowing its **destructor and heap** — all three are private implementation details you're deliberately not shown. Instead you ask the system for the object, and the *object itself* destroys itself once the last user has let go. It knows that because it counts its users (Module 1).

3. You call `QueryInterface`, asking for `IBar`. If it succeeds you get a **second, independent pointer** to the same object — and you now owe a separate "I'm finished" on it as well as on your original `IFoo` pointer. If it fails, the object genuinely doesn't offer `IBar`, and it never will (the answer can't change over the object's lifetime).

4. The client only ever calls functions through a table of pointers it was handed; it never touches the object's memory directly. So Windows can substitute a stand-in that looks identical but forwards each call somewhere else — another process, another machine — and the client's code cannot tell the difference.

5. Examples: the Explorer right-click menu, any File→Open dialog, PowerShell `Get-CimInstance`, any game or video playback (Direct3D / Media Foundation), Excel/Word, Windows Update, Task Scheduler.

**Part B**

6. The customer's code allocates `sizeof(Widget)` **as it knew it at their compile time** — the v1 size. The v2 DLL's constructor and methods then write to the new member, which lies past the end of that allocation. Result: heap corruption, typically detected far from the cause — a crash in an unrelated `free`, or silently corrupted neighbouring data.

7. Because the *contract* is the vtable, not the capability. Adding a method lengthens the vtable and shifts slot indices for anything derived from it, and already-compiled clients call by slot number. A new interface gets a new IID and a separate vtable — old clients never ask for it and are completely unaffected.

8. (a) The C++ exception ABI is compiler-specific and cannot cross a binary boundary. (b) An exception cannot propagate across a process or machine boundary over RPC, so location transparency would break. (Bonus: (c) not every client language has exceptions at all — VBScript doesn't.)

9. The client and the object may be linked against different C runtimes, each with its own private heap. Freeing a block on the wrong heap corrupts it. A single shared allocator (`CoTaskMemAlloc`/`CoTaskMemFree`) means both sides provably agree — and across a process boundary the caller's heap doesn't even exist in the callee's address space.

</details>

---

## 0.8 Tooling setup (do this now)

```powershell
# Verify you have the SDK tools on PATH from a "Developer PowerShell for VS"
midl.exe /?          # IDL compiler
guidgen.exe          # or: New-Guid
```

Install:

1. **Visual Studio** with *Desktop development with C++* (include ATL) and *.NET desktop development*.
2. **WinDbg** — from the Microsoft Store, or the Windows SDK's Debugging Tools.
3. **Sysinternals Suite** — Process Monitor, Process Explorer.
4. **OleView.NET** — https://github.com/tyranid/oleviewdotnet (releases page). The old `oleview.exe` still ships with the SDK but OleView.NET is far better.

Quick smoke test that COM works on your box:

```powershell
$fso = New-Object -ComObject Scripting.FileSystemObject
$fso.GetTempName()
[Runtime.InteropServices.Marshal]::ReleaseComObject($fso) | Out-Null
```

If that prints a random `radXXXXX.tmp` filename, your COM stack is healthy and you've just done late-bound `IDispatch` activation via ProgID.

---

**Next: [Module 1 — IUnknown, vtables, and lifetime](01-iunknown-and-lifetime.md)**
