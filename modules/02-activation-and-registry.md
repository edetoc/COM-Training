# Module 2 — Activation, registration, and the registry

In Module 1 you called `CreateCalculator` — a plain C++ function. Real clients don't do that; they know only a GUID. This module explains how a GUID becomes a running object, and every way that can fail. **It generates the single largest category of COM support tickets.**

**What this module covers**

The full path from `CoCreateInstance` to a live object: which registry keys COM reads, in what order, and how WOW64 redirection quietly turns a working component into "class not registered". You build a real in-proc server — class factory, DLL exports, self-registration — then do it again with no registry at all, using manifests. It ends with server lifetime rules and a triage table for the activation `HRESULT`s you will actually be handed.

**Contents**

- [2.1 The activation question](#21-the-activation-question)
- [2.2 The registry map](#22-the-registry-map)
- [2.3 WOW64 registry redirection — the #1 cause of "class not registered"](#23-wow64-registry-redirection--the-1-cause-of-class-not-registered)
- [2.4 The activation call chain](#24-the-activation-call-chain)
- [2.5 Building a real in-proc server](#25-building-a-real-in-proc-server)
- [2.6 LAB 2.1 — Build, register, activate, and watch it happen](#26-lab-21--build-register-activate-and-watch-it-happen)
- [2.7 LAB 2.2 — Bitness](#27-lab-22--bitness)
- [2.8 LAB 2.3 — Registration-free (Reg-Free) COM](#28-lab-23--registration-free-reg-free-com)
- [2.9 Server lifetime and unloading](#29-server-lifetime-and-unloading)
- [2.10 The HRESULT triage table](#210-the-hresult-triage-table)
- [2.11 LAB 2.4 — Reproduce every failure deliberately (support drill)](#211-lab-24--reproduce-every-failure-deliberately-support-drill)
- [2.12 Tooling: OleView.NET](#212-tooling-oleviewnet)
- [2.13 Checkpoint](#213-checkpoint)

---

## 2.1 The activation question

A client says:

```cpp
CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_INPROC_SERVER,
                 IID_ICalculator, (void**)&pCalc);
```

Two GUIDs appear in that single call, and they answer completely different questions:

| Argument | Name | Question it answers |
|---|---|---|
| `CLSID_Calculator` | **CLSID** — class ID | *Which component do I want created?* |
| `IID_ICalculator` | **IID** — interface ID (Module 1) | *Which of its interfaces do I want back?* |

Only the first one drives activation. The IID matters afterwards, once an object exists.

So Windows must answer: *given the CLSID and nothing else — 128 bits, no path, no filename, no hint — which file on this machine holds that class's code, and how do I get it running?*

The answer lives in the **registry** (or, for registration-free COM, in a **manifest**). That is the entire mechanism. Understanding the exact lookup path is what lets you diagnose activation failures in seconds instead of hours.

---

## 2.2 The registry map

Everything lives under `HKEY_CLASSES_ROOT` (`HKCR`), which is a **merged view**:

```
HKCR = HKEY_CURRENT_USER\Software\Classes      (per-user, wins on conflict)
     + HKEY_LOCAL_MACHINE\Software\Classes     (per-machine)
```

> **Support fact #1:** `HKCU` wins. A per-user registration silently shadows a per-machine one. "It works for me but not for the customer" and "it works when I RDP but not as a service" are very often this.

### The CLSID key — where an object's code lives

```
HKCR\CLSID\{A1B2C3D4-1111-4000-9000-000000000001}
    (Default)                = "Calculator Component"
    │
    ├── InprocServer32
    │       (Default)        = "C:\Components\Calc.dll"
    │       ThreadingModel   = "Apartment" | "Free" | "Both" | "Neutral"
    │
    ├── LocalServer32
    │       (Default)        = "C:\Components\CalcSrv.exe"
    │
    ├── ProgID
    │       (Default)        = "MyCompany.Calculator.1"
    │
    ├── VersionIndependentProgID
    │       (Default)        = "MyCompany.Calculator"
    │
    ├── TypeLib
    │       (Default)        = "{LIBID GUID}"
    │
    └── (Default value of an "AppID" named value) = "{APPID GUID}"
```

- **`InprocServer32`** → a DLL loaded into the *client's* process.
- **`LocalServer32`** → an EXE launched as a *separate* process.
- Both present? The **third parameter of `CoCreateInstance`** decides — the `CLSCTX` flags (§2.4). They say which kinds of server the client is willing to accept, and COM tries them in its own fixed preference order: in-proc before local, local before remote.
- **`ThreadingModel`** → which **apartment** the object is allowed to live in.

> **An apartment** is COM's thread-safety boundary: a group of threads that are permitted to call an object **directly**. A thread outside the object's apartment may not use a raw pointer to it at all — its calls have to be packaged up and handed across, which is **marshaling**. Two facts carry you through this module: **every object belongs to exactly one apartment**, and **an apartment boundary is as real as a process boundary** — COM makes a call cross both the same way. Module 3 is devoted to this, and `ThreadingModel` is its central knob; set it to `"Both"` for now.

### The ProgID keys — the human-readable alias

```
HKCR\MyCompany.Calculator.1
    (Default)    = "Calculator Component"
    └── CLSID
            (Default) = "{A1B2C3D4-1111-4000-9000-000000000001}"

HKCR\MyCompany.Calculator          <- version independent
    └── CurVer
            (Default) = "MyCompany.Calculator.1"
```

`CLSIDFromProgID(L"MyCompany.Calculator.1", &clsid)` walks exactly these keys. That's all `New-Object -ComObject MyCompany.Calculator` does before calling `CoCreateInstance`.

#### Which should a client use?

Both routes end at the same `CoCreateInstance`. The difference is who performs the lookup, and what can go wrong on the way.

| Reach for the **CLSID** when | Reach for the **ProgID** when |
|---|---|
| You are compiled code and can embed a GUID | The caller cannot hold a GUID comfortably — scripts, config files, anything a human types |
| You need certainty about *which* class you get | You want whichever version is currently installed |
| You would rather have one less lookup and one less failure mode | Readability matters more than precision |

Two consequences worth carrying forward:

- **A ProgID is an alias, not an identity.** It is a string somebody chose, and nothing enforces uniqueness. The CLSID is the identity. And since `HKCU` wins over `HKLM` (Support fact #1), a per-user registration can quietly point a well-known ProgID at a different class altogether.
- **A ProgID adds a failure mode that CLSID activation cannot have:** `CO_E_CLASSSTRING` (`0x800401F3`), when the name is not registered or its `CLSID` subkey is missing. So if a script fails and a C++ client using the raw CLSID succeeds, the ProgID keys are the problem — not the component.

The two ProgID forms differ the same way. `MyCompany.Calculator.1` pins a version; `MyCompany.Calculator` follows `CurVer` to whatever is current. The version-independent form is what scripts normally use, and it is also why installing an upgrade can silently change which class a script gets.

### The Interface key — how a call crosses a boundary

```
HKCR\Interface\{IID}
    (Default)          = "ICalculator"
    ├── ProxyStubClsid32
    │       (Default)  = "{CLSID of the proxy/stub, or {00020424-...} for typelib marshaling}"
    ├── TypeLib
    │       (Default)  = "{LIBID}"
    │       Version    = "1.0"
    └── NumMethods
            (Default)  = "5"
```

> **Support fact #2:** this key is only consulted when a call must **cross a boundary** — between apartments, between processes, or between machines. A plain in-proc call inside one apartment never touches it, which is why the classic bug is "works in-process, `E_NOINTERFACE` out-of-process": the registration was always missing, but nothing needed it until something crossed a boundary. Module 4 covers it.

### The AppID key — process-wide settings for out-of-proc servers

```
HKCR\AppID\{APPID}
    (Default)          = "Calculator Server"
    RunAs              = "Interactive User" | "NT AUTHORITY\LocalService" | "DOMAIN\user"
    DllSurrogate       = ""            <- empty string means "use dllhost.exe"
    LaunchPermission   = <binary SD>
    AccessPermission   = <binary SD>
    AuthenticationLevel = dword
```

Module 7 covers these settings properly. What matters now is **why they live on a separate key at all**.

A CLSID identifies **one class**: which DLL or EXE implements it, and which threading model it wants. But the values above are not properties of a class — they are properties of a **process**. "Run as `LocalService`" or "require this authentication level" cannot sensibly be answered per class, because one EXE can implement a dozen classes and they all share its process, its identity, and its security.

So COM splits the two:

```
HKCR\CLSID\{CLSID-A}   AppID = {APPID}      \
HKCR\CLSID\{CLSID-B}   AppID = {APPID}       >  three classes, one process
HKCR\CLSID\{CLSID-C}   AppID = {APPID}      /

HKCR\AppID\{APPID}     RunAs, permissions, authentication level
```

Each CLSID names its AppID with a value on its own key, and every class pointing at that AppID shares one set of process-wide settings.

Two practical consequences: an in-proc DLL usually has **no** AppID, because it has no process of its own — it runs in the client's. And when you change a permission in `dcomcnfg`, you are editing the **AppID**, so the change applies to every class that server implements, not just the one you were debugging.

### The TypeLib key — the machine-readable description of the interfaces

A **type library** ("typelib") is a *compiled, binary description* of a component: its coclasses, interfaces, methods, parameter types, and enums. It is the same information a C++ header or an IDL file carries, but in a form **any language can read at runtime** rather than only a C++ compiler at build time.

It is *generated from* the IDL, not written by hand: MIDL compiles one `.idl` into headers, proxy/stub code, and the `.tlb` together (Module 4). Only the part of the IDL inside its `library { }` block ends up in the typelib.

It is usually **embedded as a resource inside the DLL or EXE itself** — which is why the path below points at `Calc.dll` rather than a separate file — though it can also ship as a standalone `.tlb`.

Who reads it:

- **Scripting and late binding** — PowerShell, VBScript, VBA, C# `dynamic`. This is how `$calc.Add(2,3)` resolves a method name to something callable (Module 5).
- **`#import` in C++** and `tlbimp`/COM references in .NET, to generate wrappers at build time (Module 6).
- **COM itself**, for *typelib marshaling* — the `{00020424-…}` value in the `Interface` key above means "work out how to marshal this from the type library" instead of using a hand-built proxy/stub (Module 4).

```
HKCR\TypeLib\{LIBID}\1.0
    (Default)      = "MyCompany Calculator 1.0 Type Library"
    ├── 0\win64
    │       (Default) = "C:\Components\Calc.dll"    <- or a .tlb path
    └── FLAGS, HELPDIR
```

Two details that generate tickets:

- **Type libraries are versioned in the registry** (`\1.0`), unlike CLSIDs. A client built against 1.0 will not find 2.0.
- **`win32` vs `win64` subkeys** — the same bitness split as everything else in this module. A missing, unregistered, or wrong-bitness typelib gives you `0x80029C4A TYPE_E_CANTLOADLIBRARY`, and the symptom is usually "the script can't find the method" rather than an activation failure, because the object itself creates fine.

---

## 2.3 WOW64 registry redirection — the #1 cause of "class not registered"

On 64-bit Windows there are **two** COM registries:

| Process bitness | Reads `HKLM\Software\Classes` from |
|---|---|
| 64-bit | `HKLM\Software\Classes` |
| 32-bit | `HKLM\Software\Classes\Wow6432Node` (transparently redirected) |

Consequences you must internalize:

- A 32-bit `regsvr32` registers into `Wow6432Node`. A 64-bit client cannot see it.
- A 64-bit client **cannot** load a 32-bit DLL into its process. Ever. Not with any flag.
- `%SystemRoot%\System32\regsvr32.exe` is the **64-bit** one. `%SystemRoot%\SysWOW64\regsvr32.exe` is the **32-bit** one. (Yes, that naming is backwards. System32 = native; SysWOW64 = 32-bit-on-64.)

**Diagnostic reflex:** when you see `0x80040154 REGDB_E_CLASSNOTREG`, question #1 is *"what bitness is the client, and what bitness is the server?"* — before you look at anything else.

Check bitness of a DLL — from a **Developer PowerShell for VS** (`dumpbin` is not on the default `PATH`):

```
dumpbin /headers Calc.dll | findstr machine
```

For a 64-bit build:

```
            8664 machine (x64)
```

For a 32-bit build:

```
             14C machine (x86)
                   32 bit word machine
```

`8664` = x64, `14C` = x86, `AA64` = ARM64. The second line on the 32-bit output comes from the file characteristics, not the machine type — ignore it and read the `machine (...)` value.

---

## 2.4 The activation call chain

```
Client: CoCreateInstance(CLSID, pUnkOuter, CLSCTX, IID, &pv)
   │
   └─► CoCreateInstanceEx(...)                     ; the real entry point
          │
          └─► CoGetClassObject(CLSID, CLSCTX, ..., IID_IClassFactory, &pCF)
                 │
                 ├─ 1. Check the activation context (reg-free COM manifests)
                 ├─ 2. Check the per-process class table (CoRegisterClassObject)
                 ├─ 3. Ask the SCM (RpcSs) -> read HKCR\CLSID\{...}
                 │        InprocServer32? -> LoadLibrary + GetProcAddress("DllGetClassObject")
                 │        LocalServer32?  -> CreateProcess, wait for CoRegisterClassObject
                 │
                 └─► returns IClassFactory*
          │
          ├─► pCF->CreateInstance(pUnkOuter, IID, &pv)   ; the object is born here
          └─► pCF->Release()
```

Two calls are happening. This matters: if you create many objects of the same class, calling `CoGetClassObject` once and reusing the `IClassFactory` is significantly faster than N `CoCreateInstance` calls.

### `CLSCTX` flags

This is the **third parameter** of `CoCreateInstance` (and of `CoGetClassObject`) — the `CLSCTX_INPROC_SERVER` in the very first example of this module. It is a **bitmask of the server kinds you will accept**, not a single choice:

| Flag | Meaning |
|---|---|
| `CLSCTX_INPROC_SERVER` | DLL in my process |
| `CLSCTX_INPROC_HANDLER` | In-proc handler (rare; `InprocHandler32`) |
| `CLSCTX_LOCAL_SERVER` | EXE on this machine |
| `CLSCTX_REMOTE_SERVER` | On another machine |
| `CLSCTX_ALL` | All of the above |
| `CLSCTX_ACTIVATE_32_BIT_SERVER` / `_64_BIT_SERVER` | Force bitness (only meaningful for LocalServer) |
| `CLSCTX_ENABLE_CLOAKING` | Use the thread token for activation |
| `CLSCTX_NO_CUSTOM_MARSHAL` | Security hardening: refuse `IMarshal` |

When you pass more than one, **COM picks the order, not you.** It prefers in-proc, then in-proc handler, then local, then remote — regardless of how you wrote the flags. So `CLSCTX_ALL` does not mean "try what I listed first"; it means "any of these, cheapest first."

#### What a "handler" actually is

A **handler** is a *half object*: an in-process DLL that implements **part** of a class locally and forwards the rest to the real out-of-process server. Registered under `HKCR\CLSID\{...}\InprocHandler32`.

Do not confuse it with a proxy. A **proxy** contains no logic — it packages every call and ships it to the server. A **handler** contains real code and cached data, so it can answer some calls *without the server running at all*.

The canonical case is OLE embedding: an Excel chart inside a Word document. Word must **draw** that chart every time the document opens, and launching Excel to do it would be absurd. So the handler loads into Word and serves `IViewObject2` / `IDataObject` / `IPersistStorage` from the **presentation cache** saved inside the document. Only when you double-click to edit does it launch the real server and delegate.

In practice `InprocHandler32` is almost always `ole32.dll` — the **default handler**, which provides that cache-and-draw behaviour generically. A *custom* handler is a DLL that aggregates the default handler via `OleCreateDefaultHandler` and overrides only the interfaces it cares about.

> **Support angle:** this is legacy OLE compound-document territory, plus some shell extensions — you will rarely write one. But it explains a confusing signature: an object that renders and answers queries while its server process is nowhere in Task Manager. That's the handler serving cached state, not a ghost.

> **Support fact #3:** `CLSCTX_ALL` is convenient and dangerous. If the in-proc registration is broken but a LocalServer32 exists, you'll silently get an out-of-proc object with different threading, different security, and different performance. When diagnosing, **always** re-test with the specific flag you expect.

### The other activation route: monikers

`CoCreateInstance` says *"make me a new, empty object of class X."* Sometimes you instead want *"the object named by this string"* — a specific file, a WMI namespace, an **already-running** Excel, or an elevated instance. That's what **monikers** do, via `CoGetObject` / `MkParseDisplayName` and the Running Object Table.

You'll need them for Module 7's elevation moniker, and every WMI connection string (`winmgmts:\\.\root\cimv2`) is one. See **[Appendix A §A.1](appendix-a-monikers-and-persistence.md#a1-monikers)**.

Worth knowing now: moniker failures frequently bottom out as ordinary activation failures, so a `0x80040154` from `GetObject` sends you straight back to this module's flow.

---

## 2.5 Building a real in-proc server

A COM DLL must export exactly four functions:

```cpp
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv);
STDAPI DllCanUnloadNow(void);
STDAPI DllRegisterServer(void);
STDAPI DllUnregisterServer(void);
```

### 2.5.1 `IClassFactory`

**Why a factory at all?** Because your DLL cannot simply export `CreateCalculator()`. The client has never heard of your DLL: it has no header for it, no import library, it never links against it, and it may not even be written in C++. All it holds is a CLSID and the `CoCreateInstance` API.

So COM needs **one entry point that every COM DLL has, with a fixed name and a fixed signature** — that is `DllGetClassObject`. Given a CLSID, it returns a **class object**: a small, long-lived COM object that represents *the class itself* and knows how to stamp out instances of it. `IClassFactory` is the interface that class object almost always implements — hence the everyday name **class factory**.

```
   one CLSID  ──►  one class object (the factory)  ──►  many instances
```

**What lives where.** Note that for an *in-proc* server the DLL is loaded into the client's own process — so "sides" here means **who owns which code**, not two separate processes:

| | **Client code** | **`Calc.dll` — the code you write** |
|---|---|---|
| Starts with | two GUIDs: `CLSID_Calculator` and `IID_ICalculator` | nothing; it sits on disk until COM loads it |
| Must provide | nothing but calls to COM APIs | `DllGetClassObject` — the one fixed export |
| Ends up holding | an `ICalculator*` (and an `IClassFactory*` briefly) | the `CalculatorFactory`, and every `Calculator` instance |
| Never has | your headers, your `.lib`, your class names | any idea who called it, or in what language |

**The exchange**, with `CoCreateInstance` expanded into the calls it actually makes:

```
   CLIENT                           ║  Calc.dll  —  the code you write
   ═════════════════════════════════╬════════════════════════════════════════
                                    ║
   CoCreateInstance(CLSID, ...)     ║
        │                           ║
        │  COM: read the registry,  ║
        │       then LoadLibrary    ║
        │                           ║
        │──1────────────────────────╫─►  DllGetClassObject(CLSID, IID_IClassFactory)
        │                           ║              │  returns the singleton
        │                           ║              ▼
        │◄─2── IClassFactory* ──────╫───  [ CalculatorFactory ]
        │                           ║              │
        │──3── CreateInstance() ────╫──────────────┤  creates a new object
        │                           ║              ▼
        │◄─4── ICalculator* ────────╫───  [ Calculator instance ]
        │                           ║
        │──5── Release() factory ───╫─►  (factory gone; the instance lives on)
        │                           ║
   pCalc->Add(40, 2, &r) ───────────╫─►  Calculator::Add() executes here
                                    ║
```

Steps 1, 2 and 5 are exactly what `CoCreateInstance` hides from you. Call `CoGetClassObject` yourself instead and you keep the factory alive, so step 3 can be repeated as often as you like — which is the entire reason the factory is a separate object.

That extra hop buys four things:

| | |
|---|---|
| **Language independence** | The client calls a documented COM interface, never a named function inside your binary. |
| **Batch creation** | `CoGetClassObject` once, then `CreateInstance` N times — one registry lookup and one DLL load instead of N (§2.4). |
| **Server lifetime control** | `LockServer(TRUE)` pins the server in memory between creations, so it isn't unloaded and reloaded each time. |
| **Uniformity across server types** | An in-proc server hands its factory out through `DllGetClassObject`; an EXE server registers the same factory with `CoRegisterClassObject` at startup (Module 7). Same interface, entirely different plumbing, identical client code. |

Note what `CreateInstance` does **not** take: constructor arguments. COM objects are always created empty and initialized afterwards through an interface — which is exactly why `IPersistFile`, `IInitializeWithStream` and friends exist.

```cpp
struct IClassFactory : public IUnknown
{
    virtual HRESULT CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) = 0;
    virtual HRESULT LockServer(BOOL fLock) = 0;
};
```

- `CreateInstance` makes one object. `pUnkOuter` is for **aggregation** — if it's non-null and you don't support aggregation, you **must** return `CLASS_E_NOAGGREGATION`.
- `LockServer(TRUE)` keeps the server loaded even with zero objects, so a client can hold the factory across creations.

#### What `pUnkOuter` actually is: aggregation

You'll write `if (pUnkOuter) return CLASS_E_NOAGGREGATION;` in a moment, so here's what you're declining.

**The situation.** Suppose someone wants to ship a `SecureCalculator`: the same arithmetic your `Calculator` already does, plus one extra interface of its own, `IAuditLog`. COM has **no implementation inheritance** — you cannot subclass another vendor's coclass, because all you were given is a GUID and a vtable. So there are exactly two ways to reuse the component, and both involve the outer object creating a private instance of the inner one.

**Option 1 — containment (delegation).** The outer object re-implements every method by forwarding:

```cpp
STDMETHODIMP CSecureCalculator::Add(long a, long b, long* r)
{
    return m_pInner->Add(a, b, r);       // one forwarding stub per method, forever
}
```

Simple, always legal, and the client never learns the inner object exists. The cost is a stub per method plus an indirection per call — for methods you are not changing in any way.

**Option 2 — aggregation.** The outer object writes **no forwarding code at all**. When the client asks `SecureCalculator` for `ICalculator`, it hands back a pointer that physically belongs to the **inner** object. Calls go straight there, at full speed.

**The catch: the client must not be able to tell.** Either way the client believes it is holding *one* object that happens to support both `ICalculator` and `IAuditLog`:

```
                what the client believes            what is really there
                ------------------------            --------------------
                  ┌──────────────┐                  ┌─ OUTER ──────────┐
   pAudit ──────► │              │                  │ IAuditLog        │
                  │   ONE object │                  │                  │
   pCalc  ──────► │              │                  │  ┌─ INNER ─────┐ │
                  └──────────────┘                  │  │ ICalculator │ │
                                                    │  └─────────────┘ │
                                                    └──────────────────┘
```

And that is where the difficulty lies, because of Module 1's Rule 1: **every interface of one object must return the same `IUnknown`.** If the outer naively handed out the inner's pointer, the client could call `QI(IID_IUnknown)` on it and get the *inner's* identity — two different answers from what is supposed to be one object, with its reference count split across both halves.

The fix is that an aggregatable object implements **two** `IUnknown`s:

| | Purpose |
|---|---|
| **Non-delegating `IUnknown`** | The inner object's real one. Only the *outer* object uses it, to control the inner's lifetime. |
| **Delegating `IUnknown`** | What the inner exposes on all its *other* interfaces. Every call forwards to `pUnkOuter`. |

So when a client calls `QueryInterface`, `AddRef`, or `Release` on the `ICalculator` it got from the aggregate, the call lands on the **outer** object — one identity, one reference count, Rule 1 preserved.

```
   Client ──► ICalculator (physically implemented by the INNER object)
                 │
                 └─ its QI/AddRef/Release delegate to ──► OUTER object's IUnknown
```

**Where `pUnkOuter` comes from.** It is the outer object handing the inner a pointer to itself, at creation time:

```cpp
// inside CSecureCalculator's initialization:
CoCreateInstance(CLSID_Calculator,
                 static_cast<IUnknown*>(this),   // pUnkOuter - "I am your outer object"
                 CLSCTX_INPROC_SERVER,
                 IID_IUnknown,                   // MUST be IUnknown when aggregating
                 (void**)&m_pInnerUnknown);
```

So `pUnkOuter` is one object saying to another: *"you are being aggregated — send all your `QueryInterface`/`AddRef`/`Release` traffic to me."* A **null** `pUnkOuter` — what every ordinary `CoCreateInstance` passes — means "you are standing alone, be your own identity."

That also explains the odd `IID_IUnknown` restriction: while aggregating, the factory may return *only* the non-delegating `IUnknown`. Handing back `ICalculator` at this point would expose the inner's separate identity before the delegation is wired up, and Rule 1 would already be broken.

**In practice:** aggregation is rare in new code. It's fiddly, easy to get wrong, and containment is nearly always fast enough. ATL supports it via `DECLARE_AGGREGATABLE` / `CComAggObject` (Module 6) if you need it. **Returning `CLASS_E_NOAGGREGATION` is a perfectly respectable, and by far the most common, answer.**

What you must know for support work: if a client passes a non-null `pUnkOuter` and the server ignores it instead of refusing, you get an object that violates Rule 1 — producing intermittent `E_NOINTERFACE`, duplicate identities, and premature or missed destruction. Module 1's "intermittent `E_NOINTERFACE` is always a bug" note points here.

### 2.5.2 Complete server: `Calc.cpp`

```cpp
#include <windows.h>
#include <objbase.h>
#include <olectl.h>       // SELFREG_E_CLASS
#include <new>
#include <strsafe.h>
#include "Calculator.h"   // ICalculator from Module 1

// {A1B2C3D4-1111-4000-9000-000000000001}   <-- generate your own!
DEFINE_GUID(CLSID_Calculator,
    0xa1b2c3d4, 0x1111, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01);

static LONG   g_cObjects = 0;      // live objects
static LONG   g_cLocks   = 0;      // LockServer count
static HMODULE g_hModule = nullptr;

// ------------------------------------------------------------------ object
class Calculator : public ICalculator
{
    LONG m_cRef = 1;
public:
    Calculator()  { InterlockedIncrement(&g_cObjects); }
    ~Calculator() { InterlockedDecrement(&g_cObjects); }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == IID_ICalculator)
            *ppv = static_cast<ICalculator*>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&m_cRef); }
    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG n = InterlockedDecrement(&m_cRef);
        if (!n) delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE Add(long a, long b, long* r) override
    { if (!r) return E_POINTER; *r = a + b; return S_OK; }
    HRESULT STDMETHODCALLTYPE Subtract(long a, long b, long* r) override
    { if (!r) return E_POINTER; *r = a - b; return S_OK; }
};

// ----------------------------------------------------------------- factory
class CalculatorFactory : public IClassFactory
{
public:
    // The factory is a singleton with an artificially high ref count; it is
    // never destroyed, so QI/AddRef/Release are trivial.
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == IID_IClassFactory)
            *ppv = static_cast<IClassFactory*>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef()  override { return InterlockedIncrement(&g_cLocks); }
    ULONG STDMETHODCALLTYPE Release() override { return InterlockedDecrement(&g_cLocks); }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (pUnkOuter) return CLASS_E_NOAGGREGATION;   // we don't support aggregation

        Calculator* p = new (std::nothrow) Calculator();
        if (!p) return E_OUTOFMEMORY;
        HRESULT hr = p->QueryInterface(riid, ppv);
        p->Release();
        return hr;
    }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override
    {
        if (fLock) InterlockedIncrement(&g_cLocks);
        else       InterlockedDecrement(&g_cLocks);
        return S_OK;
    }
};

static CalculatorFactory g_factory;   // static instance; never freed

// ------------------------------------------------------------- DLL exports
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (!ppv) return E_POINTER;
    *ppv = nullptr;
    if (rclsid != CLSID_Calculator) return CLASS_E_CLASSNOTAVAILABLE;
    return g_factory.QueryInterface(riid, ppv);
}

STDAPI DllCanUnloadNow(void)
{
    return (g_cObjects == 0 && g_cLocks == 0) ? S_OK : S_FALSE;
}

// ------------------------------------------------------------ registration
static HRESULT SetKeyValue(HKEY root, PCWSTR subkey, PCWSTR name, PCWSTR value)
{
    HKEY hKey = nullptr;
    LONG rc = RegCreateKeyExW(root, subkey, 0, nullptr, REG_OPTION_NON_VOLATILE,
                              KEY_WRITE, nullptr, &hKey, nullptr);
    if (rc != ERROR_SUCCESS) return HRESULT_FROM_WIN32(rc);
    rc = RegSetValueExW(hKey, name, 0, REG_SZ,
                        reinterpret_cast<const BYTE*>(value),
                        static_cast<DWORD>((wcslen(value) + 1) * sizeof(WCHAR)));
    RegCloseKey(hKey);
    return HRESULT_FROM_WIN32(rc);
}

STDAPI DllRegisterServer(void)
{
    WCHAR modulePath[MAX_PATH];
    if (!GetModuleFileNameW(g_hModule, modulePath, ARRAYSIZE(modulePath)))
        return HRESULT_FROM_WIN32(GetLastError());

    WCHAR clsidStr[64];
    StringFromGUID2(CLSID_Calculator, clsidStr, ARRAYSIZE(clsidStr));

    WCHAR key[256];
    // HKCR\CLSID\{...}
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s", clsidStr);
    if (FAILED(SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, L"Calculator Component")))
        return SELFREG_E_CLASS;

    // HKCR\CLSID\{...}\InprocServer32
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s\\InprocServer32", clsidStr);
    if (FAILED(SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, modulePath)))
        return SELFREG_E_CLASS;
    if (FAILED(SetKeyValue(HKEY_CLASSES_ROOT, key, L"ThreadingModel", L"Both")))
        return SELFREG_E_CLASS;

    // HKCR\CLSID\{...}\ProgID
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s\\ProgID", clsidStr);
    SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, L"Training.Calculator.1");

    // HKCR\Training.Calculator.1\CLSID
    SetKeyValue(HKEY_CLASSES_ROOT, L"Training.Calculator.1", nullptr, L"Calculator Component");
    SetKeyValue(HKEY_CLASSES_ROOT, L"Training.Calculator.1\\CLSID", nullptr, clsidStr);

    return S_OK;
}

STDAPI DllUnregisterServer(void)
{
    WCHAR clsidStr[64];
    StringFromGUID2(CLSID_Calculator, clsidStr, ARRAYSIZE(clsidStr));
    WCHAR key[256];

    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s", clsidStr);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, key);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, L"Training.Calculator.1");
    return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hInst, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hModule = hInst;
        DisableThreadLibraryCalls(hInst);   // we don't need thread notifications
    }
    return TRUE;
}
```

### 2.5.3 `Calc.def` — the exports

```
LIBRARY   Calc
EXPORTS
    DllGetClassObject   PRIVATE
    DllCanUnloadNow     PRIVATE
    DllRegisterServer   PRIVATE
    DllUnregisterServer PRIVATE
```

Set **Project → Linker → Input → Module Definition File** to `Calc.def`. Without this, name decoration hides your exports and `regsvr32` reports "entry-point DllRegisterServer was not found."

### 2.5.4 The client

```cpp
#include <windows.h>
#include <objbase.h>
#include <cstdio>
#include "Calculator.h"

DEFINE_GUID(CLSID_Calculator,
    0xa1b2c3d4, 0x1111, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01);

int wmain()
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) { wprintf(L"CoInitializeEx: 0x%08X\n", hr); return 1; }

    ICalculator* pCalc = nullptr;
    hr = CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_INPROC_SERVER,
                          IID_ICalculator, reinterpret_cast<void**>(&pCalc));
    if (SUCCEEDED(hr))
    {
        long r = 0;
        pCalc->Add(40, 2, &r);
        wprintf(L"40 + 2 = %ld\n", r);
        pCalc->Release();
    }
    else
    {
        wprintf(L"CoCreateInstance failed: 0x%08X\n", hr);
    }

    // Also demonstrate ProgID -> CLSID resolution.
    CLSID fromProgID{};
    hr = CLSIDFromProgID(L"Training.Calculator.1", &fromProgID);
    wprintf(L"CLSIDFromProgID: 0x%08X\n", hr);

    CoUninitialize();
    return 0;
}
```

---

## 2.6 LAB 2.1 — Build, register, activate, and watch it happen

> **Requirements**
> - **Tools:** Visual Studio C++; `regsvr32`; **Process Monitor** (Sysinternals).
> - **Elevation:** **required.** `regsvr32` writes to `HKLM\Software\Classes`, and ProcMon needs admin to load its driver. The per-user variant at the end of the lab deliberately runs *without* elevation.
> - **Bitness:** x64 DLL and x64 client. Keep them matched — mismatching them is Lab 2.2.
> - **Depends on:** the `Calc.dll` and `CalcClient.exe` sources from §2.5.
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/) — open `Stage2.sln`, build **Debug | x64**, then register `x64\Calc.dll` from an elevated prompt.
> - **Time:** ~90 min.

This whole module turns on one question: how does a GUID become a running object? Here you answer it by doing it — build the server, register it, activate it, and then **watch the SCM work** in Process Monitor.

The trace you capture is the reference picture of a *healthy* activation. Every failure in Lab 2.4 is a deviation from it, so keep it.

1. Build `Calc.dll` (x64) and `CalcClient.exe` (x64).
2. Register from an **elevated** prompt (HKCR writes need admin):
   ```powershell
   regsvr32 C:\Components\Calc.dll
   ```
3. Run the client. Expect `40 + 2 = 42`.
4. **Inspect the registry:**
   ```powershell
   $clsid = "{A1B2C3D4-1111-4000-9000-000000000001}"
   Get-ChildItem "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid" -Recurse |
       ForEach-Object { $_.Name; $_ | Get-ItemProperty | Format-List }
   ```
5. **Watch the SCM work.** Start Process Monitor, set a filter `Path contains A1B2C3D4-1111`, run the client, and read the trace. You will see the exact probe order: `RegOpenKey` on the CLSID, `InprocServer32`, `ThreadingModel`, then `CreateFile`/`Load Image` on the DLL. **Save this trace.** It is the reference picture of a *successful* activation; every failure is a deviation from it.
6. **Break it:** `regsvr32 /u C:\Components\Calc.dll`, run the client again. Expect `0x80040154`. Look at the ProcMon trace now — you'll see `NAME NOT FOUND` on the CLSID key. That's the signature.

### Register per-user instead of per-machine

Modify `DllRegisterServer` to write to `HKEY_CURRENT_USER\Software\Classes` instead of `HKCR`, rebuild, and register **without elevation**. Confirm:

- It works for your user.
- It does *not* work for another user account.
- `HKCU` shadows `HKLM` if both exist (register different DLL paths in each and see which wins).

This experiment is worth an hour; it explains a whole family of tickets.

---

## 2.7 LAB 2.2 — Bitness

> **Requirements**
> - **Tools:** Visual Studio with both **x86 and x64** configurations; both copies of `regsvr32` (`System32` = 64-bit, `SysWOW64` = 32-bit); Process Explorer.
> - **Elevation:** required.
> - **Bitness:** you need **all four** binaries — x86 and x64 of both DLL and client.
> - **Depends on:** Lab 2.1.
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/) — open `Stage2.sln` and build **both** platforms: once with the dropdown on **x64**, once on **x86**.
> - **Expected to fail:** step 6 (DLL surrogate) *cannot* succeed yet — the interface has no marshaling support. Record the failure and finish it in Lab 7.2.
> - **Time:** ~60 min.

A bitness mismatch is the most common activation failure there is, and it reports the *same* `0x80040154` as "never registered at all". Telling those two apart from the error code alone is impossible — you have to look at where the registration landed.

So cause it deliberately, confirm in the registry exactly which hive received the keys, then fix it two different ways.

1. Build `Calc.dll` as **x86**. Register with the 32-bit regsvr32:
   ```powershell
   C:\Windows\SysWOW64\regsvr32.exe C:\Components\x86\Calc.dll
   ```
2. Run the **x64** client. Expect `0x80040154`.
3. Confirm the cause: the registration went to `HKLM\Software\Classes\Wow6432Node\CLSID\{...}`:
   ```powershell
   Get-Item "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{A1B2C3D4-1111-4000-9000-000000000001}"
   Get-Item "HKLM:\SOFTWARE\Classes\CLSID\{A1B2C3D4-1111-4000-9000-000000000001}"   # not found
   ```
4. Run the **x86** client. It works.
5. **Fix option A:** build and register both bitnesses.
6. **Fix option B (the interesting one):** use a DLL surrogate so the 32-bit DLL runs out-of-process and any client bitness can reach it. Add an AppID and point the CLSID at it:

```powershell
$clsid = "{A1B2C3D4-1111-4000-9000-000000000001}"
$appid = "{B1B2C3D4-2222-4000-9000-000000000002}"

New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$appid" -Force |
    Set-ItemProperty -Name "(default)" -Value "Calculator Surrogate"
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\AppID\$appid" -Name "DllSurrogate" -Value ""
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\$clsid" -Name "AppID" -Value $appid
```

Now call with `CLSCTX_LOCAL_SERVER`. The object runs in `dllhost.exe`. Verify with Process Explorer.

> **Caveat:** surrogate activation requires the interface to be marshalable — a registered proxy/stub or a type library. Your `ICalculator` has neither yet, so this lab will fail with `E_NOINTERFACE` at the `QueryInterface` step. **That is the intended lesson.** Come back and finish this lab at the end of Module 4. Write the failure down now.

---

## 2.8 LAB 2.3 — Registration-free (Reg-Free) COM

> **Requirements**
> - **Tools:** Visual Studio C++; **`mt.exe`** (Windows SDK) to embed the manifests, or the linker's *Manifest Tool* property page; **`sxstrace.exe`** to diagnose activation-context failures.
> - **Elevation:** not required to *run* the lab — that is the entire point. You do need admin **once**, up front, to unregister the component.
> - **Bitness:** the manifest's `processorArchitecture` must match the build exactly (`amd64` for x64, `x86` for 32-bit). A mismatch fails silently and falls back to the registry.
> - **Depends on:** Lab 2.1 binaries, then **fully unregistered** (`regsvr32 /u`) — otherwise the registry satisfies the activation and you prove nothing.
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/), built but **not** registered.
> - **Time:** ~90 min.

Modern deployment avoids the registry entirely: no admin rights, no machine-wide state, side-by-side versions, and clean uninstall. Activation data comes from **manifests** read into the process's **activation context**.

### Step 1 — the server's assembly manifest, `Calc.manifest`

```xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
      type="win32"
      name="Training.Calc"
      version="1.0.0.0"
      processorArchitecture="amd64" />
  <file name="Calc.dll">
    <comClass
        clsid="{A1B2C3D4-1111-4000-9000-000000000001}"
        threadingModel="Both"
        progid="Training.Calculator.1"
        description="Calculator Component" />
  </file>
</assembly>
```

### Step 2 — the client's application manifest, `CalcClient.exe.manifest`

```xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
      type="win32"
      name="Training.CalcClient"
      version="1.0.0.0"
      processorArchitecture="amd64" />
  <dependency>
    <dependentAssembly>
      <assemblyIdentity
          type="win32"
          name="Training.Calc"
          version="1.0.0.0"
          processorArchitecture="amd64" />
    </dependentAssembly>
  </dependency>
</assembly>
```

### Step 3 — deploy and test

```
C:\RegFreeTest\
    CalcClient.exe
    CalcClient.exe.manifest
    Calc.dll
    Calc.manifest              <- name must match assemblyIdentity name + ".manifest"
```

1. **Fully unregister** the DLL: `regsvr32 /u Calc.dll`, and delete any leftover keys.
2. Confirm the client now fails with `0x80040154`.
3. Drop in the manifests. Run again — it works, with **nothing in the registry**.
4. Verify in Process Monitor: no `HKCR\CLSID` probe at all. The activation context short-circuits it.

### Common reg-free failures (memorize these)

| Symptom | Cause |
|---|---|
| `0x80040154` still | Manifest not found (wrong filename), or the EXE has an *embedded* manifest which takes precedence over the external `.exe.manifest` file |
| `0x800736B1` `ERROR_SXS_...` / app won't start at all | Malformed XML, mismatched `assemblyIdentity`, wrong `processorArchitecture` |
| Works for the EXE but not from a DLL in the same process | Activation context is per-thread; a DLL must use `CreateActingContext`/`ActivateActCtx`, or embed its own dependency |
| Works in debug, fails when deployed | The manifest wasn't copied, or VS embedded a manifest that omits the dependency |

Diagnostics: check the **Event Viewer → Applications and Services Logs → Microsoft → Windows → SideBySide** log. `sxstrace.exe trace -logfile:sxs.etl` then `sxstrace parse` gives an exact reason.

> **Note on embedded manifests:** if the linker embedded a manifest into `CalcClient.exe` (VS does by default), the external `.exe.manifest` is ignored. Either disable manifest embedding (`Linker → Manifest File → Generate Manifest = No`) or add your dependency to the embedded one via *Additional Manifest Files*.

---

## 2.9 Server lifetime and unloading

### In-proc

- COM calls `DllCanUnloadNow` when it feels like it — mainly from `CoFreeUnusedLibraries` / `CoFreeUnusedLibrariesEx`.
- Return `S_OK` only when `g_cObjects == 0 && g_cLocks == 0`.
- **Never** call `FreeLibrary` on yourself. **Never** unload while a call is in flight.
- A DLL that never returns `S_OK` simply stays loaded — an annoyance, not a bug.
- A DLL that returns `S_OK` too eagerly gets unloaded with live objects → the client's next vtable call jumps into unmapped memory. Classic `0xC0000005` with a garbage return address. **Support signature:** crash address is in no module; `!address <ptr>` says "free".

### Out-of-proc

- The EXE calls `CoRegisterClassObject` for each CLSID at startup, then pumps messages.
- It calls `CoRevokeClassObject` and exits when object and lock counts hit zero.
- Timing bug: exiting between "count hits zero" and "revoke completes" races with an incoming activation → `CO_E_SERVER_EXEC_FAILURE` or `RPC_E_DISCONNECTED` for the unlucky client. Real servers use `CoAddRefServerProcess`/`CoReleaseServerProcess`, which handle this correctly:

```cpp
// In every object's constructor / factory LockServer(TRUE):
CoAddRefServerProcess();

// In every destructor / LockServer(FALSE):
if (CoReleaseServerProcess() == 0)
    PostQuitMessage(0);      // safe: COM has already revoked the class objects
```

---

## 2.10 The HRESULT triage table

This is the heart of the support track. For each row: know the symptom, the first check, and the fix.

| HRESULT | Symbol | Real-world causes | First diagnostic |
|---|---|---|---|
| `0x80040154` | `REGDB_E_CLASSNOTREG` | Not registered; **bitness mismatch**; HKCU vs HKLM; missing manifest; wrong CLSID | ProcMon filter on the CLSID → look for `NAME NOT FOUND`; check client & server bitness |
| `0x80040155` | `REGDB_E_IIDNOTREG` | Interface has no proxy/stub or typelib registration; only surfaces across a boundary | Check `HKCR\Interface\{IID}\ProxyStubClsid32` |
| `0x80040111` | `CLASS_E_CLASSNOTAVAILABLE` | DLL loaded fine but `DllGetClassObject` doesn't recognize that CLSID | Registration points at the wrong DLL, or CLSID typo |
| `0x8007007E` | `ERROR_MOD_NOT_FOUND` | The server DLL was found but **one of its dependencies** wasn't (VC++ redist, a helper DLL) | ProcMon for `NAME NOT FOUND` on `*.dll`; check with Dependencies.exe |
| `0x8007000E` | `E_OUTOFMEMORY` | Genuine OOM; or 32-bit address space exhaustion | Check commit/private bytes |
| `0x80070005` | `E_ACCESSDENIED` | DCOM Launch/Activation permission; integrity level mismatch (medium client → high server); file ACL on the DLL | Event Viewer → System → DistributedCOM **10016**; `dcomcnfg` |
| `0x80080005` | `CO_E_SERVER_EXEC_FAILURE` | EXE server crashed at startup, or didn't `CoRegisterClassObject` within the timeout; wrong `RunAs` identity/password; session 0 issue | Try launching the EXE manually; check Application event log for the server's crash |
| `0x800706BA` | `RPC_S_SERVER_UNAVAILABLE` | Remote machine unreachable; firewall blocking TCP 135 or the dynamic port range; `RpcSs` stopped | `Test-NetConnection host -Port 135` |
| `0x80010108` | `RPC_E_DISCONNECTED` | The server process died while you held a proxy | Check for a crash dump of the server |
| `0x8001010E` | `RPC_E_WRONG_THREAD` | Raw interface pointer used on a different apartment | Module 3 |
| `0x800401F0` | `CO_E_NOTINITIALIZED` | `CoInitializeEx` not called on this thread, or already `CoUninitialize`d | Check every thread that touches COM |
| `0x80004002` | `E_NOINTERFACE` | Object genuinely lacks it; **or** marshaling isn't registered and you crossed a boundary | Does it work in-proc? If yes → marshaling |
| `0x80070422` | service disabled | `RpcSs`/`DcomLaunch` disabled | `Get-Service RpcSs, DcomLaunch` |
| `0x80029C4A` | `TYPE_E_CANTLOADLIBRARY` | Type library missing/unregistered/wrong bitness | Check `HKCR\TypeLib\{LIBID}` |

### Decode any HRESULT fast

```powershell
# PowerShell
[ComponentModel.Win32Exception]::new(0x8007007E).Message

# or
certutil -error 0x8007007e
```

```
0:000> !error 0x8007007e        ; WinDbg
```

---

## 2.11 LAB 2.4 — Reproduce every failure deliberately (support drill)

> **Requirements**
> - **Tools:** `regsvr32`, Registry Editor, **Process Monitor**, Event Viewer, `icacls`.
> - **Elevation:** required — you edit `HKCR` values and registry ACLs.
> - **Bitness:** x86 and x64 builds (row 2).
> - **Depends on:** Labs 2.1 and 2.2.
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/), both bitnesses built. Copy the folder first — this lab breaks things on purpose.
> - **Caution:** run this on a **VM or dedicated test machine**. Rows 8 and 9 deny *your own account* Read access to a registry key and a file. Export the key and record the original ACL (`icacls <file> /save`) before you change anything, and restore it at the end.
> - **Time:** ~2 h.

This is the most valuable lab in the module. For each row below, **cause it on purpose** and record the ProcMon/Event Viewer signature in your notes.

| # | How to cause it | Expected |
|---|---|---|
| 1 | Unregister the DLL | `0x80040154` |
| 2 | Register x86 DLL, run x64 client | `0x80040154` |
| 3 | Point `InprocServer32` at a non-existent path | `0x8007007E` or `0x80070002` |
| 4 | Point `InprocServer32` at a real DLL that isn't a COM server | `0x80040111` (no `DllGetClassObject`) or `0x8007007F` |
| 5 | Build the DLL against a VC++ runtime not installed on the box | `0x8007007E` |
| 6 | Make `DllGetClassObject` return `CLASS_E_CLASSNOTAVAILABLE` for the right CLSID | `0x80040111` |
| 7 | Remove `CoInitializeEx` from the client | `0x800401F0` |
| 8 | Deny your own account Read on the `HKCR\CLSID\{...}` key | `0x80040154` (ProcMon shows `ACCESS DENIED`, not `NAME NOT FOUND` — learn the difference!) |
| 9 | Deny Read on the DLL file itself | `0x80070005` |
| 10 | Register a `LocalServer32` pointing at an EXE that exits immediately | `0x80080005` |

Row 8 vs row 1 is the key discrimination skill: **`NAME NOT FOUND` = not registered; `ACCESS DENIED` = registered but you can't read it.** Both surface as `0x80040154` to the client. Only ProcMon distinguishes them.

---

## 2.12 Tooling: OleView.NET

Install from https://github.com/tyranid/oleviewdotnet. Then practise:

- **Registry → CLSIDs** — find your `Training.Calculator.1`, inspect its server, threading model, AppID.
- **Registry → CLSIDs by Server** — "which classes does this DLL implement?" Enormously useful when a customer says "installing X broke Y."
- **Registry → Interfaces** — check whether an IID has a proxy/stub.
- **Registry → AppIDs with Access permissions** — the security view for Module 7.
- **Object → Create Instance** — activate a class interactively and browse its interfaces without writing a client. This alone saves hours.
- **Diff two machines** — export the registry view from a working and a broken machine and compare. This is the fastest route on "works on my machine" cases.

### A useful PowerShell reconnaissance snippet

```powershell
function Get-ComRegistration {
    param([string]$Clsid)
    $paths = @(
        "HKLM:\SOFTWARE\Classes\CLSID\$Clsid",
        "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\$Clsid",
        "HKCU:\SOFTWARE\Classes\CLSID\$Clsid"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            [pscustomobject]@{
                Hive    = $p
                Inproc  = (Get-ItemProperty "$p\InprocServer32" -EA SilentlyContinue).'(default)'
                Threading = (Get-ItemProperty "$p\InprocServer32" -EA SilentlyContinue).ThreadingModel
                Local   = (Get-ItemProperty "$p\LocalServer32" -EA SilentlyContinue).'(default)'
                AppID   = (Get-ItemProperty $p -EA SilentlyContinue).AppID
            }
        }
    }
}

Get-ComRegistration "{A1B2C3D4-1111-4000-9000-000000000001}" | Format-List
```

Keep this in your toolkit. Run it as the *first* step on any `REGDB_E_CLASSNOTREG` ticket — it answers "which hive, which bitness, which server" in one shot.

---

## 2.13 Checkpoint

1. A customer's 64-bit app fails with `0x80040154`. The vendor insists the component is registered, and `regedit` on the customer's box shows the CLSID under `HKEY_CLASSES_ROOT\CLSID`. What's your next question, and why might `regedit` be misleading you?
2. What's the difference between `CoGetClassObject` and `CoCreateInstance`, and when would you deliberately use the former?
3. Why must `IClassFactory::CreateInstance` return `CLASS_E_NOAGGREGATION` when `pUnkOuter` is non-null and you don't support aggregation?
4. Your DLL returns `S_OK` from `DllCanUnloadNow` while a client still holds an object. Describe the crash and how you'd recognize it in a dump.
5. Under what circumstances does `HKCR\Interface\{IID}` get read at all?
6. A component works when the test app runs interactively but fails with `0x80070005` when the same code runs in a Windows service. Name three plausible causes.
7. In ProcMon, you see `RegOpenKey HKCR\CLSID\{...} ACCESS DENIED`. What does that rule out?

<details>
<summary>Answers</summary>

1. **"What bitness is the app, and what bitness is the registered DLL?"** `regedit` running as 64-bit shows `HKCR\CLSID` = the 64-bit view; a 32-bit-only registration lives under `Wow6432Node` and would *also* appear at `HKCR\CLSID` when viewed by a 32-bit tool. Also ask whether the registration is in HKCU for a *different* user than the one the app runs as.

2. `CoCreateInstance` = `CoGetClassObject` + `CreateInstance` + `Release` of the factory. Use `CoGetClassObject` when creating many instances (amortize the SCM round-trip and registry lookup), when you need `IClassFactory2` licensing, or when you want to hold the server loaded via `LockServer`.

3. Because the aggregator has already committed to the aggregation contract — it will delegate its own `IUnknown` to the inner object and expects the inner object to delegate back. If you ignore `pUnkOuter` and return a normal object, you'd break `QueryInterface` identity (Rule 1) for the aggregate, producing an object that violates COM's rules in ways that fail unpredictably later.

4. The DLL is unmapped while the client holds an interface pointer. The client's next call loads the vptr (still pointing into the now-unmapped image) and jumps to unmapped memory → `0xC0000005` with an instruction pointer that belongs to **no loaded module**. In a dump: `lm` doesn't cover the faulting address, `!address <eip>` reports FREE/RESERVE, and the stack shows the client calling through an interface. Fix: correct the object/lock counting.

5. Only when a call must cross an apartment or process/machine boundary — i.e. when COM needs to build a proxy/stub pair. Purely in-proc, same-apartment calls never touch it.

6. (a) **Session 0 isolation** — a service can't drive an interactive-user COM server. (b) **DCOM Launch/Activation permissions** don't grant the service account. (c) The registration is under **HKCU of the interactive user**, invisible to `LocalSystem`/`NetworkService`. (Bonus: the service account lacks NTFS read on the DLL, or the service runs at a different integrity level.)

7. It rules out "not registered." The key **exists**; the calling process's token can't read it. Look at the key's ACL and the caller's identity/integrity level — this is a permissions ticket, not a deployment ticket.

</details>

---

**Next: [Module 3 — Threading and apartments](03-apartments-and-threading.md)**
