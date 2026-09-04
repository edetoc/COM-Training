# Module 4 — Interfaces, IDL, MIDL, and marshaling

Module 3 showed that crossing an apartment boundary requires a proxy. This module explains where proxies come from, how COM knows how to package your parameters, and who owns the memory.

Two names appear throughout, and they are a language and its compiler:

| | | |
|---|---|---|
| **IDL** | *Interface Definition Language* | A small, C-like language for describing an interface **completely** — not just method names and types, but which parameters go in, which come out, how big the arrays are, and who frees what. You write `.idl` files by hand. |
| **MIDL** | *Microsoft Interface Definition Language compiler* | `midl.exe`, the tool that reads your `.idl` and generates C/C++ headers, GUID definitions, the proxy/stub source, and a type library. You never write its output by hand. |

The relationship is the ordinary one between source and compiler: **you write IDL, MIDL compiles it.** A third term, **NDR** (*Network Data Representation*), is the wire format MIDL's generated code uses to lay parameters out as bytes; it appears in §4.10 when you learn to read its failures.

**What this module covers**

Why a C++ header cannot describe an interface completely, and what IDL adds: parameter direction, array sizes, pointer semantics, and the memory-ownership rules that settle who allocates and who frees. You run MIDL and read everything it generates — including the type library, decompiled back into IDL, to see which parts of your interface a scripting language can and cannot reach. Then the three kinds of marshaling and when COM picks each, the one versioning rule you must never break, and how to read an NDR failure when the wire format and the code no longer agree.

**Contents**

- [4.1 Why IDL exists](#41-why-idl-exists)
- [4.2 IDL by example](#42-idl-by-example)
- [4.3 The memory-management rules](#43-the-memory-management-rules)
- [4.4 The three kinds of marshaling](#44-the-three-kinds-of-marshaling)
- [4.5 The "works in-proc, fails out-of-proc" bug](#45-the-works-in-proc-fails-out-of-proc-bug)
- [4.6 LAB 4.1 — Author IDL, compile with MIDL, read the output](#46-lab-41--author-idl-compile-with-midl-read-the-output)
- [4.7 LAB 4.2 — Break and fix marshaling](#47-lab-42--break-and-fix-marshaling)
- [4.8 LAB 4.3 — Memory ownership, verified](#48-lab-43--memory-ownership-verified)
- [4.9 Versioning: never modify a published interface](#49-versioning-never-modify-a-published-interface)
- [4.10 Reading NDR failures](#410-reading-ndr-failures)
- [4.11 Checkpoint](#411-checkpoint)
- [4.12 Rules to carry forward](#412-rules-to-carry-forward)

---

## 4.1 Why IDL exists

### First: why the labs so far did *not* need it

You have built a working COM component across three modules with no IDL anywhere. That is worth explaining before adding a new tool, because it tells you exactly what IDL is for.

COM's runtime contract is smaller than people assume. To create and call an object, all COM requires is:

- a **vtable** with the methods in a known order, and
- a **registered CLSID** pointing at the DLL.

Your hand-written `Calculator.h` — an ordinary C++ header file — supplied both: the `struct ICalculator : IUnknown` gave the C++ compiler the vtable layout, and `DEFINE_GUID` gave you the IIDs. Nothing else was needed, because **every call you made was a direct vtable call** — client and object in the same process, in the same apartment. No description of the parameters is required to jump to a function pointer with the arguments already on the stack.

IDL becomes necessary the moment something has to *understand* the call rather than merely make it:

| Situation | Who needs the description | Where you met it |
|---|---|---|
| Call crosses an apartment or process boundary | The marshaler, to pack parameters into bytes | Lab 3.1, experiment 4 — `REGDB_E_IIDNOTREG` |
| Caller is not C++ | A type library, read at runtime | Modules 5 and 6 |
| Caller binds by name at runtime | A scripting engine | Module 5 |

That failed line at the end of Lab 3.1 was the first time the header was not enough. This module is the fix.

> **In real projects the order is usually reversed:** you write the IDL first and let MIDL generate the header. This course wrote the header by hand precisely so that you can now see exactly which parts MIDL takes over — and which extra facts it captures that a header cannot.

### What `Calculator.h` cannot say

By "a C++ header" this module means an ordinary `.h` file — concretely, the [`Calculator.h`](../labs/stage-2-inproc-server/Calculator.h) you have been including since Module 1, with its `DEFINE_GUID` lines and its `struct ICalculator : public IUnknown`. Nothing more exotic than that.

Such a file describes an interface to a *C++ compiler*, and that is all. COM needs more:

- **The marshaler** needs to know that `[out] BSTR* p` means "the callee allocates a string and the caller frees it," and that `[in, size_is(n)] BYTE* buf` means "copy `n` bytes across the wire."
- **Other languages** need a machine-readable description. A `.h` file is useless to PowerShell or C#.
- **Scripting engines** need names and parameter types at runtime, long after your compiler has finished.

A header can't express any of that. Look at a line like `long* result` in your own interface: is it in or out? One integer or an array? If an array, how long? Your C++ compiler does not care — it only needs the size of a pointer — but anything that has to *transmit* that parameter needs every one of those answers.

**IDL** (Interface Definition Language) is the language that removes the ambiguity. **MIDL** is the compiler that turns IDL into:

| Output | Purpose |
|---|---|
| `Foo.h` | C/C++ interface declarations |
| `Foo_i.c` | GUID definitions (`IID_IFoo`, `CLSID_Bar`) |
| `Foo_p.c` | **Proxy and stub code** — the marshaling implementation |
| `dlldata.c` | The boilerplate that turns that code into a real COM in-proc server: `DllGetClassObject`, `DllRegisterServer`, and friends |
| `Foo.tlb` | **Type library** — binary metadata for late binding and interop |

The last two combine into a **proxy/stub DLL** — commonly shortened to **PS DLL**, and named `CalcPS.dll` in this course. It is an ordinary COM in-proc server; the only unusual thing about it is that the objects it creates are proxies and stubs rather than components of your own.

---

## 4.2 IDL by example

```idl
// Calculator.idl
import "oaidl.idl";
import "ocidl.idl";

[
    object,                                 // this is a COM interface (derives IUnknown)
    uuid(A1B2C3D4-0001-4000-9000-000000000001),
    pointer_default(unique),
    helpstring("Basic calculator interface")
]
interface ICalculator : IUnknown
{
    HRESULT Add([in] LONG a, [in] LONG b, [out, retval] LONG* result);

    HRESULT Subtract([in] LONG a, [in] LONG b, [out, retval] LONG* result);

    // Caller supplies a buffer; callee writes into it. Caller owns it.
    HRESULT Checksum([in, size_is(cb)] const BYTE* data,
                     [in] ULONG cb,
                     [out, retval] ULONG* checksum);

    // Callee allocates; CALLER frees with SysFreeString.
    HRESULT Describe([out, retval] BSTR* description);

    // Callee allocates an array; caller frees with CoTaskMemFree.
    HRESULT GetHistory([out] ULONG* count,
                       [out, size_is(, *count)] LONG** values);

    // Property-style accessors.
    [propget] HRESULT Precision([out, retval] LONG* value);
    [propput] HRESULT Precision([in] LONG value);
}

[
    uuid(A1B2C3D4-9999-4000-9000-000000000099),
    version(1.0),
    helpstring("Training Calculator 1.0 Type Library")
]
library TrainingCalcLib
{
    importlib("stdole2.tlb");

    [
        uuid(A1B2C3D4-1111-4000-9000-000000000001),
        helpstring("Calculator Component")
    ]
    coclass Calculator
    {
        [default] interface ICalculator;
    };
};
```

### Interface attributes

| Attribute | Meaning |
|---|---|
| `object` | This is a COM interface (as opposed to plain DCE RPC). Implies it derives from `IUnknown`. |
| `uuid(...)` | The IID. |
| `pointer_default(unique)` | Default pointer semantics for unattributed pointers. `unique` = may be null, no aliasing. |
| `dual` | Derives from `IDispatch` *and* has a vtable. Module 5. |
| `oleautomation` | Restricted to Automation-compatible types → can use the universal (typelib) marshaler. |
| `local` | Not marshaled; no proxy/stub generated. |
| `helpstring` | Documentation, surfaced in type library browsers. |

### Parameter direction attributes

| Attribute | Direction | Who allocates | Who frees |
|---|---|---|---|
| `[in]` | client → server | caller | caller |
| `[out]` | server → client | **callee** | **caller** |
| `[in, out]` | both | caller allocates initial; callee may replace | caller frees final |
| `[out, retval]` | like `[out]`, and it's the "return value" for script/VB/C# | callee | caller |

Rules:

- `[out]` parameters must be **pointers to something** (`LONG*`, `BSTR*`, `IFoo**`).
- Only the **last** parameter may be `[out, retval]`, and there can be only one.
- `[out]` parameters must be **set to null/zero on every failure path**. Non-negotiable — marshaling code and callers both depend on it.

### Array attributes

| Attribute | Meaning |
|---|---|
| `size_is(n)` | The buffer holds `n` elements (allocated size) |
| `length_is(n)` | Only the first `n` elements are meaningful (transmitted count) |
| `size_is(, *pn)` | For `Type**`: the *pointee* array has `*pn` elements — used with `[out]` arrays |
| `first_is`, `last_is`, `max_is` | Sub-range transmission (rare) |
| `string` | Null-terminated; the marshaler computes the length |

The empty first slot in `size_is(, *count)` is not a typo: each comma-separated position corresponds to a level of pointer indirection. `LONG** values` has two levels; the first level is a single pointer (no size), the second has `*count` elements.

### Pointer attributes

| Attribute | Meaning | Cost |
|---|---|---|
| `ref` | Must not be null, no aliasing | cheapest |
| `unique` | May be null, no aliasing | cheap |
| `ptr` | Full pointer: may be null, may alias, cycles allowed | **expensive** — the marshaler builds an alias table |

Default to `unique` unless you genuinely need aliasing.

### `iid_is` — runtime-typed interface pointers

```idl
HRESULT QueryService([in] REFGUID guidService,
                     [in] REFIID riid,
                     [out, iid_is(riid)] void** ppv);
```

`iid_is(riid)` tells the marshaler "the interface type of `ppv` is whatever GUID is in `riid`." Without it, the marshaler cannot know what to marshal. Every `QueryInterface`-shaped method needs this.

---

## 4.3 The memory-management rules

These are absolute, and violating them corrupts the heap.

### Rule 1 — The COM task allocator

```cpp
void* p = CoTaskMemAlloc(cb);
CoTaskMemFree(p);           // ok on NULL, like free()
void* q = CoTaskMemRealloc(p, cb2);
```

Anything a callee allocates for a caller uses this allocator. It's process-wide and shared by all COM code, which is exactly why it exists (recall Module 0: the client and server may have different CRT heaps).

### Rule 2 — Who frees what

| Case | Callee | Caller |
|---|---|---|
| `[in]` buffer | reads only | allocates and frees |
| `[out]` buffer | **allocates** with `CoTaskMemAlloc` | **frees** with `CoTaskMemFree` |
| `[in, out]` buffer | may free the old and allocate a new one | frees the final value |
| `[out]` interface pointer | `AddRef`s | `Release`s |
| Failure | must set `[out]` to null and free anything it allocated | frees nothing |

### Rule 3 — `BSTR` has its own allocator

```cpp
BSTR bstr = SysAllocString(L"hello");
SysFreeString(bstr);
```

A `BSTR` is a `WCHAR*` that points **4 bytes past** the start of its allocation; the preceding `DWORD` holds the byte length. So:

- `SysStringLen(bstr)` is O(1) and **may include embedded nulls**.
- `CoTaskMemFree(bstr)` frees the *wrong address* → heap corruption.
- `SysFreeString(NULL)` is legal and does nothing.
- An empty `BSTR` and a `NULL` `BSTR` are different things; many APIs treat `NULL` as "".

```
  allocation start
        │
        ▼
   ┌────────┬──────────────────────────────┬────┐
   │ DWORD  │  h e l l o                   │ \0 │
   │ len=10 │                              │    │
   └────────┴──────────────────────────────┴────┘
                ▲
                └── the BSTR pointer you receive
```

Use `CComBSTR` / `_bstr_t` / `wil::unique_bstr` and you never touch this.

### Rule 4 — `SAFEARRAY` has its own allocator

`SafeArrayCreate` / `SafeArrayDestroy`. Use `CComSafeArray<T>`.

### Rule 5 — `VARIANT` owns whatever it holds

```cpp
VARIANT v;
VariantInit(&v);        // sets VT_EMPTY - ALWAYS do this first
// ... fill it ...
VariantClear(&v);       // frees BSTR / releases IUnknown / destroys SAFEARRAY as appropriate
```

`VariantClear` on an uninitialized `VARIANT` will free a garbage pointer. Use `CComVariant`, which does this correctly.

### The one-line summary

> **`BSTR` → `SysFreeString`. `SAFEARRAY` → `SafeArrayDestroy`. `VARIANT` → `VariantClear`. Interface → `Release`. Everything else `[out]` → `CoTaskMemFree`.**

---

## 4.4 The three kinds of marshaling

### Standard marshaling (MIDL proxy/stub)

MIDL generates NDR (Network Data Representation) format strings and proxy/stub code, compiled into a separate DLL.

```
Client apartment                        Server apartment
┌──────────────┐                        ┌──────────────┐
│   Client     │                        │   Object     │
│      │       │                        │      ▲       │
│      ▼       │                        │      │       │
│   Proxy      │  ── NDR-encoded ──►    │   Stub       │
│ (ICalculator │      buffer            │              │
│  vtable)     │  ◄── results ──        │              │
└──────────────┘                        └──────────────┘
       │                                        ▲
       └──► CRpcChannelBuffer ── LRPC/ALPC ─────┘
```

Registration:

```
HKCR\Interface\{IID_ICalculator}
    (Default)               = "ICalculator"
    NumMethods              = "8"
    ProxyStubClsid32        = "{CLSID of your PS DLL}"

HKCR\CLSID\{PS-CLSID}\InprocServer32
    (Default)               = "C:\Components\CalcPS.dll"
    ThreadingModel          = "Both"
```

- Handles **any** data type MIDL can describe.
- Fastest marshaling.
- Requires shipping and registering an extra DLL **in every process on both sides**.

### Type library marshaling (a.k.a. the universal or Automation marshaler)

If the interface is `[oleautomation]` or `[dual]`, and its type library is registered, COM uses `oleaut32.dll`'s built-in universal marshaler, which reads the type library at runtime.

```
HKCR\Interface\{IID}
    ProxyStubClsid32 = "{00020424-0000-0000-C000-000000000046}"   <- the universal marshaler
    TypeLib          = "{LIBID}"
        Version      = "1.0"
```

That GUID `{00020424-...}` is worth memorizing — when you see it under `ProxyStubClsid32`, you're looking at typelib marshaling, and the thing to verify is the **TypeLib** registration, not a proxy DLL.

- **No proxy/stub DLL to build, ship, or register.**
- **Restricted to Automation-compatible types**: `BSTR`, `VARIANT`, `SAFEARRAY`, `LONG`, `DOUBLE`, `DATE`, `CURRENCY`, `VARIANT_BOOL`, `IUnknown*`, `IDispatch*`, and interface pointers to other `oleautomation` interfaces. **No raw `BYTE*` + `size_is` arrays, no arbitrary structs.**
- Slower than standard marshaling.
- `MIDL` will error out if you mark an interface `[oleautomation]` and use a non-conforming type — a useful compile-time check.

### Custom marshaling (`IMarshal`)

The object implements `IMarshal` and decides for itself what crosses the wire. Rare, but powerful: an immutable object can marshal *by value*, so the "proxy" is actually a full local copy and calls never cross a boundary at all.

```cpp
struct IMarshal : IUnknown
{
    HRESULT GetUnmarshalClass(REFIID, void*, DWORD, void*, DWORD, CLSID*);
    HRESULT GetMarshalSizeMax(REFIID, void*, DWORD, void*, DWORD, DWORD*);
    HRESULT MarshalInterface(IStream*, REFIID, void*, DWORD, void*, DWORD);
    HRESULT UnmarshalInterface(IStream*, REFIID, void**);
    HRESULT ReleaseMarshalData(IStream*);
    HRESULT DisconnectObject(DWORD);
};
```

**Security note:** custom marshaling means "the server tells the client which CLSID to instantiate locally." That has been the basis of real privilege-escalation attacks. `CLSCTX_NO_CUSTOM_MARSHAL` and the `EOAC_NO_CUSTOM_MARSHAL` flag to `CoInitializeSecurity` exist to block it. Modern hardened code sets them.

### Which one will you actually meet?

The first two are both common; they simply serve different kinds of interface. The third is rare in code you write, but you have already used one without knowing it.

| Kind | How common | Where it shows up |
|---|---|---|
| **Standard** (MIDL proxy/stub) | Very common | C++ components with rich parameter types — byte buffers, structs, counted arrays. Anything that cannot be expressed in Automation types has no alternative. |
| **Type library** | Very common | Anything `[dual]` or `[oleautomation]`: components meant to be driven from scripting, VBA, Office, or older VB. Preferred when it is possible, because there is no extra DLL to build, ship, register, or keep in sync. |
| **Custom** (`IMarshal`) | Rare | Almost never written by hand. But the **free-threaded marshaler** from Module 3 is a custom marshaler — that is exactly how an agile object says "hand over the raw pointer, no proxy needed." |

**The practical rule when writing an interface:** if it can be Automation-compatible, make it `[dual]` and use typelib marshaling — you get scripting support and cross-apartment marshaling from one registration. Reach for a proxy/stub DLL when the types will not fit, which in this course is `Checksum`'s `BYTE*` buffer and `GetHistory`'s counted array.

**Telling them apart on a customer's machine** is a single registry read, and it decides what you go looking for next:

```powershell
$iid = "{A1B2C3D4-0001-4000-9000-000000000001}"
Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Interface\$iid\ProxyStubClsid32"
```

| Value | Meaning | What to check next |
|---|---|---|
| `{00020424-0000-0000-C000-000000000046}` | Typelib marshaling | Is the **TypeLib** key present, and does the `.tlb` exist at that path, in the right bitness? |
| Any other CLSID | Standard marshaling | Is that CLSID's `InprocServer32` present, and is the PS DLL there in **both** bitnesses if both are used? |
| Key missing entirely | No marshaling support | This is Lab 3.1's `REGDB_E_IIDNOTREG`, and Lab 2.2's failed surrogate |

---

## 4.5 The "works in-proc, fails out-of-proc" bug

You must be able to recognize this instantly.

**Symptom:** `CoCreateInstance` with `CLSCTX_INPROC_SERVER` works. The identical code with `CLSCTX_LOCAL_SERVER` (or from a different apartment) returns `E_NOINTERFACE` (`0x80004002`) — or `REGDB_E_IIDNOTREG` (`0x80040155`).

**Cause:** the interface has no registered marshaling. In-proc, same-apartment, COM hands you the raw pointer and never needs a proxy. The moment a boundary appears, COM looks up `HKCR\Interface\{IID}\ProxyStubClsid32`, finds nothing, and fails at the `QueryInterface` step of the activation.

**Diagnosis, in order:**

```powershell
$iid = "{A1B2C3D4-0001-4000-9000-000000000001}"
Get-ChildItem "Registry::HKEY_CLASSES_ROOT\Interface\$iid" -Recurse -EA SilentlyContinue
Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Interface\$iid\ProxyStubClsid32" -EA SilentlyContinue
```

1. Key missing entirely → marshaling was never registered.
2. `ProxyStubClsid32` present → follow it to `HKCR\CLSID\{ps}\InprocServer32` and check the DLL **exists** and is the **right bitness**.
3. `ProxyStubClsid32 = {00020424-...}` → typelib marshaling; verify `HKCR\TypeLib\{LIBID}\<ver>\0\win64` (or `win32`) points at a file that exists.
4. Everything present but still failing → bitness mismatch on the PS DLL, or the typelib is for a different version.

This is also the reason **Lab 2.2's surrogate experiment failed**: `dllhost.exe` puts the object in another process, so every call has to be marshaled, and `ICalculator` had no marshaling to offer. You will finish that lab in **[Lab 4.2](#47-lab-42--break-and-fix-marshaling)**, once you have built the proxy/stub in Lab 4.1 — and note that it needs registering in **both bitnesses**, since that experiment deliberately pairs a 32-bit DLL with a 64-bit client.

---

## 4.6 LAB 4.1 — Author IDL, compile with MIDL, read the output

> **Requirements**
> - **Tools:** a **Developer PowerShell for VS** or *x64 Native Tools Command Prompt* — this is what puts `midl.exe` on `PATH`. A plain PowerShell window will fail with "midl is not recognized." Plus **Tools → Create GUID**, and **OleView.NET** for step 4.
> - **Elevation:** not required — this lab only compiles.
> - **Bitness:** `/env x64`, matching your build.
> - **Depends on:** Lab 2.1 — you will reuse the **same GUIDs** you generated there.
> - **What you write:** `Calculator.idl`, by hand, from the listing in §4.2. Type it out; do not copy the finished file.
> - **Reference copy:** [`labs/stage-3-idl-marshaling/Calculator.idl`](../labs/stage-3-idl-marshaling/Calculator.idl) is the completed version — use it to check your work if MIDL rejects something, not as a starting point. The same folder has `CalcPS.vcxproj`, which automates step 5 for later labs.
> - **Time:** ~90 min — most of it spent reading the generated files, which is the point.

Until now `ICalculator` has existed only as a C++ header — readable by one compiler, and by nothing else. This lab describes the same interface in **IDL** and runs MIDL over it.

The build takes a minute; the lab is what comes out. Those four generated files are what COM actually uses to carry a call across a boundary, and reading them is the difference between marshaling being magic and marshaling being obvious.

### Step 1 — write it

In a new folder, type the IDL from §4.2 into a file called `Calculator.idl`.

**The GUIDs matter, and most of them are not new ones.** A proxy/stub is registered against a specific IID, so the interface described here must carry the *same* IID as the interface your DLL already implements — otherwise you will register marshaling for an interface nobody uses, and the failure will look exactly like having no marshaling at all.

| In the IDL | Which GUID to use |
|---|---|
| `interface ICalculator` → `uuid(...)` | **Your `IID_ICalculator`** from Lab 2.1 — copy it out of `Calculator.h` |
| `coclass Calculator` → `uuid(...)` | **Your `CLSID_Calculator`** from Lab 2.1 |
| `library TrainingCalcLib` → `uuid(...)` | A **new** GUID — the type library has its own identity (its LIBID), and you have not generated one before |

Note the format difference: IDL wants the GUID bare, `uuid(A1B2C3D4-0001-...)`, with no braces and no `0x` — not the `DEFINE_GUID` form in your header.

### Step 2 — compile

```powershell
# From a Developer PowerShell for VS
midl /nologo /W1 /char signed /env x64 `
     /h Calculator.h `
     /iid Calculator_i.c `
     /proxy Calculator_p.c `
     /dlldata dlldata.c `
     /tlb Calculator.tlb `
     Calculator.idl
```

In Visual Studio, adding the `.idl` to the project does this automatically (Project → MIDL properties control the output names).

### Step 3 — read the generated code

This is the point of the lab. Open each file:

**`Calculator.h`** — your interface, in three forms: C++ (`struct ICalculator : public IUnknown`), C (a manual vtable struct plus `#define`d helper macros), and `CINTERFACE` macros. Note `MIDL_INTERFACE("A1B2...")` expands to `struct __declspec(uuid(...)) __declspec(novtable)` — which is what makes `__uuidof(ICalculator)` work.

**`Calculator_i.c`** — nothing but GUID definitions:
```c
const IID IID_ICalculator = {0xA1B2C3D4,0x0001,0x4000,{0x90,0x00,...}};
```
Exactly one translation unit in your program should compile this, or you get duplicate symbols. (The alternative is `#include <initguid.h>` before your header.)

**`Calculator_p.c`** — the interesting one. Find:

```c
static const unsigned short ICalculator_FormatStringOffsetTable[] = { 0, 38, 76, ... };
static const MIDL_STUB_DESC Object_StubDesc = { ... };
static const unsigned char __MIDL_ProcFormatString[] = { ... };  /* NDR bytecode */
```

The **format strings** are a compact bytecode describing every parameter of every method. `NdrClientCall3`/`NdrStubCall3` interpret it at runtime. That's why proxy/stub DLLs are small even for big interfaces — the marshaling is table-driven, not code-generated per method.

Also find:

```c
CINTERFACE_PROXY_VTABLE(8) _ICalculatorProxyVtbl =
{
    &ICalculator_ProxyInfo,
    &IID_ICalculator,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy,
    (void*) (INT_PTR) -1 /* ICalculator::Add */,
    ...
};
```

The `-1` entries mean "use the generic NDR forwarder with this method's format string offset." **Count the entries: 3 for `IUnknown` + your methods.** That number must equal the `NumMethods` value in the registry.

**`dlldata.c`** — provides `DllGetClassObject`, `DllCanUnloadNow`, `DllRegisterServer`, `DllUnregisterServer` for the proxy/stub DLL itself. A proxy/stub DLL is just an ordinary COM in-proc server whose objects happen to be proxies.

### Step 4 — read the type library, and see what got left out

Four of MIDL's outputs are for the C++ toolchain. `Calculator.tlb` is the one **everybody else** reads — scripting engines, `#import`, `tlbimp`, VBA's object browser (Modules 5 and 6). It is worth ten minutes now.

Open it in **OleView.NET** (File → Open Type Library), which decompiles it back into IDL-like text. Put that next to the `Calculator.idl` you wrote and compare:

1. **Find what survived.** The `library TrainingCalcLib` block, the `coclass Calculator`, and `ICalculator` with its methods.
2. **Find what did not.** Anything outside the `library { }` block never reaches the typelib — a typelib is a *subset* of your IDL, not a translation of it.
3. **Look at how the parameter types are described.** `[out, retval] LONG*` becomes a return value, which is why PowerShell can write `$calc.Add(2,3)` and get a number rather than an `HRESULT`. That single detail is most of Module 5.
4. **Try `Checksum`.** Its `[in, size_is(cb)] const BYTE*` has no Automation equivalent. This is the concrete reason `[oleautomation]` exists, and you will hit it as a hard error in Lab 4.2.

> **This is a support skill, not just a lab step.** A registered type library lets you read a third-party component's entire interface surface with **no source and no documentation** — which is often exactly the position you are in on a ticket.

### Step 5 — build and register the proxy/stub DLL

Create a DLL project with `Calculator_p.c`, `dlldata.c`, `Calculator_i.c`, and:

```
# CalcPS.def
LIBRARY CalcPS
EXPORTS
    DllGetClassObject   PRIVATE
    DllCanUnloadNow     PRIVATE
    DllRegisterServer   PRIVATE
    DllUnregisterServer PRIVATE
    GetProxyDllInfo     PRIVATE
```

Compile with `REGISTER_PROXY_DLL` defined:

```
Preprocessor Definitions: WIN32;REGISTER_PROXY_DLL;_WINDLL
Additional Dependencies:  rpcrt4.lib
```

Then:

```powershell
regsvr32 CalcPS.dll
```

Verify:

```powershell
Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Interface\{A1B2C3D4-0001-4000-9000-000000000001}\ProxyStubClsid32"
```

---

## 4.7 LAB 4.2 — Break and fix marshaling

> **Requirements**
> - **Tools:** Developer PowerShell (for `midl`), Visual Studio C++, `regsvr32`, Process Explorer.
> - **Elevation:** required — registering `CalcPS.dll` and writing the AppID keys.
> - **Bitness:** build and register the proxy/stub DLL for **both x86 and x64**; the surrogate step deliberately pairs a 32-bit DLL with a 64-bit client, and each side loads its own proxy.
> - **Depends on:** Lab 4.1 (proxy/stub built) and Lab 2.2 (the surrogate registration you left failing).
> - **Starting point:** [`labs/stage-3-idl-marshaling/`](../labs/stage-3-idl-marshaling/) built for **both** bitnesses and registered, plus [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/) x86 registered.
> - **Time:** ~2 h.

Marshaling is invisible while it works. This lab switches it on and off underneath a client that never changes, so you can see precisely **which call fails and when** — activation or first method — and how that differs between a proxy/stub DLL and typelib marshaling.

It also closes Lab 2.2: the surrogate that could not work then works now, and you will know exactly which registration made the difference.

1. With the PS DLL registered, redo **[Lab 2.2](02-activation-and-registry.md#27-lab-22--bitness) steps 6–8** — *Fix option B*, the DLL surrogate: the AppID with its empty `DllSurrogate` value, the client switched to `CLSCTX_LOCAL_SERVER`, and the run. It now works, where before it stopped at `E_NOINTERFACE`: your 32-bit DLL runs inside `dllhost.exe` and a 64-bit client talks to it. Confirm with Process Explorer that `dllhost.exe` has loaded your DLL.

   Nothing about the AppID registration changed — the only difference is that `ICalculator` can now be marshaled, so COM is finally able to build the proxy that a separate process requires.

2. `regsvr32 /u CalcPS.dll`. Re-run. Record the exact HRESULT and where it fails (activation vs. first call).

3. **Switch to typelib marshaling.** Same interface, same client, a completely different marshaler — and no `CalcPS.dll` at all. Yes, this means editing the IDL and re-running MIDL; work through these in order.

   **3a. Edit `Calculator.idl`.** Add `oleautomation` to the interface attributes, and delete the two methods whose types Automation cannot express — `Checksum` (a `BYTE*` buffer) and `GetHistory` (a `LONG**` array). **Keep the same `uuid`:** you are re-describing the existing interface, not inventing a new one.

   ```idl
   [
       object,
       oleautomation,                       // <- added
       uuid(A1B2C3D4-0001-...),             // <- unchanged, your IID from Lab 2.1
       pointer_default(unique)
   ]
   interface ICalculator : IUnknown
   {
       HRESULT Add([in] LONG a, [in] LONG b, [out, retval] LONG* result);
       HRESULT Describe([out, retval] BSTR* description);
   }
   ```

   **3b. Re-run MIDL.** Exactly the command from Lab 4.1 step 2. This regenerates `Calculator.h`, `Calculator_i.c`, `Calculator_p.c` and — the one that matters now — `Calculator.tlb`.

   **3c. Rebuild `Calc.dll` against the new header.** You removed two methods, so the vtable changed; your C++ class must match it or the object will not compile. Delete the corresponding implementations.

   **3d. Embed the type library in the DLL.** Add a `.rc` file to the `Calc` project containing one line, so the `.tlb` travels inside the DLL rather than as a loose file:

   ```
   1 typelib "Calculator.tlb"
   ```

   **3e. Register the type library from `DllRegisterServer`.** `LoadTypeLibEx` with `REGKIND_REGISTER` writes the `HKCR\TypeLib` keys *and* points the interface at the universal marshaler:

   ```cpp
   // In DllRegisterServer, after your existing CLSID registration:
   ITypeLib* pTypeLib = nullptr;
   HRESULT hr = LoadTypeLibEx(modulePath, REGKIND_REGISTER, &pTypeLib);   // modulePath = this DLL
   if (pTypeLib) pTypeLib->Release();
   ```

   **3f. Remove the competition.** Unregister the proxy/stub so it cannot be responsible for any success you see:

   ```powershell
   regsvr32 /u CalcPS.dll        # elevated
   regsvr32 Calc.dll             # re-register, now with the typelib
   ```

   **3g. Verify the switch happened.** The interface should now point at the universal marshaler rather than at your PS DLL:

   ```powershell
   $iid = "{A1B2C3D4-0001-4000-9000-000000000001}"   # <- your IID
   Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Interface\$iid\ProxyStubClsid32"
   Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Interface\$iid\TypeLib"
   ```

   Expect `ProxyStubClsid32` = `{00020424-0000-0000-C000-000000000046}` and a `TypeLib` value naming your LIBID.

   **3h. Re-run the surrogate test from step 1.** It still works — cross-process calls, **with no proxy/stub DLL anywhere on the machine**. That is the whole point of this step: two entirely different mechanisms, one unchanged client.

4. **Prove the restriction.** Put `Checksum([in, size_is(cb)] const BYTE* data, ...)` back while keeping `oleautomation`. MIDL fails with something like:

```
error MIDL2311: this type is not supported by automation: [ Parameter 'data' of Procedure 'Checksum' ]
```

   **Why it fails.** Typelib marshaling has no generated code — `oleaut32.dll`'s universal marshaler reads the type library at runtime and marshals each parameter according to the **`VARTYPE`** recorded there. So every parameter must be describable as one of the Automation types: `VT_I4`, `VT_BSTR`, `VT_VARIANT`, `VT_ARRAY | ...`, an interface pointer, and the rest of the short list in §4.4.

   `[in, size_is(cb)] const BYTE* data` is not on that list, and cannot be. It says *"a pointer to `cb` bytes"* — a length carried in a **separate parameter**. A type library describes each parameter independently; there is nowhere to record "this one's length lives in that one." Standard marshaling handles it easily, because MIDL bakes that relationship into the NDR format string it generates. The universal marshaler has no such string to read.

   Adding `oleautomation` is you promising the interface stays inside that type set, so MIDL checks the promise and refuses at compile time rather than letting you ship an interface that fails to marshal at runtime. `GetHistory`'s `[out, size_is(, *count)] LONG**` is rejected for exactly the same reason.

   **The Automation-compatible way to send a byte buffer** is a `SAFEARRAY` of `VT_UI1`, which carries its own length and so needs no companion parameter:

```idl
HRESULT Checksum([in] SAFEARRAY(BYTE) data, [out, retval] ULONG* checksum);
```

Write that error into your notes — you *will* see a developer hit it.

   Then take `oleautomation` back out, restore `Checksum` and `GetHistory`, re-run MIDL and rebuild — step 5 needs the full interface and the proxy/stub route again.

5. **Watch `ThreadingModel` decide whether you get a proxy.** Now that marshaling works, one small experiment settles §3.3's table for good. Re-register `CalcPS.dll`, then add this helper to your client — it identifies a proxy without a debugger, by asking which module the vtable's code belongs to:

```cpp
// Your DLL -> you hold the real object. CalcPS/combase -> you hold a proxy.
void WhoOwnsVtable(const char* label, void* pItf)
{
    void** vtbl = *reinterpret_cast<void***>(pItf);
    HMODULE mod = nullptr;
    GetModuleHandleExA(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                       GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                       static_cast<LPCSTR>(vtbl[0]), &mod);
    char path[MAX_PATH] = "??";
    if (mod) GetModuleFileNameA(mod, path, MAX_PATH);
    const char* leaf = strrchr(path, '\\');
    printf("%s vtable code lives in: %s\n", label, leaf ? leaf + 1 : path);
}
```

Call `CoCreateInstance` and then `WhoOwnsVtable` from an **STA** thread and from an **MTA** thread. Then change `ThreadingModel` in `DllRegisterServer` — `Apartment`, `Free`, `Both`, `Neutral` — re-registering between runs, and fill this in:

| `ThreadingModel` | STA client gets | MTA client gets |
|---|---|---|
| `Apartment` | | |
| `Free` | | |
| `Both` | | |
| `Neutral` | | |

With `Apartment` and an MTA client you should see `CalcPS.dll` — COM created the object in a host STA and handed you a proxy. That same case is exactly what fails with `E_NOINTERFACE` when no proxy/stub is registered, which is one of the most common activation tickets there is.

**The cost, measured.** In any cell where you got a proxy, time the difference:

```cpp
LARGE_INTEGER f, t0, t1; QueryPerformanceFrequency(&f);
QueryPerformanceCounter(&t0);
for (int i = 0; i < 100000; ++i) { long r; sp->Add(i, 1, &r); }
QueryPerformanceCounter(&t1);
printf("  %.2f us/call\n", 1e6 * double(t1.QuadPart - t0.QuadPart) / f.QuadPart / 100000);
```

Expect single-digit **nanoseconds** for a direct call against **tens of microseconds** across an apartment boundary — a factor of roughly 1000. **That number is your argument** the next time someone asks why a component is slow inside a service.

---

## 4.8 LAB 4.3 — Memory ownership, verified

> **Requirements**
> - **Tools:** Visual Studio C++ with **ATL** (`atlbase.h`); **Application Verifier** — already on Windows, run `appverif.exe`, no SDK component required — and **WinDbg** for step 3.
> - **Elevation:** only for Application Verifier, at the end. The programming half needs none.
> - **Bitness:** x64.
> - **Depends on:** §4.3's rules. **No registration, no MIDL, no DLL** — see the note below.
> - **Starting point:** a new, empty C++ console project.
> - **Time:** ~60 min.

The ownership rules in §4.3 are about **which allocator frees a block**. That is fixed by the interface contract — `BSTR` means `SysFreeString`, `[out]` array means `CoTaskMemFree` — and none of it depends on how the object was *created*. So this lab skips the activation machinery entirely: no IDL, no MIDL, no proxy/stub DLL, no class factory, no registry entry, no `CoCreateInstance`. Just `new Calculator()` and an interface pointer.

What it does still use is the part that matters here: COM's allocators. `SysAllocString`, `CoTaskMemAlloc` and their matching frees behave identically whether the caller is across a proxy or one stack frame away, which is why testing them in-process proves the same thing.

You write **both sides** — the method that allocates and the caller that frees — because that is the only way to feel where the boundary is. Then you break each rule deliberately and watch the tooling catch it.

### Step 1 — the whole program

Paste this into the project's `.cpp` file. It compiles as-is.

```cpp
#include <windows.h>
#include <objbase.h>
#include <atlbase.h>
#include <cassert>
#include <cstdio>
#include <vector>

// The same interface as Calculator.idl, declared inline so this lab needs no build plumbing.
struct __declspec(uuid("A1B2C3D4-0001-4000-9000-000000000001")) ICalculator : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Add(LONG a, LONG b, LONG* result) = 0;
    virtual HRESULT STDMETHODCALLTYPE Describe(BSTR* description) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetHistory(ULONG* count, LONG** values) = 0;
    virtual HRESULT STDMETHODCALLTYPE Checksum(const BYTE* data, ULONG cb, ULONG* checksum) = 0;
    virtual HRESULT STDMETHODCALLTYPE put_Precision(LONG value) = 0;
};

class Calculator : public ICalculator
{
    LONG              m_cRef = 1;
    std::vector<LONG> m_history;
    LONG              m_precision = 2;

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == __uuidof(ICalculator))
            *ppv = static_cast<ICalculator*>(this);
        else return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&m_cRef); }
    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG n = InterlockedDecrement(&m_cRef);
        if (n == 0) delete this;
        return n;
    }

    HRESULT STDMETHODCALLTYPE Add(LONG a, LONG b, LONG* result) override
    {
        if (!result) return E_POINTER;
        *result = a + b;
        m_history.push_back(*result);
        return S_OK;
    }

    // RULE 1: callee allocates a BSTR, CALLER frees it with SysFreeString.
    HRESULT STDMETHODCALLTYPE Describe(BSTR* description) override
    {
        if (!description) return E_POINTER;
        *description = nullptr;                  // null first, before anything can fail
        if (m_precision < 0) return E_FAIL;      // failure path, used by rule 4
        *description = SysAllocString(L"Training Calculator v1");
        return *description ? S_OK : E_OUTOFMEMORY;
    }

    // RULE 2: callee allocates with CoTaskMemAlloc, CALLER frees with CoTaskMemFree.
    HRESULT STDMETHODCALLTYPE GetHistory(ULONG* count, LONG** values) override
    {
        if (!count || !values) return E_POINTER;
        *count = 0; *values = nullptr;           // null EVERYTHING first
        const ULONG n = static_cast<ULONG>(m_history.size());
        if (n == 0) return S_OK;                 // zero items is not an error
        auto* p = static_cast<LONG*>(CoTaskMemAlloc(n * sizeof(LONG)));
        if (!p) return E_OUTOFMEMORY;
        memcpy(p, m_history.data(), n * sizeof(LONG));
        *values = p; *count = n;
        return S_OK;
    }

    // RULE 3: an [in] buffer belongs to the CALLER. Read it; never free it, never keep it.
    HRESULT STDMETHODCALLTYPE Checksum(const BYTE* data, ULONG cb, ULONG* checksum) override
    {
        if (!checksum) return E_POINTER;
        *checksum = 0;
        if (cb && !data) return E_POINTER;
        ULONG sum = 0;
        for (ULONG i = 0; i < cb; ++i) sum += data[i];
        *checksum = sum;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE put_Precision(LONG value) override
    {
        m_precision = value;
        return S_OK;
    }
};

void ExerciseMemoryRules(ICalculator* p)
{
    // RULE 1 - BSTR
    BSTR desc = nullptr;
    if (SUCCEEDED(p->Describe(&desc)))
    {
        wprintf(L"desc = %s (len %u)\n", desc, SysStringLen(desc));
        SysFreeString(desc);
    }

    // RULE 2 - [out] array
    LONG r = 0;
    p->Add(40, 2, &r);
    ULONG count = 0; LONG* values = nullptr;
    if (SUCCEEDED(p->GetHistory(&count, &values)))
    {
        for (ULONG i = 0; i < count; ++i) wprintf(L"history[%u] = %ld\n", i, values[i]);
        CoTaskMemFree(values);
    }

    // RULE 3 - [in] buffer, owned by us start to finish
    BYTE data[] = { 1, 2, 3, 4 };
    ULONG sum = 0;
    p->Checksum(data, ARRAYSIZE(data), &sum);
    wprintf(L"checksum = %lu\n", sum);

    // RULE 4 - an [out] parameter must be nulled even when the call fails
    p->put_Precision(-1);
    BSTR shouldBeNull = reinterpret_cast<BSTR>(static_cast<UINT_PTR>(0xDEADBEEF));
    HRESULT hr = p->Describe(&shouldBeNull);
    assert(FAILED(hr) && shouldBeNull == nullptr);
    p->put_Precision(2);
    wprintf(L"failure path left the [out] parameter null - correct\n");
}

int main()
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    {
        CComPtr<ICalculator> sp;
        sp.Attach(new Calculator());
        ExerciseMemoryRules(sp);
    }
    CoUninitialize();
    return 0;
}
```

Run it. The correct output is:

```
desc = Training Calculator v1 (len 22)
history[0] = 42
checksum = 10
failure path left the [out] parameter null - correct
```

Read `ExerciseMemoryRules` against the `RULE` comments in the class above it. Each pair is one rule from §4.3, written from both ends — the side that allocates and the side that frees.

### Step 2 — break each rule, one at a time

Make one change, run, then put it back. Run each one in **Debug** and in **Release**, with **Ctrl+F5** (Start Without Debugging) so the console stays open and Visual Studio prints the exit code.

| # | Change | Which rule it breaks |
|---|---|---|
| 1 | In `ExerciseMemoryRules`, replace `SysFreeString(desc)` with `CoTaskMemFree(desc)` | `BSTR` must be freed with `SysFreeString` |
| 2 | In `ExerciseMemoryRules`, replace `CoTaskMemFree(values)` with `delete[] values` | An `[out]` array must be freed with `CoTaskMemFree` |
| 3 | In `Checksum`, add `CoTaskMemFree((void*)data)` before returning | An `[in]` buffer belongs to the caller; the callee must not free it |
| 4 | In `Describe`, delete the `*description = nullptr;` line | An `[out]` parameter must be nulled even on failure |
| 5 | In `GetHistory`, add `p[n] = 0;` immediately after the `memcpy` | An `[out]` array must not be written past the size that was allocated for it |

**Do not expect a particular result. Write down what you actually get.** Every one of these is undefined behaviour, and undefined behaviour has no fixed symptom. Across the machines this lab was written on, the same five changes produced *all* of the following:

| What you may see | What it means |
|---|---|
| The program runs and prints correct output | The damage was done and nobody noticed. **This is the most dangerous result**, and the most common in Release. |
| Exit code `0xC0000374` (`STATUS_HEAP_CORRUPTION`) | The heap manager read a block header, found nonsense, and killed the process on the spot. |
| Exit code `0xC0000005` (access violation) | A wild pointer was dereferenced. |
| What looks like a **hang** | Almost always a modal dialog waiting behind the console — the debug CRT's *invalid heap pointer* check for #2, or your own `assert` for #4. Alt-Tab to it. |

Whether you see a crash depends on the build configuration, the C runtime version, which heap the process happens to be using, and what else was allocated first. Break 1 crashed reliably on one machine and completed cleanly on another, with the same compiler flags.

That variability is not a nuisance to work around — it *is* the lesson, and it is worth stating in one sentence you can reuse on a ticket:

> **A mismatched free is not reliably fatal. It is reliably a bug.**

The absence of a crash tells you nothing about whether the code is correct, which is why "we've shipped it for years and it's fine" is not evidence. It also means you cannot find these by running the program and watching. You need a tool that checks every free as it happens — which is step 3.

**Break #2 deserves a special mention: it may be undetectable.** `CoTaskMemAlloc` and the CRT's `new[]` both end up calling `HeapAlloc` on the **same** heap — you can prove it with `GetProcessHeaps` and `HeapValidate`. So at the Win32 level, freeing a `CoTaskMemAlloc` block with `delete[]` is a perfectly matched allocate/free pair on one heap, and nothing in the operating system has grounds to object. No crash, no corruption, and — as you will see in step 3 — not even page heap will flag it.

It is still a bug. The contract says `CoTaskMemFree`, and the moment that object lives in another process the memory is freed by a different allocator entirely, at which point the mismatch becomes real. **A rule you cannot test for is a rule you have to follow deliberately**, and that is the most uncomfortable idea in this module.

**Two things you can predict**, because they do not depend on the heap:

- **#4 in Debug** trips the `assert`, because a failed `Describe` leaves `0xDEADBEEF` in the caller's variable.
- **#4 in Release** does not, because `assert` is compiled out — leaving a caller holding a garbage pointer it believes it must free. That is how this bug reaches a customer.

### Step 3 — troubleshoot break #5 with page heap

Break #5 is the one to chase, for three reasons. It is a **server-side** bug, so it is in someone else's code as far as the caller is concerned. It is an ordinary off-by-one, the kind that survives code review. And it produces **no symptom whatsoever** — which is exactly the ticket you get handed.

Page heap changes that. It gives every allocation its own page and leaves the **next page unmapped**, so a write that runs off the end has a wall to hit instead of harmless slack — and it checks a fill pattern around each block when the block is freed. One of those two catches the overrun, and §3c explains which and why.

Everything here uses the **Release** configuration — that is what customers run, and Debug's C runtime does its own checking, which would muddy the question of what is catching the bug.

**3a. Reproduce the "nothing is wrong" report.**

Set the dropdown to **Release | x64**, and make sure break #5 is in place:

```cpp
    memcpy(p, m_history.data(), n * sizeof(LONG));
    p[n] = 0;                    // BREAK 5: one LONG past the end of the allocation
    *values = p;
    *count  = n;
```

Run with **Ctrl+F5**. It prints the correct history and exits cleanly, with code 0. Run it a few more times — still fine. **This is the state you are in when a customer says it crashes twice a week and you cannot reproduce it.** The stray write landed in padding the allocator happened to leave after the block; a different allocation size on a different day, and it lands on live data instead.

**3b. Turn on page heap.** Launch `appverif.exe` **as administrator** (it ships with Windows; use the `System32` copy for an x64 build). Then:

- **File → Add Application** → browse to your **Release** build's `.exe`, e.g. `...\x64\Release\YourApp.exe`
- In the **Tests** pane tick **Basics → Heaps**, and untick everything else
- **File → Save**, then close Application Verifier

> Application Verifier stores this under *Image File Execution Options* keyed on the **image name**, not the full path. So it now applies to any executable called `YourApp.exe` on this machine — including your Debug build. Convenient here; a genuine surprise months later when a forgotten entry is still slowing something down.

**3c. Run it under WinDbg.** From here on, use **WinDbg** rather than the Visual Studio debugger — the commands you need are WinDbg extensions and Visual Studio cannot run them. Get it from the Microsoft Store, from *Debugging Tools for Windows* in the SDK, or with `winget install Microsoft.WinDbg`.

**File → Launch Executable**, pick your **Release** `.exe`. WinDbg stops immediately with:

```
Break instruction exception - code 80000003 (first chance)
```

**That first one is not your bug.** WinDbg always halts at a loader breakpoint before the target's own code has run — note the stack is in `ntdll`. Let the program continue:

```
0:000> g
```

Now watch for a *second* `80000003`. That one is the verifier, and you can tell them apart because it arrives with a `VERIFIER STOP` block printed alongside it. Both events report the same exception code, which is why the report text — not the code — is what identifies it.

A debugger is required for any of this: a verifier stop is a breakpoint exception, so with no debugger attached the process simply exits with `0x80000003` and tells you nothing.

Where it stops depends on **how far past the end you wrote**, and that is worth understanding rather than shrugging at.

Ticking **Basics → Heaps** enables *full* page heap — Application Verifier records `PageHeapFlags = 0x3` for your image, meaning full page heap plus stack-trace collection. Full page heap leaves the page after each block unmapped, but the block must still satisfy the allocator's alignment, so there is normally **padding** between the end of your data and that guard page. Page heap fills the padding with a known pattern and checks it when the block is freed.

| How far the write lands | When you find out |
|---|---|
| Inside the alignment padding | **At `CoTaskMemFree`.** The write touches a mapped page, so nothing faults; the verifier notices the disturbed pattern at free time and reports a *corrupted suffix pattern*. |
| Past the padding, onto the guard page | **At the write itself**, as an access violation on that line. |

Break #5 gives you the first case, and that is the normal outcome for a realistic off-by-one: the padding on x64 can be a dozen bytes or more, so overrunning a small array by one or two elements stays inside it. Do not chase the second case by nudging the index — how far you must go depends on the allocation size and the alignment page heap chose. If you want to see a guard-page fault, overrun by an amount that obviously clears the padding, such as `memset(p, 0, (n + 32) * sizeof(LONG));`.

**Stopping at the free is just as useful as stopping at the write.** Either way the verifier hands you the block, its size, and the exact offset of the damage:

```
=======================================
VERIFIER STOP 000000000000000F: pid 0x55B8: Corrupted suffix pattern for heap block.

    000001956B381000 : Heap handle used in the call.
    000001956FBD9FF0 : Heap block involved in the operation.
    0000000000000004 : Size of the heap block.
    000001956FBD9FF4 : Corruption address.
```

Read it as a sentence, because the numbers corroborate each other:

| Line | What it tells you |
|---|---|
| *Corrupted suffix pattern* | The fill pattern immediately **after** the block was disturbed. The damage is an **overrun** — not a bad pointer, not a double free. |
| Heap block `...D9FF0`, size `4` | The allocation was **4 bytes**: one `LONG`. So `m_history` held a single entry and `CoTaskMemAlloc(n * sizeof(LONG))` asked for exactly four bytes. |
| Corruption address `...D9FF4` | Block start **+ 4** — the first byte past the end. That is `p[1]` on a one-element array: exactly the `p[n] = 0;` you wrote. |

Those addresses also explain why nothing faulted at the write. The block sits at `...D9FF0` and ends at `...D9FF4`; the page ends at `...DA000`. Everything from `...D9FF4` to `...D9FFF` is alignment padding on a **mapped** page, filled with the verifier's pattern. The write landed inside it — harmless to the hardware, invisible until the block was freed and the pattern checked. To fault on the write itself it would have to reach `...DA000`, the guard page, twelve bytes further on.

The **size** tells you how many elements were expected; the **offset** tells you how far past the end the write went. Together they identify an off-by-one without reading a line of source — and with no repro steps and no luck.

Now ask where the block came from, using the address from that report:

```
0:000> !ext.heap -p -a 000001956FBD9FF0
    address 000001956fbd9ff0 found in
    _DPH_HEAP_ROOT @ 1956b381000
    in busy allocation (  DPH_HEAP_BLOCK:         UserAddr         UserSize -         VirtAddr         VirtSize)
                             1956b3abea0:      1956fbd9ff0                4 -      1956fbd9000             2000
    ...
    00007ff8bab9e71c ntdll!RtlAllocateHeap+0xabc
    00007ff881b6d411 vrfcore!VfCoreRtlAllocateHeap+0x21
    00007ff67bdf1481 COMmemOwnership!Calculator::GetHistory+0x71 [...\COMmemOwnership.cpp @ 68]
    00007ff67bdf16d7 COMmemOwnership!main+0xd7                  [...\COMmemOwnership.cpp @ 140]
    00007ff67bdf1fbf COMmemOwnership!__scrt_common_main_seh+0x10f
    00007ff8ba3accb7 KERNEL32!BaseThreadInitThunk+0x17
```

This is the whole point of the exercise. Three things to read out of it:

- **`Calculator::GetHistory ... @ 68`** — the line that allocated the block, with file and line number. You did not have to guess which allocation the corrupt block belonged to; page heap recorded the stack at allocation time and kept it.
- **`UserSize 4 - VirtSize 2000`** — you asked for 4 bytes and page heap gave the block its own **0x2000** of address space, so it could sit against a guard page. That is the cost of the technique, and why you turn it off afterwards.
- **`vrfcore!VfCoreRtlAllocateHeap`** — the verifier's hook, sitting between your `CoTaskMemAlloc` and the real `RtlAllocateHeap`. That interception is what makes all of this possible.

Put the two reports together and the bug is fully described without opening the source: *a 4-byte block allocated at `GetHistory` line 68 was written past its end.*

**3d. Turn it off when you finish.** Back in Application Verifier, select the image, choose **Delete**, and save. Page heap left enabled slows that executable down permanently, and because the setting is keyed on the image name it is easy to forget it is there.

> **Why this matters more than the code.** In a real ticket you are handed a crash *inside someone else's component*, with a stack that has nothing to do with the mistake. The mistake was a mismatched free that happened minutes earlier. Page heap is how you move the fault back to the culprit, and knowing that these five rules exist is how you guess what to look for.

## 4.9 Versioning: never modify a published interface

Module 0 explained why. Here's the practice.

**Wrong:**

```idl
interface ICalculator : IUnknown
{
    HRESULT Add([in] LONG a, [in] LONG b, [out, retval] LONG* r);
    HRESULT Multiply([in] LONG a, [in] LONG b, [out, retval] LONG* r);  // ADDED in v2
}
```

Every already-compiled client that calls `Subtract` at slot 4 now calls something else. Also `NumMethods` in the registry changes, and the old proxy/stub is silently wrong.

**Right:**

```idl
[object, uuid(...OLD...), pointer_default(unique)]
interface ICalculator : IUnknown
{
    HRESULT Add([in] LONG a, [in] LONG b, [out, retval] LONG* r);
    HRESULT Subtract([in] LONG a, [in] LONG b, [out, retval] LONG* r);
}

[object, uuid(...NEW...), pointer_default(unique)]
interface ICalculator2 : ICalculator          // extends, doesn't modify
{
    HRESULT Multiply([in] LONG a, [in] LONG b, [out, retval] LONG* r);
}
```

Client:

```cpp
CComPtr<ICalculator2> sp2;
if (SUCCEEDED(spCalc->QueryInterface(IID_ICalculator2, (void**)&sp2)))
    sp2->Multiply(6, 7, &r);          // new server
else
    /* fall back to the v1 path */;
```

**This `QI`-and-fall-back pattern is how all of Windows does feature detection.** `IShellFolder2`, `IFileDialog2`, `IPropertyStore2`, `IClassFactory2` — every one of these exists because somebody needed a new method.

### Two more rules

- **Never change a GUID's meaning.** If you change an interface during development, change the IID too. Shipping two different vtables under one IID is the worst COM bug there is: it produces stack corruption on call, with no diagnostic.
- **Do not derive from a `dual` interface** across versions if you want scripting to work cleanly; prefer a separate `dual` interface. (Module 5 explains why `IDispatch` inheritance is awkward.)

---

## 4.10 Reading NDR failures

| HRESULT | Symbol | Cause |
|---|---|---|
| `0x800706F7` | `RPC_X_BAD_STUB_DATA` | The wire data didn't match the format string — usually a **mismatched IDL between client and server**, or a rebuilt interface with an unchanged IID |
| `0x800706F4` | `RPC_X_NULL_REF_POINTER` | A `[ref]` pointer was null |
| `0x800706C6` | `RPC_X_BYTE_COUNT_TOO_SMALL` | `size_is`/`length_is` inconsistent with the actual buffer |
| `0x80070057` | `E_INVALIDARG` | Often a `size_is` mismatch surfacing generically |
| `0x8007000E` | `E_OUTOFMEMORY` | Or a wildly wrong `size_is` causing a huge allocation |
| `0x80010105` | `RPC_E_SERVERFAULT` | The server threw an exception inside the stub |
| `0x800706BE` | `RPC_S_CALL_FAILED` | Generic transport failure; check whether the server died |

> **`RPC_X_BAD_STUB_DATA` is almost always "someone changed the interface without changing the IID."** Ask for both sides' build numbers before anything else.

---

## 4.11 Checkpoint

1. Why can't a C++ header alone describe a COM interface well enough to marshal it? Give three specific ambiguities.
2. What does `[out, size_is(, *count)] LONG** values` mean, and why is the leading comma there?
3. A method returns `E_FAIL`. What must be true of every `[out]` parameter, and why does the marshaler care?
4. You free a `BSTR` with `CoTaskMemFree`. Describe precisely what goes wrong at the memory level.
5. A component works in-proc and returns `0x80004002` out-of-proc. Give the three-step diagnosis.
6. When would you choose typelib marshaling over a MIDL proxy/stub, and what do you give up?
7. A customer reports `RPC_X_BAD_STUB_DATA` after a partial upgrade. What is your first hypothesis and your first question to them?
8. Why does `CLSCTX_NO_CUSTOM_MARSHAL` exist?

<details>
<summary>Answers</summary>

1. (a) **Direction** — `int* p` doesn't say whether the callee reads or writes it. (b) **Cardinality** — is it one `int` or an array, and if an array, how long? (c) **Ownership** — who allocated it and who frees it. (Also: null-ability, aliasing, and whether a `void**` holds an interface pointer that needs `AddRef`.)

2. `values` is a pointer to a pointer. Each comma-separated slot in `size_is` corresponds to one level of indirection. The empty first slot means "the outer pointer is a single pointer, no array size"; `*count` applies to the second level — the array of `LONG` that the callee allocates. So: the callee allocates `*count` `LONG`s, sets `*values` to point at them, and the caller frees with `CoTaskMemFree`.

3. Every `[out]` parameter must be set to null/zero. The marshaler walks the `[out]` parameters after the call to package results; garbage in an `[out]` pointer means it dereferences a wild address in the *server* process. Callers also rely on it to avoid double-freeing stack garbage.

4. A `BSTR` points 4 bytes past the start of its allocation (the preceding `DWORD` is the byte length). `CoTaskMemFree(bstr)` passes an address that is not an allocation base, so the allocator either faults immediately or, worse, corrupts its own metadata and the process crashes later somewhere unrelated. Additionally `BSTR`s come from `oleaut32`'s allocator/cache, not the task allocator, so even the correct base address would be the wrong heap.

5. (a) Check `HKCR\Interface\{IID}` exists. (b) If `ProxyStubClsid32` is present, follow it to `HKCR\CLSID\{ps}\InprocServer32` and verify the file exists **and matches the bitness of both processes**. (c) If it's `{00020424-...}`, verify `HKCR\TypeLib\{LIBID}\<version>\0\win32|win64` points at an existing file. Use ProcMon to see which lookup actually fails.

6. Choose typelib marshaling when the interface is Automation-compatible and you want to avoid shipping/registering an extra DLL — especially for components consumed by script, VB, or .NET. You give up: non-Automation types (`size_is` arrays, structs, `BYTE*`), and some performance.

7. Hypothesis: **client and server were built from different IDL revisions while keeping the same IID** — the vtable/format strings disagree. First question: "what are the exact file versions of the client binary, the server binary, and the proxy/stub DLL on the affected machine?" Frequently one of the three didn't get updated.

8. Because custom marshaling lets the *server* dictate which CLSID the *client* loads and instantiates in its own process, via `GetUnmarshalClass`. A malicious or compromised lower-privileged server can therefore influence code loading in a higher-privileged client. The flag refuses `IMarshal` and forces standard marshaling.

</details>

---

## 4.12 Rules to carry forward

1. IDL is the contract. Write it first; generate the header from it, never the reverse.
2. `[out]` parameters: null on entry, null on every failure path.
3. `BSTR` → `SysFreeString`; `SAFEARRAY` → `SafeArrayDestroy`; `VARIANT` → `VariantClear`; interface → `Release`; everything else `[out]` → `CoTaskMemFree`.
4. `[in]` buffers belong to the caller. The callee never frees them.
5. `iid_is` on every `void**` that carries an interface pointer.
6. Never modify a published interface. Add `IFoo2` and `QI` for it.
7. Change the interface, change the IID. Always.
8. "Works in-proc, `E_NOINTERFACE` out-of-proc" = missing marshaling registration. Check `HKCR\Interface\{IID}` first.
9. `{00020424-0000-0000-C000-000000000046}` under `ProxyStubClsid32` means typelib marshaling — go verify the TypeLib key.
10. `RPC_X_BAD_STUB_DATA` = mismatched builds. Get version numbers before theorizing.

---

**Next: [Module 5 — Automation, `IDispatch`, and scripting](05-automation-and-idispatch.md)**
