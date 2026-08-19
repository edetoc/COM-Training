# Module 4 — Interfaces, IDL, MIDL, and marshaling

**Time: 1 week.**

Module 3 showed that crossing an apartment boundary requires a proxy. This module explains where proxies come from, how COM knows how to package your parameters, and who owns the memory.

---

## 4.1 Why IDL exists

A C++ header describes an interface to a *C++ compiler*. But COM needs more:

- **The marshaler** needs to know that `[out] BSTR* p` means "the callee allocates a string and the caller frees it," and that `[in, size_is(n)] BYTE* buf` means "copy `n` bytes across the wire."
- **Other languages** need a machine-readable description.
- **Scripting engines** need names and parameter types at runtime.

A C++ header can't express any of that. `int* p` is ambiguous: is it in or out? One integer or an array? If an array, how long?

**IDL** (Interface Definition Language) is the language that removes the ambiguity. **MIDL** is the compiler that turns IDL into:

| Output | Purpose |
|---|---|
| `Foo.h` | C/C++ interface declarations |
| `Foo_i.c` | GUID definitions (`IID_IFoo`, `CLSID_Bar`) |
| `Foo_p.c` | **Proxy and stub code** — the marshaling implementation |
| `dlldata.c` | Proxy/stub DLL boilerplate (`DllGetClassObject` etc. for the ps DLL) |
| `Foo.tlb` | **Type library** — binary metadata for late binding and interop |

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

This is also the reason **Lab 2.2's surrogate experiment failed**. Go back and finish it now: register the proxy/stub, and the `dllhost.exe` surrogate will work.

---

## 4.6 LAB 4.1 — Author IDL, compile with MIDL, read the output

> **Requirements**
> - **Tools:** a **Developer PowerShell for VS** or *x64 Native Tools Command Prompt* — this is what puts `midl.exe` on `PATH`. A plain PowerShell window will fail with "midl is not recognized." Plus `New-Guid` or `guidgen.exe`.
> - **Elevation:** not required — this lab only compiles.
> - **Bitness:** `/env x64`, matching your build.
> - **Depends on:** the IDL from §4.2. Generate **real** GUIDs; do not reuse the placeholders.
> - **Time:** ~90 min — most of it spent reading the generated files, which is the point.

### Step 1 — write it

Save the IDL from §4.2 as `Calculator.idl`. Generate real GUIDs.

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

### Step 4 — build and register the proxy/stub DLL

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
> - **Time:** ~2 h.

1. With the PS DLL registered, redo **Lab 2.2's surrogate experiment**. It now works: your 32-bit DLL runs in `dllhost.exe` and a 64-bit client talks to it. Confirm with Process Explorer that `dllhost.exe` has loaded your DLL.

2. `regsvr32 /u CalcPS.dll`. Re-run. Record the exact HRESULT and where it fails (activation vs. first call).

3. **Switch to typelib marshaling.** Add `oleautomation` to the interface attributes and remove the non-Automation methods (`Checksum` with its `BYTE*`, `GetHistory` with its `LONG**`):

```idl
[
    object,
    oleautomation,
    uuid(...),
    pointer_default(unique)
]
interface ICalculator : IUnknown
{
    HRESULT Add([in] LONG a, [in] LONG b, [out, retval] LONG* result);
    HRESULT Describe([out, retval] BSTR* description);
}
```

Register the type library instead of a PS DLL:

```cpp
// In DllRegisterServer:
ITypeLib* pTypeLib = nullptr;
HRESULT hr = LoadTypeLibEx(modulePath, REGKIND_REGISTER, &pTypeLib);
if (pTypeLib) pTypeLib->Release();
```

(Embed the `.tlb` as resource `1` in the DLL: add `1 typelib "Calculator.tlb"` to your `.rc` file.)

Confirm `ProxyStubClsid32` is now `{00020424-0000-0000-C000-000000000046}` and that cross-apartment calls work **with no proxy/stub DLL at all**.

4. **Prove the restriction.** Put `Checksum([in, size_is(cb)] const BYTE* data, ...)` back while keeping `oleautomation`. MIDL fails with something like:

```
error MIDL2311: this type is not supported by automation: [ Parameter 'data' of Procedure 'Checksum' ]
```

Write that error into your notes — you *will* see a developer hit it.

---

## 4.8 LAB 4.3 — Memory ownership, verified

> **Requirements**
> - **Tools:** Visual Studio C++ with **ATL** (`atlbase.h`, `atlsafe.h`); **page heap** via `gflags.exe` or Application Verifier (Windows SDK); WinDbg for `!heap`.
> - **Elevation:** required — `gflags` and Application Verifier write machine-wide image-execution options. Turn the flags **off** when you finish; page heap left enabled will slow the image down permanently.
> - **Bitness:** x64.
> - **Depends on:** Lab 4.1's `ICalculator` (you need `Describe`, `Checksum`, and `GetHistory`).
> - **Time:** ~90 min.

Write a client that exercises every allocation rule, then prove correctness with tooling.

```cpp
#include <atlbase.h>
#include <atlsafe.h>

void ExerciseMemoryRules(ICalculator* p)
{
    // --- BSTR: callee allocates, caller frees with SysFreeString ---
    BSTR desc = nullptr;
    if (SUCCEEDED(p->Describe(&desc)))
    {
        wprintf(L"desc = %s (len %u)\n", desc, SysStringLen(desc));
        SysFreeString(desc);          // correct
        // CoTaskMemFree(desc);       // <-- HEAP CORRUPTION: wrong base address
    }

    // --- [out] array: callee allocates with CoTaskMemAlloc, caller frees ---
    ULONG count = 0;
    LONG* values = nullptr;
    if (SUCCEEDED(p->GetHistory(&count, &values)))
    {
        for (ULONG i = 0; i < count; ++i) wprintf(L"%ld ", values[i]);
        CoTaskMemFree(values);        // correct
    }

    // --- [in] buffer: caller owns it start to finish ---
    BYTE data[] = { 1, 2, 3, 4 };
    ULONG sum = 0;
    p->Checksum(data, ARRAYSIZE(data), &sum);   // callee must NOT free 'data'

    // --- Failure path: [out] must be nulled ---
    BSTR shouldBeNull = reinterpret_cast<BSTR>(0xDEADBEEF);
    HRESULT hr = p->DescribeThatAlwaysFails(&shouldBeNull);
    assert(FAILED(hr) && shouldBeNull == nullptr);
}
```

Server side, showing correct allocation:

```cpp
HRESULT STDMETHODCALLTYPE Calculator::Describe(BSTR* description)
{
    if (!description) return E_POINTER;
    *description = nullptr;                       // null first
    *description = SysAllocString(L"Training Calculator v1");
    return *description ? S_OK : E_OUTOFMEMORY;
}

HRESULT STDMETHODCALLTYPE Calculator::GetHistory(ULONG* count, LONG** values)
{
    if (!count || !values) return E_POINTER;
    *count = 0; *values = nullptr;                // null EVERYTHING first

    const ULONG n = static_cast<ULONG>(m_history.size());
    if (n == 0) return S_OK;                      // zero items is not an error

    auto* p = static_cast<LONG*>(CoTaskMemAlloc(n * sizeof(LONG)));
    if (!p) return E_OUTOFMEMORY;

    memcpy(p, m_history.data(), n * sizeof(LONG));
    *values = p;
    *count  = n;
    return S_OK;
}
```

### Verify with tooling

1. **Application Verifier** → your client EXE → enable **Basics: Heaps** (page heap) and **COM**. Run under a debugger.
2. Introduce the `CoTaskMemFree(desc)` bug. With page heap on, it faults **immediately, at the bad free**, with a clear "invalid address" — instead of corrupting the heap and crashing somewhere unrelated later. **This is why page heap is the first thing you turn on for a heap-corruption ticket.**
3. Introduce a leak (skip `SysFreeString`) and detect it:
   ```
   0:000> !heap -s
   0:000> !heap -stat -h <heap>          ; allocation stats by size
   0:000> !heap -p -a <address>          ; who allocated this block (needs page heap)
   ```
4. Check `BSTR` accounting specifically: `oleaut32` keeps a cache. `!bstr` isn't standard, but you can set `OANOCACHE=1` in the environment to disable the `BSTR` cache so leaks show up immediately in heap statistics. **Remember this trick** — `BSTR` leaks are otherwise masked by the cache.

---

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
