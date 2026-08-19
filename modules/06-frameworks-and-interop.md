# Module 6 — Frameworks: ATL, WRL, WIL, and .NET interop

**Time: 1 week.**

Modules 1–5 built everything by hand so you'd understand the machinery. Nobody writes production COM that way. This module covers the frameworks that remove the boilerplate — and, more importantly, what each one is doing underneath so you can debug it.

---

## 6.1 What the frameworks actually replace

Count what you hand-wrote for one object:

| Hand-written | Lines | Replaced by |
|---|---|---|
| `QueryInterface` if/else chain | ~15 | `BEGIN_COM_MAP` / `RuntimeClass` |
| `AddRef`/`Release` + `delete this` | ~10 | `CComObjectRootEx` / `RuntimeClass` |
| `IClassFactory` implementation | ~40 | `CComCoClass` + object map |
| `DllGetClassObject`/`DllCanUnloadNow` | ~20 | `CAtlDllModuleT` / `Module<InProc>` |
| `DllRegisterServer` registry writing | ~60 | `.rgs` script / `Module::RegisterServer` |
| `IDispatch` implementation | ~50 | `IDispatchImpl` |
| Connection points | ~80 | `IConnectionPointContainerImpl` + `CProxy_*` |
| `BSTR`/`VARIANT`/`SAFEARRAY` lifetime | everywhere | `CComBSTR`/`CComVariant`/`CComSafeArray` |

Roughly 275 lines of error-prone boilerplate → about 20 lines of declarations.

---

## 6.2 ATL — the Active Template Library

ATL is template-based, has no runtime dependency beyond the CRT, and produces small binaries. It remains the standard for C++ COM servers.

### The class declaration

```cpp
#include <atlbase.h>
#include <atlcom.h>

class ATL_NO_VTABLE CCalculator :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<CCalculator, &CLSID_Calculator>,
    public ISupportErrorInfoImpl<&IID_ICalculator>,
    public IConnectionPointContainerImpl<CCalculator>,
    public CProxy_ICalculatorEvents<CCalculator>,
    public IDispatchImpl<ICalculator, &IID_ICalculator, &LIBID_TrainingCalcLib, 1, 0>
{
public:
    CCalculator() = default;

    DECLARE_REGISTRY_RESOURCEID(IDR_CALCULATOR)      // registration from a .rgs script
    DECLARE_NOT_AGGREGATABLE(CCalculator)
    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CCalculator)
        COM_INTERFACE_ENTRY(ICalculator)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY(ISupportErrorInfo)
        COM_INTERFACE_ENTRY(IConnectionPointContainer)
    END_COM_MAP()

    BEGIN_CONNECTION_POINT_MAP(CCalculator)
        CONNECTION_POINT_ENTRY(__uuidof(_ICalculatorEvents))
    END_CONNECTION_POINT_MAP()

    HRESULT FinalConstruct() { return S_OK; }        // real init goes here, not the ctor
    void    FinalRelease()   {}                      // real teardown goes here

    // ICalculator
    STDMETHOD(Add)(LONG a, LONG b, LONG* result)
    {
        if (!result) return E_POINTER;
        *result = a + b;
        Fire_OnCalculated(*result);
        return S_OK;
    }
};

OBJECT_ENTRY_AUTO(__uuidof(Calculator), CCalculator)   // adds it to the object map
```

### What each piece does

| Element | Purpose |
|---|---|
| `ATL_NO_VTABLE` | `__declspec(novtable)` — suppresses vtable init in the abstract base's ctor. Smaller code, faster construction. Safe because ATL never instantiates the class directly. |
| `CComObjectRootEx<ThreadModel>` | Provides the ref count and the `QueryInterface` engine. `CComMultiThreadModel` = `InterlockedIncrement`; `CComSingleThreadModel` = plain `++` (STA only, faster). |
| `CComCoClass<T, &CLSID>` | Provides `CreateInstance`, the class factory hookup, and `Error()` for `IErrorInfo`. |
| `BEGIN_COM_MAP` | Builds a static table of `{IID, offset, function}`. `QueryInterface` becomes a table walk. |
| `OBJECT_ENTRY_AUTO` | Registers the class in the module's object map so `DllGetClassObject` can find it. |
| `FinalConstruct` / `FinalRelease` | Init/teardown that can **fail** and can safely call `QueryInterface` on `this` — a C++ ctor can't do either. |
| `DECLARE_PROTECT_FINAL_CONSTRUCT` | Guards against the object being destroyed inside `FinalConstruct` if it hands out a reference that gets released. |

### `COM_INTERFACE_ENTRY` variants worth knowing

```cpp
BEGIN_COM_MAP(CFoo)
    COM_INTERFACE_ENTRY(IFoo)                             // plain
    COM_INTERFACE_ENTRY2(IDispatch, IFoo)                 // disambiguate multiple IDispatch paths
    COM_INTERFACE_ENTRY_IID(IID_IBar, CMyBarImpl)         // explicit IID
    COM_INTERFACE_ENTRY_AGGREGATE(IID_IBaz, m_pInner)     // delegate to an aggregated object
    COM_INTERFACE_ENTRY_NOINTERFACE(IMarshal)             // explicitly refuse (blocks custom marshaling)
    COM_INTERFACE_ENTRY_FUNC(IID_IQux, 0, &CFoo::QueryQux) // custom handler / tear-off
    COM_INTERFACE_ENTRY_CHAIN(CBaseClass)                 // continue in a base class's map
END_COM_MAP()
```

`COM_INTERFACE_ENTRY2` matters: if your class inherits `IDispatch` through two paths (e.g. a dual interface *and* an event interface), the compiler can't pick, and you must say which.

### The module class

```cpp
class CCalcModule : public ATL::CAtlDllModuleT<CCalcModule>
{
public:
    DECLARE_LIBID(LIBID_TrainingCalcLib)
    DECLARE_REGISTRY_APPID_RESOURCEID(IDR_CALC, "{APPID-GUID}")
};
CCalcModule _AtlModule;

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{ return _AtlModule.DllGetClassObject(rclsid, riid, ppv); }

STDAPI DllCanUnloadNow(void)      { return _AtlModule.DllCanUnloadNow(); }
STDAPI DllRegisterServer(void)    { return _AtlModule.DllRegisterServer(); }
STDAPI DllUnregisterServer(void)  { return _AtlModule.DllUnregisterServer(); }
```

For an EXE server, use `CAtlExeModuleT`, which handles `CoRegisterClassObject`, the message loop, and `CoAddRefServerProcess` shutdown logic.

### `.rgs` registration scripts

ATL registers via a script resource instead of hand-written `RegSetValueEx` calls:

```
HKCR
{
    NoRemove CLSID
    {
        ForceRemove {A1B2C3D4-1111-4000-9000-000000000001} = s 'Calculator Component'
        {
            ProgID = s 'Training.Calculator.1'
            VersionIndependentProgID = s 'Training.Calculator'
            ForceRemove 'Programmable'
            InprocServer32 = s '%MODULE%'
            {
                val ThreadingModel = s 'Both'
            }
            val AppID = s '%APPID%'
            'TypeLib' = s '{A1B2C3D4-9999-4000-9000-000000000099}'
        }
    }
    Training.Calculator.1 = s 'Calculator Component'
    {
        CLSID = s '{A1B2C3D4-1111-4000-9000-000000000001}'
    }
}
```

Keywords: `NoRemove` (don't delete on unregister), `ForceRemove` (delete the whole subtree first — prevents stale values), `val` (a named value), `s`/`d`/`b` (string/dword/binary). `%MODULE%` and `%APPID%` are substituted at runtime.

**Support tip:** when a component "half-registers," read its `.rgs` (visible in the DLL's resources via Resource Hacker or `dumpbin /section:.rsrc`). It's the authoritative statement of what *should* be in the registry — diff it against reality.

### ATL smart types

| Type | Wraps | Notes |
|---|---|---|
| `CComPtr<T>` | interface pointer | `operator&` asserts if already non-null |
| `CComQIPtr<T, &IID>` | interface pointer | assigning from another pointer does an implicit `QI` |
| `CComBSTR` | `BSTR` | `SysAllocString`/`SysFreeString` |
| `CComVariant` | `VARIANT` | `VariantInit`/`VariantClear`, `ChangeType` |
| `CComSafeArray<T>` | `SAFEARRAY*` | bounds, locking, element cleanup |
| `CComCritSecLock<T>` | critical section | RAII lock |
| `CComGITPtr<T>` | GIT cookie | RAII `RegisterInterfaceInGlobal`/`Revoke` — use this in Module 3 scenarios |

```cpp
CComQIPtr<IPersistFile> spFile = spLink;    // implicit QueryInterface
if (spFile) spFile->Save(path, TRUE);       // null check == QI succeeded
```

---

## 6.3 WRL — Windows Runtime C++ Template Library

Header-only, no ATL dependency, works in app containers, and is what WinRT components use. Also perfectly usable for classic COM.

```cpp
#include <wrl.h>
using namespace Microsoft::WRL;

class Calculator : public RuntimeClass<RuntimeClassFlags<ClassicCom>, ICalculator>
{
    InspectableClass(nullptr, TrustLevel::BaseTrust)   // omit for pure classic COM

public:
    HRESULT RuntimeClassInitialize(LONG precision)     // fallible init, like FinalConstruct
    {
        m_precision = precision;
        return S_OK;
    }

    IFACEMETHODIMP Add(LONG a, LONG b, LONG* result) override
    {
        if (!result) return E_POINTER;
        *result = a + b;
        return S_OK;
    }
private:
    LONG m_precision = 0;
};

// Creation, with fallible init:
ComPtr<ICalculator> calc;
HRESULT hr = MakeAndInitialize<Calculator>(&calc, /*precision*/ 4);
```

| WRL flag | Meaning |
|---|---|
| `ClassicCom` | Plain `IUnknown`-based COM |
| `WinRt` | `IInspectable`-based (WinRT) |
| `WinRtClassicComMix` | Both |
| `InhibitWeakReference` | No `IWeakReference` support |
| `Delegate` | For WinRT delegates |

`RuntimeClass` generates `QueryInterface` from the template parameter list — no macro map. `ComPtr<T>` is the WRL smart pointer; note it uses `.Get()`, `.GetAddressOf()`, and `.ReleaseAndGetAddressOf()` rather than `operator&` overloading tricks.

---

## 6.4 WIL — Windows Implementation Library

Not a COM framework; a set of RAII and error-handling helpers that make COM code dramatically safer. Header-only, from https://github.com/microsoft/wil.

```cpp
#include <wil/com.h>
#include <wil/result.h>
#include <wil/resource.h>

HRESULT DoWork()
{
    auto coInit = wil::CoInitializeEx(COINIT_APARTMENTTHREADED);   // RAII, correct S_FALSE handling

    wil::com_ptr<ICalculator> calc;
    RETURN_IF_FAILED(CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_INPROC_SERVER,
                                      IID_PPV_ARGS(&calc)));

    LONG r = 0;
    RETURN_IF_FAILED(calc->Add(2, 3, &r));

    auto adv = calc.query<IAdvancedCalculator>();      // throws on failure
    auto advOpt = calc.try_query<IAdvancedCalculator>(); // returns null instead

    wil::unique_bstr desc;
    RETURN_IF_FAILED(calc->Describe(&desc));           // SysFreeString automatic

    wil::unique_variant v;                             // VariantClear automatic

    return S_OK;
}
```

The error macros are the real win:

| Macro | Behaviour |
|---|---|
| `RETURN_IF_FAILED(hr)` | returns `hr` and **logs file/line/HRESULT** |
| `THROW_IF_FAILED(hr)` | throws `wil::ResultException` |
| `LOG_IF_FAILED(hr)` | logs and continues |
| `FAIL_FAST_IF_FAILED(hr)` | terminates immediately (for invariant violations) |
| `RETURN_HR_IF(hr, cond)` | conditional return |
| `RETURN_LAST_ERROR_IF(cond)` | converts `GetLastError()` |

Because WIL logs the **originating** file and line, a failure deep in a call chain tells you exactly where it started — instead of an `HRESULT` that bubbled up from somewhere unknown. Hook it up:

```cpp
wil::SetResultLoggingCallback([](const wil::FailureInfo& fi) noexcept {
    wchar_t msg[2048];
    wil::GetFailureLogString(msg, ARRAYSIZE(msg), fi);
    OutputDebugStringW(msg);
});
```

**For support work this is gold:** ask the developer to enable WIL logging and re-run; you get a precise origin instead of a generic `E_FAIL`.

---

## 6.5 C++/WinRT for classic COM

```cpp
#include <winrt/base.h>

winrt::com_ptr<ICalculator> calc;
winrt::check_hresult(CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_INPROC_SERVER,
                                      IID_PPV_ARGS(calc.put())));

auto adv  = calc.as<IAdvancedCalculator>();      // throws winrt::hresult_error
auto adv2 = calc.try_as<IAdvancedCalculator>();  // returns empty com_ptr

winrt::agile_ref<ICalculator> agile{ calc };     // GIT-backed, safe across apartments
```

Note `put()` vs `get()` vs `put_void()` — C++/WinRT deliberately avoids `operator&` so the `CComPtr` double-assign trap can't happen.

---

## 6.6 .NET interop: RCW and CCW

### The two wrappers

```
  .NET calling COM:                     COM calling .NET:

  ┌──────────────┐                      ┌──────────────┐
  │  C# code     │                      │  C++ client  │
  │      │       │                      │      │       │
  │      ▼       │                      │      ▼       │
  │     RCW      │  Runtime Callable    │     CCW      │  COM Callable
  │      │       │  Wrapper             │      │       │  Wrapper
  └──────┼───────┘                      └──────┼───────┘
         ▼                                     ▼
   COM object                            .NET object
```

- **RCW** — a .NET object that holds the COM interface pointer, translates calls, and `Release`s in its finalizer. **One RCW per COM object identity per apartment** — the CLR uses `QI(IID_IUnknown)` (Module 1, Rule 1!) to detect that two pointers are the same object and returns the same RCW.
- **CCW** — a COM object that implements `IUnknown`, `IDispatch`, `ISupportErrorInfo`, etc., and forwards to a .NET object. Keeps the .NET object alive via a GC handle.

### The `ReleaseComObject` minefield

```csharp
var excel = new Excel.Application();
var books = excel.Workbooks;              // <- this creates an RCW you forgot about
var book  = books.Open(path);
// ...
Marshal.ReleaseComObject(book);
Marshal.ReleaseComObject(books);
Marshal.ReleaseComObject(excel);
```

Semantics you must know:

| Call | What it does |
|---|---|
| `Marshal.ReleaseComObject(o)` | Decrements the **RCW's internal reference count**. Only when it hits 0 does it `Release` the underlying COM pointer. Returns the new count. |
| `Marshal.FinalReleaseComObject(o)` | Forces the RCW count to 0 and releases immediately, regardless of other holders. |
| Doing nothing | The GC finalizes the RCW eventually and releases. Correct, but non-deterministic. |

The danger: **the RCW is shared**. If two places in your code obtained the "same" COM object, they share one RCW. `ReleaseComObject` from one place breaks the other:

```csharp
Marshal.ReleaseComObject(book);
book.Save();     // System.Runtime.InteropServices.InvalidComObjectException:
                 // "COM object that has been separated from its underlying RCW cannot be used."
```

Also `0x80004002` / `NullReferenceException` in odd places.

**The "two dots" rule** you'll see in Office interop guidance:

```csharp
// BAD: creates an unreferenced intermediate RCW for Workbooks
var book = excel.Workbooks.Open(path);

// GOOD: every intermediate object is named and can be released
var books = excel.Workbooks;
var book = books.Open(path);
```

**Modern guidance:** for most code, don't call `ReleaseComObject` at all — let the GC handle it, and use `GC.Collect(); GC.WaitForPendingFinalizers();` only if you genuinely must force teardown (e.g. making Excel.exe exit). Use `ReleaseComObject` only in tight, well-understood scopes where you own every reference.

### Exposing .NET to COM

```csharp
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid("B1B2C3D4-0001-4000-9000-000000000001")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface ICalculator
{
    [DispId(1)] int Add(int a, int b);
    [DispId(2)] int Subtract(int a, int b);
    [DispId(3)] int Precision { get; set; }
}

[ComVisible(true)]
[Guid("B1B2C3D4-1111-4000-9000-000000000001")]
[ProgId("Training.NetCalculator.1")]
[ClassInterface(ClassInterfaceType.None)]      // <- ALWAYS None
public class Calculator : ICalculator
{
    public int Add(int a, int b) => a + b;
    public int Subtract(int a, int b) => a - b;
    public int Precision { get; set; }
}
```

**`ClassInterfaceType.None` is not optional advice.** The alternatives auto-generate an interface from the class's public members, which means:

- Adding a public method changes the generated vtable → breaks compiled clients. You've reintroduced exactly the fragility COM was invented to prevent.
- DISPIDs shift between builds.

Always declare an explicit interface and set `ClassInterface(ClassInterfaceType.None)`.

### Registration

| Runtime | Tool | Notes |
|---|---|---|
| .NET Framework | `regasm Calc.dll /codebase /tlb` | `/codebase` writes the assembly path (needed outside the GAC); `/tlb` generates and registers a type library |
| .NET Core / 5+ | `dotnet build` with `<EnableComHosting>true</EnableComHosting>` | Produces `Calc.comhost.dll`; register that with `regsvr32` |

```xml
<!-- .csproj for .NET 5+ COM server -->
<PropertyGroup>
  <TargetFramework>net8.0-windows</TargetFramework>
  <EnableComHosting>true</EnableComHosting>
</PropertyGroup>
```

> **Support fact:** `regasm` **does not exist** for .NET Core/5+. If a customer is trying to `regasm` a modern .NET assembly, that's the bug. They need COM hosting and `regsvr32` on the generated `*.comhost.dll`.

### Type libraries and PIAs

| Concept | Meaning |
|---|---|
| `tlbimp Foo.tlb` | Type library → interop assembly (COM → .NET) |
| `tlbexp Foo.dll` | Assembly → type library (.NET → COM) |
| **PIA** (Primary Interop Assembly) | The vendor's official, signed interop assembly (e.g. `Microsoft.Office.Interop.Excel`). Ensures everyone uses the same types. |
| **NoPIA / Embed Interop Types** | The compiler embeds only the interop types you actually use, directly into your assembly. **No PIA deployment required.** |

Embed Interop Types is the modern default (`<EmbedInteropTypes>true</EmbedInteropTypes>` or the "Embed Interop Types" property = True). It eliminates a whole class of "PIA not installed" deployment tickets.

### `ComWrappers` — the modern API

.NET 5+ introduced `ComWrappers`, giving explicit control over RCW/CCW creation, and .NET 8 added **source-generated COM interop**:

```csharp
using System.Runtime.InteropServices.Marshalling;

[GeneratedComInterface]
[Guid("A1B2C3D4-0001-4000-9000-000000000001")]
internal partial interface ICalculator
{
    int Add(int a, int b);
}
```

The source generator emits the marshaling code at compile time — AOT-compatible, trimmable, no runtime reflection. This is the direction .NET COM interop is going; built-in `ComImport` interop is not supported in Native AOT.

### .NET interop error codes

| Exception / HRESULT | Meaning |
|---|---|
| `InvalidComObjectException` | RCW was separated (over-released) |
| `COMException 0x80040154` | Class not registered — check bitness, `/codebase`, comhost |
| `COMException 0x80004002` | `E_NOINTERFACE` — often a missing typelib for marshaling |
| `InvalidCastException` on a COM interface | `QI` failed; usually marshaling or apartment |
| `0x8013150A` `HOST_E_INVALIDOPERATION` | COM call on a thread with the wrong apartment/host state |
| `SafeArrayTypeMismatchException` | `SAFEARRAY` element type differs from the declared one |

---

## 6.7 LAB 6.1 — Rewrite the Module 2 server in ATL

> **Requirements**
> - **Tools:** Visual Studio with the optional component **C++ ATL for latest build tools (x86 & x64)**. It is *not* installed by default with *Desktop development with C++* — if you have no *ATL Project* template, that is why. Add it in the VS Installer.
> - **Elevation:** required for `regsvr32`.
> - **Bitness:** x64, matching the Module 2 client you reuse.
> - **Depends on:** the hand-written Module 2 server — kept, not deleted. The comparison table is the deliverable.
> - **Time:** ~2 h.

1. **File → New → Project → ATL Project**, DLL, no attributes.
2. Add a **Simple Object** (`CCalculator`) via the wizard: choose "Dual" interface, "Both" threading, "Support ISupportErrorInfo", "Support Connection Points".
3. Add `Add`, `Subtract`, `Describe`, and a `Precision` property through the IDL editor / Add Method wizard.
4. Build and `regsvr32` it. Verify with the same client from Module 2.

### Then do the comparison — this is the actual lab

Produce a table:

| Concern | Hand-written (Module 2) | ATL | What ATL generated |
|---|---|---|---|
| `QueryInterface` | lines __ | lines __ | |
| Ref counting | | | |
| Class factory | | | |
| DLL exports | | | |
| Registration | | | |
| `IDispatch` | | | |
| Connection points | | | |

Then **step into the ATL code in the debugger**:

- Break in `CComObjectRootBase::InternalQueryInterface` and look at the `_ATL_INTMAP_ENTRY` table it walks. That's what `BEGIN_COM_MAP` built.
- Break in `CComCreator::CreateInstance` and watch `FinalConstruct` get called.
- Look at `CComObject<CCalculator>` — the concrete class ATL synthesizes, which supplies the actual `AddRef`/`Release`/`delete this`.

Understanding that `CComObject<T>` is the thing actually instantiated (not `T`) explains why `T`'s constructor can't fail and why `FinalConstruct` exists.

---

## 6.8 LAB 6.2 — Cross-language interop, both directions

> **Requirements**
> - **Tools:** the **.NET SDK** (.NET 8 or later) for the `EnableComHosting` path, **and** the .NET Framework 4.x developer pack if you want to compare `regasm`/`tlbimp` — those tools are Framework-only and their absence is itself a support lesson. Visual Studio C++ for the native client.
> - **Elevation:** required — `regsvr32` on the generated `*.comhost.dll`.
> - **Bitness:** the comhost is **architecture-specific**. Publish it for the same architecture as the calling client (`-r win-x64`), or you reproduce `0x80040154`.
> - **Depends on:** the Lab 6.1 ATL server for Direction 2.
> - **Extra machine:** Direction 3 needs a **clean VM or second machine** that does not have the interop assembly deployed — on your dev box the failure will not reproduce.
> - **Time:** ~3 h.

### Direction 1: C# server, C++ client

```csharp
// NetCalc.cs, .NET Framework or .NET 8 with EnableComHosting
[ComVisible(true), Guid("B1B2C3D4-0001-...")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface INetCalculator { [DispId(1)] int Add(int a, int b); }

[ComVisible(true), Guid("B1B2C3D4-1111-..."), ProgId("Training.NetCalc.1")]
[ClassInterface(ClassInterfaceType.None)]
public class NetCalculator : INetCalculator { public int Add(int a, int b) => a + b; }
```

Register, then call from C++ with `CoCreateInstance` + `#import` of the generated TLB. Observe that from C++ this is indistinguishable from a C++ server. **That's the whole point of COM.**

### Direction 2: C++ server, C# client

Take the ATL server from Lab 6.1 and consume it three ways:

```csharp
// a) Early bound with a COM reference (Embed Interop Types = true)
var calc = new TrainingCalcLib.Calculator();
Console.WriteLine(calc.Add(2, 3));

// b) Late bound
Type t = Type.GetTypeFromProgID("Training.Calculator.1");
dynamic d = Activator.CreateInstance(t);
Console.WriteLine(d.Add(2, 3));

// c) Explicit reflection - what 'dynamic' does underneath
object o = Activator.CreateInstance(t);
object r = t.InvokeMember("Add", BindingFlags.InvokeMethod, null, o, new object[] { 2, 3 });
```

Compare: which of these needs the type library registered? Which needs the TLB *file* present at build time? Which survives the TLB being unregistered at runtime?

### Direction 3: turn EmbedInteropTypes off and on

Build the C# client with `<EmbedInteropTypes>false</EmbedInteropTypes>`, deploy to a clean machine without the interop assembly, and observe the failure. Turn it back on and confirm the deployment problem disappears.

---

## 6.9 LAB 6.3 — Reproduce the RCW bugs

> **Requirements**
> - **Tools:** the .NET SDK and the Visual Studio debugger. Run each case in the debugger — several of these bugs are only visible as a *timing* difference in when the finalizer runs.
> - **Elevation:** required once, to register `Training.Calculator.1`.
> - **Bitness:** match the registered server.
> - **Depends on:** a registered `Training.Calculator.1` (Lab 6.1) with the Module 1 ref-count tracing still compiled in.
> - **Time:** ~1 h.

```csharp
using System;
using System.Runtime.InteropServices;

class Program
{
    static void Main()
    {
        Type t = Type.GetTypeFromProgID("Training.Calculator.1");

        // --- Bug 1: use after ReleaseComObject ---
        dynamic calc = Activator.CreateInstance(t);
        Console.WriteLine(calc.Add(2, 3));
        Marshal.ReleaseComObject(calc);
        try   { Console.WriteLine(calc.Add(4, 5)); }
        catch (InvalidComObjectException e) { Console.WriteLine($"Bug 1: {e.Message}"); }

        // --- Bug 2: shared RCW ---
        dynamic a = Activator.CreateInstance(t);
        dynamic b = a;                       // SAME RCW, not a copy
        Marshal.ReleaseComObject(a);
        try   { Console.WriteLine(b.Add(1, 1)); }
        catch (InvalidComObjectException e) { Console.WriteLine($"Bug 2: {e.Message}"); }

        // --- Correct: let the GC do it ---
        dynamic c = Activator.CreateInstance(t);
        Console.WriteLine(c.Add(6, 7));
        c = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();       // now the COM Release has happened
        Console.WriteLine("clean");
    }
}
```

Run with the Module 1 ref-count tracing enabled in the C++ server so you can **see** when the actual COM `Release` happens in each case. The gap between "C# dropped the reference" and "COM `Release` fired" is the non-determinism people complain about.

### Then: the Excel-won't-exit drill

```csharp
var excel = new Excel.Application();
excel.Visible = false;
var books = excel.Workbooks;
var book = books.Add();
book.Close(false);
excel.Quit();
// Excel.exe is STILL RUNNING. Why?
```

Because RCWs for `books`, `book`, and any un-named intermediates still hold references. Fix it and verify with Task Manager that `EXCEL.EXE` exits. This exact scenario is one of the most common .NET+COM support tickets in existence.

---

## 6.10 Choosing a framework

| Situation | Use |
|---|---|
| New C++ in-proc/out-of-proc COM server, desktop | **ATL** — mature, small, wizard support, `.rgs` |
| WinRT component, or COM in an app container | **WRL** or **C++/WinRT** |
| Any modern C++ COM *client* code | **WIL** `com_ptr` + `RETURN_IF_FAILED` |
| Consuming COM from C++ with a type library | `#import` for prototyping; explicit interfaces for shipping |
| New .NET COM server | Explicit interface + `ClassInterfaceType.None`; `EnableComHosting` for .NET 5+ |
| .NET consuming COM | Embed Interop Types; avoid `ReleaseComObject` unless necessary |
| .NET 8+, Native AOT or trimming | **`[GeneratedComInterface]`** source-generated interop |

> **"WinRT" keeps appearing above — what is it?** The Windows Runtime is **COM with a stricter contract**: `IInspectable` derives from `IUnknown`, so every rule from Modules 1–7 still holds. What changed is metadata (`.winmd` instead of `.tlb`), activation (by class-name string instead of CLSID), strings (`HSTRING` instead of `BSTR`), registration (package manifest instead of registry), and a third apartment type (ASTA) that forbids the reentrancy Module 3 warned you about. See **[Appendix B §B.2](appendix-b-com-plus-and-winrt.md#b2-winrt--com-with-new-rules)**.

---

## 6.11 Checkpoint

1. What does `ATL_NO_VTABLE` do, and why is it safe when ATL uses it but dangerous if you applied it to a class you `new` directly?
2. Why does ATL have `FinalConstruct` instead of doing the work in the constructor? Give two reasons.
3. Explain `CComObject<CCalculator>` — what does it add to `CCalculator`, and why isn't `CCalculator` instantiated directly?
4. A C# COM server uses `[ClassInterface(ClassInterfaceType.AutoDual)]`. A developer adds a public method. What breaks, and why is this exactly the problem COM was designed to prevent?
5. What is the difference between `Marshal.ReleaseComObject` and `Marshal.FinalReleaseComObject`, and when is either appropriate?
6. Two C# variables reference "the same" COM object obtained by two separate calls. How does the CLR know to give them the same RCW, and which Module 1 rule does that depend on?
7. A customer runs `regasm` on a .NET 8 assembly and it fails. What's your answer?
8. Why does `COM_INTERFACE_ENTRY2` exist?

<details>
<summary>Answers</summary>

1. `ATL_NO_VTABLE` = `__declspec(novtable)`: the compiler omits vtable pointer initialization in the class's own constructor/destructor, shrinking code and speeding construction. Safe only because ATL never instantiates that class directly — it instantiates `CComObject<T>`, whose constructor *does* set the vtable. If you `new` a `novtable` class directly and call a virtual function, you jump through an uninitialized vptr.

2. (a) A C++ constructor cannot return an `HRESULT`, so fallible initialization has no clean failure path (and COM servers must not rely on exceptions crossing the boundary). (b) During the constructor the vtable and the ATL ref-count plumbing aren't fully set up, so you cannot safely call `QueryInterface` on `this` or hand out references. `FinalConstruct` runs after the object is fully formed.

3. `CComObject<T>` derives from `T` and supplies the concrete `AddRef`/`Release` (with `delete this` at zero) and the outer `QueryInterface`. `T` itself is abstract-ish and ref-count-agnostic, which lets ATL generate *variants*: `CComObject` (heap, self-deleting), `CComObjectStack` (stack), `CComObjectGlobal` (static, never deleted), `CComAggObject` (aggregatable), `CComObjectNoLock` (doesn't lock the module). Same `T`, different lifetime policies.

4. The auto-generated class interface's vtable and DISPIDs are derived from the class's public members in declaration order. Adding a public method shifts slots and DISPIDs, so already-compiled clients call the wrong function — silently wrong results or a crash. This is precisely the fragile-binary-interface problem from Module 0, reintroduced by convenience.

5. `ReleaseComObject` decrements the RCW's internal count and only calls COM `Release` when it reaches zero; `FinalReleaseComObject` forces it to zero and releases immediately. Use `ReleaseComObject` in narrow scopes where you own every reference; use `FinalReleaseComObject` only when you must guarantee teardown (e.g., forcing an out-of-proc server to exit) and you're certain nothing else uses the object. Prefer neither — let the GC do it — unless you have a concrete deterministic-release requirement.

6. The CLR calls `QueryInterface(IID_IUnknown)` on both pointers and compares the results. Identical `IUnknown` pointers means the same object, so the same RCW is returned from the per-apartment RCW cache. This depends entirely on Module 1's Rule 1 (reflexive `QI` returns a canonical `IUnknown`). An object that violates that rule breaks .NET interop in bewildering ways.

7. `regasm` is .NET Framework only and doesn't exist for .NET Core/5+. They need `<EnableComHosting>true</EnableComHosting>` in the project, which produces `<AssemblyName>.comhost.dll`, and then `regsvr32` that file. Also check bitness — the comhost is architecture-specific.

8. When a class inherits `IDispatch` through more than one path (e.g., a dual interface and an event dispinterface), `COM_INTERFACE_ENTRY(IDispatch)` is ambiguous and won't compile. `COM_INTERFACE_ENTRY2(IDispatch, ICalculator)` says "when asked for `IDispatch`, return the one reached via `ICalculator`."

</details>

---

## 6.12 Rules to carry forward

1. Use ATL for C++ COM servers; use WIL for COM client code. Don't hand-write `IUnknown` in production.
2. Fallible initialization goes in `FinalConstruct`/`RuntimeClassInitialize`, never the constructor.
3. In .NET COM servers: explicit interface + `[ClassInterface(ClassInterfaceType.None)]` + explicit `[DispId]`. Always.
4. `regasm` for .NET Framework; `EnableComHosting` + `regsvr32` for .NET 5+.
5. Prefer Embed Interop Types over deploying PIAs.
6. Don't call `ReleaseComObject` reflexively. Understand RCW sharing before you do.
7. `.rgs` files are the authoritative record of what a component registers — read them when diagnosing registration.
8. `CComQIPtr` assignment is a hidden `QueryInterface`; `CComPtr::operator&` asserts on a non-null pointer.
9. Turn on WIL failure logging when diagnosing an opaque `E_FAIL`.
10. For new .NET 8+ interop, use `[GeneratedComInterface]`.

---

**Next: [Module 7 — DCOM, security, and out-of-proc servers](07-dcom-and-security.md)**
