# Module 6 — Choosing COM tools: ATL, WRL, WIL, C++/WinRT, and .NET

In Module 2 you implemented a COM server by hand. In Module 5 you used the ATL wizard to build a server with Automation and events. Both followed the same COM rules; ATL supplied much of the code you would otherwise have written yourself.

This module explains what ATL and the other tools provide, how they differ, and how to choose between them. You do not need to learn every library before writing your next COM program.

**What this module covers**

We start with ATL, which you already used, then introduce WRL, WIL, and C++/WinRT through their main uses. Next we explain how .NET connects managed code to COM. The labs reuse your existing servers to examine generated code and object lifetimes; the final decision guide brings the choices together.

**Contents**

- [6.1 The jobs these tools do](#61-the-jobs-these-tools-do)
- [6.2 ATL — the Active Template Library](#62-atl--the-active-template-library)
- [6.3 WRL — Windows Runtime C++ Template Library](#63-wrl--windows-runtime-c-template-library)
- [6.4 WIL — Windows Implementation Library](#64-wil--windows-implementation-library)
- [6.5 C++/WinRT: a Windows Runtime projection](#65-cwinrt-a-windows-runtime-projection)
- [6.6 .NET interop: RCW and CCW](#66-net-interop-rcw-and-ccw)
- [6.7 Choosing a framework](#67-choosing-a-framework)
- [6.8 LAB 6.1 — The hand-written server vs the ATL one](#68-lab-61--the-hand-written-server-vs-the-atl-one)
- [6.9 LAB 6.2 — Cross-language interop, both directions](#69-lab-62--cross-language-interop-both-directions)
- [6.10 LAB 6.3 — Reproduce the RCW bugs](#610-lab-63--reproduce-the-rcw-bugs)
- [6.11 Checkpoint](#611-checkpoint)
- [6.12 Rules to carry forward](#612-rules-to-carry-forward)

---

## 6.1 The jobs these tools do

Before choosing a tool, distinguish three jobs:

- **Implementing a COM object:** supplying its interfaces, `QueryInterface`, and reference counting. A registered server also needs activation and deployment support.
- **Calling a COM object:** obtaining interface pointers, calling methods, checking errors, and releasing resources.
- **Crossing a language boundary:** making a COM object usable from .NET, or exposing a .NET object to COM callers.

The tools overlap, but they are not five alternatives for the same job:

- **ATL** provides extensive support for classic C++ COM servers, especially Automation and connection points. This is what you used in Stage 4.
- **WRL** provides C++ helpers for implementing COM and Windows Runtime objects, plus smart pointers for callers.
- **WIL** helps C++ code clean up resources and report errors. It complements an object framework rather than replacing one.
- **C++/WinRT** makes Windows Runtime APIs feel like modern C++ APIs. It also has helpers for classic COM.
- **.NET interop** connects managed languages such as C# to COM through wrappers.

> **Windows Runtime (WinRT)** is a COM-based API model used by many Windows features. It adds conventions and machine-readable metadata so languages can present those APIs naturally. It is not the same thing as the .NET runtime.

These choices are independent on either side of a call. Your ATL server can be used by a WRL client or a C# client without changing the server. You can also keep ATL for the server and use WIL inside its methods.

None of these tools removes COM's contracts: you still need correct interface definitions, ownership, threading, and registration where activation requires it.

---

## 6.2 ATL — the Active Template Library

### What it is for

**ATL helps you build classic COM components in C++.** Instead of writing all the server machinery, you describe your classes and interfaces using templates and macros, then implement the methods that do the actual work.

Your Stage 4 calculator is the example: you wrote `Add`, `Divide`, and `SumTo`; ATL supplied reference counting, activation support, and much of the Automation and event machinery. ATL also has client helpers, including `CComPtr` for interface ownership and event-sink helpers. The hand-written sink in Module 5 was a learning exercise, not an ATL limitation.

### What makes it different

ATL's strength is the **breadth of its classic COM support**. It brings together class factories, DLL or EXE hosting, registry scripts, type-library-based `IDispatch`, and connection points. The Visual Studio wizard connects these pieces for you. MIDL still generates the type library from IDL; ATL does not replace that tool.

That convenience comes with conventions to learn: base classes, COM maps, and generated files. ATL does not implement your business logic or make it thread-safe. For example, `CComMultiThreadModel` protects the reference count, not concurrent access to your calculator's fields.

### Recognize it in your project

This is the COM map inside your calculator class, not a new class to paste into the project:

```cpp
BEGIN_COM_MAP(CCalculator)
    COM_INTERFACE_ENTRY(ICalculator)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    COM_INTERFACE_ENTRY(IConnectionPointContainer)
END_COM_MAP()
```

It lists the interfaces that `QueryInterface` can return. Compare that with the `if` statements you wrote in Module 1. Other useful landmarks are:

- `CComObject<CCalculator>` supplies the concrete `IUnknown` implementation around your class, using `CComObjectRootEx` and the COM map.
- `CComCoClass` and the module's object map connect the class to its factory. `CAtlDllModuleT` supports the DLL entry points.
- `IDispatchImpl` uses type information to service Automation calls. The connection point container and `CProxy_*` helpers manage event subscriptions and delivery.
- The `.rgs` resource describes registry entries. `FinalConstruct` is an initialization hook that can return an `HRESULT`; `FinalRelease` is a cleanup hook before destruction.

**Choose it when:** you are building or maintaining a classic C++ COM server, especially one with Automation, events, or existing ATL code. It is a good default for extending this course's calculator.

**Setup:** install Visual Studio's optional **C++ ATL for latest build tools (x86 & x64)** component. You already did this for Stage 4.

**Support angle:** follow the failing operation to its owner: COM map for `E_NOINTERFACE`, `.rgs` and deployment for activation, connection point code for event subscriptions. You can step into ATL's headers in the debugger.

---

## 6.3 WRL — Windows Runtime C++ Template Library

### What it is for

**WRL provides C++ building blocks for COM and WinRT code without depending on ATL.** It was introduced with Windows 8 for low-level WinRT development, but it also supports classic COM.

You will often encounter it in graphics and media clients that use `ComPtr` to own interface pointers, or in code implementing an interface that a Windows API calls back. A callback object is still a COM object even if you create it directly and never register it.

### What makes it different

Like ATL, WRL can implement `IUnknown` for you. Unlike ATL, it does not provide comparable ready-made implementations of classic Automation features such as dual-interface `IDispatch` and connection points. Building a calculator with the scripting and events support from Module 5 would therefore require more work in WRL.

**WRL can also build DLL and EXE servers.** Its `Module` helpers support factories and hosting; it does not stop at implementing individual objects. The reason to prefer ATL for the course's server is convenience, not an inability to build it with WRL.

WRL keeps the underlying COM calling style visible: interface pointers, output parameters, and `HRESULT` checks. It also supports WinRT-specific concepts. Those are separate from classic Automation events.

### Recognize it in a project

Three names are enough to start:

- `ComPtr<ICalculator>` owns an interface pointer and releases it when the smart pointer no longer holds it. Using this alone does not make your application a WRL server.
- `RuntimeClass<RuntimeClassFlags<ClassicCom>, ICalculator>` is a base class that supplies the `IUnknown` machinery. You still implement **every** method required by the interface, including inherited methods.
- `Make` and `MakeAndInitialize` create your implementation directly. That differs from `CoCreateInstance`, which asks COM to activate a registered class.

`ClassicCom` selects ordinary COM support. WinRT objects generally also implement `IInspectable`, an extension of `IUnknown` that exposes type information. Do not add WinRT's `InspectableClass` macro to a `ClassicCom` implementation.

**Choose it when:** existing native code already uses WRL, or you need a COM interface implementation or smart pointer without ATL's broader infrastructure. For example, implementing a small callback in an existing WRL-based media client is a natural fit. For new WinRT development, start with C++/WinRT instead.

**Setup:** WRL headers ship with the Windows SDK. Include `<wrl/client.h>` for `ComPtr`, or `<wrl.h>` for the broader library. Required link libraries depend on the Windows APIs you call.

**Support angle:** distinguish pointer ownership from object implementation. Seeing `ComPtr` in a caller tells you nothing about which framework built the server, or whether the object may be used from another apartment.

---

## 6.4 WIL — Windows Implementation Library

### What it is for

**WIL makes everyday Windows C++ code easier to clean up and diagnose.** It contains helpers for COM pointers, strings, handles, locks, and errors. It is not a replacement for ATL's COM server infrastructure.

Its central idea is **RAII (Resource Acquisition Is Initialization)**: a C++ object owns a resource and releases it in its destructor. Cleanup then happens on normal returns, early error returns, and exception unwinding.

### What makes it different

ATL and WRL help you implement objects; WIL's main role here is managing resources and failures **inside either clients or servers**. You can add WIL to your ATL calculator without changing its interfaces, registration, or clients.

WIL supports both exceptions and `HRESULT`-based error handling. Choose the helpers that match your function's policy. A function returning `HRESULT` must not accidentally let a C++ exception escape through the COM interface.

### Recognize it in a project

- `wil::com_ptr` and `wil::com_ptr_nothrow` manage COM interface references; their fallible helper operations use different error policies.
- `wil::unique_bstr` frees a `BSTR`; `wil::unique_variant` clears a `VARIANT`. These prevent the ownership mistakes from Module 4.
- `RETURN_IF_FAILED(expression)` evaluates an `HRESULT` expression and returns a failure from the current function. `THROW_IF_FAILED(expression)` turns a failure into a C++ exception instead.

For example, a client calling `Describe` needs to free the returned `BSTR`, even if a later operation fails. A `wil::unique_bstr` can own that string. WIL changes the cleanup code, not how the calculator implements `Describe`.

**Choose it when:** your C++ code needs consistent resource ownership or error reporting. Use it alongside the object framework that suits the component. Keeping existing `CComPtr` or WRL `ComPtr` usage is also reasonable; changing smart pointer libraries is not a goal by itself.

**Setup:** add the [WIL package](https://github.com/microsoft/wil) through vcpkg (`wil`) or NuGet (`Microsoft.Windows.ImplementationLibrary`). Do not assume it is installed with the Windows SDK.

**Support angle:** WIL failure reporting can capture the file, line, and `HRESULT` where a helper observes a failure. Configure a logging callback to collect that information. It identifies your failing call site, not necessarily the root cause inside a third-party server.

---

## 6.5 C++/WinRT: a Windows Runtime projection

### What it is for

**C++/WinRT is the usual starting point for new C++ code that consumes or implements WinRT APIs.** It is a *language projection*: a library that presents the underlying COM-based API as convenient C++ types and methods.

For example, a WinRT property getter can appear as a method returning a string, instead of an `HRESULT` plus a string output parameter. You are still calling the same Windows API; the projection handles the conversion.

### What makes it different

Both WRL and C++/WinRT are C++ libraries. The important difference is the **level at which you work**, not whether one uses standard C++:

- WRL leaves most COM-style calls visible, including interface pointers and explicit `HRESULT` checks.
- C++/WinRT's projected WinRT APIs manage those details, return values naturally, translate failures into exceptions, and support `co_await` for asynchronous operations.

At the binary boundary, COM still uses interfaces, reference counting, and `HRESULT`s. C++ exceptions must not cross that boundary. The projection adapts the two sides.

Unlike ATL, C++/WinRT's primary purpose is not building classic Automation servers. Unlike WIL, it supplies a language view of WinRT APIs, not just resource helpers.

### Recognize it in a project

`winrt::Windows::...` types are projected Windows APIs; `winrt::hstring` is an owning WinRT string type. For components, `winrt::implements` supplies interface implementation machinery, and the C++/WinRT tools generate code from WinRT metadata.

The library also provides **classic COM helpers** in `<winrt/base.h>`. For example, `winrt::com_ptr<ICalculator>` can own a pointer to your existing calculator, and `winrt::implements` can implement ordinary `IUnknown`-based interfaces. This does **not** convert the calculator into a WinRT API: calls through a raw `ICalculator` interface still follow that interface's original `HRESULT` contract.

**Choose it when:** you are writing new C++ WinRT code, or adding classic COM interactions to a project already using C++/WinRT. Do not rewrite a working ATL Automation server just because C++/WinRT is newer. Microsoft's recommendation to prefer it over WRL concerns WinRT development, not a requirement to replace every WRL smart pointer.

**Setup:** for new projects, use the `Microsoft.Windows.CppWinRT` NuGet package and appropriate project tooling. The Windows SDK also includes headers. C++/WinRT is based on C++17; use C++20 for the VS 2026 toolchain used in this course to avoid its deprecated coroutine-header path. Link requirements depend on the APIs used.

**Support angle:** a projected exception can carry the underlying `HRESULT`. Recover that code and investigate the COM or WinRT operation that failed; the C++ surface does not remove activation or apartment requirements. See [Appendix B.2](appendix-b-com-plus-and-winrt.md#b2-winrt--com-with-new-rules) for the WinRT model.

---

## 6.6 .NET interop: RCW and CCW

### What it is for

**.NET interop lets managed code and COM call each other.** You do not choose ATL or WRL for a C# client. Instead, .NET adapts the COM interface to managed code. You already used this when the C# client in Module 5 called your ATL calculator.

With .NET's built-in COM interop, two wrappers do the adapting:

```text
C# caller  -> RCW (Runtime Callable Wrapper) -> COM object
COM caller -> CCW (COM Callable Wrapper)    -> C# object
```

The **RCW** holds COM interface references and translates calls and data. For example, a returned `BSTR` can appear as a C# `string`. The **CCW** exposes a managed object's declared COM interfaces and keeps the managed object alive while native clients hold references.

### What makes it different

Native COM uses reference counting; .NET uses **garbage collection (GC)** to reclaim unreachable managed objects. The wrapper connects these lifetime systems, but it does not make them identical. Setting a C# variable to `null` is not an immediate COM `Release`.

**Two C# variables can refer to the same wrapper.** Suppose `calculator` already refers to your COM calculator:

```csharp
dynamic otherCalculator = calculator;
```

You now have **two variables, but still only one RCW and one COM object**. The assignment copies a managed reference; it does not call COM `AddRef` or increase the RCW's internal count.

Setting `otherCalculator = null` removes only that variable's reference; `calculator` remains usable. But if `Marshal.ReleaseComObject(otherCalculator)` reduces the shared RCW's internal count to zero, it disconnects the wrapper from COM. Calls through **either** variable then fail. Manual release acts on the wrapper, not just the variable passed to it.

Sharing can also happen when separate COM calls return the same object. .NET recognizes that object through its **COM identity**: asking its different interfaces for `IUnknown` must return the same identity pointer. This lets the runtime reuse a wrapper instead of treating each returned interface as a new object.

### Recognize it in a project

**On the client side**, a COM reference supplies managed interface declarations generated from a type library. This is *early binding*: the compiler knows the methods. A C# `dynamic` call can instead use `IDispatch` at runtime, as in Module 5. .NET can also call ordinary `IUnknown`-based interfaces when it has matching declarations; COM interop does not always require Automation.

An **interop assembly** contains these managed declarations, not the COM server. A **PIA (Primary Interop Assembly)** is the vendor's official interop assembly.

**Embed Interop Types is a setting on an individual reference in your C# project.** In Visual Studio's **Solution Explorer**, select the COM or interop reference under **Dependencies** (or **References** in older projects), then press **F4** to open its Properties window. Find **Embed Interop Types**:

- **True:** the compiler copies the supported interop type definitions your code uses into your client assembly. For those types, you no longer need to deploy a separate interop DLL.
- **False:** your client depends on the interop assembly, so that DLL must also be available at runtime.

In the project file, this is `<EmbedInteropTypes>true</EmbedInteropTypes>` inside the relevant `COMReference` or `Reference` item, not a project-wide `PropertyGroup` setting.

It embeds **type definitions, not the COM server**. The server still needs to be deployed, with registration and marshaling support where required. Some uses, such as the typed event-helper classes from Lab 5.2, still need the interop assembly.

**On the server side**, use explicit interfaces with stable GUIDs and `[ClassInterface(ClassInterfaceType.None)]`. This prevents the public members of a C# class from automatically becoming an unstable COM contract. For Automation interfaces, assign explicit DISPIDs too.

**Choose it when:** the client or server belongs in a managed application. Keep the language that fits the application; COM allows the other side to use a different one. Account for runtime deployment and nondeterministic cleanup when making that choice.

### Setup and deployment

- **.NET Framework servers:** use its `regasm` tool. It can also generate a type library with `/tlb`.
- **Modern .NET servers:** enable `EnableComHosting`, build, and register the generated `*.comhost.dll`, not the managed DLL. The matching .NET runtime must be installed; this hosting path is framework-dependent.
- **Modern .NET type libraries:** `EnableComHosting` does **not** generate one. Supply a matching IDL/type library when needed, or use explicit native declarations as in Lab 6.2.

Keep the native host and caller architecture aligned. A 64-bit .NET assembly accompanied by a 64-bit COM host cannot be loaded into a 32-bit client.

### Lifetime: do not treat an RCW as a smart pointer

`Marshal.ReleaseComObject` decrements the **RCW's internal count**, not the number of C# variables or the native object's reference count. At zero, it releases the COM interface references held by that wrapper. `FinalReleaseComObject` forces that wrapper count to zero. Neither guarantees destruction if other native references remain.

Manually disconnecting a shared wrapper can break another caller with `InvalidComObjectException`. Avoid either API unless you have a specific cleanup requirement and can prove exclusive ownership, including that no calls are in progress. Prefer the component's documented shutdown operations, unsubscribe events, and let the runtime manage ordinary wrapper cleanup. Forced garbage collection is a diagnostic experiment, not a routine shutdown strategy.

### Advanced option: source-generated interop

`ComWrappers` gives advanced code explicit control over wrappers. Starting in .NET 8, `[GeneratedComInterface]` can generate much of the interop code at build time, supporting **Native AOT** (ahead-of-time compilation) and **trimming** (removing unused code).

This is not a universal replacement for the examples above: source-generated COM supports `IUnknown`-based interfaces, not dual interfaces or `IDispatch`. Choose it when deployment requirements call for it and the contract is supported. The Automation labs in this course use built-in COM interop.

**Support angle:** first separate activation failures, interface mismatches, and lifetime failures. For `0x80040154`, check registration and bitness; for `E_NOINTERFACE`, check the requested IID and marshaling path; for `InvalidComObjectException`, look for premature manual release.

---

## 6.7 Choosing a framework

Start with the **API contract and the job**, not which library is newest.

1. **Are you implementing an object or calling one?** A caller does not need to use the server's framework. Adding a smart pointer library does not change the server.
2. **Is the API classic COM or WinRT?** Existing IIDs and type libraries describe the classic contract; WinRT metadata and projected Windows types point toward WinRT tooling. Your clients' requirements decide the contract.
3. **What does your project already use?** Extending a working ATL, WRL, or C++/WinRT component usually costs less than changing frameworks. Check target-app restrictions and dependencies; packaging alone does not decide the answer.
4. **What additional constraints matter?** Automation, callbacks, exception policy, threading, deployment, and Native AOT can change the recommendation.

| Your task | Starting choice | Why / main caveat |
|---|---|---|
| Build a classic C++ server with scripting and events | **ATL** | Integrated Automation, connection points, hosting, and registration support; you still own the contract and threading. |
| Implement a small native COM callback in SDK-based code | **WRL**, or the project's existing object framework | `RuntimeClass` handles `IUnknown` without requiring the full ATL server setup. |
| Call classic COM from C++ | **Existing smart pointer library; consider WIL for new resource/error handling** | ATL, WRL, WIL, and C++/WinRT pointers can all own classic COM references. Keep one consistent ownership style. |
| Consume or author WinRT APIs in new C++ code | **C++/WinRT** | Projected APIs reduce low-level calling code; learn WinRT's contracts and tooling. |
| Call COM from C# | **Built-in .NET interop** for ordinary Windows Automation use | Use a COM reference or `dynamic`; understand RCW sharing and deployment. |
| Expose an existing C# library to native callers | **Explicit COM interfaces and modern .NET COM hosting** | Reuse managed logic, but deploy its runtime and supply any required type library. |
| Need Native AOT or trimming with .NET COM interop | **Evaluate source-generated COM** | Suitable for supported `IUnknown` contracts, not a drop-in replacement for `IDispatch`. |

### Apply it to the calculator

**Keep ATL for the Module 5 server.** Its clients already depend on Automation and events, which ATL supports directly. Moving it to WRL would add implementation work without improving that contract.

**Choose client tools independently.** A C++ client can keep `CComPtr` or use WIL for cleanup. A C# client can keep its COM reference. Neither choice requires rebuilding the server.

**Add C++/WinRT only for a relevant job**, such as calling a WinRT API from the application. Its presence does not mean the calculator itself must become a WinRT component.

Before committing to a new framework, make a small prototype of the hardest requirement: an event subscription, cross-apartment call, or clean-machine activation. Successful compilation alone does not establish that the deployment and lifetime design works.

---

## 6.8 LAB 6.1 — The hand-written server vs the ATL one

> **Requirements**
> - **Tools:** Visual Studio with the optional component **C++ ATL for latest build tools (x86 & x64)**. It is *not* installed by default with *Desktop development with C++* — if you have no *ATL Project* template, that is why. Add it in the VS Installer.
> - **Elevation:** none. Both servers are already registered from Modules 2 and 5.
> - **Bitness:** x64.
> - **Depends on:** the hand-written Module 2 server — kept, not deleted — and the ATL server from Lab 5.1. The comparison table is the deliverable.
> - **Starting point:** [`labs/stage-2-inproc-server/`](../labs/stage-2-inproc-server/) for the "before", and your `TrainingCalc` project for the "after".
> - **Time:** ~1 h, using existing Debug builds.

Open the hand-written server and your existing `TrainingCalc` project. Do not recreate the ATL server. The goal is to locate the code responsible for the same COM duties in each implementation.

> If you skipped Module 5, first complete the server setup in [`labs/stage-4-atl-server/README.md`](../labs/stage-4-atl-server/README.md).

### Open the two implementations

Keep [Stage 2's Calc.cpp](../labs/stage-2-inproc-server/Calc.cpp) open beside your ATL project. Nearly all of the manual server is in that one file. In the ATL project, the responsibilities are spread across the calculator class, module class, and registration resources. Use **Ctrl+Shift+F** to find the symbols below; use **F12** on an ATL class to inspect its definition in the installed headers.

The manual class is named `Calculator`; your ATL implementation is `CCalculator`. Their similar names do not mean they have the same IID or CLSID. Compare responsibilities, not identifier values.

### 1. Interface discovery: a function becomes a map

**Hand-written:** find `Calculator::QueryInterface`. It checks the requested IID, selects the correct interface pointer, calls `AddRef` on success, and returns `E_NOINTERFACE` for unsupported interfaces.

**ATL:** find `BEGIN_COM_MAP(CCalculator)` in your calculator header. Each `COM_INTERFACE_ENTRY` declares an interface the object exposes. ATL's `QueryInterface` implementation searches that map and supplies the pointer and reference-count handling.

**Notice the difference:** you no longer write the search, but you still specify the supported interfaces. Inheriting an interface is not enough to make ATL return it from `QueryInterface`; it also needs an appropriate map entry. You will not see an explicit `IUnknown` entry in this map because ATL handles it specially, using the first interface entry for the object's identity.

**Check your understanding:** which entry makes `QueryInterface(IID_IDispatch)` succeed in the ATL server, and why would that query fail in Stage 2?

### 2. Object lifetime: a counter becomes a template implementation

**Hand-written:** find `m_cRef`, `Calculator::AddRef`, and `Calculator::Release`. You wrote the interlocked increments and decrements, including `delete this` at zero.

**ATL:** find the `CComObjectRootEx<CComMultiThreadModel>` base class, then use **F12** to inspect `CComObject` in ATL's headers. For this non-aggregated server, ATL creates a `CComObject<CCalculator>`: a concrete class derived from your `CCalculator` that supplies `QueryInterface`, `AddRef`, and `Release`. The root class supplies the reference-count machinery; the thread model selects interlocked operations.

**Notice the difference:** the ownership rule has not changed; the code enforcing it moved into ATL. Use the existing `FinalConstruct` and `FinalRelease` hooks for initialization and cleanup rather than adding your own `AddRef` or `Release` overrides. `FinalConstruct` can return an `HRESULT` if initialization fails. `CComMultiThreadModel` protects the counter, not your calculator's fields.

**Check your understanding:** when a client releases its last reference, which code decides to delete the object, and where can you observe its final cleanup in your own class?

### 3. Object creation: your factory becomes ATL's factory

**Hand-written:** find `CalculatorFactory::CreateInstance`. Follow its sequence: reject aggregation, allocate a calculator, query the requested interface, and release the temporary construction reference.

**ATL:** find `CComCoClass<CCalculator, &CLSID_Calculator>`, `DECLARE_NOT_AGGREGATABLE(CCalculator)`, and `OBJECT_ENTRY_AUTO`. Together they connect the class to ATL's factory and creation machinery. The default factory is `CComClassFactory`; the creation path uses `CComCreator` to construct and initialize the object.

**Notice the difference:** you no longer implement `IClassFactory` yourself. You declare which class can be created and whether it supports aggregation; ATL handles allocation, initialization, the requested interface, and failure cleanup. `OBJECT_ENTRY_AUTO` adds the class to an **in-memory object map**, not the Windows registry.

**Check your understanding:** identify where the manual server rejects aggregation and the ATL declaration that selects the equivalent policy.

### 4. DLL exports: hand-written decisions become module calls

**Hand-written:** find `DllGetClassObject` and `DllCanUnloadNow`. The first checks a CLSID and returns `g_factory`. The second checks `g_cObjects` and `g_cLocks`. Also find where object construction/destruction and the factory's `LockServer` change those counters.

**ATL:** search your project for the same export names. Their bodies delegate to `_AtlModule`. Find `CTrainingCalcModule`, derived from `CAtlDllModuleT`, and the `OBJECT_ENTRY_AUTO` entry from the previous step. The module uses its object map to locate a factory and its lock count to decide whether the DLL can unload.

**Notice the difference:** the exports still exist because COM requires them. ATL replaces their implementation, not the entry points. Its object and factory machinery maintains the module's locks instead of your two global counters. An object's reference count and the module's unload-prevention count are different things.

**Check your understanding:** why can releasing one calculator be insufficient to make `DllCanUnloadNow` return `S_OK`?

### 5. Registration: registry API calls become resource data

**Hand-written:** find `DllRegisterServer`, `SetKeyValue`, and `DllUnregisterServer`. Follow how they write or remove the CLSID, `InprocServer32`, `ThreadingModel`, and ProgID entries.

**ATL:** find the calculator's `.rgs` resource and `DECLARE_REGISTRY_RESOURCEID` in its class. The declaration connects the class to that resource; ATL's registration code interprets it when the module registers the server. Find `InprocServer32` and `%MODULE%` in the script: `%MODULE%` is replaced with the DLL's path.

**Notice the difference:** you describe the entries instead of writing the registry API calls. You still own their correctness. This is why leaving the wizard's ProgID blank could produce a registered CLSID without a usable name in Stage 4. The object map enables factory lookup inside the DLL; registration enables COM to find that DLL in the first place.

**Check your understanding:** find where each server declares `ThreadingModel = Both`, and where each associates its ProgID with its CLSID. Inspect only; do not copy GUIDs between the servers or register them again.

### 6. Automation and events: additions, not replacements

**Hand-written:** inspect [Stage 2's interface](../labs/stage-2-inproc-server/Calculator.h). It derives from `IUnknown` and declares only `Add` and `Subtract`. There is no `IDispatch` or connection point implementation to replace.

**ATL:** find `IDispatchImpl` in the calculator's base classes and the matching interface definition in the project's IDL. For events, find `IConnectionPointContainerImpl`, `BEGIN_CONNECTION_POINT_MAP`, and `CProxy_ICalculatorEvents`. In the event proxy header, inspect the `Fire_OnProgress` helper you added in Stage 4.

**Notice the difference:** Stage 4 added capabilities as well as removing boilerplate. `IDispatchImpl` services Automation calls using type information. The connection point classes manage subscriptions; the `Fire_*` helpers call the sinks. You still declare the methods and events, implement their behavior, and decide when to fire an event. The wizard did not write `SumTo` or your event-delivery helpers for you.

**Check your understanding:** locate the `Add` method body in both implementations. The pointer check and arithmetic are still your code. Which ATL base class lets a script reach that same method through `Invoke`?

### Follow one activation in the debugger

1. Run your existing native client under the debugger using the ATL server's **Debug** build. Put a breakpoint in `CCalculator::FinalConstruct` before activation. If it stays unbound, check **Debug > Windows > Modules** after the DLL loads: the loaded path must be your Debug DLL and its matching symbols must be loaded.
2. When the breakpoint hits, open **Call Stack** and inspect the ATL callers leading into your hook. Look for the creation machinery from step 3. Disable **Just My Code** if it hides those frames. You do not need to understand every template parameter.
3. Add a function breakpoint for `ATL::CComObjectRootBase::InternalQueryInterface`, then restart the client. Inspect the requested IID and the `_ATL_INTMAP_ENTRY` map. Expect several queries during activation, not just one call for `ICalculator`.
4. Break in `CCalculator::Add` and `FinalRelease`, then let the client run to completion. Observe where control reaches your method and where ATL performs final cleanup after the client releases its references.

**Deliverable:** make a comparison table with three columns: **responsibility**, **manual function or field**, and **ATL counterpart / what I still supply**. Use steps 1–6 as the rows; mark Automation and events as absent in Stage 2. The goal is to explain the change in responsibility, not count lines or memorize the ATL headers.

<details>
<summary>Answers to the six "Check your understanding" questions</summary>

1. **Interface discovery:** `COM_INTERFACE_ENTRY(IDispatch)` exposes the `IDispatch` implementation inherited through `IDispatchImpl`. Stage 2 implements only its `IUnknown`-based `ICalculator`; its `QueryInterface` recognizes only `IID_IUnknown` and `IID_ICalculator`, so requesting `IID_IDispatch` returns `E_NOINTERFACE`.

2. **Object lifetime:** in Stage 2, `Calculator::Release` decrements the counter and deletes the object at zero. In ATL, `CComObject<CCalculator>::Release` makes that decision using the inherited reference-count machinery. Observe your cleanup in `CCalculator::FinalRelease`. A client's last reference is not necessarily the object's last reference: another client or owner may still keep it alive.

3. **Object creation:** `CalculatorFactory::CreateInstance` rejects a non-null `pUnkOuter` with `CLASS_E_NOAGGREGATION`. The corresponding ATL policy is `DECLARE_NOT_AGGREGATABLE(CCalculator)`, which selects creation machinery that rejects aggregation too.

4. **DLL lifetime:** the DLL serves more than one reference or object. The calculator may still have other references, another calculator may exist, or a factory reference or unmatched `LockServer(TRUE)` may keep the module in use. Stage 2 requires both `g_cObjects` and `g_cLocks` to be zero; ATL checks its module lock count. Destroying one object does not prove the whole DLL is unused.

5. **Registration:** Stage 2's `DllRegisterServer` writes `ThreadingModel = Both` under `CLSID\{...}\InprocServer32` through `SetKeyValue`. It writes `Training.Calculator.1` under the CLSID's `ProgID` key and the reverse mapping under `Training.Calculator.1\CLSID`. In ATL, find the equivalent entries in the calculator's `.rgs` resource: `val ThreadingModel = s 'Both'`, the CLSID's `ProgID`, and `TrainingCalc.Calculator.1\CLSID`. Each server uses its own CLSID and ProgID.

6. **Automation:** Stage 2's `Calculator::Add` is inline in [Calc.cpp](../labs/stage-2-inproc-server/Calc.cpp); the ATL project's method is `CCalculator::Add`. Both contain your method logic. The `IDispatchImpl<ICalculator, ...>` base class supplies `Invoke` and uses type information to dispatch a scripting call to that implementation. It does not supply the arithmetic itself.

</details>

---

## 6.9 LAB 6.2 — Cross-language interop, both directions

> **Requirements**
> - **Tools:** Visual Studio with C++ and ATL, the **x64 .NET 10 SDK**, and the x64 .NET 10 runtime. No .NET Framework developer pack is needed for this path.
> - **Elevation:** required only for registering and unregistering the new .NET COM host.
> - **Bitness:** x64 for both clients and servers.
> - **Depends on:** the ATL server from Lab 5.1 for Direction 2.
> - **Time:** ~1–2 h.

This lab tests a framework choice in both directions. First, expose a small C# object to C++. Then reuse your C# clients against the ATL server. The calls follow COM contracts even though deployment and lifetime differ.

### Direction 1: C# server, C++ client

Here, **C# implements the calculator and C++ calls it through COM**. The calculator is a managed object; .NET exposes its `INetCalculator` interface through a **COM Callable Wrapper (CCW)**:

```text
C++ client -> INetCalculator on the CCW -> C# NetCalculator.Add
```

The native COM host enables activation; the CCW forwards method calls and keeps the managed object alive while C++ holds COM references. Everything runs inside the C++ client's process, not in a separate C# application.

**1. Create the managed server.** In Visual Studio, create a **Class Library (C#)** project named `NetCalc`, targeting .NET 10. Replace the project file's contents with:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0-windows</TargetFramework>
    <EnableComHosting>true</EnableComHosting>
    <PlatformTarget>x64</PlatformTarget>
    <RuntimeIdentifier>win-x64</RuntimeIdentifier>
  </PropertyGroup>
</Project>
```

Setting `EnableComHosting` to `true` makes the build generate `NetCalc.comhost.dll`, a **native COM host** alongside your managed DLL. COM loads this host to start .NET as needed and create your COM-visible C# objects. It does not register the server or generate a type library; registration is step 2.

Replace the generated class file's entire contents with:

```csharp
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid("182de894-9f6f-478e-a6ab-30c113a6da9a")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface INetCalculator
{
    int Add(int left, int right);
}

[ComVisible(true)]
[Guid("1e0e9bf9-7075-4933-8918-3c5237764cc5")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(INetCalculator))]
public class NetCalculator : INetCalculator
{
    public int Add(int left, int right) => left + right;
}
```

This is a **new, separate interface**, not the Automation `ICalculator` from Module 5. It derives from `IUnknown`, so it does not promise scripting or events. The native signature is `HRESULT Add(LONG left, LONG right, LONG* result)`: built-in interop maps the C# return value to the final output parameter.

**2. Build and register.** From the directory containing the project file, run:

```powershell
dotnet build -c Debug
```

Keep the output files together, including the managed DLL, COM host, runtime configuration, and dependencies file. In an **elevated x64 PowerShell** opened at the same project directory, run:

```powershell
& "$env:WINDIR\System32\regsvr32.exe" ".\bin\Debug\net10.0-windows\win-x64\NetCalc.comhost.dll"
```

**3. Create the native client.** Add a **Console App (C++)** project named `Client_NetCalc`, select **x64**, and replace its generated `.cpp` file with the complete program below. Use the default precompiled-header-free console template; if the project enables precompiled headers, disable them for this file.

```cpp
#include <Windows.h>
#include <atlbase.h>
#include <cstdio>

MIDL_INTERFACE("182de894-9f6f-478e-a6ab-30c113a6da9a")
INetCalculator : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Add(LONG left, LONG right, LONG* result) = 0;
};

class DECLSPEC_UUID("1e0e9bf9-7075-4933-8918-3c5237764cc5") NetCalculator;

HRESULT CallCalculator()
{
    CComPtr<INetCalculator> calculator;
    HRESULT status = calculator.CoCreateInstance(__uuidof(NetCalculator));
    if (FAILED(status)) return status;

    LONG result = 0;
    status = calculator->Add(2, 3, &result);
    if (SUCCEEDED(status)) std::printf("Add(2,3) = %ld\n", result);
    return status;
}

int main()
{
    HRESULT status = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(status))
    {
        std::printf("CoInitializeEx failed: 0x%08lX\n", static_cast<unsigned long>(status));
        return 1;
    }

    status = CallCalculator();
    CoUninitialize();
    if (FAILED(status))
        std::printf("Calculator call failed: 0x%08lX\n", static_cast<unsigned long>(status));
    return FAILED(status) ? 1 : 0;
}
```

Add `ole32.lib` to **Linker > Input > Additional Dependencies**. Build and run `Client_NetCalc`. Expected output: `Add(2,3) = 5`.

There is no `#import`: modern .NET did not generate a type library. The explicit C++ declaration matches the C# interface's IID, method order, and native signature. This lab calls it within the same apartment; crossing apartments or processes with this custom interface would also need marshaling support.

**4. Clean up this registration** after closing the client. Run in the same elevated shell before moving or deleting the build output:

```powershell
& "$env:WINDIR\System32\regsvr32.exe" /u ".\bin\Debug\net10.0-windows\win-x64\NetCalc.comhost.dll"
```

### Direction 2: C++ server, C# client

Here, **C++ implements the calculator and C# calls it through COM**. This time, the server is your existing ATL `TrainingCalc` DLL from Module 5, not the managed `NetCalc` library from Direction 1. .NET gives the C# caller a **Runtime Callable Wrapper (RCW)**:

```text
C# client -> RCW -> native C++ CCalculator.Add
```

The ATL DLL loads inside the C# client's process. You will call it in two ways: **early binding**, where the compiler knows the interface, and **late binding**, where method names are resolved at runtime.

**1. Open or create the client and check its settings.** Reuse the `Client_CSharp` project from Lab 5.1, or create a new **Console App (C#)** project targeting .NET 10. Name the new project `Client_CSharp` if that name is not already used in your solution; otherwise use `Client_CSharp_Module6`.

Right-click the C# project and choose **Edit Project File**. Ensure its existing `PropertyGroup` contains these values; keep the rest of the project file:

```xml
<PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0-windows</TargetFramework>
    <PlatformTarget>x64</PlatformTarget>
</PropertyGroup>
```

A **new project has no `COMReference` item yet**, which is expected. Step 2 adds it through Visual Studio; you do not need to write it by hand. If you are **reusing the Lab 5.1 client**, preserve any existing `COMReference` item. In the remaining steps, use whichever client project you chose here.

Do **not** add `EnableComHosting`: this project calls COM; it does not expose a C# class as a COM server. There is no client DLL to register. Your x64 ATL server should still be registered as `TrainingCalc.Calculator.1` from Lab 5.1; do not rebuild or register it again for this exercise. Unregistering `NetCalc` at the end of Direction 1 does not unregister `TrainingCalc`.

**2. Add or check the COM reference for early binding.** In **Solution Explorer**, expand the C# project's **Dependencies** (or **References**) and look for `TrainingCalcLib`. If it is already present, reuse it; do not add a duplicate.

For a new project, or if the reference is missing, right-click the project, select **Add > COM Reference**, select **TrainingCalc 1.0 Type Library**, and confirm with **OK**. Visual Studio adds the `COMReference` item to your project file. This is the ATL server's registered type library, not the C# library from Direction 1. If it is absent from the list, check Stage 4's registration before continuing.

The build uses that type library to generate C# declarations such as `TrainingCalcLib.ICalculator`. You do not copy the C++ header into this project. Select the reference and press **F4** to inspect **Embed Interop Types**. Leave its current value: both `True` and `False` work for these calls.

**3. Replace the client program.** Replace the generated program file's **entire contents**, including any top-level statements or previous `Main`, with this complete program:

```csharp
using System;

class Program
{
        [STAThread]
        static void Main()
        {
                Console.WriteLine("Early bound");
                TrainingCalcLib.ICalculator earlyCalculator = new TrainingCalcLib.Calculator();
                Console.WriteLine($"Add(2,3) = {earlyCalculator.Add(2, 3)}");
                Console.WriteLine($"Describe() = {earlyCalculator.Describe()}");

                Console.WriteLine("Late bound");
                Type calculatorType = Type.GetTypeFromProgID("TrainingCalc.Calculator.1", true)!;
                dynamic lateCalculator = Activator.CreateInstance(calculatorType)!;
                Console.WriteLine($"Add(2,3) = {lateCalculator.Add(2, 3)}");
                Console.WriteLine($"Describe() = {lateCalculator.Describe()}");
        }
}
```

`[STAThread]` selects an STA for the main thread. .NET handles COM initialization for this managed entry point; do not copy `CoInitializeEx` from the native client. The `!` operators tell the C# nullable checker that the activation results are expected to be non-null; they do not suppress runtime errors.

This creates **two separate calculator instances from the same native server**, not two names for one RCW. No manual `ReleaseComObject` is needed for this short client; .NET manages the wrappers' cleanup.

**4. Build and run the C# client.** Right-click `Client_CSharp` and choose **Build**, then **Set as Startup Project**. Run with **Ctrl+F5**. Building only this project avoids rebuilding the ATL DLL and triggering its registration step.

Use **Visual Studio** for the build. This project has a COM reference, whose `ResolveComReference` task requires Visual Studio's .NET Framework MSBuild. `dotnet build` and `dotnet run` cannot perform that task, even though the application itself targets .NET 10. A late-bound-only project without a COM reference would not have that restriction.

Expected output with the Stage 4 implementation:

```text
Early bound
Add(2,3) = 5
Describe() = Training Calculator 1.0
Late bound
Add(2,3) = 5
Describe() = Training Calculator 1.0
```

If you kept Lab 5.2's lifetime tracing, additional `[source]` lines are normal. If you customized `Describe`, its text will reflect your implementation.

**5. Compare how the calls reach the server.** Hover over `earlyCalculator` in the editor: its declared type is `TrainingCalcLib.ICalculator`. The compiler can check that `Add` exists and that its arguments match; the RCW calls this dual interface through its vtable.

Hover over `lateCalculator`: its declared type is `dynamic`. The compiler does not check the calculator's members. At runtime, .NET resolves the call through `IDispatch`, which ATL's `IDispatchImpl` supplies. The COM reference is needed by the early-bound code, not by these `dynamic` calls.

Both routes reach the same **C++ method implementation**, and both use an RCW. The early-bound declarations provide type information to C#, not a second implementation of `Add`. A returned `BSTR` becomes a C# `string`, which is why `Describe()` prints naturally in both cases.

**If a step fails:**

- `TrainingCalcLib` or `ICalculator` is unknown at build time: check the reference in step 2 and confirm you selected the ATL calculator's type library.
- `MSB4803` mentions `ResolveComReference`: build the client from Visual Studio, not `dotnet build`.
- Activation fails with `0x80040154`: verify `PlatformTarget = x64` and the existing ATL server's registration and DLL location. Registering Direction 1's COM host does not fix this server's registration.
- `Describe()` returns unexpected text: confirm which ATL DLL is registered; both halves should use the same server build.

**Deliverable:** explain why the native client did not need WRL or C++/WinRT to call the C# server, and why the C# client did not need ATL to call the native server.

---

## 6.10 LAB 6.3 — Reproduce the RCW bugs

> **Requirements**
> - **Tools:** Visual Studio and .NET 10. Run without the debugger for the final lifetime observation; debugging can keep managed objects alive longer.
> - **Elevation:** none if Lab 5.1 is complete.
> - **Bitness:** x64, matching the registered server.
> - **Depends on:** a registered `TrainingCalc.Calculator.1` (Lab 5.1) with the Lab 5.2 lifetime tracing still compiled in.
> - **Starting point:** your ATL server with the existing `FinalConstruct` and `FinalRelease` prints. Do not add new `AddRef` or `Release` overrides to the ATL class.
> - **Time:** ~1 h.

Create a **Console App (C#)** named `Client_RcwLifetime`, target `net10.0-windows`, and set **Platform target = x64**. No COM reference is required because this client uses `dynamic`. Replace the generated program with the complete code below.

The first two cases deliberately release a COM wrapper and then try to call the calculator again: first through the original variable, then through another variable sharing that wrapper. Both calls should fail.

The third case lets garbage collection clean up the wrapper instead. `CreateAndDrop` creates and uses a calculator without returning it or storing a reference elsewhere. Once that method returns, its local variable no longer keeps the wrapper alive. `Main` then requests garbage collection so you can observe the cleanup; returning from the helper does not itself destroy the COM object.

The `[MethodImpl(MethodImplOptions.NoInlining)]` attribute keeps `CreateAndDrop` as a separate method call. Without it, .NET could copy the helper's code into `Main` as an optimization and keep the local reference alive longer than expected, making this experiment harder to interpret.

```csharp
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

class Program
{
    [STAThread]
    static void Main()
    {
        Type calculatorType = Type.GetTypeFromProgID("TrainingCalc.Calculator.1", true)!;

        dynamic calculator = Activator.CreateInstance(calculatorType)!;
        Console.WriteLine(calculator.Add(2, 3));
        Marshal.ReleaseComObject(calculator);
        try { Console.WriteLine(calculator.Add(4, 5)); }
        catch (InvalidComObjectException) { Console.WriteLine("Case 1: disconnected wrapper"); }

        dynamic original = Activator.CreateInstance(calculatorType)!;
        dynamic alias = original;
        Marshal.ReleaseComObject(original);
        try { Console.WriteLine(alias.Add(1, 1)); }
        catch (InvalidComObjectException) { Console.WriteLine("Case 2: alias also disconnected"); }

        CreateAndDrop(calculatorType);
        Console.WriteLine("Requesting collection for this experiment");
        GC.Collect();
        GC.WaitForPendingFinalizers();
        Console.WriteLine("Collection experiment complete");
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    static void CreateAndDrop(Type calculatorType)
    {
        dynamic calculator = Activator.CreateInstance(calculatorType)!;
        Console.WriteLine(calculator.Add(6, 7));
    }
}
```

Build **Release** and run with **Ctrl+F5**. Expect `5`, the two disconnected-wrapper messages, and `13`, together with your server's lifetime prints. The first two objects lose their wrapper-held references during manual release. For the third, observe the destruction print around collection rather than assuming method return immediately destroys the COM object. Exact timing can vary.

### Case 3: what happens under the hood

This case demonstrates **successful cleanup, not a third bug**. There are two objects with different lifetime rules: the managed **RCW**, tracked by .NET's garbage collector, and the native **C++ calculator**, kept alive by COM references.

While `CreateAndDrop` is using the calculator, the relationship is:

```text
C# local variable -> RCW -> COM interface references -> native calculator
```

Follow the third case's output:

```text
[source] constructed
13
Requesting collection for this experiment
[source] destroyed
Collection experiment complete
```

1. **The helper creates and uses the object.** `Activator.CreateInstance` activates the native calculator and returns its RCW to C#. The wrapper holds COM interface references that keep the calculator alive. Calling `Add(6, 7)` prints `13`.

2. **The helper returns, but that does not call COM `Release`.** It did not return or store the calculator, so no application reference keeps this RCW reachable. *Reachable* means .NET can still find the object through a live variable or another reachable object. The wrapper is now eligible for cleanup, but it can still hold its native COM references until that cleanup happens. In this run, the lack of a destruction print before `Requesting collection` shows that the native calculator is still alive after the helper returns.

3. **The program requests collection and waits for finalization.** `GC.Collect()` asks .NET to find unreachable managed objects. `GC.WaitForPendingFinalizers()` waits for pending finalization work, which can run separately from collection. These calls make runtime cleanup observable in this experiment; they do not tell COM to destroy every object in the process.

4. **Wrapper cleanup releases the native references.** When .NET cleans up the RCW, it calls COM `Release` on the interface references the wrapper holds. In this run, that removes the calculator's last references. ATL's `CComObject<CCalculator>::Release` sees a zero count and deletes the object; its destruction calls your `FinalRelease` hook, producing `[source] destroyed`. `Main` then prints `Collection experiment complete`.

**The garbage collector does not directly delete the C++ calculator.** The sequence is: the wrapper becomes unreachable, .NET cleans up its COM references, and COM reference counting allows ATL to destroy the calculator. If another native owner still held a reference, cleaning up this RCW would not destroy it.

Without the explicit GC calls, .NET can perform this cleanup during a later collection, but there is no promise about when. A collection can also happen earlier than the one requested here, so the exact output ordering is an observation, not a lifetime guarantee. Cases 1 and 2 forced a release while C# still wanted to use the wrapper; case 3 stops using it and lets .NET manage the release.

**Deliverable:** explain why the alias does not protect the second object and why the last case uses forced collection only to observe runtime cleanup. Remove the explicit collection calls to see ordinary nondeterministic lifetime; do not adopt them as application shutdown code.

---

## 6.11 Checkpoint

1. A C++ server needs PowerShell callers and connection-point events. What would you choose, and why?
2. A graphics client already uses WRL `ComPtr`. Does that tell you how its COM server was implemented? Must you replace those pointers with WIL?
3. Can WRL build a registered COM server? What would ATL save you when implementing the Module 5 calculator?
4. How does C++/WinRT differ from WRL for WinRT calls? Does `winrt::com_ptr<ICalculator>` turn a classic calculator into a projected WinRT API?
5. Can ATL and WIL be used together? Give one responsibility for each.
6. Two C# variables share an RCW. Why can calling `ReleaseComObject` through one break the other? Does a return value of zero prove the native object was destroyed?
7. A .NET 10 COM server built successfully, but its author cannot find a generated TLB. Is that a build failure? What gets registered?
8. A C# Automation client must move to Native AOT. Is adding `[GeneratedComInterface]` enough?

<details>
<summary>Answers</summary>

1. ATL is a good starting choice because it provides type-library-based `IDispatch`, connection points, and classic server infrastructure. The interface and event design remain your responsibility.

2. No to both. The smart pointer owns a client-side reference; the server can use any compatible implementation. Keep a working ownership convention unless changing it solves a concrete problem.

3. Yes. WRL has factory and hosting support. ATL saves the work of supplying classic Automation and connection-point implementations and integrating registration.

4. WRL keeps ABI-style calls visible. C++/WinRT projects WinRT APIs into C++ values, methods, exceptions, and asynchronous operations. Its classic `com_ptr` is an ownership helper, not a converter of arbitrary COM interfaces into WinRT APIs.

5. Yes. ATL can provide the object's COM map and `IDispatch`; WIL can own strings or handles and report errors inside its methods.

6. Both variables use the same wrapper. Disconnecting it invalidates access through either variable. Zero is the wrapper's internal count, not the native object's count; other native clients may still hold references.

7. No. Modern .NET COM hosting does not generate a type library from the assembly. Register the generated `*.comhost.dll`; supply matching IDL/type information separately if the clients need it.

8. No. Built-in COM interop is not supported by Native AOT, and source-generated COM does not support `IDispatch` or dual interfaces. The interop design needs reassessment, not just a new attribute.

</details>

---

## 6.12 Rules to carry forward

1. Choose by contract, job, and existing code, not library age. Client and server choices are independent.
2. ATL is a strong default for classic C++ Automation servers; WRL remains useful for native COM building blocks; C++/WinRT is the starting point for new C++ WinRT work.
3. WIL complements these choices with resource and error helpers. Do not introduce multiple smart pointer styles without a reason.
4. No framework makes arbitrary object state thread-safe or removes apartment and marshaling rules.
5. Managed wrappers connect two lifetime systems. Understand sharing before forcing a release.
6. Keep explicit, stable COM interfaces and verify deployment. Modern .NET hosting, type libraries, and source-generated interop are separate decisions.

### Further reading

- [ATL overview](https://learn.microsoft.com/en-us/cpp/atl/atl-overview?view=msvc-170)
- [WRL overview and comparison with ATL](https://learn.microsoft.com/en-us/cpp/cppcx/wrl/windows-runtime-cpp-template-library-wrl?view=msvc-170)
- [WIL documentation](https://github.com/microsoft/wil/wiki)
- [Introduction to C++/WinRT](https://learn.microsoft.com/en-us/windows/uwp/cpp-and-winrt-apis/intro-to-using-cpp-with-winrt)
- [.NET COM hosting](https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com) and [source-generated COM limitations](https://learn.microsoft.com/en-us/dotnet/standard/native-interop/comwrappers-source-generation)

---

**Next: [Module 7 — DCOM, security, and out-of-proc servers](07-dcom-and-security.md)**
