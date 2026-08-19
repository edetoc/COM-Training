# COM Ramp-Up Training Plan

> Audience: Developer **and** Support/Escalation Engineer, starting from beginner.
> Format: 9 progressive modules. Each has **Concepts → Lab → Checkpoint questions → Support angle**.
> Suggested pace: one module per week, but go strictly in order — COM is unforgiving of skipped fundamentals.

## Modules

| # | Module | Focus |
|---|---|---|
| 0 | [Why COM exists](modules/00-why-com-exists.md) | Binary compatibility, the three pillars, vocabulary, tooling setup |
| 1 | [`IUnknown`, vtables, and lifetime](modules/01-iunknown-and-lifetime.md) | Vtables, `QueryInterface` rules, reference counting, HRESULTs |
| 2 | [Activation, registration, and the registry](modules/02-activation-and-registry.md) | CLSID/ProgID/AppID keys, bitness, reg-free COM, HRESULT triage |
| 3 | [Threading and apartments](modules/03-apartments-and-threading.md) | STA/MTA/NA, `ThreadingModel`, marshaling, reentrancy, deadlocks |
| 4 | [Interfaces, IDL, MIDL, and marshaling](modules/04-idl-and-marshaling.md) | IDL attributes, memory ownership, proxy/stub vs typelib, versioning |
| 5 | [Automation, `IDispatch`, and scripting](modules/05-automation-and-idispatch.md) | Dual interfaces, `VARIANT`/`BSTR`/`SAFEARRAY`, events, enumerators, `IErrorInfo` |
| 6 | [Frameworks: ATL, WRL, WIL, .NET interop](modules/06-frameworks-and-interop.md) | What the macros generate, RCW/CCW, `ComWrappers` |
| 7 | [DCOM, security, and out-of-proc](modules/07-dcom-and-security.md) | AppID, session 0, `CoInitializeSecurity`, UAC, Event 10016, hardening |
| 8 | [Debugging, diagnostics, and capstone](modules/08-debugging-and-capstone.md) | ProcMon, WinDbg, ETW, AppVerifier, triage flowchart, capstone project |

### Appendices

Reference material for topics the main modules use but don't stop to teach. Read them when the cross-reference sends you there, not up front.

| | Appendix | Covers | Read before |
|---|---|---|---|
| A | [Monikers, persistence, and structured storage](modules/appendix-a-monikers-and-persistence.md) | `IMoniker`, `CoGetObject`, the Running Object Table, `IPersist*`, `IStream`/`IStorage` | Module 7 §7.5 (elevation moniker); any WMI or `GetObject` ticket |
| B | [COM+ and WinRT](modules/appendix-b-com-plus-and-winrt.md) | MTS/COM+ contexts, transactions, DTC; `IInspectable`, `.winmd`, ASTA, `HSTRING` | Supporting server estates (B.1) or anything modern on Windows (B.2) |

---

## How to use this plan

- **Do every lab.** COM is a contract-driven technology; reading about it produces false confidence.
- Keep a **`COM-Notes.md` scratch file** with every HRESULT you hit and what actually caused it. That file becomes your support cheat sheet.
- Build all labs as **C++ Desktop / Console** projects first (raw COM shows you the machinery), then repeat key ones with smart pointers and C#.
- Environment: Visual Studio (Desktop development with C++ + .NET desktop), Windows SDK, and the **Debugging Tools for Windows** (WinDbg).

### Tooling checklist (install before Module 1)

| Tool | Purpose |
|---|---|
| Visual Studio + Windows SDK | Compile, MIDL, `#import`, ATL/WRL |
| WinDbg (Store version or from SDK) | Ref-count and marshaling debugging |
| **OleView.NET** (James Forshaw, GitHub) | Modern registry/type-library/proxy browser — replaces the old `oleview.exe` |
| Process Monitor (Sysinternals) | Catch registry lookups and `REGDB_E_CLASSNOTREG` root causes |
| Process Explorer (Sysinternals) | Find surrogate/host processes, loaded DLLs |
| `dcomcnfg.exe` / Component Services | DCOM identity, launch/activation permissions |
| `RegDllView` or `regsvr32` | Registration inspection |
| Windows Performance Recorder/Analyzer (optional) | Cross-apartment call latency |

---

## Module 0 — Why COM exists (½ day)

**Concepts**

- The problem COM solves: **binary interoperability across compilers, languages, versions, and processes**, without source or recompilation.
- The three pillars: **interfaces are immutable contracts**, **objects are reference-counted**, **location is transparent**.
- Where COM lives today: Shell extensions, Office automation, DirectX/Direct2D/Media Foundation, WMI, WinRT (COM-based), .NET COM interop, MSXML, ADO/OLE DB, Windows Runtime Broker, browser hosting.
- Terminology map: *component*, *coclass*, *interface*, *CLSID*, *IID*, *ProgID*, *type library*, *apartment*, *marshaling*, *proxy/stub*.

**Reading**

- *Essential COM* by Don Box — Chapters 1–2 (the single best conceptual on-ramp).
- Microsoft Learn: "Component Object Model (COM)" overview.

**Checkpoint**

- Explain in 3 sentences why a C++ class exported from a DLL is *not* safe across compiler versions, but a COM interface is.
- Name three Windows features you use daily that are COM under the hood.

---

## Module 1 — `IUnknown`, vtables, and lifetime (1 week)

This is the module that matters most. 70% of real-world COM bugs are lifetime bugs.

**Concepts**

- The **vtable layout**: what an interface pointer actually points to in memory.
- `IUnknown`'s three methods and the **three rules of `QueryInterface`**:
  1. **Reflexive** — `QI(IID_IUnknown)` on any interface must return the *same* pointer value (object identity test).
  2. **Symmetric** — if you can get from A to B, you can get back from B to A.
  3. **Transitive** — if A→B and B→C, then A→C.
  4. (Corollary) **Static set** — the set of interfaces must not change over the object's lifetime.
- `AddRef` / `Release` rules — reference counting is per-*reference*, not per-*object*.
- **When the ref count is incremented** (memorize this list):
  - Activation APIs (`CoCreateInstance`, `IClassFactory::CreateInstance`, `CoGetClassObject`).
  - Every successful `QueryInterface`.
  - Any interface returned via an `[out]` / `[out, retval]` parameter (enumerators, property getters).
  - Copying a pointer into storage that outlives the current scope (member, global, container).
  - Marshaling: `CoMarshalInterface`, GIT registration, proxy creation.
  - Smart-pointer copy/assign (`CComPtr`, `_com_ptr_t`, `winrt::com_ptr`).
- **When it is not**: `[in]` parameters used only for the call duration, `Attach`/`Detach` ownership transfer, deliberate weak references.
- Reference cycles and why they leak; `IWeakReference` (WinRT), and the "parent holds strong / child holds weak" pattern.
- `HRESULT` anatomy: severity bit, facility, code. `SUCCEEDED`/`FAILED` — and why `if (hr == S_OK)` is a bug (`S_FALSE` is a success).

**Lab 1.1 — Implement `IUnknown` by hand**

Write a C++ class with no ATL, no WRL:

```cpp
// Goal: feel the machinery. No frameworks.
class Counter : public ICounter   // ICounter : public IUnknown
{
    LONG m_ref = 1;
public:
    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == IID_ICounter)
            *ppv = static_cast<ICounter*>(this);
        else
            return E_NOINTERFACE;
        reinterpret_cast<IUnknown*>(*ppv)->AddRef();   // ALWAYS AddRef on success
        return S_OK;
    }
    STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
    STDMETHODIMP_(ULONG) Release() override
    {
        ULONG n = InterlockedDecrement(&m_ref);
        if (n == 0) delete this;
        return n;
    }
    // ICounter
    STDMETHODIMP Increment(LONG* out) override { /* ... */ }
};
```

Then:
1. Add a second interface (`ICounter2`) and prove all three `QI` rules with assertions.
2. Add `OutputDebugString` tracing to `AddRef`/`Release` printing the resulting count and a call-site tag.
3. **Deliberately introduce a leak** (forget one `Release`) and observe the object never destructing.
4. **Deliberately over-release** and watch the crash. Note the callstack shape — you will see this in real dumps.

**Lab 1.2 — Smart pointers**

Rewrite the client code three ways and compare: raw pointers with `goto Cleanup`, `CComPtr<T>` (ATL), and `winrt::com_ptr<T>`. Note where `&ptr` vs `ptr.put()` vs `ptr.GetAddressOf()` differ and why double-assignment through `&` leaks in some libraries.

**Support angle**

- Symptom → cause table you should start building:
  - *Handle/memory grows steadily, never returns* → missing `Release` (leak).
  - *Crash in `Release` or random vtable call, `0xC0000005`* → over-release / use-after-free.
  - *`E_NOINTERFACE` intermittently* → object identity or aggregation bug, or wrong apartment proxy.
- WinDbg: `!heap`, `!address`, and setting a breakpoint on the object's `Release` with `bp <addr> "kb; g"` to log the callstack of every release.

**Checkpoint**

- Why must `QueryInterface` `AddRef` before returning, even when returning `this`?
- What does `S_FALSE` mean, and name two APIs that return it legitimately?
- Draw the memory layout of an object implementing two interfaces via multiple inheritance.

---

## Module 2 — Activation, registration, and the registry (1 week)

**Concepts**

- **CLSID vs IID vs ProgID vs LIBID vs AppID.** GUID generation and why you never reuse one.
- Registration keys and what each does:
  - `HKCR\CLSID\{...}\InprocServer32` → DLL path + `ThreadingModel`
  - `HKCR\CLSID\{...}\LocalServer32` → EXE path
  - `HKCR\CLSID\{...}\ProgID` / `VersionIndependentProgID`
  - `HKCR\<ProgID>\CLSID` (reverse lookup)
  - `HKCR\Interface\{IID}\ProxyStubClsid32` and `\TypeLib`
  - `HKCR\AppID\{...}` (DCOM identity, surrogate, permissions)
- **HKCU vs HKLM registration** and the `HKCR` merge — a huge source of "works for me" support cases.
- **Registry-free (Reg-Free) COM**: application and assembly manifests, `<comClass>`, `<file>`, activation context. Why modern deployment prefers it.
- Server types: **in-proc (DLL)**, **local (EXE)**, **remote**, **DLL surrogate (`dllhost.exe`)**.
- The DLL server's four exports: `DllGetClassObject`, `DllCanUnloadNow`, `DllRegisterServer`, `DllUnregisterServer`. What `CoFreeUnusedLibraries` does.
- `IClassFactory` / `IClassFactory2`, `CoRegisterClassObject`, `REGCLS_*` flags, `CoInitializeEx` vs `CoInitialize`.
- Activation call chain: `CoCreateInstance` → `CoCreateInstanceEx` → `CoGetClassObject` → SCM (`rpcss` service) → server → `IClassFactory::CreateInstance` → `QueryInterface`.
- `CLSCTX_*` flags and how `CLSCTX_ALL` can silently pick a different server than you expect.

**Lab 2.1 — Build a real in-proc server**

1. Create a DLL exporting all four entry points, implementing `IClassFactory`.
2. Register with `regsvr32`, activate from a separate console client with `CoCreateInstance`.
3. Inspect your CLSID in **OleView.NET**; then find the same keys in `regedit`.
4. Unregister and capture the failure: note the exact HRESULT (`0x80040154 REGDB_E_CLASSNOTREG`).
5. Run **Process Monitor** filtered on `Path contains {your-CLSID}` and trace the exact registry probe sequence the SCM performs — including the WOW6432Node redirection if you build x86 on x64.

**Lab 2.2 — Bitness**

Build the server **x86** and the client **x64**. Observe the failure. Explain it. Then fix it two ways: (a) matching bitness, (b) `DllSurrogate` under the AppID so the x86 DLL runs out-of-proc.

**Lab 2.3 — Registry-free activation**

Take the Lab 2.1 DLL, unregister it completely, and activate it purely from an application manifest + assembly manifest. This is a top-tier support skill.

**Support angle — the HRESULT triage table**

| HRESULT | Name | Usual real cause |
|---|---|---|
| `0x80040154` | `REGDB_E_CLASSNOTREG` | Not registered, **wrong bitness**, wrong hive (HKCU vs HKLM), missing manifest |
| `0x80040155` | `REGDB_E_IIDNOTREG` | Interface has no proxy/stub or typelib registration |
| `0x8007007E` | `ERROR_MOD_NOT_FOUND` | Server DLL found in registry but its **dependencies** are missing |
| `0x80070005` | `E_ACCESSDENIED` | DCOM launch/activation permissions, or integrity level mismatch |
| `0x80080005` | `CO_E_SERVER_EXEC_FAILURE` | EXE server failed to start / register class objects in time |
| `0x800706BA` | `RPC_S_SERVER_UNAVAILABLE` | Remote host, firewall, or `rpcss` issue |
| `0x80010108` | `RPC_E_DISCONNECTED` | Server died; you're holding a stale proxy |
| `0x8001010E` | `RPC_E_WRONG_THREAD` | Apartment violation — see Module 3 |
| `0x800401F0` | `CO_E_NOTINITIALIZED` | `CoInitializeEx` not called on this thread |
| `0x80070422` | service disabled | `rpcss`/DCOM service state |

Practise: for each row, **reproduce it deliberately** in your lab and record the ProcMon/Event Viewer signature. Also learn to read **Event Viewer → System → DistributedCOM (Event IDs 10016, 10001, 10005)**.

---

## Module 3 — Threading and apartments (1.5 weeks — the hardest module)

**Concepts**

- Why apartments exist: COM objects are not automatically thread-safe; the apartment is a **thread-safety contract**.
- **STA** (single-threaded apartment): one thread, calls serialized via a **hidden message window** and the message pump. Requires `CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)`.
- **MTA** (multi-threaded apartment): many threads, object must be thread-safe, no pumping. `COINIT_MULTITHREADED`.
- **NA** (neutral apartment / TNA) and `ThreadingModel=Neutral`.
- `ThreadingModel` values in the registry — `Apartment`, `Free`, `Both`, `Neutral`, and *absent* (= legacy single-threaded, "Main STA"). Know what each does to activation placement.
- **Cross-apartment calls require marshaling.** Never pass a raw interface pointer between apartments.
  - `CoMarshalInterThreadInterfaceInStream` / `CoGetInterfaceAndReleaseStream`
  - **Global Interface Table (GIT)**: `IGlobalInterfaceTable::RegisterInterfaceInGlobal` / `GetInterfaceFromGlobal` — the safe, reusable option.
- **Message pumping and reentrancy**: while an STA thread waits on an outgoing COM call, it *pumps* — so incoming calls can reenter your code. This is the source of the nastiest COM bugs.
  - `IMessageFilter` (`HandleIncomingCall`, `RetryRejectedCall`, `MessagePending`) — the "Server Busy"/"switch to" dialogs.
  - `CoWaitForMultipleHandles` vs raw `WaitForSingleObject` (deadlock).
- Deadlock patterns: STA thread blocking on a mutex while a server calls back; UI thread doing `WaitForSingleObject` on a thread that is calling into the UI apartment.
- `RPC_E_WRONG_THREAD`, `RPC_E_CANTCALLOUT_ININPUTSYNCCALL`, `RPC_E_SERVERCALL_RETRYLATER`, `RPC_E_CALL_REJECTED` — what each tells you.
- .NET specifics: `[STAThread]`, `Thread.SetApartmentState`, why WinForms/WPF are STA, `ConfigureAwait` and COM context, `SynchronizationContext`.

**Lab 3.1 — Prove marshaling is required**

1. Create an object on an STA thread. Pass the **raw pointer** to an MTA thread and call a method. Observe corrupted behaviour or `RPC_E_WRONG_THREAD`.
2. Repeat correctly with `CoMarshalInterThreadInterfaceInStream`.
3. Repeat with the **GIT**. Compare ergonomics.

**Lab 3.2 — Build a deadlock, then fix it**

STA UI thread calls a slow out-of-proc server method; server calls back into the client during the call; client's callback tries to take a lock held by the UI thread. Reproduce the hang. Then:
- Capture it with WinDbg: `~*kb`, `!locks`, and identify the RPC wait frames.
- Fix it with `IMessageFilter` and by removing the lock from the callback path.

**Lab 3.3 — ThreadingModel experiment**

Register the same DLL server four times under four CLSIDs with `Apartment`, `Free`, `Both`, `Neutral`. From an MTA client and an STA client, activate each and log the thread ID the object's methods actually run on. Build a 4×2 matrix of results. **Keep this matrix** — it explains most "why is my object slow / why is it on the wrong thread" cases.

**Support angle**

- Hang triage recipe:
  1. Capture a full user-mode dump (`procdump -ma`, or Task Manager → Create dump file).
  2. `~*kb` — look for `NdrClientCall`, `CoWaitForMultipleHandles`, `ModalLoop`, `SwitchTo`.
  3. `!cs -l` / `!locks` for critical section owners.
  4. `!runaway` for CPU.
  5. In cross-process cases, dump **both** processes and match RPC call IDs.
- Learn the classic dump signature of "STA thread blocked without pumping."

---

## Module 4 — Interfaces, IDL, MIDL, and marshaling (1 week)

**Concepts**

- **IDL** syntax: `interface`, `[object]`, `[uuid]`, `[pointer_default]`, `library`, `coclass`, `dispinterface`.
- Parameter attributes: `[in]`, `[out]`, `[in,out]`, `[retval]`, `[size_is]`, `[length_is]`, `[string]`, `[unique]`, `[ref]`, `[ptr]`, `[iid_is]`.
- **Memory ownership rules** (the "COM memory management rules"):
  - `[out]` and `[in,out]` buffers are allocated with `CoTaskMemAlloc` and freed by the **caller** with `CoTaskMemFree`.
  - On failure, `[out]` params must be set to `NULL`/zero.
  - `BSTR` uses `SysAllocString`/`SysFreeString` — **not** `CoTaskMemFree`.
- Marshaling flavours:
  - **Standard marshaling** — MIDL-generated proxy/stub DLL, registered under `HKCR\Interface\{IID}\ProxyStubClsid32`.
  - **Type-library marshaling (oleautomation)** — uses `oleaut32`'s universal marshaler; restricted to Automation-compatible types.
  - **Custom marshaling** — `IMarshal` (rare, e.g., for immutable objects that copy themselves across).
- `NdrProxy`/`NdrStub`, and why a missing proxy/stub gives `E_NOINTERFACE` **only when crossing an apartment** (the "works in-proc, fails out-of-proc" classic).
- Versioning: **never modify a published interface**. Create `IFoo2`, `IFoo3`. Understand why adding a method breaks binary compatibility.

**Lab 4.1 — Author an IDL and compile it**

Define `ICalculator` with methods taking `[in] LONG`, `[out] BSTR*`, `[in, size_is(count)] BYTE*`, and `[out, retval] LONG*`. Run MIDL. Read the generated `_i.c`, `_p.c`, and `dlldata.c`. Build and register the proxy/stub DLL.

**Lab 4.2 — Break marshaling on purpose**

Unregister the proxy/stub, call the object from a different apartment, and record the exact error. Then re-register and confirm. Then mark the interface `[oleautomation]`, register the **type library** instead, and prove it works without a proxy/stub DLL.

**Lab 4.3 — Memory ownership**

Write a method returning a `BSTR` and an array. Verify with Application Verifier / CRT leak detection that the caller frees correctly. Deliberately free a `BSTR` with `CoTaskMemFree` and observe the heap corruption.

**Checkpoint**

- Why can't `[oleautomation]` interfaces take arbitrary structs?
- Who frees an `[in, out]` `BSTR*` on the way in and on the way out?

---

## Module 5 — Automation, `IDispatch`, and scripting (1 week)

**Concepts**

- `IDispatch`: `GetIDsOfNames`, `Invoke`, `GetTypeInfo`, `GetTypeInfoCount`. Late binding vs early binding vs **dual interfaces**.
- `VARIANT`, `BSTR`, `SAFEARRAY`, `DISPPARAMS`, `EXCEPINFO`, `DISPID` (and the well-known negative DISPIDs like `DISPID_VALUE`, `DISPID_NEWENUM`, `DISPID_PROPERTYPUT`).
- Type libraries: `.tlb`, `LoadTypeLib`, `ITypeInfo`, `ITypeLib`, registration under `HKCR\TypeLib`.
- Connection points and events: `IConnectionPointContainer`, `IConnectionPoint`, `Advise`/`Unadvise`, sink objects, and the **classic leak: forgetting `Unadvise`** (a reference cycle).
- Enumerators: `IEnumVARIANT`, `IEnumUnknown`, the `Next/Skip/Reset/Clone` pattern, and `for each` support via `_NewEnum`.
- Error reporting: `ISupportErrorInfo`, `ICreateErrorInfo`, `IErrorInfo`, `SetErrorInfo`/`GetErrorInfo`. Why rich errors beat bare HRESULTs.

**Lab 5.1 — Dual interface**

Extend `ICalculator` to derive from `IDispatch`, implement it via `ITypeInfo::Invoke`, and drive it from:
- C++ with `#import` (early bound)
- VBScript / PowerShell `New-Object -ComObject` (late bound)
- C# `dynamic` and via a COM reference

**Lab 5.2 — Events**

Implement a connection point, subscribe a sink, and confirm callbacks. Then remove the `Unadvise` and prove with your `AddRef` tracing that neither object ever dies.

**Support angle**

- `0x80020006 DISP_E_UNKNOWNNAME`, `0x80020005 DISP_E_TYPEMISMATCH`, `0x8002000E DISP_E_BADPARAMCOUNT` — mapping these to customer script errors.
- Office automation from a service/session 0 — why it's unsupported, and how to recognize the symptom set.

---

## Module 6 — Frameworks: ATL, WRL, WIL, and .NET interop (1 week)

**Concepts**

- **ATL**: `CComObjectRootEx`, `CComCoClass`, `BEGIN_COM_MAP`, `CComPtr`, `CComBSTR`, `CComVariant`, `CComSafeArray`, object map, `CComModule`/`CAtlDllModuleT`. Aggregation and `CComAggObject`.
- **WRL** (`Microsoft::WRL`): `RuntimeClass`, `ComPtr`, `MakeAndInitialize`, `RuntimeClassFlags<ClassicCom>`.
- **WIL** (`wil::com_ptr`, `wil::unique_couninitialize_call`, error macros `RETURN_IF_FAILED`, `THROW_IF_FAILED`).
- **C++/WinRT** `winrt::com_ptr`, `winrt::check_hresult`.
- **.NET interop**:
  - RCW (Runtime Callable Wrapper) and CCW (COM Callable Wrapper).
  - `Marshal.ReleaseComObject` vs `FinalReleaseComObject` vs letting the GC do it — and when each is right/dangerous.
  - `[ComVisible]`, `[Guid]`, `[InterfaceType]`, `[DispId]`, `tlbexp`/`tlbimp`, **Primary Interop Assemblies**, "Embed Interop Types" / NoPIA.
  - `ComWrappers` (modern, .NET 5+) and source-generated COM interop (`GeneratedComInterface`) — the direction .NET is moving.
  - Registration for .NET servers: `regasm`, and why COM-visible .NET is different on .NET Core/5+.

**Lab 6.1** — Rewrite the Module 2 server in ATL. Compare the line count and note exactly which boilerplate each macro replaced.

**Lab 6.2** — Build a C# COM server, consume it from C++; build a C++ COM server, consume it from C#. Register both ways. Then repeat with **NoPIA / embedded interop types**.

**Lab 6.3** — Reproduce the classic RCW bug: call `Marshal.ReleaseComObject` on an object still referenced elsewhere, then use it → `InvalidComObjectException` / `COMException 0x80004002`. Understand the "RCW is per-object, not per-reference" model.

---

## Module 7 — DCOM, security, and out-of-proc (1 week)

**Concepts**

- Local server activation and the **SCM (`rpcss`)**; the class object registration handshake and the `CO_E_SERVER_EXEC_FAILURE` timeout.
- **AppID** settings: `RunAs` identity (`Interactive User`, `Launching User`, a specific account, `NT AUTHORITY\LocalService`), `DllSurrogate`, `AccessPermission`, `LaunchPermission`, `AuthenticationLevel`.
- **Session 0 isolation** — services cannot show UI or drive interactive-user COM servers. Recognize this instantly in support cases.
- `CoInitializeSecurity`: authentication level, impersonation level, cloaking, and why calling it twice fails.
- Identity/impersonation: `CoImpersonateClient`, `CoRevertToSelf`, `IServerSecurity`.
- **UAC and integrity levels**: a medium-IL client cannot drive a high-IL server; the "COM Elevation Moniker" (`Elevation:Administrator!new:{CLSID}`) and its registry requirements (`LocalizedString`, `Elevation` key).
- **Event 10016 (DCOM permission)** — how to read it, how to actually fix it (and when it's benign noise).
- Remote DCOM: endpoint mapper (TCP 135), dynamic RPC port range, firewall rules, `DCOM hardening` (CVE-2021-26414 / `RequireIntegrityActivationAuthenticationLevel`) — a very common modern support topic.
- Windows Firewall + `dcomcnfg` + `netsh` diagnostics.

**Lab 7.1** — Convert your in-proc server to an out-of-proc EXE server with `CoRegisterClassObject`. Watch the SCM start it. Kill the server mid-call from the client and observe `RPC_E_DISCONNECTED` / `0x800706BA`.

**Lab 7.2** — Configure a `DllSurrogate` for the in-proc server. Confirm it now runs in `dllhost.exe` (Process Explorer).

**Lab 7.3** — Set a restrictive Launch permission via `dcomcnfg`, reproduce `E_ACCESSDENIED` + Event 10016, then fix it properly (grant to the right principal, not "Everyone").

**Lab 7.4** — Remote activation between two machines/VMs with `CoCreateInstanceEx` + `COSERVERINFO`. Break it with a firewall rule, diagnose with the error code alone, then fix.

---

## Module 8 — Debugging, diagnostics, and support mastery (ongoing)

**Skills to build**

- **WinDbg for COM**:
  - `!error <hr>` — decode any HRESULT instantly.
  - `~*kb` / `!uniqstack` — hang analysis.
  - `dps <interface-ptr>` — dump the vtable and identify the implementing module (identifies *which* component you actually have).
  - `bp <module>!*::Release` with logging — ref-count archaeology.
  - `!htrace`, `!heap -p -a <addr>` — handle/heap leaks.
  - Loading `ole32`/`combase` symbols and reading `CStdIdentity`, `OXID`, `IPID` structures at a basic level.
- **ETW / tracing**: `Microsoft-Windows-COM`, `Microsoft-Windows-COMRuntime`, `Microsoft-Windows-RPC` providers; capture with `wpr`/`tracelog`, view in WPA.
- **Application Verifier** with the COM checks enabled — catches apartment violations and bad ref counting at runtime.
- **Process Monitor** patterns for activation failures (the exact key probe order).
- **OleView.NET** for: enumerating CLSIDs by server, viewing proxy/stub registration, checking a component's `ThreadingModel` and AppID, and diffing registration between a working and a broken machine.

**Capstone project**

Build a small but complete component and then support it:

1. `IDocumentStore` — an out-of-proc COM server (EXE) with a dual interface, a connection-point event for change notifications, an `IEnumVARIANT` enumerator, and rich `IErrorInfo` errors.
2. Clients: C++ (raw + ATL), C# (both classic interop and `ComWrappers`), and PowerShell.
3. Ship it three ways: `regsvr32`-registered, reg-free with manifests, and via a DLL surrogate.
4. **Then write the support runbook**: for each of the 12 HRESULTs in the Module 2 table, document how it would manifest for *this* component and the exact diagnostic steps.

That runbook is the deliverable that proves you've made it.

---

## Reference shelf

**Books (in priority order)**

1. *Essential COM* — Don Box. Still the definitive explanation of *why*.
2. *Inside COM* — Dale Rogerson. Gentlest possible on-ramp; builds `IUnknown` from nothing.
3. *ATL Internals* — Rector/Sells (2nd ed. Tavares et al.).
4. *Advanced Windows Debugging* — Hewardt/Pravat, for the support track.
5. *Windows Internals* (Part 1 & 2) — for RPC, ALPC, sessions, and integrity levels.

**Online**

- Microsoft Learn → "Component Object Model (COM)" and "COM Fundamentals".
- Raymond Chen's *The Old New Thing* — search for "apartment", "STA", "message pump", "COM".
- OleView.NET wiki (James Forshaw) — modern COM security research and tooling.
- Microsoft Learn → "Introduction to COM interop" and "COM Wrappers" for the .NET side.

---

## Progress tracker

| # | Module | Read | Labs done | Checkpoint passed | Notes written |
|---|---|---|---|---|---|
| 0 | [Why COM exists](modules/00-why-com-exists.md) | ☐ | ☐ | ☐ | ☐ |
| 1 | [IUnknown & lifetime](modules/01-iunknown-and-lifetime.md) | ☐ | ☐ | ☐ | ☐ |
| 2 | [Activation & registry](modules/02-activation-and-registry.md) | ☐ | ☐ | ☐ | ☐ |
| 3 | [Threading & apartments](modules/03-apartments-and-threading.md) | ☐ | ☐ | ☐ | ☐ |
| 4 | [IDL & marshaling](modules/04-idl-and-marshaling.md) | ☐ | ☐ | ☐ | ☐ |
| 5 | [Automation & IDispatch](modules/05-automation-and-idispatch.md) | ☐ | ☐ | ☐ | ☐ |
| 6 | [ATL / WRL / .NET interop](modules/06-frameworks-and-interop.md) | ☐ | ☐ | ☐ | ☐ |
| 7 | [DCOM & security](modules/07-dcom-and-security.md) | ☐ | ☐ | ☐ | ☐ |
| 8 | [Debugging & capstone](modules/08-debugging-and-capstone.md) | ☐ | ☐ | ☐ | ☐ |
| A | [Monikers & persistence](modules/appendix-a-monikers-and-persistence.md) | ☐ | — | ☐ | ☐ |
| B | [COM+ & WinRT](modules/appendix-b-com-plus-and-winrt.md) | ☐ | — | ☐ | ☐ |

---

## The 10 rules to internalize

1. Every `AddRef` needs exactly one `Release`; out-params always carry a reference you own.
2. Never compare interface pointers for identity except via `QI(IID_IUnknown)`.
3. Check `FAILED(hr)`, never `hr == S_OK`.
4. Never pass a raw interface pointer across apartments — marshal it.
5. Published interfaces are immutable. Version by adding new interfaces.
6. `[out]` params must be nulled on failure.
7. `BSTR` → `SysFreeString`; everything else `[out]` → `CoTaskMemFree`.
8. An STA thread must pump messages; blocking waits on an STA are deadlocks waiting to happen.
9. Bitness and registry hive mismatches cause more `REGDB_E_CLASSNOTREG` than actual missing registration.
10. Always `Unadvise` what you `Advise`, and break cycles deliberately.
