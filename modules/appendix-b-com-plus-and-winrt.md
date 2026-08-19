# Appendix B — COM+ and WinRT

**Reference material. Read B.1 if you support server/enterprise estates; read B.2 if you support anything modern on Windows.**

Two ecosystems built *on* COM that you'll meet in tickets. Neither replaces what you learned — both are COM underneath, so every rule from Modules 1–7 still applies.

---

# B.1 COM+ (and its ancestor, MTS)

## B.1.1 What it is and why it exists

Classic COM gives you activation, lifetime, and marshaling. It gives you nothing for the problems every business application has: transactions spanning multiple resources, connection pooling, role-based authorization, asynchronous invocation.

**Microsoft Transaction Server (MTS)**, shipped for NT 4, bolted those on. In Windows 2000 it merged with COM to become **COM+**. It is still present in every Windows Server release today.

> **The COM+ idea: declare your requirements; let the runtime supply the plumbing.**

You mark a component "requires a transaction," and COM+ interposes itself between the caller and your object, starting and committing a DTC transaction around your calls. Your code contains no transaction API calls at all.

## B.1.2 What COM+ adds

| Service | What it does |
|---|---|
| **Automatic transactions** | Declarative `Required` / `Requires New` / `Supported` / `Not Supported` / `Disabled`, backed by MS DTC. Spans SQL Server, MSMQ, other resource managers. |
| **JIT activation** | The client holds a reference; the actual object is created on first call and destroyed on `SetComplete`. Scales stateless components. |
| **Object pooling** | Instances are reused rather than recreated. Requires a thread-agnostic (usually Neutral/`Both`) component. |
| **Role-based security** | Declarative: define roles in the catalog, assign users, mark methods. No security code in the component. |
| **Queued components** | Calls are recorded to MSMQ and replayed later. Enables genuinely disconnected operation. |
| **Loosely-coupled events (LCE)** | Publish/subscribe with the subscription list in the catalog, so publisher and subscriber never reference each other (unlike connection points, Module 5). |
| **Synchronization** | Declarative concurrency control — activity-based, so you don't hand-roll locks. |
| **CRM** | Compensating Resource Manager, for making non-transactional resources participate. |

## B.1.3 Applications, contexts, and the catalog

**COM+ application** = a configured group of components sharing settings and identity.

| Type | Runs in | Use |
|---|---|---|
| **Server application** | Its own `dllhost.exe` | Isolation, separate identity, remote access, full COM+ services |
| **Library application** | The **caller's** process | Speed; loses process isolation and its own identity |

A server application is essentially a managed `DllSurrogate` (Module 7 §7.10) with services layered on. Everything you know about AppID identity and permissions applies.

**Context** is the mechanism. When a COM+ component is activated, the runtime places it in a **context** carrying its declared attributes. Calls crossing a context boundary are **intercepted** — and that interception is where transactions, security checks, and JIT activation happen.

```
   Caller's context                Interception             Object's context
 ┌──────────────────┐        ┌────────────────────┐      ┌──────────────────┐
 │  client object   │───────►│ begin transaction? │─────►│  your component  │
 │                  │        │ check role?        │      │  (no plumbing    │
 │                  │◄───────│ JIT-create?        │◄─────│   code at all)   │
 └──────────────────┘        └────────────────────┘      └──────────────────┘
```

This is the same proxy/interception idea as Module 3's apartment marshaling — a different boundary, same shape.

The **catalog** is COM+'s configuration store (`comexp.msc` / *Component Services*), scriptable through the `COMAdmin` objects:

```powershell
$admin = New-Object -ComObject COMAdmin.COMAdminCatalog
$apps  = $admin.GetCollection("Applications")
$apps.Populate()
$apps | ForEach-Object { "{0,-40} {1}" -f $_.Name, $_.Value("Identity") }
```

## B.1.4 Programming model

```cpp
// Classic: vote on the transaction's outcome.
IObjectContext* pCtx = nullptr;
GetObjectContext(&pCtx);

if (SUCCEEDED(DoTheWork()))
    pCtx->SetComplete();     // "I'm done AND I vote commit"
else
    pCtx->SetAbort();        // "I'm done AND I vote abort" - transaction WILL roll back

pCtx->Release();
```

`SetComplete`/`SetAbort` do two things at once: cast a transaction vote **and** signal that the object can be deactivated (JIT). `EnableCommit`/`DisableCommit` vote without deactivating, for stateful objects.

.NET equivalent (`System.EnterpriseServices`):

```csharp
[Transaction(TransactionOption.Required)]
[SecurityRole("Managers")]
[JustInTimeActivation(true)]
public class OrderProcessor : ServicedComponent
{
    [AutoComplete]                       // SetComplete on return, SetAbort on exception
    public void PlaceOrder(int id) { /* no transaction code needed */ }
}
```

> **Note for modern work:** `System.EnterpriseServices` is **.NET Framework only**. It does not exist in .NET Core/5+. A customer porting a serviced component to modern .NET must move to `System.Transactions`, a message broker, or a database-native transaction — there's no direct successor. Knowing this saves a long, doomed investigation.

## B.1.5 COM+ support essentials

### Where to look

| Artifact | Location |
|---|---|
| Configuration | `comexp.msc` → Component Services → Computers → My Computer → COM+ Applications |
| Events | Event Viewer → Application, source **COM+**, **COM+ SVC**, **MSDTC** |
| DTC | `dcomcnfg` → Component Services → … → Distributed Transaction Coordinator |
| Identity | Application → Properties → Identity tab (same trap as Module 7's `RunAs`) |
| Process | `dllhost.exe /ProcessID:{APPID}` — the AppID identifies which application |

Map a running `dllhost.exe` to its application:

```powershell
Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" |
    Select-Object ProcessId, CommandLine
```

The `/ProcessID:{GUID}` in the command line is the COM+ application's AppID. Look it up in the catalog.

### Common failures

| Symptom / code | Cause |
|---|---|
| `8004E00F` `CONTEXT_E_ABORTED` | Something in the transaction voted abort — find *which* component |
| `8004E002` `CONTEXT_E_ABORTING` | The transaction is already aborting; your call arrived too late |
| `8004D01B` `XACT_E_TMNOTAVAILABLE` | **MSDTC service not running** |
| `8004D00E` `XACT_E_NOTRANSACTION` | Expected a transaction, none present — check the declared option |
| `8004E00C` `CONTEXT_E_NOCONTEXT` | `GetObjectContext` outside COM+ — component isn't actually configured |
| DTC failures between machines | **Network DTC access** disabled, or firewall. Test with `DTCPing`. |
| App won't start | Identity password stale — identical to Module 7's `RunAs` trap |
| Random `E_ACCESSDENIED` | Role-based security: the caller isn't in a required role |

> **The single most common COM+ ticket is a DTC one**, and the two questions that resolve most of them are: *is the MSDTC service running on **both** machines*, and *is Network DTC Access enabled with the right inbound/outbound settings*. `DTCPing` on both ends answers both.

---

# B.2 WinRT — COM with new rules

## B.2.1 WinRT *is* COM

This is the fact that matters. The Windows Runtime is not a replacement for COM; it's COM with a stricter contract and better metadata.

```cpp
struct IInspectable : public IUnknown          // <- IUnknown. Still.
{
    HRESULT GetIids(ULONG* iidCount, IID** iids);
    HRESULT GetRuntimeClassName(HSTRING* className);
    HRESULT GetTrustLevel(TrustLevel* trustLevel);
};
```

Every WinRT object is a COM object. `QueryInterface`, `AddRef`, `Release` all work exactly as Module 1 described. Reference counting, apartment rules, marshaling — all still apply.

`IInspectable` adds what classic COM lacked: an object can **describe itself** at runtime (`GetIids`, `GetRuntimeClassName`) without a separate type library.

## B.2.2 What changed

| Concern | Classic COM | WinRT |
|---|---|---|
| Base interface | `IUnknown` | `IInspectable` (: `IUnknown`) |
| Metadata | `.tlb` type library, registered in `HKCR\TypeLib` | **`.winmd`** files — ECMA-335 (the .NET metadata format) |
| Activation | `CoCreateInstance` + `IClassFactory` | `RoActivateInstance` + **`IActivationFactory`** |
| Naming | CLSID (a GUID) | **Runtime class name** (a string): `Windows.Storage.StorageFile` |
| Registration | Registry (`HKCR\CLSID\…`) | **Package manifest** (`AppxManifest.xml`); no registry for in-package classes |
| Strings | `BSTR` (`SysAllocString`) | **`HSTRING`** (`WindowsCreateString`) — immutable, ref-counted |
| Errors | `HRESULT` + `IErrorInfo` | `HRESULT` + **`IRestrictedErrorInfo`** (carries a message *and* a call stack) |
| Collections | `IEnumVARIANT` | `IVector<T>`, `IMap<K,V>`, `IIterable<T>` — genuinely generic (parameterized IIDs) |
| Async | ad hoc | `IAsyncOperation<T>` / `IAsyncAction` — a uniform, mandated pattern |
| Events | connection points | `add_X` / `remove_X` returning an `EventRegistrationToken` |
| Apartments | STA / MTA / NA | STA / MTA / **ASTA** |

### `HSTRING` in one line

Immutable and reference-counted, so passing one around doesn't copy. `WindowsCreateString` / `WindowsDeleteString`. A "fast pass" variant (`HSTRING_HEADER`) lets you wrap an existing buffer with no allocation — which is why WinRT string passing is cheaper than `BSTR`.

### Events

```cpp
EventRegistrationToken token{};
button.add_Click(handler, &token);
// ...
button.remove_Click(token);        // <- the Module 5 Unadvise rule, renamed
```

**Same leak, new spelling.** Forget `remove_X` and you've built the exact reference cycle from Module 5 §5.7. C++/WinRT's `event_revoker` and `auto_revoke` are the RAII fix:

```cpp
auto revoker = button.Click(winrt::auto_revoke, handler);   // revoked on destruction
```

## B.2.3 ASTA — the new apartment

**ASTA** (Application STA) is a third apartment type introduced for WinRT/UWP UI threads. It is an STA that **forbids nested pumping**.

Recall Module 3 §3.6: a classic STA pumps messages while waiting on an outbound call, which permits arbitrary reentrancy and produces the nastiest COM bugs. ASTA closes that hole:

| | Classic STA | ASTA |
|---|---|---|
| One thread | yes | yes |
| Pumps while waiting on an outbound call | **yes** — arbitrary reentrancy | **no** |
| Blocking waits allowed | discouraged | **actively blocked** |
| Reentrancy | possible at any call | only via explicitly awaited continuations |

Attempting to block an ASTA thread, or to make a synchronous cross-apartment call that would require nested pumping, fails rather than deadlocking or silently re-entering. That's why WinRT's async model is *mandatory* rather than merely encouraged — the apartment enforces it.

> **Support signature:** `RPC_E_CANTCALLOUT_ININPUTSYNCCALL` or a WinRT "call was made on the wrong thread"/reentrancy error in a UWP or WinUI app usually means someone did `.get()` / `.wait()` / `Task.Wait()` on a UI thread. The fix is `co_await` / `await`, not a workaround.

## B.2.4 Activation

```cpp
// Raw ABI - what the projections do underneath
HSTRING className;
WindowsCreateString(L"Windows.Globalization.Calendar", 30, &className);

IInspectable* pInst = nullptr;
HRESULT hr = RoActivateInstance(className, &pInst);      // ~ CoCreateInstance

// Or get the factory first - ~ CoGetClassObject
IActivationFactory* pFactory = nullptr;
hr = RoGetActivationFactory(className, IID_IActivationFactory, (void**)&pFactory);
```

`IActivationFactory::ActivateInstance` is the direct analogue of `IClassFactory::CreateInstance`.

You will almost never write this. **Language projections** generate it:

```cpp
// C++/WinRT
winrt::Windows::Globalization::Calendar cal;      // that's the activation
auto now = cal.GetDateTime();
```

```csharp
// C#/WinRT
var cal = new Windows.Globalization.Calendar();
```

Projections are compile-time generated from `.winmd`, so there's no runtime reflection cost — this is the same "early binding beats late binding" lesson from Module 5, applied at a different layer.

## B.2.5 Registration and packaging

For an app package, WinRT classes are declared in `AppxManifest.xml`, not the registry:

```xml
<Extensions>
  <Extension Category="windows.activatableClass.inProcessServer">
    <InProcessServer>
      <Path>MyComponent.dll</Path>
      <ActivatableClass ActivatableClassId="MyCompany.MyComponent"
                        ThreadingModel="both" />
    </InProcessServer>
  </Extension>
</Extensions>
```

Outside a package, **registration-free WinRT** uses an application manifest — the direct descendant of Module 2 §2.8's reg-free COM, with `<activatableClass>` instead of `<comClass>`.

## B.2.6 WinRT support essentials

| Symptom | Meaning | Check |
|---|---|---|
| `0x80040154` `REGDB_E_CLASSNOTREG` from `RoActivateInstance` | Class not found | Is the app packaged? Is `ActivatableClass` in the manifest? Is the `.winmd` deployed? |
| `0x80073D54` / `APPX` errors | Package registration/deployment | `Get-AppxPackage`, `Get-AppxLog` |
| `RPC_E_CANTCALLOUT_ININPUTSYNCCALL` | Blocking on an ASTA | Look for `.get()`, `.wait()`, `Task.Wait()`, `.Result` on a UI thread |
| `0x8000000B` `E_ILLEGAL_METHOD_CALL` | Wrong state — e.g. reusing a completed `IAsyncOperation` | The async object's lifecycle |
| `0x80070005` in an app container | Missing **capability**, not a DCOM permission | The manifest's `<Capabilities>` |
| Generic `E_FAIL` with no detail | Rich error info wasn't fetched | `GetRestrictedErrorInfo` / `RoGetErrorReportingFlags` |

```powershell
Get-AppxPackage -Name *MyApp* | Format-List Name, InstallLocation, Status
Get-AppxLog -ActivityId <id>          # deployment failures, in detail
```

> **The most important WinRT diagnostic habit:** always retrieve `IRestrictedErrorInfo`. WinRT captures a description *and the originating call stack* at the point of failure — but only if you ask for it. C++/WinRT's `winrt::hresult_error::message()` and .NET's `Exception.Message` do this for you; raw ABI code that ignores it discards the useful half of the error.

## B.2.7 Everything from Modules 1–7 still applies

| Module | Still true in WinRT? |
|---|---|
| 1 — `IUnknown`, ref counting, `QueryInterface` rules | **Yes, unchanged.** `IInspectable` derives from `IUnknown`. |
| 2 — Activation & registration | Yes, with manifests replacing the registry for packaged apps |
| 3 — Apartments & marshaling | Yes, plus ASTA and its stricter rules |
| 4 — Interfaces are immutable, version by adding | **Yes** — WinRT adds `IFoo2`, `IFoo3` for exactly the Module 4 reason |
| 5 — Events leak without unsubscribing | Yes — `remove_X` is the new `Unadvise` |
| 6 — Frameworks hide the boilerplate | Yes — projections instead of ATL |
| 7 — Security is a boundary | Yes, plus app-container capabilities |

---

## B.3 Checkpoint

1. What does COM+ add over classic COM, and what's the mechanism that makes it possible without changing your component's code?
2. Server application vs library application — what do you gain and lose with each?
3. `SetComplete` does two things. What are they?
4. A customer's distributed transaction fails between two servers. What are your first two questions?
5. A .NET Framework app uses `ServicedComponent`. They're porting to .NET 8. What do you tell them?
6. In what sense is WinRT "just COM"? Name three things that changed and one that didn't.
7. What is ASTA, and which Module 3 problem does it exist to eliminate?
8. A WinUI app throws `RPC_E_CANTCALLOUT_ININPUTSYNCCALL`. What did the developer almost certainly do?
9. `remove_Click` is forgotten. What Module 5 bug have you just reproduced?

<details>
<summary>Answers</summary>

1. Declarative transactions, JIT activation, object pooling, role-based security, queued components, loosely-coupled events, and declarative synchronization. The mechanism is **context and interception**: COM+ activates the object in a context carrying its declared attributes, and intercepts every call crossing that context boundary — starting transactions, checking roles, creating/destroying the instance. The component contains no plumbing code, exactly as a marshaling proxy requires no code in the object (Module 3).

2. **Server application** runs in its own `dllhost.exe`: process isolation, its own identity, remotable, full services — at the cost of cross-process marshaling on every call. **Library application** runs in the caller's process: fast, no marshaling — but no isolation (a crash takes the host down), no separate identity, and not remotely activatable.

3. It casts a **commit vote** for the transaction, and it signals that the object may be **deactivated** (JIT activation destroys it, returning the client a still-valid reference to a now-empty context). `EnableCommit` votes without deactivating.

4. (a) Is the **MSDTC service running on both machines**? (b) Is **Network DTC Access** enabled with the correct inbound/outbound/authentication settings, and is the firewall permitting it? `DTCPing` from each end answers both. (Bonus third: does name resolution work in both directions — DTC is sensitive to it.)

5. `System.EnterpriseServices` is **.NET Framework only** and has no .NET Core/5+ equivalent. There is no direct port. They must re-architect: `System.Transactions` with a single resource manager, database-native transactions, a saga/compensation pattern, or a message broker for the queued-components scenarios. Telling them this early prevents a long doomed migration attempt.

6. WinRT objects are COM objects: `IInspectable` derives from `IUnknown`, and `QueryInterface`/`AddRef`/`Release` behave exactly as in Module 1. **Changed:** metadata is `.winmd` (ECMA-335) not `.tlb`; activation is `RoActivateInstance`/`IActivationFactory` by *class name string* rather than `CoCreateInstance` by CLSID; strings are `HSTRING` not `BSTR`; registration is by package manifest not registry; ASTA is added. **Unchanged:** reference counting and the `QueryInterface` rules — including that interfaces are immutable and you version by adding `IFoo2`.

7. ASTA (Application STA) is an STA that **does not pump messages while waiting on an outbound call** and blocks attempts to wait synchronously. It exists to eliminate Module 3 §3.6's **arbitrary reentrancy**: in a classic STA, any outbound cross-apartment call can let unrelated code — even a second entry into the same method — run on your thread mid-call. ASTA makes that structurally impossible, which is why WinRT's async model is enforced rather than advised.

8. Blocked the UI thread on an asynchronous operation — `.get()`, `.wait()`, `Task.Wait()`, or `.Result` on an ASTA thread. The fix is `co_await` / `await`, not a workaround; the apartment is correctly refusing an operation that would require nested pumping.

9. The connection-point leak from Module 5 §5.7. The event source holds a strong reference to the handler, the handler typically captures the page/window, and nothing reaches zero — so the entire UI tree leaks. The fix is the same shape as RAII `Unadvise`: C++/WinRT's `winrt::auto_revoke` / `event_revoker`, or unsubscribing explicitly on unload.

</details>

---

## B.4 Rules

1. COM+ is context + interception. Your component stays plumbing-free; the boundary does the work.
2. Server vs library application is the isolation/speed trade-off, and it's the same one as out-of-proc vs in-proc.
3. Most COM+ tickets are DTC tickets: service running on both ends, Network DTC Access, firewall.
4. COM+ application identity has the same stale-password trap as Module 7's `RunAs`.
5. `System.EnterpriseServices` does not exist on .NET 5+. Say so early.
6. WinRT is COM. Every Module 1 rule holds unchanged.
7. `.winmd` replaces `.tlb`; class-name strings replace CLSIDs; manifests replace the registry.
8. ASTA forbids nested pumping — that's a feature, and it's why async is mandatory.
9. `remove_X` is `Unadvise`. Use `auto_revoke`.
10. Always fetch `IRestrictedErrorInfo`; it carries the originating stack that a bare HRESULT throws away.

---

**Back to:** [Module 6 — Frameworks & interop](06-frameworks-and-interop.md) · [Module 7 — DCOM & security](07-dcom-and-security.md) · [Curriculum](../README.md)
