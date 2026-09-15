# Module 7 — DCOM, security, and out-of-proc servers

Once a COM object lives in another process — or on another machine — activation stops being a registry lookup and becomes a *security decision*. Nearly every "access denied," "server execution failed," and Event 10016 ticket lives here.

**What is Event 10016?** It is a Windows event-log entry from the **DistributedCOM** source, visible in **Event Viewer > Windows Logs > System**. The number `10016` identifies the event type; it is not an `HRESULT`. The message reports that a caller lacked permission to launch or activate a COM component and identifies the caller, component, and permission involved.

**It does not necessarily mean an application is broken.** Some Windows components encounter this denial and then succeed through another supported path. Correlate the event with an actual application failure before considering permission changes. [Section 7.7](#77-event-10016--reading-it-correctly) explains how to read the message and decide whether to investigate.

**What this module covers**

What changes once an object lives in another process: the AppID and the process-wide settings configured under its registry key, session 0 isolation, and the authentication and impersonation levels set by `CoInitializeSecurity`. Then the security decisions themselves — UAC and integrity levels, Launch versus Access permissions and how to tell them apart from the symptom alone, and Event 10016 read correctly, including when *not* to act on it. It finishes with remote DCOM: the ports, the endpoint mapper, and the failures that only appear across a network.

> **DCOM** stands for **Distributed COM**. It is not a separate technology you opt into — it is the same COM from Modules 1–6, with the proxy/stub plumbing of Modules 3 and 4 carried over a **network transport** (Microsoft RPC) instead of staying inside one process. Your client code does not change at all: still `CoCreateInstance`, still `QueryInterface`, still the same interfaces. This illustrates **location transparency**, introduced in Module 0: once the client has an interface pointer, it calls the same methods whether the object runs in the same process, another local process, or on another machine.
>
> **RPC (Remote Procedure Call)** lets a program call a function in another process, on the same computer or across a network.
>
> What DCOM adds is everything a boundary forces you to answer:
>
> | Question | Answered by |
> |---|---|
> | How does the call physically get there? | MSRPC — TCP port 135 plus a dynamic port range (§7.8) |
> | **Who** is calling? | authentication levels (§7.4) |
> | May they **start** the server? May they **call** it? | launch and access permissions (§7.1, §7.6) |
> | Under **which account** does the server run? | the AppID's `RunAs` (§7.2) |
>
> One naming quirk to expect: in practice "DCOM" is used loosely for **any out-of-process COM configuration**, remote or not. That is why `dcomcnfg` and "DCOM permissions" (§7.6) apply just as much to a local `LocalServer32` on your own machine as to a server across the network.

**Contents**

- [7.1 The out-of-proc picture](#71-the-out-of-proc-picture)
- [7.2 The AppID](#72-the-appid)
- [7.3 Session 0 isolation](#73-session-0-isolation)
- [7.4 `CoInitializeSecurity`](#74-coinitializesecurity)
- [7.5 UAC, integrity levels, and elevation](#75-uac-integrity-levels-and-elevation)
- [7.6 DCOM permissions in practice](#76-dcom-permissions-in-practice)
- [7.7 Event 10016 — reading it correctly](#77-event-10016--reading-it-correctly)
- [7.8 Remote DCOM](#78-remote-dcom)
- [7.9 LAB 7.1 — Out-of-process hosting](#79-lab-71--out-of-process-hosting)
- [7.10 LAB 7.2 — Permissions and Event 10016](#710-lab-72--permissions-and-event-10016)
- [7.11 LAB 7.3 — Remote DCOM](#711-lab-73--remote-dcom)
- [7.12 Security checklist for reviewing a COM server](#712-security-checklist-for-reviewing-a-com-server)
- [7.13 The DCOM support triage flow](#713-the-dcom-support-triage-flow)
- [7.14 Checkpoint](#714-checkpoint)
- [7.15 Rules to carry forward](#715-rules-to-carry-forward)

---

## 7.1 The out-of-proc picture

**Out-of-proc** is short for **out-of-process**: the COM object runs in a separate process from the client application, for example in `CalcSrv.exe` or `dllhost.exe`. An **in-proc** object instead runs inside the client's own process, typically in a DLL loaded by the client.

```
     Client process                    SCM (RpcSs / DcomLaunch)           Server process
 ┌──────────────────┐              ┌───────────────────────────┐     ┌───────────────────────────┐
 │ CoInitializeEx   │              │                           │     │                           │
 │ CoCreateInstance │─────────────►│ 1. read HKCR\CLSID        │     │                           │
 │                  │              │ 2. read HKCR\AppID        │     │                           │
 │                  │              │ 3. check LaunchPerm       │     │                           │
 │                  │              │ 4. CreateProcessAsUser    │────►│  CoInitializeEx           │
 │                  │              │ 5. wait for class reg     │     │  CoRegisterClassObject()  │
 │                  │              │                           │◄────│                           │
 │                  │              └───────────────────────────┘     │                           │
 │      proxy       │◄── OXID/IPID: marshaled interface reference ───│  message loop             │
 │        │         │                                                │                           │
 │        └────────── RPC (local) / RPC over TCP (remote) ──────────►│  stub -> object           │
 └──────────────────┘                                                └───────────────────────────┘
```

**Where the stub fits:** the **proxy** in the client process packages the method arguments for transmission. The **stub** in the server process unpacks them and calls the real object's method. Results travel back through the stub and proxy. The SCM helps with activation; it is not in the path of each method call.

**Why both sides call `CoInitializeEx`:** it initializes COM for the **calling thread**, not for the whole application. The client's thread calls it before `CoCreateInstance`. When COM starts the server EXE, the server's own startup code calls it on its thread before `CoRegisterClassObject`. The client's initialization does not initialize the server. The SCM starts the process; it is the server's code that calls `CoInitializeEx`.

**What `CoRegisterClassObject` does**

The EXE server calls it to tell COM: **"This running process can supply objects for this CLSID; here is the factory to use."** A **class factory** is a COM object implementing `IClassFactory`; its `CreateInstance` method creates or supplies an instance of a class, such as your calculator. Here, the "class object" being registered is that factory, **not a calculator instance**.

This fills in the gap between starting the server and returning a proxy in the diagram:

1. The server initializes COM and constructs its class factory.
2. It calls `CoRegisterClassObject`, passing the **CLSID**, a pointer to the **factory**, and `CLSCTX_LOCAL_SERVER` to register it for out-of-process activation.
3. Once the registration is available to clients, COM can connect the client's activation request to the factory. The client's `CoCreateInstance` uses `IClassFactory::CreateInstance` to request the calculator interface.
4. The interface is marshaled back to the client, which receives a **proxy** for calling the calculator in the server process.

If the server registers with `REGCLS_SUSPENDED`, it must also call `CoResumeClassObjects` to make the suspended registrations available. This lets it finish initialization before accepting activation requests.

**Do not confuse the two kinds of registration:**

| Registration | What it tells COM | When it happens |
|---|---|---|
| Registry entries such as `LocalServer32` | Where to find the EXE to start | Installation, or the server's `-RegServer` command |
| `CoRegisterClassObject` | Which live factory can supply objects for a CLSID | Each time the EXE server starts and prepares to serve clients |

`CoRegisterClassObject` does **not** write those registry entries, and it does **not** create the calculator itself. For comparison, an in-process DLL exposes its factory through `DllGetClassObject`; an EXE server makes its factory available through this runtime registration.

On success, the function writes a **registration cookie** (a `DWORD` identifier) to its final output parameter. The server saves it and later passes it to `CoRevokeClassObject` to withdraw the factory registration. Revoking it does not delete the installed registry entries or destroy calculator objects already created.

**Support clue:** seeing the EXE running is not enough. If it never makes the expected factory registration available, activation can time out with `CO_E_SERVER_EXEC_FAILURE` (`0x80080005`). That is why step 5 in the diagram waits for class registration.

> **"SCM" here is COM's Service Control Manager** — the broker that maps a CLSID to a server, runs the security checks, starts the process, and hands the client back a marshaled reference.
>
> **It is a different component from the Windows Service Control Manager**, the one you drive with `sc.exe` and `Get-Service`. They share a name and an abbreviation and nothing else:
>
> | | COM's SCM | Windows' SCM |
> |---|---|---|
> | Job | turn a CLSID into a running object | start/stop/configure Windows **services** |
> | Lives in | `rpcss.dll`, hosted by the `RpcSs` / `DcomLaunch` services | `services.exe` |
> | Reads | `HKCR\CLSID`, `HKCR\AppID` | `HKLM\SYSTEM\CurrentControlSet\Services` |
> | You reach it via | `CoCreateInstance`, `CoGetClassObject` | `sc.exe`, `Get-Service`, `OpenSCManager` |
>
> **They meet in exactly one place:** the `LocalService` value on an AppID (§7.2). When a COM server is packaged as a Windows service, COM's SCM does not call `CreateProcess` — it asks *Windows'* SCM to start that service, then waits for it to call `CoRegisterClassObject`. That single hand-off is why a COM activation can fail with a plain service error such as `0x80070422` (service disabled).

> **The identifiers in that diagram.** An **OXID** (Object Exporter ID) identifies the *apartment* that hosts an object, and is what the client's RPC layer resolves into an actual binding — a machine, a protocol, and an endpoint. An **IPID** (Interface Pointer ID) identifies one specific *interface pointer* on one specific object within that apartment. Together they are the wire form of "which object, where." You'll meet both in WinDbg when inspecting proxies (Module 8), and `RPC_E_DISCONNECTED` means the OXID no longer resolves — the hosting apartment is gone.

Two distinct security checks happen:

1. **Launch/Activation permission** — step 3 in the diagram. *May this client **start** this server, or activate an object inside it?* The SCM checks the caller's token against the **`LaunchPermission`** security descriptor stored on the AppID key (§7.2), and it does so **before** `CreateProcess` — so a failure here means the server process never even starts. If the AppID carries no `LaunchPermission` value, the machine-wide default under `HKLM\SOFTWARE\Microsoft\Ole` applies instead.
2. **Access permission** — may this client *call* into the running server? Checked per-call by the RPC layer, against `AccessPermission`.

> **The `LaunchPermission` setting covers four separate permissions:** *Local Launch*, *Remote Launch*, *Local Activation*, and *Remote Activation*. They are granted and denied independently, which is why "it works on the box but not from another machine" is a permissions answer at least as often as it is a firewall one. `dcomcnfg` shows all four as separate checkboxes (§7.6).

They are configured separately and fail differently. Confusing them is the most common diagnostic error.

---

## 7.2 The AppID

Recall from Module 2: **CLSID is per-class; AppID is per-process.** Multiple CLSIDs implemented by the same EXE share one AppID, and process-wide settings live there.

```
HKCR\CLSID\{CLSID}
    AppID = "{APPID}"                <- named value pointing at the AppID

HKCR\AppID\{APPID}
    (Default)             = "My Server"
    RunAs                 = "Interactive User" | "NT AUTHORITY\LocalService" | "DOMAIN\svcacct"
    DllSurrogate          = ""                       <- empty = use dllhost.exe
    LocalService          = "MyServiceName"           <- run as a Windows service
    ServiceParameters     = "-Embedding"
    LaunchPermission      = <binary security descriptor>
    AccessPermission      = <binary security descriptor>
    AuthenticationLevel   = dword
    RunAsSystemService    = ...
    SRPTrustLevel         = ...

HKCR\AppID\MyServer.exe
    AppID = "{APPID}"                <- lets the SCM map an EXE name to its AppID
```

### `RunAs` values

| Value | Server runs as | Notes |
|---|---|---|
| *(absent)* | **The activating user (launching user)** | A separate server instance per user. Default. |
| `Interactive User` | The user logged on at the console | **One instance**, in the interactive session. Fails if nobody is logged on. Cannot be reached from session 0 reliably. |
| `NT AUTHORITY\LocalService` | LocalService | Low privilege, no network identity |
| `NT AUTHORITY\NetworkService` | NetworkService | Machine account on the network |
| `NT AUTHORITY\SYSTEM` | LocalSystem | Only via `LocalService`/service configuration in modern Windows |
| `DOMAIN\user` | That account | **Password stored in LSA secrets** as `SCM:{appid}`. Password expiry silently breaks activation → `CO_E_SERVER_EXEC_FAILURE`. |

> **Support fact:** a service account password change that nobody updated in `dcomcnfg` is a classic cause of `0x80080005` appearing "suddenly, with no changes."

---

## 7.3 Session 0 isolation

Since Windows Vista:

- **Session 0** contains services and system processes. **No interactive desktop.**
- **Sessions 1, 2, …** contain user logons.

Consequences you must know cold:

- A service **cannot** display UI. A modal dialog in session 0 blocks forever, invisible to everyone.
- A service activating a COM server with `RunAs = Interactive User` gets an object in a *different session*. Cross-session activation is restricted and frequently fails or hangs.
- If no user is logged on, `Interactive User` activation fails outright (`CO_E_SERVER_EXEC_FAILURE` or `E_ACCESSDENIED`).
- This is a primary reason **server-side Office automation is unsupported** (Module 5).

**Diagnostic:** in Task Manager, add the **Session** column. A COM server sitting in session 0 that was supposed to be interactive — or a `dllhost.exe` piling up in session 0 — is the signature.

---

## 7.4 `CoInitializeSecurity`

Sets the process-wide security policy for COM: access permissions and authentication requirements for incoming calls, plus authentication and impersonation defaults for outgoing calls. It establishes policy at startup; it does not itself impersonate anyone.

### How it differs from `CoInitializeEx`

Two unrelated jobs that happen to sit next to each other in startup code:

| | `CoInitializeEx` | `CoInitializeSecurity` |
|---|---|---|
| Scope | per **thread** | per **process** |
| Decides | which apartment this thread joins (§3.4) | incoming-call security and default settings for outgoing calls |
| Who calls it | **every** thread that touches COM | **one** thread, once |
| Mandatory? | yes — or you get `CO_E_NOTINITIALIZED` | no — COM will pick defaults for you |

### Who should call it

- **Out-of-proc server EXEs — yes, explicitly.** You are a callable surface, so state your own terms instead of inheriting whatever the registry happens to say.
- **Clients — only when they need to.** To raise the authentication level, pick an impersonation level, or set `EOAC_*` flags. Note that a client receiving **callbacks** (a Module 5 sink) is itself a server, and needs this too.
- **In-proc DLLs — no.** You do not own the process. It is usually too late by the time you load, and if it isn't, you have just silently rewritten your host's security policy.

### When to call it

After `CoInitializeEx` on that thread, and **before any other COM call** — before activating anything or marshaling any interface.

```
CoInitializeEx(nullptr, COINIT_MULTITHREADED);   // 1. per-thread: pick an apartment
CoInitializeSecurity(...);                       // 2. once per process: set the policy
// ... only now do any COM work
```

**If you never call it,** COM initializes security automatically when an interface is first marshaled or unmarshaled, using the applicable AppID settings and registry defaults under `HKLM\SOFTWARE\Microsoft\Ole`. That is perfectly acceptable for many clients, but it is **one-shot**. Once initialized, explicitly or implicitly, the process defaults cannot be reinitialized: a later call returns **`RPC_E_TOO_LATE` (`0x80010119`)**. A library may already have initialized security or caused COM to initialize it implicitly before your code gets here.

### Does this override the AppID settings?

**Yes, for the process's COM call-security policy, but not for every AppID setting.** When an early, successful `CoInitializeSecurity` call supplies the security values directly, those values take precedence over the corresponding registry settings. The call affects **its own process**; it neither edits the registry nor changes another process's policy.

| Setting | Can `CoInitializeSecurity` replace it? |
|---|---|
| AppID `AccessPermission` | **Yes.** The server can supply its own access policy for incoming calls instead. |
| AppID `AuthenticationLevel` | **Yes.** The process can specify its authentication default in code, subject to Windows' minimum requirements. |
| AppID `LaunchPermission` | **No.** COM still checks launch/activation permissions independently of the server's call-security policy. |
| AppID `RunAs`, `LocalService`, or `DllSurrogate` | **No.** This function does not choose the server's account or how it is hosted. |
| `MachineAccessRestriction` and `MachineLaunchRestriction` | **No.** Machine-wide restrictions remain an upper limit on what the application can allow. |

For example, a server can use code to decide who may call its objects, but that does not grant a denied client **Remote Activation** permission. Likewise, a client's call to `CoInitializeSecurity` does not override the server's access policy.

**You can also explicitly choose the AppID policy:** set `EOAC_APPID` and pass a pointer to the AppID GUID as the first argument. In that mode, COM reads the registry policy and ignores the other arguments apart from those selecting the AppID mode; it does not combine registry settings with authentication levels supplied in the call.

### Example: client authentication defaults

This is a **startup excerpt for a native client EXE** that makes outgoing COM calls and does not expose objects or receive callbacks. Place it after successful `CoInitializeEx` and before activation or marshaling. It configures this client's defaults, not the server's policy.

```cpp
HRESULT hr = CoInitializeSecurity(
    nullptr,
    -1,
    nullptr,
    nullptr,
    RPC_C_AUTHN_LEVEL_PKT_INTEGRITY,
    RPC_C_IMP_LEVEL_IDENTIFY,
    nullptr,
    EOAC_NONE,
    nullptr);
```

Check `hr` and stop initialization on failure rather than continuing to activate objects with an unintended security policy.

- **`RPC_C_AUTHN_LEVEL_PKT_INTEGRITY`** requests authenticated calls with tamper detection. It does not encrypt the call data.
- **`RPC_C_IMP_LEVEL_IDENTIFY`** lets the receiving server identify this client and check its permissions, but not open files or other resources as this client. It does not grant the client permission to impersonate the server.
- These are **process defaults**. A client can configure a particular proxy with `CoSetProxyBlanket`, which sets that proxy's call-security settings without reinitializing the process. The receiving server's requirements and Windows restrictions still apply.

**This is not a server access-policy example.** With `EOAC_NONE`, the first `nullptr` supplies no application-level access restriction; it does **not** mean "use AppID/defaults." Authentication requirements and machine-wide restrictions still apply. A process accepting calls, including callbacks, must also choose an appropriate incoming-call access policy. The server example below shows how to select that policy from the AppID.

See the [CoInitializeSecurity API reference](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-coinitializesecurity) for the `EOAC_APPID` and null-descriptor rules.

### Authentication levels

| Level | Value | Meaning |
|---|---|---|
| `RPC_C_AUTHN_LEVEL_NONE` | 1 | No authentication |
| `RPC_C_AUTHN_LEVEL_CONNECT` | 2 | Authenticate at connect only |
| `RPC_C_AUTHN_LEVEL_CALL` | 3 | Authenticate each call |
| `RPC_C_AUTHN_LEVEL_PKT` | 4 | Authenticate each packet |
| **`RPC_C_AUTHN_LEVEL_PKT_INTEGRITY`** | **5** | **+ tamper detection — the modern minimum** |
| `RPC_C_AUTHN_LEVEL_PKT_PRIVACY` | 6 | + encryption |

### Impersonation levels

The impersonation argument to `CoInitializeSecurity` sets what other servers may do on behalf of **this process when it makes outgoing calls**. A server cannot increase the authority granted by an incoming caller merely by setting a higher level in its own initialization.

| Level | Server may… |
|---|---|
| `RPC_C_IMP_LEVEL_ANONYMOUS` | not identify the client |
| `RPC_C_IMP_LEVEL_IDENTIFY` | see the client's identity and ACL-check, but not act as them |
| `RPC_C_IMP_LEVEL_IMPERSONATE` | act as the client **on the local machine** |
| `RPC_C_IMP_LEVEL_DELEGATE` | act as the client **on other machines** (requires delegation config; grant sparingly) |

Rule: **give the minimum that works.** `IDENTIFY` is enough for "check whether this caller is allowed"; `IMPERSONATE` is needed only if the server opens resources as the client.

### Useful `EOAC_*` capability flags

| Flag | Effect |
|---|---|
| `EOAC_STATIC_CLOAKING` / `EOAC_DYNAMIC_CLOAKING` | Outgoing calls use the **thread** token rather than the process token |
| `EOAC_NO_CUSTOM_MARSHAL` | Refuse `IMarshal` — blocks a real privilege-escalation vector (Module 4) |
| `EOAC_DISABLE_AAA` | Disallow "activate-as-activator" — prevents a server from being launched under the caller's identity |
| `EOAC_SECURE_REFS` | Authenticate `AddRef`/`Release` so an unauthenticated party can't release your objects out from under you |

Hardened services set `EOAC_NO_CUSTOM_MARSHAL | EOAC_DISABLE_AAA | EOAC_SECURE_REFS` with `PKT_INTEGRITY` or higher.

### Example: server initialization using the AppID

An EXE server also initializes security **once at startup**, after `CoInitializeEx` and before making its class factories available. It can supply its own access policy in code or explicitly select the administrator-configured AppID policy. This excerpt demonstrates the second choice:

```cpp
HRESULT hr = CoInitializeSecurity(
    nullptr,
    -1,
    nullptr,
    nullptr,
    RPC_C_AUTHN_LEVEL_DEFAULT,
    RPC_C_IMP_LEVEL_IDENTIFY,
    nullptr,
    EOAC_APPID,
    nullptr);
```

With **`EOAC_APPID`**, the first `nullptr` tells COM to find the AppID through the executable-name mapping, for example the named `AppID` value under `HKCR\AppID\CalcSrv.exe`. That mapping must already exist; this call does not create it. Alternatively, pass a pointer to the AppID GUID as the first argument.

In this mode the remaining security arguments are ignored, including the authentication and impersonation values shown above. COM reads the applicable registry policy instead. Check `hr` and do not publish the factories if security initialization fails. The server cannot use this call to bypass launch permissions or machine-wide restrictions.

**Initialization is separate from handling an individual call.** Later, inside a server method, `CoQueryClientBlanket` can inspect the current call's security information. `CoImpersonateClient` lets the executing thread use the caller's identity to the extent that caller allowed; opening a file as the caller requires at least `IMPERSONATE`, not the client example's `IDENTIFY`. Every successful impersonation must be paired with `CoRevertToSelf`, including on error paths. These are **per-call operations**, not replacements for `CoInitializeSecurity`, and initializing security does not automatically perform them.

---

## 7.5 UAC, integrity levels, and elevation

Windows integrity levels: **Untrusted < Low < Medium < High < System.**

- A normal user process runs at **Medium**.
- An elevated (admin) process runs at **High**.
- AppContainer/sandboxed processes run at **Low** or **AppContainer**.

**A lower-integrity process cannot open a higher-integrity process's objects.** Applied to COM:

- A Medium-IL client activating an in-proc server: fine (same process).
- A Medium-IL client calling into a High-IL running server: the server's ACL and the process-level mandatory policy typically block it → `E_ACCESSDENIED`.
- A Medium-IL client **launching** a server that runs at High: requires elevation.

### The COM Elevation Moniker

The supported way to obtain an elevated COM object from a non-elevated process.

> **Moniker refresher.** A moniker is an object that *names* another object and knows how to resolve that name into it. `CoGetObject` parses a display name into a moniker and binds it in one call. Here the display name `Elevation:Administrator!new:{CLSID}` means "a new instance of this CLSID, elevated." Full treatment in [Appendix A §A.1](appendix-a-monikers-and-persistence.md#a1-monikers).

```cpp
HRESULT CoCreateInstanceAsAdmin(HWND hwnd, REFCLSID rclsid, REFIID riid, void** ppv)
{
    WCHAR clsidStr[64];
    StringFromGUID2(rclsid, clsidStr, ARRAYSIZE(clsidStr));

    WCHAR monikerName[300];
    StringCchPrintfW(monikerName, ARRAYSIZE(monikerName),
                     L"Elevation:Administrator!new:%s", clsidStr);

    BIND_OPTS3 bo = {};
    bo.cbStruct = sizeof(bo);
    bo.hwndApp = hwnd;                       // parent for the UAC prompt
    bo.dwClassContext = CLSCTX_LOCAL_SERVER;

    return CoGetObject(monikerName, &bo, riid, ppv);
}
```

Requirements on the server side, or the moniker fails:

```
HKCR\CLSID\{CLSID}
    (Default)                = "My Elevated Component"
    LocalizedString          = "@C:\Path\Server.exe,-101"    <- REQUIRED: shown in the UAC prompt
    └── Elevation
            Enabled          = dword:00000001                <- REQUIRED
            IconReference    = "@C:\Path\Server.exe,-201"

HKCR\AppID\{APPID}
    (Default)                = "My Elevated Component"
```

- `LocalizedString` must be an **indirect string** (`@file,-resourceID`) so the UAC dialog can display a trustworthy name. A plain string is a common mistake and causes the elevation to fail.
- The server EXE should be **Authenticode signed**, or the UAC prompt shows the scary "unidentified publisher" banner.
- The server must be a `LocalServer32` — you cannot elevate an in-proc DLL.

**Failure codes:** `E_ACCESSDENIED` (`0x80070005`) if the registration is incomplete; `ERROR_CANCELLED` (`0x800704C7`) if the user declined the prompt. Distinguishing those two is a common support question.

---

## 7.6 DCOM permissions in practice

### Default machine-wide settings

```
HKLM\SOFTWARE\Microsoft\Ole
    DefaultLaunchPermission       REG_BINARY  <SD>
    DefaultAccessPermission       REG_BINARY  <SD>
    MachineLaunchRestriction      REG_BINARY  <SD>   <- a CEILING, applied first
    MachineAccessRestriction      REG_BINARY  <SD>   <- a CEILING, applied first
    EnableDCOM                    REG_SZ      "Y"
    LegacyAuthenticationLevel     REG_DWORD
    LegacyImpersonationLevel      REG_DWORD
```

**The evaluation order matters and is frequently misunderstood:**

1. **Machine restriction** (`MachineLaunchRestriction` / `MachineAccessRestriction`) — a hard ceiling. If it denies, nothing else can grant.
2. **AppID-specific permission** (`LaunchPermission` / `AccessPermission` on the AppID).
3. **Default permission** (`DefaultLaunchPermission` / `DefaultAccessPermission`) if the AppID has none.

So granting a user permission on the AppID has **no effect** if the machine-wide limit denies them. In `dcomcnfg` these are the "Limits" vs "Default" tabs — and support engineers regularly edit the wrong one.

### `dcomcnfg` walkthrough

**`dcomcnfg`** is a Windows command that opens the **Component Services** administration console. It provides a graphical interface for configuring COM/DCOM settings, including launch and access permissions and the account a server runs under.

`dcomcnfg` → Component Services → Computers → My Computer:

- **Right-click → Properties → Default Properties**: `EnableDCOM`, default authentication/impersonation levels.
- **Default COM Security** tab: four buttons — Access Permissions (Edit Limits / Edit Default), Launch and Activation Permissions (Edit Limits / Edit Default). *Limits = machine ceiling. Default = fallback when the AppID has none.*
- **DCOM Config** node: per-AppID settings. Right-click a component → Properties → Security tab → per-AppID Launch/Access/Configuration permissions; Identity tab → `RunAs`.

Each permission set has four rights:

| Right | Meaning |
|---|---|
| **Local Launch** | Start the server from this machine |
| **Remote Launch** | Start the server from another machine |
| **Local Activation** | Activate an object in an already-running server, locally |
| **Remote Activation** | Same, remotely |

Access permissions have **Local Access** and **Remote Access**.

### Reading permissions from PowerShell

```powershell
$appid = "{APPID-GUID}"
$key = Get-ItemProperty "HKLM:\SOFTWARE\Classes\AppID\$appid" -EA SilentlyContinue
if ($key.LaunchPermission) {
    $sd = New-Object System.Security.AccessControl.RawSecurityDescriptor(
              $key.LaunchPermission, 0)
    $sd.DiscretionaryAcl | ForEach-Object {
        [pscustomobject]@{
            Identity = try { $_.SecurityIdentifier.Translate(
                             [System.Security.Principal.NTAccount]) }
                       catch { $_.SecurityIdentifier }
            Type     = $_.AceType
            Mask     = "0x{0:X}" -f $_.AccessMask
        }
    }
}
```

Access-mask bits for DCOM: `1` = Execute, `2` = Local Launch (Execute_Local), `4` = Remote Launch, `8` = Local Activation, `16` = Remote Activation.

---

## 7.7 Event 10016 — reading it correctly

The most-reported and most-misunderstood DCOM event.

```
Log Name:      System
Source:        Microsoft-Windows-DistributedCOM
Event ID:      10016
Description:
The application-specific permission settings do not grant Local Activation
permission for the COM Server application with CLSID
    {2593F8B9-4EAF-457C-B68A-50F6B8EA6B54}
 and APPID
    {15C20B67-12E7-4BB6-92BB-7AFF07997402}
to the user NT AUTHORITY\SYSTEM SID (S-1-5-18) from address LocalHost (Using LRPC)
running in the application container Unavailable SID (Unavailable).
This security permission can be corrected using the Component Services
administrative tool.
```

### How to read it

| Field | Tells you |
|---|---|
| `Local Activation` vs `Local Launch` vs `Remote …` | **Which right** is missing |
| CLSID | The class — look it up to identify the component |
| APPID | The application's configuration to inspect |
| user | **Who** was denied |
| address / LRPC vs TCP | Local or remote |
| application container | Whether an AppContainer/packaged app was involved |

### The critical judgement call

**Many 10016 events are benign and should NOT be "fixed."** Windows components frequently attempt an activation, get denied, and fall back to a supported path. Microsoft has explicitly documented several 10016 events as safe to ignore.

Loosening permissions on a system AppID — which requires taking ownership of the registry key from `TrustedInstaller` — **weakens the machine's security posture** and can break Windows updates. Do not recommend it reflexively.

**Decision procedure:**

1. Is there a **functional symptom** (a feature actually failing) correlated in time with the event? If not → document as benign, do not change permissions.
2. Identify the component from the CLSID:
   ```powershell
   $clsid = "{2593F8B9-4EAF-457C-B68A-50F6B8EA6B54}"
   (Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\$clsid" -EA SilentlyContinue).'(default)'
   (Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\$clsid\LocalServer32" -EA SilentlyContinue).'(default)'
   (Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32" -EA SilentlyContinue).'(default)'
   ```
3. If it's a **Microsoft OS component** → check Microsoft's documented list of ignorable 10016s. Default answer: leave it alone.
4. If it's a **third-party or in-house component** and the intended caller should have the missing permission → follow the component's documented requirements and grant only the required right. Never "Everyone / Full Control."

### How to fix a confirmed permission failure

For the event above, the fix logic would be:

```text
If the event is harmless:
    Leave the permissions unchanged.
If the denial causes a real failure and component-specific guidance supports a change:
    Find AppID {15C20B67-12E7-4BB6-92BB-7AFF07997402} in dcomcnfg.
    Check that machine-wide launch/activation limits permit the request.
    Record the existing AppID permissions.
    In that AppID's Launch and Activation Permissions, allow:
        Account:    NT AUTHORITY\SYSTEM (S-1-5-18)
        Permission: Local Activation
    Preserve the other permissions and retest the failing operation.
    If it still fails, undo the change and investigate further.
```

The event therefore tells you **which application**, **which account**, and **which permission** to investigate. It does not call for granting every permission. Because this example concerns a Windows component, follow Microsoft's guidance rather than changing protected permissions just to remove the event.

---

## 7.8 Remote DCOM

**Remote DCOM** lets a client use a COM object running on **another computer**, rather than in another process on the same computer. For example, a client on machine A can call a calculator hosted on machine B. The client calls a local proxy, which communicates with the remote server over RPC. This requires network connectivity, authentication, and appropriate remote launch, activation, and access permissions.

### Activation

```cpp
COSERVERINFO si = {};
si.pwszName = L"RemoteMachine";      // or an FQDN / IP

COAUTHINFO ai = {};
COAUTHIDENTITY id = {};
// optionally fill in explicit credentials:
// id.User = ..., id.Domain = ..., id.Password = ..., id.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
ai.dwAuthnSvc = RPC_C_AUTHN_WINNT;
ai.dwAuthzSvc = RPC_C_AUTHZ_NONE;
ai.dwAuthnLevel = RPC_C_AUTHN_LEVEL_PKT_INTEGRITY;
ai.dwImpersonationLevel = RPC_C_IMP_LEVEL_IMPERSONATE;
ai.pAuthIdentityData = &id;
si.pAuthInfo = &ai;

MULTI_QI mqi[1] = {};
mqi[0].pIID = &IID_ICalculator;

HRESULT hr = CoCreateInstanceEx(CLSID_Calculator, nullptr,
                                CLSCTX_REMOTE_SERVER, &si, 1, mqi);
if (SUCCEEDED(hr) && SUCCEEDED(mqi[0].hr))
    auto* p = static_cast<ICalculator*>(mqi[0].pItf);   // remember to Release
```

**`MULTI_QI` is a performance feature:** it fetches several interfaces in one network round trip instead of N `QueryInterface` calls. Each entry has its own `hr` — check both the outer `hr` and each `mqi[i].hr`.

### Network requirements

| Port / range | Purpose |
|---|---|
| **TCP 135** | RPC Endpoint Mapper — the initial contact |
| **Dynamic range**: TCP 49152–65535 (Win Vista+); 1024–5000 (legacy) | Actual DCOM traffic, negotiated after 135 |
| TCP 445 | SMB, often needed alongside for authentication |

The dynamic range is why "just open 135" never works. Restrict it explicitly:

```powershell
# Narrow the dynamic RPC port range (requires a restart)
netsh int ipv4 set dynamicport tcp start=50000 num=100
netsh int ipv4 show dynamicport tcp

# Then allow 135 and 50000-50099 through the firewall.
```

Or configure DCOM's own range via `HKLM\SOFTWARE\Microsoft\Rpc\Internet` (`Ports`, `PortsInternetAvailable`, `UseInternetPorts`).

### Diagnostics for remote DCOM

```powershell
Test-NetConnection RemoteMachine -Port 135

# Enumerate RPC endpoints on the remote machine (needs RPC tools / PortQry)
portqry -n RemoteMachine -e 135

# WMI over DCOM is a convenient end-to-end test of the whole stack:
Get-CimInstance -ComputerName RemoteMachine -ClassName Win32_OperatingSystem `
                -Protocol DCOM
```

If WMI-over-DCOM works but your component doesn't, the transport is fine and the problem is AppID permissions or registration. **That single test bisects the problem space in one command** — make it your first move on remote DCOM tickets.

### DCOM hardening (CVE-2021-26414)

Microsoft raised the minimum authentication level for DCOM activation. Since the phased rollout completed, servers **reject activations below `RPC_C_AUTHN_LEVEL_PKT_INTEGRITY`**.

Registry control (client and server):

```
HKLM\SOFTWARE\Microsoft\Ole\AppCompat
    RequireIntegrityActivationAuthenticationLevel   REG_DWORD
        0 = disabled (hardening off)   - the temporary escape hatch, now removed
        1 = enabled  (hardening on)
```

**Symptom:** after patching, remote activation fails with `E_ACCESSDENIED` (`0x80070005`) or `RPC_E_ACCESS_DENIED`, and the **server** logs:

```
Event 10036, DistributedCOM:
"The server-side authentication level policy does not allow the user
 <domain\user> SID (…) from address <ip> to activate DCOM server.
 Please raise the activation authentication level at least to
 RPC_C_AUTHN_LEVEL_PKT_INTEGRITY in client application."
```

**Event 10037** is logged on the *client* side when it is the one making the too-weak request.

**The correct fix is on the client**: raise the authentication level, either via `CoInitializeSecurity` in code, or per-AppID/machine defaults. Disabling the hardening is not an option on current builds. Older third-party products that hardcode `RPC_C_AUTHN_LEVEL_CONNECT` need a vendor update — recognizing this saves days of misdirected work.

---

## 7.9 LAB 7.1 — Out-of-process hosting

> **Requirements**
> - **Tools:** Visual Studio with Desktop development with C++ and a Windows SDK; [Process Explorer](https://learn.microsoft.com/en-us/sysinternals/downloads/process-explorer); 64-bit PowerShell; `dcomcnfg` for inspection.
> - **Elevation:** required for EXE, DLL, proxy/stub, and surrogate registration.
> - **Bitness:** x64 throughout the main lab. The x86 comparison is optional.
> - **Depends on:** the earlier concepts, but **no earlier lab projects, binaries, or registrations**.
> - **Starting point:** [`labs/stage-5-exe-server/`](../labs/stage-5-exe-server/) contains one solution with an EXE server, a DLL, a client, and their matching proxy/stub.
> - **Caution:** the setup changes only this lab's registration. Do not change system-component permissions, `RunAs`, or machine-wide DCOM settings.
> - **Time:** ~2–3 h; allow another ~30 min for the optional cross-bitness comparison.

This lab compares two ways to host an object outside its client: **your own EXE server**, and a **DLL loaded by Windows' `dllhost.exe` surrogate**. Both require marshaling across the process boundary. The main difference is who supplies the host process and its lifetime machinery.

### Part A: run the EXE server

1. Open **Stage5.sln** from the starting folder. Build **Debug | x64**. It builds all four projects, including the proxy/stub generated from this lab's two-method IDL.

    | Project | Purpose |
    |---|---|
    | `CalcSrv` | The calculator EXE server: runs in its own process and manages its factories and lifetime. |
    | `CalcDll` | The calculator DLL: runs inside the client or in `dllhost.exe` for Part B's comparison. |
    | `CalcSrvClient` | The test client: selects a host and calls the calculator's `Add` and `Subtract` methods. |
    | `Stage5PS` | The proxy/stub DLL: generated from the IDL, it marshals calculator calls across process boundaries. |

2. In an **elevated 64-bit PowerShell** window, change to the [labs/stage-5-exe-server](../labs/stage-5-exe-server/) folder (not its `x64` subfolder) and run `.\Setup.ps1 -Action Register`. This registers both hosting options and the proxy/stub; no registration from previous labs is used.
3. In an **ordinary PowerShell** window in the same folder, run `.\x64\CalcSrvClient.exe`. Both `Add(40,2)` and `Subtract(44,2)` should return `42` with `hr=0x00000000`. The client then displays `Press Enter to release the object...` and pauses while still holding the calculator object, keeping the server alive.

    **Before pressing Enter**, open Process Explorer and find `CalcSrvClient.exe` and `CalcSrv.exe`. They have different process IDs: the calculator runs in the server process, not inside the client. Return to the client's PowerShell window and press Enter. The client releases the object and exits; if no other client is using this server, `CalcSrv.exe` should exit too.

The complete build, troubleshooting, and cleanup instructions are in [Stage 5's README](../labs/stage-5-exe-server/README.md). Follow the startup path in [CalcSrv.cpp](../labs/stage-5-exe-server/CalcSrv.cpp):

| Code to find | Responsibility |
|---|---|
| `wWinMain`'s `-Embedding` check | Distinguishes a COM launch from a manual launch |
| `CoInitializeEx` | Initializes the main thread's MTA participation |
| `CoInitializeSecurity` with `EOAC_APPID` | Reads this EXE's AppID call-security policy before marshaling begins |
| `PeekMessageW` | Creates the main thread's message queue before calls can arrive |
| `CoRegisterClassObject` | Publishes the factory in a suspended state |
| `CoResumeClassObjects` | Starts accepting activations after initialization succeeds |
| `GetMessageW` | Keeps the main thread alive until it receives the shutdown message |
| `CoRevokeClassObject` and `CoUninitialize` | Withdraw the factory and finish COM cleanup |

The calculator's constructor/destructor and the factory's `LockServer` method update the process lock count through `ServerLock` and `ServerUnlock`. When the count reaches zero, `ServerUnlock` suspends activation and posts the shutdown message to the **main thread**, even when the last release arrives on a COM worker thread. Creating that thread's message queue before publishing the factory prevents a startup race in which the shutdown message could be lost.

### `REGCLS` flags

| Flag | Meaning |
|---|---|
| `REGCLS_SINGLEUSE` | Removes the class registration after one connection to its class object; this is not an object-count limit |
| `REGCLS_MULTIPLEUSE` | Many objects per process; also registers for in-proc use |
| `REGCLS_MULTI_SEPARATE` | Many objects, but **not** registered for in-proc — the usual choice for an EXE server |
| `REGCLS_SUSPENDED` | Register but don't serve until `CoResumeClassObjects` |
| `REGCLS_SURROGATE` | For DLL surrogates |

`REGCLS_SUSPENDED` + `CoResumeClassObjects` exists to close a real race: register all your class objects first, then start serving, so a client can't activate class A while class B is still unregistered.

### Registration

The EXE registration has this shape; `{CLSID}` and `{APPID}` are this lab's dedicated EXE identifiers, listed in the [registration reference](../labs/stage-5-exe-server/README.md#registration-reference). The path below is illustrative; setup records the actual build path.

```text
HKCR\CLSID\{CLSID}\LocalServer32
    (Default) = "C:\Components\CalcSrv.exe"
HKCR\CLSID\{CLSID}
    AppID     = "{APPID}"
HKCR\AppID\{APPID}
    (Default) = "Stage 5 Calculator EXE"
```

Unlike the `nullptr` example in §7.4, this lab passes its AppID GUID directly to `CoInitializeSecurity` with `EOAC_APPID`. It therefore does not need an `AppID\CalcSrv.exe` lookup entry and cannot overwrite another sample's filename-based mapping. The CLSID-to-AppID link above is still needed for activation policy.

### Part A exercises

1. **Test lifetime with two clients.** From the lab folder, run `.\x64\CalcSrvClient.exe` in two ordinary PowerShell windows, leaving both at the Enter prompt. In Process Explorer, confirm there are two client processes but only one `CalcSrv.exe`. Press Enter in one client: the server should stay alive because the other still holds an object. Press Enter in the remaining client: the server should now exit.
2. **Follow activation in source.** Find these functions in [CalcSrv.cpp](../labs/stage-5-exe-server/CalcSrv.cpp) and match them to their roles:

    | Function | What it does in this server |
    |---|---|
    | `CoRegisterClassObject` | Registers the existing factory for the calculator's CLSID. Because this call uses `REGCLS_SUSPENDED`, clients cannot use the registration yet. |
    | `CoResumeClassObjects` | Makes the suspended factory registration available so COM can route activation requests to it. |
    | `CalculatorFactory::CreateInstance` | Creates a calculator object and uses `QueryInterface` to return the interface the client requested. |
    | `ServerUnlock` | Decrements the process lock count. When it reaches zero, suspends new activations and posts `WM_QUIT` to the main thread, letting it leave the message loop, clean up, and exit. |

3. **Confirm the marshaling dependency is local to this lab.** After closing clients, temporarily unregister **only** this lab's x64 proxy from elevated 64-bit PowerShell:

    ```powershell
    & "$env:SystemRoot\System32\regsvr32.exe" /u "$PWD\x64\Stage5PS.dll"
    ```

    Run the EXE client normally and record the failure. An expected result is:

    ```text
    CoCreateInstance failed: 0x80004002
    ```

    `0x80004002` is **`E_NOINTERFACE`**. In this experiment, the calculator class is still registered and implements `ICalculator`, but COM cannot marshal that interface back to the client without the proxy/stub registration. `CoCreateInstance` must return the requested interface, not just start the server. The client's next three lines are always-printed troubleshooting hints, **not three additional failures**.

    Restore the proxy/stub from **elevated 64-bit PowerShell in the same lab folder**:

    ```powershell
    & "$env:SystemRoot\System32\regsvr32.exe" "$PWD\x64\Stage5PS.dll"
    ```

    Run `.\x64\CalcSrvClient.exe` again from ordinary PowerShell; both calls should return `42` with `hr=0x00000000`. Do not unregister an earlier module's proxy. The exact failure stage may vary, but this custom interface cannot cross the process boundary without its marshaler.
4. **Observe a server failure.** Add `Sleep(10000);` at the beginning of `Calculator::Add` in [CalcSrv.cpp](../labs/stage-5-exe-server/CalcSrv.cpp) and rebuild **Debug | x64** with clients closed. Run `.\x64\CalcSrvClient.exe` and, during that ten-second delay, terminate **that CalcSrv.exe instance** using **Task Manager > Details > End task** or **Process Explorer > Kill Process**. Verify its executable path is your lab's `x64\CalcSrv.exe`; do not terminate the client. Do this **before the client prints the `Add` result**, not at its later Enter prompt.

    **Expected behavior:** the interrupted `Add` returns a failing HRESULT. The client then attempts `Subtract` on the same disconnected proxy, which should also fail. In a local test, `Add` returned `0x800706BE` (RPC call failed) and `Subtract` returned `0x800706BA` (RPC server unavailable). The exact codes can vary with timing; record yours and ignore the arithmetic result values when a call failed.

    **The client still waits for Enter after these failures.** Once `Press Enter to release the object...` appears, it is waiting for keyboard input, not for the dead server. Press Enter to release the proxy and exit, or run `.\x64\CalcSrvClient.exe --auto` to skip this final keyboard pause. `--auto` does not add a timeout to COM calls.

    If you terminate the server **after** both results and the Enter prompt appear, no method call is in progress: the client is already waiting for keyboard input, and killing the server does not end that wait or change the completed results. If there is **no result and no Enter prompt** after terminating the correct server, that is a different symptom from this intentional pause; record the last output and confirm the killed process's path and PID.

    A disconnected proxy cannot reconnect itself; a fresh activation is required. Remove the delay, rebuild, and run a new client to confirm both calls succeed again.

### Part B: compare with a DLL surrogate

A **surrogate** is a process that hosts a COM DLL on a client's behalf. This solution supplies its own DLL and matching marshaler, so you can revisit the idea from Lab 2.2 without depending on that lab's setup.

1. Run `.\x64\CalcSrvClient.exe --surrogate`. Both arithmetic results should still be `42`. Leave the client at its Enter prompt while identifying the host in Process Explorer:

    Press **Ctrl+F** (**Find > Find Handle or DLL**), search for `CalcDll.dll`, and find the **DLL** result whose full path matches this lab's x64 build, for example `C:\tmp\stage-5-exe-server\x64\CalcDll.dll`. Its **Process** should be `dllhost.exe`; note its **PID** (process ID). Do not select a result from another lab copy or the `x86` folder.

    Double-click that result to select the hosting process, then close the search window. **Ctrl+D** shows its DLL list; confirm the same full DLL path there. This identifies the correct host among the other `dllhost.exe` processes. The client's printed `Host: dllhost.exe` label alone is not proof.

    As a cross-check, open that process's **Properties > Image** and inspect **Command line**. For this lab's surrogate, it should contain `/Processid:{C714BFAA-5711-46A7-91E1-42E2B9A14920}`. Despite the switch's name, that GUID is the DLL's **AppID**, not the numeric PID. Both architectures use this AppID, so the loaded DLL's full path remains the distinguishing check.
2. Release the object with Enter. Run `.\x64\CalcSrvClient.exe --inproc` and inspect the client process: the same DLL is now loaded inside `CalcSrvClient.exe` instead of `dllhost.exe`. Windows manages the surrogate's idle shutdown; do not expect the EXE server's immediate exit timing or terminate unrelated `dllhost.exe` processes.
3. Compare [CalcSrvClient.cpp](../labs/stage-5-exe-server/CalcSrvClient.cpp)'s three activation choices. The EXE and DLL use different CLSIDs; the DLL's in-process and surrogate calls use the **same CLSID** with different `CLSCTX` flags. The setup remains unchanged throughout.
4. Run the default client again to confirm the EXE still works. Keep this registration for Labs 7.2 and 7.3. When finished with Module 7, run `.\Setup.ps1 -Action Unregister` from elevated 64-bit PowerShell, with clients closed.

**Optional cross-bitness extension:** build **Debug | x86** in the same solution, run `.\Setup.ps1 -Action Register -Architecture x86` from elevated 64-bit PowerShell, then run `.\x64\CalcSrvClient.exe --surrogate --x86` normally. Confirm a 64-bit client calls the DLL in a 32-bit `dllhost.exe`. Unregister the optional x86 build before the final x64 cleanup. The [Stage 5 README](../labs/stage-5-exe-server/README.md#optional-64-bit-client-32-bit-surrogate) has the complete commands.

**What this proves:** the client-facing interface can stay the same while the host changes. Out-of-process calls need marshaling; a DLL loaded directly into the client does not cross that process boundary. The optional extension additionally shows why marshaled calls can cross bitness.

**Deliverable:** compare the in-process DLL, EXE server, and DLL in `dllhost.exe`: name the host process, state whether marshaling is needed across processes, and explain who manages host lifetime. No custom surrogate, server-identity change, or VM checkpoint is needed.

---

## 7.10 LAB 7.2 — Permissions and Event 10016

> **Requirements**
> - **Tools:** `dcomcnfg` (Component Services), Event Viewer / `Get-WinEvent`, and a test account for the `RunAs` steps.
> - **Elevation:** required throughout.
> - **Bitness:** `dcomcnfg` shows the **64-bit** DCOM config. For a 32-bit AppID run `mmc comexp.msc /32` — a component that "isn't in the list" is usually this.
> - **Depends on:** Lab 7.1's x64 setup, which remains registered after the hosting comparison.
> - **Starting point:** [`labs/stage-5-exe-server/`](../labs/stage-5-exe-server/) — select **Stage 5 Calculator EXE**, AppID `{A57CC9FD-4A41-4579-9D75-97ECC20C504B}`, in `dcomcnfg`. Export that key before you touch it. The EXE reads its AppID policy through `CoInitializeSecurity` with `EOAC_APPID`; close existing clients and let it exit between tests so a new process reads the changed policy.
> - **Caution:** **VM or dedicated test machine only.** You are editing machine-wide DCOM ACLs. Export `HKLM\SOFTWARE\Classes\AppID\{your-appid}` before you start, and never "fix" a Microsoft-owned AppID this way — that is the single most common bad advice in COM support, and it is what §7.7 tells you not to do.
> - **Time:** ~2 h.

Launch and Access permissions are configured in different places, checked at different times, and confused constantly — including in a great deal of published advice.

Here you break each one deliberately so you learn to tell them apart **from the symptom alone**, before reading any log. Then you find the Event 10016 your own machine just generated and map every field back to what you changed.

1. In `dcomcnfg`, find your Calculator AppID → Properties → Security → **Launch and Activation Permissions** → Customize → Edit. **Remove your user account.**
2. Run the client. Expect `E_ACCESSDENIED` (`0x80070005`).
3. Open Event Viewer → Windows Logs → System, filter Source = `DistributedCOM`. Find **your** Event 10016. Read every field and map it to what you changed.
4. Now the discrimination drill — cause each of these and record how they differ:
   - Remove **Launch** permission → fails at activation, before the process starts.
   - Grant Launch but remove **Access** permission → the server *starts*, then the first call fails.
   - Set `RunAs` to a nonexistent account → `0x80080005 CO_E_SERVER_EXEC_FAILURE`.
   - Set `RunAs` to a valid account with a wrong stored password → also `0x80080005`, but the Application log shows a logon failure (Event 4625 in Security).
5. Fix it **correctly**: grant *Local Activation* to your specific account — not "Everyone", not "Full Control".
6. Write the resulting decision tree into your notes:

```
E_ACCESSDENIED on activation
├── Event 10016 present?
│   ├── Yes → read the missing right + principal → check MachineLaunchRestriction FIRST,
│   │         then the AppID permission, then the default
│   └── No  → is it an integrity-level problem? (medium client, high server)
│             or an NTFS ACL on the server binary?
└── Remote? → also check MachineAccessRestriction, firewall, and
              the authentication-level hardening (Event 10036/10037)
```

---

## 7.11 LAB 7.3 — Remote DCOM

> **Requirements**
> - **Machines:** **two** machines or VMs on the same network or domain — there is no single-box substitute for this lab.
> - **Tools:** firewall control on B (`New-NetFirewallRule`), `Test-NetConnection`, PortQry or `rpcdump`, Event Viewer on B, `dcomcnfg` on both.
> - **Elevation:** required on **both** machines.
> - **Bitness:** identical on both ends.
> - **Depends on:** Lab 7.1's x64 EXE server and this lab's `Stage5PS.dll` registered on **A and B**. Undo Lab 7.2's deliberate permission and identity failures before starting. No Stage 3 registration is required.
> - **Starting point:** [`labs/stage-5-exe-server/`](../labs/stage-5-exe-server/) on B, registered with `Setup.ps1 -Action Register`; its x64 client and matching `Stage5PS.dll` on A. Register the proxy on A with x64 `regsvr32`. Build from the same IDL on both machines and use the EXE CLSID `{647E1E3B-42CF-4A82-B91E-6DCADB217D9B}`.
> - **Caution:** isolated lab network. Opening TCP 135 plus the dynamic RPC range, and loosening authentication levels, is a lab configuration and **not** a production one. Revert every change afterwards.
> - **Time:** ~3 h.

Local out-of-proc already works. This lab puts a network in the middle and shows exactly what that adds: an endpoint mapper on port 135, a dynamic port range behind it, a **second** marshaling registration on the client machine, and four permission bits instead of two.

The method is one variable at a time — break a single thing, record the HRESULT and which machine logged it, restore it. "It fails remotely" is not a diagnosis; this lab is how you turn it into one.

1. Register the Stage 5 server on machine B, and **its matching Stage5PS.dll on both A and B**. (Forgetting the client-side marshaling registration is a classic remote-DCOM failure.)
2. From machine A, activate with `CoCreateInstanceEx` + `COSERVERINFO`.
3. Grant **Remote Launch** and **Remote Activation** to the calling principal on B's AppID, and **Remote Access** for calls.
4. **Break it and diagnose, one variable at a time:**

| Break | Expected | How you confirm |
|---|---|---|
| Block TCP 135 on B's firewall | `0x800706BA RPC_S_SERVER_UNAVAILABLE` | `Test-NetConnection B -Port 135` fails |
| Allow 135 but block the dynamic range | Activation succeeds at the endpoint mapper, then fails/hangs | ProcMon/netmon shows a connect attempt to a high port |
| Remove Remote Activation | `0x80070005` + Event 10016 on B | Event log on **B**, not A |
| Force `RPC_C_AUTHN_LEVEL_CONNECT` on the client | `0x80070005` + **Event 10036** on B | The hardening scenario |
| Stop `RpcSs` on B | `0x800706BA` | `Get-Service RpcSs` |

5. Use `Get-CimInstance -ComputerName B -Protocol DCOM` as the control test at each step. If it works and yours doesn't, the transport is fine.

---

## 7.12 Security checklist for reviewing a COM server

Use this when a customer asks "is our component configured safely?"

- [ ] `CoInitializeSecurity` called explicitly, with `RPC_C_AUTHN_LEVEL_PKT_INTEGRITY` or higher.
- [ ] Impersonation level is the minimum needed (`IDENTIFY` unless the server genuinely acts as the client).
- [ ] `EOAC_NO_CUSTOM_MARSHAL` set — blocks the `IMarshal` escalation vector.
- [ ] `EOAC_DISABLE_AAA` set if the server should never run as the activator.
- [ ] `EOAC_SECURE_REFS` set so unauthenticated parties can't `Release` your objects.
- [ ] Launch/Access permissions grant **specific principals**, not `Everyone`/`Authenticated Users`.
- [ ] `RunAs` uses the **least-privileged** account that works; not `Interactive User` for a service-consumed component.
- [ ] Every `CoImpersonateClient` is paired with `CoRevertToSelf` via RAII.
- [ ] All `[in]` parameters validated — a COM server is a **trust boundary**. Sizes, indexes, string lengths, null checks.
- [ ] `size_is`/`length_is` values validated against the actual buffer before use (a lying client is a real threat model).
- [ ] The server binary is Authenticode signed, and its directory is not user-writable (else: DLL/EXE planting).
- [ ] Registration writes to `HKLM`, and the keys aren't writable by non-admins (else: CLSID hijacking).
- [ ] Elevated components: `LocalizedString` is an indirect string; the EXE is signed; the elevated surface is minimal and validated.

That last group is the one that matters most in security reviews: **an out-of-proc COM server is a privilege boundary, and its interface is an attack surface.** Treat `[in]` parameters exactly as you would treat network input.

---

## 7.13 The DCOM support triage flow

```
Activation or call failure in an out-of-proc / remote scenario
│
├─ Get the exact HRESULT.  !error / certutil -error
│
├─ 0x80040154 REGDB_E_CLASSNOTREG
│     → Module 2 flow: bitness, hive, ProcMon
│
├─ 0x80070005 E_ACCESSDENIED
│     ├─ Event 10016 on the SERVER machine?
│     │    → note the right (Launch vs Activation, Local vs Remote) and the principal
│     │    → check MachineLaunchRestriction (Limits) FIRST, then the AppID, then Default
│     ├─ Event 10036/10037?  → authentication-level hardening; raise the client's level
│     ├─ Integrity level mismatch? (medium client → high server)
│     └─ NTFS ACL on the server binary, or on the CLSID/AppID registry key
│
├─ 0x80080005 CO_E_SERVER_EXEC_FAILURE
│     ├─ Can you launch the EXE manually as the RunAs account?
│     ├─ RunAs account valid?  Password changed?  (LSA secret SCM:{appid})
│     ├─ Server crashing at startup?  → Application event log, WER, procdump -e
│     ├─ Server failing to CoRegisterClassObject within the timeout?
│     └─ Session 0 / no interactive user for "Interactive User"
│
├─ 0x800706BA RPC_S_SERVER_UNAVAILABLE
│     ├─ Test-NetConnection <host> -Port 135
│     ├─ Dynamic port range open?
│     ├─ RpcSs / DcomLaunch running?
│     └─ Name resolution correct?  (short name vs FQDN vs IP changes auth!)
│
├─ 0x80010108 RPC_E_DISCONNECTED
│     → the server died. Get a dump of the SERVER, not the client.
│
└─ Hang, not an error
      → dumps of BOTH processes; look for combase!...SendReceive (Module 3)
```

Print this. It is the module's deliverable.

---

## 7.14 Checkpoint

1. Distinguish Launch permission from Access permission: when is each evaluated, and how does the failure differ?
2. In what order are `MachineLaunchRestriction`, the AppID's `LaunchPermission`, and `DefaultLaunchPermission` evaluated? Why does that order defeat a common "fix"?
3. A component works when a user runs the test app, but fails with `0x80080005` under a service. Give four hypotheses.
4. What does `RunAs = Interactive User` mean, and name two situations where it cannot work.
5. Why is `EOAC_NO_CUSTOM_MARSHAL` a security control? What attack does it stop?
6. After patching, a customer's remote DCOM app fails with `E_ACCESSDENIED` and the server logs Event 10036. What happened, where is the fix, and why can't you just disable the hardening?
7. Why is opening TCP 135 insufficient for remote DCOM?
8. A support engineer "fixes" Event 10016 for a Windows component by taking ownership of the AppID key from TrustedInstaller and granting Everyone Full Control. Give three reasons this is wrong.
9. An out-of-proc COM server accepts `[in, size_is(cb)] BYTE* buf, [in] ULONG cb`. What must the server do before using `buf`, and why is this a security boundary?

<details>
<summary>Answers</summary>

1. **Launch** is checked by the SCM *before* starting the server process (and for "Activation," before creating an object in an existing one). **Access** is checked by the RPC layer on each *call* into a running server. Failure difference: no Launch → the server process never starts, activation fails immediately; no Access → the server starts fine (you can see it in Task Manager) and then the first method call fails with `E_ACCESSDENIED`. Seeing whether the server process appears is the fastest discriminator.

2. Machine restriction (Limits) → AppID-specific → Default. The machine restriction is a **ceiling**: if it denies, nothing downstream can grant. The common broken "fix" is editing the AppID's permission (or the Default) while the machine Limits still deny the principal — the change has no effect and the engineer concludes the permission model is broken.

3. (a) `RunAs` account invalid or its stored password is stale (LSA secret `SCM:{appid}`). (b) Session 0 isolation — `RunAs = Interactive User` with no interactive user, or a cross-session activation. (c) The server EXE crashes at startup under the service account (missing profile, missing HKCU registration, missing environment). (d) The server didn't `CoRegisterClassObject` within the SCM's timeout. (Also: the service account lacks NTFS execute on the binary.)

4. The server runs in the console user's session under that user's identity, as a **single shared instance**. It cannot work when (a) no user is logged on, and (b) the activator is in session 0 (a service) — cross-session activation of an interactive server is unreliable/blocked. Also fails on multi-session hosts (RDS) where "the interactive user" is ambiguous.

5. Custom marshaling lets the *server* specify, via `IMarshal::GetUnmarshalClass`, which CLSID the *client* must instantiate in its own process to build the proxy. A malicious or compromised low-privilege server can therefore cause a higher-privilege client to load and run code of the server's choosing. `EOAC_NO_CUSTOM_MARSHAL` refuses `IMarshal` and forces standard marshaling, closing that vector.

6. The DCOM hardening (CVE-2021-26414) raised the minimum activation authentication level to `RPC_C_AUTHN_LEVEL_PKT_INTEGRITY`; the client is still requesting something lower (typically `CONNECT`). Event 10036 is logged **on the server**, which is why it's easy to miss when you only collect client logs. The fix belongs on the **client**: raise the level via `CoInitializeSecurity`, or per-AppID/machine defaults. You can't disable the hardening because the `RequireIntegrityActivationAuthenticationLevel = 0` escape hatch was removed in the final rollout phase. A third-party app that hardcodes a low level needs a vendor update.

7. Port 135 is only the **endpoint mapper**. The client asks it where the service actually lives, and is redirected to a **dynamically assigned high port** (49152–65535 by default) for the real traffic. Without that range open — or without pinning it via `netsh int ipv4 set dynamicport` or the `HKLM\SOFTWARE\Microsoft\Rpc\Internet` settings — the initial contact succeeds and everything after it fails.

8. (a) The event is very likely **benign** — Windows components routinely attempt an activation, get denied, and use a supported fallback; nothing was actually broken. (b) Taking ownership from TrustedInstaller leaves the key with non-default ownership, which can break servicing/updates and will be flagged by security baselines. (c) Granting `Everyone` `Full Control` on a system AppID is a real privilege-escalation risk and violates least privilege. Correct approach: correlate with an actual functional symptom first; if genuinely needed and third-party, grant the *specific* right to the *specific* principal.

9. The server must validate `cb` against its own limits and never trust it to describe `buf` correctly — check for zero, for absurdly large values, for integer overflow in `cb * sizeof(element)`, and bound every index derived from it. It's a security boundary because the client may be a different, less-trusted, possibly malicious process (and with a remote server, a different *machine*), and the server may be running with higher privilege. `[in]` parameters are untrusted input, exactly like data off a socket.

</details>

---

## 7.15 Rules to carry forward

1. CLSID is per-class; AppID is per-process. Security lives on the AppID.
2. Launch ≠ Access. Check which one the event names before touching anything.
3. Machine Limits are a ceiling evaluated before AppID and Default permissions.
4. Most Event 10016s are benign. Require a correlated functional symptom before changing permissions, and never grant Everyone/Full Control.
5. Session 0 isolation: services get no desktop, and `Interactive User` is a trap.
6. Least privilege for `RunAs`; remember that a changed password silently breaks activation.
7. Set `PKT_INTEGRITY`, `EOAC_NO_CUSTOM_MARSHAL`, `EOAC_SECURE_REFS`; impersonate only when necessary, and always revert with RAII.
8. Remote DCOM needs 135 **and** the dynamic range; pin the range explicitly.
9. Event 10036/10037 = authentication-level hardening; fix the client, and read the *server's* event log.
10. A COM server's interface is a trust boundary. Validate every `[in]` parameter, especially sizes.

---

**Next: [Module 8 — Debugging, diagnostics, and the capstone](08-debugging-and-capstone.md)**
