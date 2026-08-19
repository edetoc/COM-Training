# Module 8 — Debugging, diagnostics, and the capstone

**Time: ongoing. This module turns knowledge into diagnostic reflexes.**

Modules 1–7 taught you how COM works and how it breaks. This one is about **evidence**: what to collect, which tool answers which question, and how to reason from a dump you didn't produce, for code you don't have, on a machine you can't touch.

---

## 8.1 The evidence hierarchy

When a ticket arrives, collect in this order. Each level is cheaper than the next and often sufficient.

| Level | Artifact | Answers |
|---|---|---|
| 0 | **The exact HRESULT** and the API that returned it | 60% of tickets, instantly |
| 1 | **Event logs** (System → DistributedCOM, Application, Security) | Permission and launch failures |
| 2 | **Registry state** (CLSID/AppID/Interface/TypeLib, all hives, both bitnesses) | Registration failures |
| 3 | **Process Monitor trace** | *Why* a lookup failed — missing vs denied |
| 4 | **Process/module state** (Process Explorer, Task Manager sessions) | Wrong process, wrong session, wrong bitness |
| 5 | **User-mode dump** (client, server, or both) | Hangs, crashes, leaks |
| 6 | **ETW trace** (COM/RPC providers) | Intermittent, timing-dependent, or cross-process flows |
| 7 | **Application Verifier / page heap** run | Latent corruption and lifetime bugs |

**The most common support mistake is skipping to level 5.** A dump takes 30 minutes to analyze; the HRESULT plus a ProcMon filter takes 3 minutes and solves the same ticket most of the time.

---

## 8.2 Level 0 — HRESULT decoding

```powershell
certutil -error 0x80070005
[ComponentModel.Win32Exception]::new(-2147024891).Message
```

```
0:000> !error 0x80070005
```

Reconstruct from structure when no tool is handy (Module 1):

- `0x8007xxxx` → FACILITY_WIN32; low word is a Win32 error. `net helpmsg <decimal>`.
- `0x8004xxxx` → FACILITY_ITF; **interface-specific — the meaning depends on which interface returned it.** `0x80040154` from the SCM is `REGDB_E_CLASSNOTREG`, but `0x80040xxx` from a third-party interface may mean something entirely different. Always ask *which call* returned it.
- `0x8001xxxx` → RPC-related (Module 3/7).
- `0x8002xxxx` → Automation/dispatch (Module 5).
- `0x800Fxxxx`, `0x8013xxxx` → setup, .NET.

> **Ask "which API returned this?" before "what does this code mean?"** Facility ITF codes are ambiguous without that context.

---

## 8.3 Level 3 — Process Monitor for COM

The single highest-value tool for activation failures.

### Filter recipe

```
Path contains {CLSID-with-braces}          -> Include
Operation is RegOpenKey                    -> Include
Operation is RegQueryValue                 -> Include
Result is NAME NOT FOUND                   -> Highlight
Result is ACCESS DENIED                    -> Highlight
```

Also add `Process Name is <client.exe>` to cut noise.

### What a healthy activation looks like

```
RegOpenKey   HKCR\CLSID\{...}                          SUCCESS
RegQueryValue HKCR\CLSID\{...}\InprocServer32\(Default) SUCCESS  C:\...\Calc.dll
RegQueryValue HKCR\CLSID\{...}\InprocServer32\ThreadingModel SUCCESS "Both"
CreateFile   C:\...\Calc.dll                            SUCCESS
Load Image   C:\...\Calc.dll                            SUCCESS
```

Capture this once from a *working* system (Lab 2.1 asked you to). Every failure is a deviation from this baseline.

### Reading failures

| ProcMon result | Meaning | Not |
|---|---|---|
| `NAME NOT FOUND` on the CLSID key | Not registered in the hive/bitness this process sees | a permission problem |
| `ACCESS DENIED` on the CLSID key | Registered, but this token can't read it | a deployment problem |
| `NAME NOT FOUND` on the DLL path | Registration points at a missing file | |
| `ACCESS DENIED` on the DLL file | NTFS ACL | |
| `NAME NOT FOUND` on a *different* DLL right after Load Image | **Missing dependency** → `0x8007007E` | |
| Probe under `Wow6432Node` | The client is 32-bit | |

That last one is a free bitness check: **if ProcMon shows `Wow6432Node` in the path, the client is 32-bit.** No other tool needed.

### Dependency failures

When `Load Image` on your DLL is followed by `NAME NOT FOUND` on `MSVCP140.dll` or similar, you have `ERROR_MOD_NOT_FOUND` (`0x8007007E`). Confirm with:

```powershell
dumpbin /dependents C:\Components\Calc.dll
```

or the Dependencies tool (the modern Dependency Walker replacement).

---

## 8.4 Level 5 — WinDbg for COM

### Setup

```
.symfix
.sympath+ C:\MySymbols
.reload /f
```

Get a dump:

```powershell
procdump -ma -o <pid> C:\dumps\out.dmp          # full dump, now
procdump -ma -e -x C:\dumps app.exe             # on unhandled exception
procdump -ma -h app.exe C:\dumps\hang.dmp       # on window hang
procdump -ma -s 5 -n 3 <pid> C:\dumps\seq.dmp   # 3 dumps, 5s apart - great for leaks/hangs
```

For a **hang**, take **three dumps 10 seconds apart**. Stacks that are identical across all three are genuinely stuck; stacks that move are just slow. This distinction saves enormous time.

### Command reference for COM work

| Command | Use |
|---|---|
| `!error <hr>` | Decode an HRESULT |
| `~*kb` | All thread stacks — the first command for any hang |
| `!uniqstack` | Deduplicated stacks; much faster to scan in a 200-thread process |
| `!runaway` | CPU time per thread — distinguishes a hang (all idle) from a spin |
| `dps <ptr> L8` | **Dump a vtable** — identifies the real implementation behind an interface pointer |
| `!cs -l` | Locked critical sections and their owners |
| `!locks` | Same, with wait chains |
| `!handle 0 f Event` | Event handles and signaled state |
| `!heap -s` | Heap summary |
| `!heap -stat -h <heap>` | Allocation stats by size — leak hunting |
| `!heap -p -a <addr>` | Allocation stack for a block (needs page heap) |
| `!address <addr>` | Is this memory committed, free, or in a module? |
| `lm` / `lmvm <mod>` | Loaded modules; verify versions and symbol status |
| `!analyze -v` | Automated first pass on a crash |
| `!gle` | Last error on the current thread |

### The most useful trick: `dps` on an interface pointer

```
0:000> dps 000001f2`3a4b5c60 L4
000001f2`3a4b5c60  00007ffb`12345678 Calc!CCalculator::`vftable'
```

or, for a proxy:

```
0:000> dps 000001f2`3a4b5c60 L4
000001f2`3a4b5c60  00007ffb`aabbccdd combase!CStdProxyBuffer_QueryInterface
```

**This one command answers "am I holding the object or a proxy?"** — which resolves apartment questions (Module 3) and "which vendor's component did I actually get?" questions instantly.

To go from the pointer to the vtable to the module:

```
0:000> dq <interface-ptr> L1        ; get the vptr
0:000> dps <vptr> L8                ; dump the slots with symbols
0:000> lm a <slot0-address>          ; which module owns that code
```

### Recognizing the standard COM stacks

**Client blocked in an outbound cross-apartment call:**
```
ntdll!NtWaitForSingleObject
combase!CSyncClientCall::SwitchAptAndDispatchCall
combase!CSyncClientCall::SendReceive2
combase!CAptRpcChn::SendReceive
combase!CCtxComChnl::SendReceive
RPCRT4!NdrpClientCall3
<YourApp>!ICalculator_Add_Proxy
```

**Server thread executing an inbound call:**
```
<YourServer>!CCalculator::Add
RPCRT4!NdrStubCall3
combase!CStdStubBuffer_Invoke
combase!SyncStubInvoke
combase!ThreadInvoke
RPCRT4!LrpcIoComplete
```

**STA thread correctly pumping while waiting:**
```
ntdll!NtWaitForMultipleObjects
combase!CCliModalLoop::BlockFn
combase!ModalLoop
combase!ThreadSendReceive
```

**STA thread incorrectly blocked (the deadlock signature):**
```
ntdll!NtWaitForSingleObject
KERNELBASE!WaitForSingleObjectEx
<YourApp>!SomeFunction              <- no ModalLoop, no CoWaitForMultipleHandles
```

> **Memorize the discrimination:** `CCliModalLoop::BlockFn` in an STA's stack = pumping correctly. `WaitForSingleObjectEx` directly under app code in an STA = the deadlock from Module 3.

### Cross-process hang analysis

For an out-of-proc hang you need **both** dumps.

1. Client: find the thread in `SendReceive`.
2. Get the target: the RPC call carries a **causality ID**. In practice, the faster route is to look at what the *server* is doing — dump it and look for a thread in `CStdStubBuffer_Invoke`, then see what *it* is blocked on.
3. Common shape: server thread is blocked calling *back* into the client, and the client's STA isn't pumping. That's a distributed version of Module 3's Pattern 2.

---

## 8.5 Level 6 — ETW tracing

For intermittent or timing-dependent problems where you can't catch a dump at the right moment.

### Relevant providers

| Provider | GUID | Use |
|---|---|---|
| `Microsoft-Windows-COM` | `{d4263c98-310c-4d97-ba39-b55354f08584}` | COM activation |
| `Microsoft-Windows-COM-Perf` | `{b8d6861b-d20f-4eec-bbae-87e0dd80602b}` | Call-level perf |
| `Microsoft-Windows-COMRuntime` | `{bf406804-6afa-46e7-8a48-6c357e1d6d61}` | Runtime internals |
| `Microsoft-Windows-RPC` | `{6ad52b32-d609-4be9-ae07-ce8dae937e39}` | RPC calls |
| `Microsoft-Windows-RPCSS` | `{d8975f88-7ddb-4ed0-91bf-3adf48c48e0c}` | The SCM itself |
| `Microsoft-Windows-DistributedCOM` | | The source behind Event 10016 etc. |

### Capture

```powershell
# Simple, using logman
logman create trace ComTrace -o C:\traces\com.etl -ets
logman update trace ComTrace -p "{d4263c98-310c-4d97-ba39-b55354f08584}" 0xffffffff 5 -ets
logman update trace ComTrace -p "{6ad52b32-d609-4be9-ae07-ce8dae937e39}" 0xffffffff 5 -ets
# ... reproduce the problem ...
logman stop ComTrace -ets
```

Then open `com.etl` in **Windows Performance Analyzer** (WPA) or convert:

```powershell
tracerpt C:\traces\com.etl -o C:\traces\com.xml -of XML
```

### What ETW tells you that a dump can't

- **Activation timing** — how long the SCM took, whether it launched a process.
- **Which CLSID** was requested when the error occurred, in a process that requests hundreds.
- **Call frequency** — is the "slow component" being called 10 times or 100,000 times? (Module 3's marshaling overhead × call count is often the real answer.)
- **Cross-process correlation** via the RPC causality ID.

**A very common outcome:** the component isn't slow; it's being called in a loop across an apartment boundary. ETW shows the call count; no dump would.

---

## 8.6 Level 7 — Application Verifier and page heap

### Application Verifier

`appverif.exe` → add your EXE → enable:

- **Basics → COM**: catches `CoInitialize` imbalance, apartment violations, use of an interface after `CoUninitialize`, and some over-release patterns.
- **Basics → Heaps** (page heap): each allocation gets a guard page, so overruns and use-after-free fault **at the moment of the bad access**, not later.
- **Basics → Handles**: invalid handle use.
- **Basics → Locks**: critical section misuse (a good companion to Module 3's deadlock work).

Then run under a debugger. Verifier breaks in with a diagnosis instead of a mystery crash.

```powershell
# Enable page heap for a single process without the GUI
gflags /p /enable MyApp.exe /full
gflags /p /disable MyApp.exe        # ALWAYS turn it off afterwards - it's slow and memory-hungry
```

### The `BSTR` cache trick

`oleaut32` caches freed `BSTR`s, which **masks leaks** in heap statistics. Disable it:

```powershell
$env:OANOCACHE = 1
.\MyApp.exe
```

Now `BSTR` leaks show up immediately in `!heap -stat`. Remember this — it turns an invisible leak into a visible one.

---

## 8.7 Leak hunting, end to end

**Symptom:** private bytes grow monotonically; the process never releases memory even at idle.

### Step 1 — Confirm it's a COM leak

```
0:000> !heap -s
0:000> !heap -stat -h 0            ; all heaps, allocation stats by size
```

Look for one size bucket dominating and growing across sequential dumps. A COM object leak usually shows a fixed allocation size repeating.

Also check:
```
0:000> lm                          ; is a server DLL still loaded that shouldn't be?
0:000> !handle 0 0                 ; handle count growing?
```

A DLL that never unloads is direct evidence that `DllCanUnloadNow` never returned `S_OK` — i.e. live objects or locks (Module 2).

### Step 2 — Find the allocation site

With page heap enabled:

```
0:000> !heap -p -a <address-of-a-leaked-block>
```

This prints the **allocation stack**. Do this for several leaked blocks; the common frame is your leak.

### Step 3 — Find the missing `Release`

Break on the object's `Release` and log stacks:

```
0:000> bp Calc!CCalculator::AddRef  "kb 6; gc"
0:000> bp Calc!CCalculator::Release "kb 6; gc"
0:000> g
```

Dump the log, count stacks. An `AddRef` stack with no matching `Release` stack is the culprit.

Conditional variant — break only when the count crosses a threshold:

```
0:000> bp Calc!CCalculator::Release ".if (poi(@rcx+8) > 0n10) { kb } .else { gc }"
```

(Offset 8 assumes `m_cRef` follows the vptr; confirm with `dt` if you have symbols.)

### Step 4 — Check the usual suspects, in order

1. **`Advise` without `Unadvise`** (Module 5) — by far the most common.
2. **GIT registration without `RevokeInterfaceFromGlobal`** (Module 3).
3. **Reference cycle** — parent/child both strong (Module 1).
4. **`[out]` interface pointer never released** — enumerators are notorious.
5. **`QueryInterface` result discarded** without `Release`.
6. **.NET RCW never collected** — check with `!dumpheap -type System.__ComObject` in SOS.

### Managed leaks

```
0:000> .loadby sos clr             ; or: .loadby sos coreclr
0:000> !dumpheap -stat
0:000> !dumpheap -type System.__ComObject
0:000> !gcroot <address>           ; what's keeping this RCW alive?
```

---

## 8.8 The complete triage flowchart

```
COM PROBLEM
│
├── Does it FAIL with an error?
│   │
│   ├── At ACTIVATION (CoCreateInstance)
│   │   ├── 0x80040154  → Module 2: bitness → hive → ProcMon (NAME NOT FOUND vs ACCESS DENIED)
│   │   ├── 0x8007007E  → missing dependency; ProcMon Load Image + dumpbin /dependents
│   │   ├── 0x80070005  → Module 7: Event 10016 → Limits → AppID → Default → integrity level
│   │   ├── 0x80080005  → Module 7: RunAs account, server crash at startup, session 0
│   │   ├── 0x800706BA  → Module 7: network, port 135 + dynamic range, RpcSs
│   │   └── 0x800401F0  → CoInitializeEx missing on this thread
│   │
│   ├── At QUERYINTERFACE
│   │   ├── 0x80004002 in-proc too   → the object genuinely lacks it
│   │   ├── 0x80004002 only out-of-proc → Module 4: marshaling not registered
│   │   └── 0x80040155  → Module 4: HKCR\Interface\{IID}\ProxyStubClsid32
│   │
│   ├── At CALL TIME
│   │   ├── 0x8001010E  → Module 3: wrong apartment, raw pointer shared
│   │   ├── 0x80010108  → server died; dump the SERVER
│   │   ├── 0x800706F7  → Module 4: mismatched IDL builds; get all three version numbers
│   │   ├── 0x80020009  → Module 5: open EXCEPINFO, the real error is inside
│   │   └── 0x80070005  → Module 7: Access (not Launch) permission
│   │
│   └── Intermittent / under load only
│       └── Module 3: threading. ThreadingModel + client apartment + shared raw pointers
│
├── Does it HANG?
│   ├── 3 dumps, 10s apart. Same stacks = truly stuck.
│   ├── ~*kb → look for combase!...SendReceive and a non-pumping STA
│   ├── !cs -l / !locks → lock held across an outbound call?
│   ├── Out-of-proc? Dump BOTH processes.
│   └── Module 3 §3.8 patterns
│
├── Does it CRASH?
│   ├── !analyze -v
│   ├── Faulting address in no module?  → DLL unloaded with live objects (Module 2)
│   ├── Crash inside Release / garbage vptr? → over-release (Module 1)
│   ├── Heap corruption? → page heap, then rerun
│   └── Wrong function called for the name? → interface changed without changing the IID (Module 4)
│
├── Does it LEAK?
│   └── §8.7: !heap -stat → !heap -p -a → bp on AddRef/Release → check Advise/GIT/cycles
│
└── Is it just SLOW?
    ├── ETW: how many calls actually happen?
    ├── ThreadingModel + client apartment → is every call marshaled? (Module 3 Lab 3.3)
    ├── Late-bound IDispatch instead of vtable? (Module 5 Lab 5.1)
    └── Out-of-proc where in-proc would do? CLSCTX_ALL picking the wrong server?
```

---

## 8.9 LAB 8.1 — Blind dump analysis

A drill. Have a colleague (or your future self, a week later) produce dumps from a modified version of your labs without telling you what was changed. For each dump, in under 15 minutes, produce:

1. The failure mode (hang / crash / leak).
2. The specific COM concept involved.
3. The evidence, quoted as WinDbg output.
4. The fix.

Suggested scenarios to prepare:

| # | Injected bug | Expected diagnosis path |
|---|---|---|
| 1 | STA blocks on `WaitForSingleObject` while a worker calls in | `~*kb` → `SendReceive` + non-pumping STA |
| 2 | Extra `Release` in a callback | crash in `Release`, garbage vptr, page heap shows freed block |
| 3 | `DllCanUnloadNow` returns `S_OK` unconditionally | faulting address in no module; `lm` doesn't cover it |
| 4 | `Advise` without `Unadvise` in a loop | growing heap; `!heap -stat` one dominant size; `AddRef` stacks in the sink |
| 5 | Interface changed, IID unchanged, one binary stale | `RPC_X_BAD_STUB_DATA`, or a call landing in the wrong function |
| 6 | GIT cookie never revoked | leak; `!heap -p -a` points at the GIT registration path |
| 7 | Lock held across an outbound call, server calls back | hang; `!cs -l` shows the STA owning the CS it's waiting on |

Log every run in your notes with: *symptom seen first → tool that localized it → time taken*. Watch the time drop.

---

## 8.10 LAB 8.2 — Build the support runbook (deliverable)

For each HRESULT below, write a runbook entry with **exactly** these fields:

```markdown
### 0x80040154 — REGDB_E_CLASSNOTREG

**What the user sees:** "Class not registered" / the app fails to start a feature.

**What it means:** The SCM could not resolve the CLSID to a server in the registry
view visible to that process.

**Data to collect (in order):**
1. Exact HRESULT and the API that returned it.
2. Bitness of the client process (Task Manager -> "*32" or Process Explorer).
3. `Get-ComRegistration <clsid>` across HKLM, HKLM\Wow6432Node, HKCU.
4. ProcMon trace filtered on the CLSID, noting NAME NOT FOUND vs ACCESS DENIED.

**Discriminators:**
- ProcMon shows Wow6432Node in the path -> the client is 32-bit.
- NAME NOT FOUND -> not registered in that view.
- ACCESS DENIED  -> registered but unreadable by this token.
- Registration present + file missing -> the path is stale.

**Fixes:**
- Register the matching bitness, or provide both.
- Register in HKLM if multiple users need it.
- Reg-free COM via manifests if admin rights are unavailable.
- DllSurrogate to bridge a bitness mismatch (requires marshaling registration).

**Do NOT:** blindly re-run the installer; it usually re-registers the same bitness.
```

Cover at minimum: `0x80040154`, `0x80040155`, `0x8007007E`, `0x80070005`, `0x80080005`, `0x800706BA`, `0x80010108`, `0x8001010E`, `0x800401F0`, `0x80004002`, `0x800706F7`, `0x80020009`.

That document is the tangible output of this course.

---

## 8.11 CAPSTONE — `IDocumentStore`

Build one complete component, then support it.

### Part 1 — The component

An out-of-proc (EXE) COM server exposing a small document store.

```idl
import "oaidl.idl";
import "ocidl.idl";

[object, uuid(...), dual, nonextensible, pointer_default(unique)]
interface IDocument : IDispatch
{
    [id(DISPID_VALUE), propget] HRESULT Name([out, retval] BSTR* name);
    [id(1), propget] HRESULT Content([out, retval] BSTR* content);
    [id(1), propput] HRESULT Content([in] BSTR content);
    [id(2), propget] HRESULT Modified([out, retval] DATE* when);
    [id(3)] HRESULT Save();
}

[object, uuid(...), dual, nonextensible, pointer_default(unique)]
interface IDocumentStore : IDispatch
{
    [id(1)] HRESULT Add([in] BSTR name, [out, retval] IDocument** doc);
    [id(2)] HRESULT Remove([in] VARIANT index);
    [id(DISPID_VALUE), propget] HRESULT Item([in] VARIANT index, [out, retval] IDocument** doc);
    [id(3), propget] HRESULT Count([out, retval] LONG* count);
    [id(DISPID_NEWENUM), propget, restricted] HRESULT _NewEnum([out, retval] IUnknown** e);
    [id(4)] HRESULT FindByContent([in] BSTR pattern,
                                  [out] ULONG* count,
                                  [out, size_is(, *count)] BSTR** names);
}

[uuid(...)]
dispinterface _IDocumentStoreEvents
{
    properties:
    methods:
        [id(1)] void OnDocumentAdded([in] BSTR name);
        [id(2)] void OnDocumentRemoved([in] BSTR name);
        [id(3)] void OnError([in] BSTR message);
};
```

Requirements — each exercises a specific module:

| Requirement | Module |
|---|---|
| Correct `IUnknown`, no leaks, no cycles between store and documents | 1 |
| Out-of-proc EXE server with `CoAddRefServerProcess` lifetime | 2, 7 |
| `ThreadingModel`/apartment chosen deliberately and documented | 3 |
| Dual interfaces + registered type library (typelib marshaling) | 4, 5 |
| Correct memory ownership for `BSTR`, `BSTR**`, `VARIANT` | 4 |
| `IEnumVARIANT` + `_NewEnum` so `For Each` works | 5 |
| Connection points with RAII `Unadvise` on the client side | 5 |
| `ISupportErrorInfo` + rich `IErrorInfo` errors | 5 |
| Built with ATL | 6 |
| `CoInitializeSecurity` with `PKT_INTEGRITY`, `EOAC_NO_CUSTOM_MARSHAL` | 7 |
| Every `[in]` parameter validated (sizes, nulls, indexes) | 7 |
| `IDocumentStore2` adding one method, without touching `IDocumentStore` | 4 |

### Part 2 — The clients

| Client | Must demonstrate |
|---|---|
| C++ raw | Manual `QI`/`Release`, `goto Cleanup` discipline |
| C++ ATL/WIL | `CComPtr`, `RETURN_IF_FAILED`, RAII connection cookie |
| C# classic interop | Embed Interop Types, `foreach` over the store, event subscription |
| C# `ComWrappers`/`[GeneratedComInterface]` | Modern .NET interop |
| PowerShell | `New-Object -ComObject`, `Register-ObjectEvent`, pipeline iteration |
| VBScript | Late binding, `Err.Description` from your `IErrorInfo` |

### Part 3 — Three deployment modes

1. `regsvr32` / `-RegServer`, per-machine.
2. Registration-free via manifests (client + server).
3. In-proc DLL variant behind a `DllSurrogate`, bridging 32/64-bit.

Document what changes between them and what breaks if you get it wrong.

### Part 4 — The support runbook (the real deliverable)

For **this specific component**, document how each failure would manifest and the exact diagnostic steps:

| Scenario | Your runbook must cover |
|---|---|
| Type library not registered | Which client breaks first, which error, which registry key |
| 32-bit client, 64-bit server | Exact error, ProcMon signature, both fixes |
| Client forgets `Unadvise` | How you'd prove it from a dump |
| Server killed mid-call | Client-side error, recovery strategy |
| Launch permission removed | Event 10016 fields, correct fix |
| `RunAs` password stale | Error, event log location, fix |
| Remote activation, firewall blocks the dynamic range | Symptom, `Test-NetConnection`, the WMI control test |
| Client requests `AUTHN_LEVEL_CONNECT` | Event 10036, where the fix belongs |
| `IDocumentStore2` added, old client on new server (and vice versa) | Why both must work, and how you'd verify |
| Store and documents form a reference cycle | Heap evidence, the fix |

**When you can hand that runbook to another engineer and they can resolve tickets with it, you're done.**

---

## 8.12 Reference shelf

### Books

| Book | Why |
|---|---|
| **Essential COM** — Don Box | The *why*. Chapters 1–3 are the best explanation of the model ever written. |
| **Inside COM** — Dale Rogerson | The gentlest ramp; builds `IUnknown` from nothing, step by step. |
| **ATL Internals** (2nd ed.) — Tavares, Rector, Sells et al. | What the ATL macros actually generate. |
| **Advanced Windows Debugging** — Hewardt & Pravat | The support track: heaps, handles, locks, dumps. |
| **Windows Internals, Parts 1 & 2** — Russinovich et al. | Sessions, integrity levels, ALPC, RPC, the SCM. |

### Online

- Microsoft Learn → **Component Object Model (COM)** and **COM Fundamentals**.
- Microsoft Learn → **Introduction to COM interop**, **COM Wrappers**, **Source-generated COM interop**.
- **The Old New Thing** (Raymond Chen) — search "apartment", "STA", "message pump", "COM". Decades of hard-won detail.
- **OleView.NET** wiki (James Forshaw) — modern COM security research and the best tooling.
- Microsoft Learn → **DCOM authentication hardening** (CVE-2021-26414) for the current state of Event 10036/10037.

### Tools, one line each

| Tool | Use |
|---|---|
| OleView.NET | Browse/diff COM registration; activate objects interactively |
| Process Monitor | Why activation failed |
| Process Explorer | Which process, which session, which DLLs |
| WinDbg | Hangs, crashes, leaks, vtable identification |
| Application Verifier | Latent lifetime and heap bugs |
| `procdump` | Capture dumps on crash, hang, or a timer |
| `certutil -error` | Decode HRESULTs without a debugger |
| WPA / `logman` | ETW: call counts, activation timing |
| `dcomcnfg` | AppID identity and permissions |
| `sxstrace` | Reg-free COM manifest failures |

---

## 8.13 Final self-assessment

You've completed the course when you can do all of these **without looking anything up**:

- [ ] Explain why COM exists, from binary-compatibility first principles.
- [ ] List every case where a reference count is incremented.
- [ ] State the `QueryInterface` rules and explain why identity uses `IID_IUnknown`.
- [ ] Decode `0x8007xxxx` in your head.
- [ ] Name the registry keys involved in activation and their order of consultation.
- [ ] Diagnose `REGDB_E_CLASSNOTREG` in under 3 minutes with ProcMon.
- [ ] Fill in the `ThreadingModel` × client-apartment matrix from memory.
- [ ] Recognize an STA deadlock in a dump in under 30 seconds.
- [ ] Explain why an STA pumps messages and what that costs you.
- [ ] Write IDL with correct `[in]`/`[out]`/`size_is` and state who frees what.
- [ ] Explain the "works in-proc, fails out-of-proc" bug and diagnose it in three steps.
- [ ] Explain `IDispatch` binding and why scripts are orders of magnitude slower.
- [ ] Identify the `Advise`/`Unadvise` leak from a heap trace.
- [ ] Explain what ATL's macros generate and step into them.
- [ ] Explain RCW sharing and the `InvalidComObjectException`.
- [ ] Distinguish Launch from Access permission and state the evaluation order.
- [ ] Judge whether an Event 10016 is benign, and justify it.
- [ ] Explain the DCOM hardening, Events 10036/10037, and where the fix belongs.
- [ ] Produce the triage flowchart from memory.

---

## 8.14 The ten rules, consolidated

1. Every `AddRef` gets exactly one `Release`; `[out]` interface pointers always carry a reference you own.
2. `QI(IID_IUnknown)` is the only identity test; the interface set never changes.
3. Check `FAILED(hr)`, never `hr == S_OK` — `S_FALSE` is success.
4. Never share a raw interface pointer across apartments; marshal via GIT, stream, or `agile_ref`.
5. Published interfaces are immutable. Add `IFoo2`. Change the interface, change the IID.
6. `[out]` params: null on entry, null on every failure path.
7. `BSTR` → `SysFreeString`; `SAFEARRAY` → `SafeArrayDestroy`; `VARIANT` → `VariantClear`; interface → `Release`; everything else → `CoTaskMemFree`.
8. An STA must never block without pumping, and must never hold a lock across an outbound call.
9. Bitness and registry-hive mismatches cause more `REGDB_E_CLASSNOTREG` than genuine missing registration.
10. Always `Unadvise` what you `Advise`; revoke what you register in the GIT; break cycles deliberately.

---

**End of course.** Return to the [curriculum overview](../README.md) and fill in the progress tracker.
