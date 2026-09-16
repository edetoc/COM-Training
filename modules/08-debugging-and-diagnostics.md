# Module 8 — Debugging and diagnostics

Modules 1–7 taught you how COM works and how it breaks. This one is about **evidence**: what to collect, which tool answers which question, and how to reason from a dump you didn't produce, for code you don't have, on a machine you can't touch.

**What this module covers**

A symptom-based guide to choosing evidence, followed by the tools aimed squarely at COM: decoding any `HRESULT` by hand, Process Monitor, WinDbg, Event Tracing for Windows (ETW), Application Verifier and page heap. It works one leak through from dump to root cause and provides a triage flowchart covering the failures discussed in the course.

**Contents**

- [8.1 What to collect for each symptom](#81-what-to-collect-for-each-symptom)
- [8.2 HRESULT decoding](#82-hresult-decoding)
- [8.3 Process Monitor for COM](#83-process-monitor-for-com)
- [8.4 WinDbg for COM](#84-windbg-for-com)
- [8.5 ETW tracing](#85-etw-tracing)
- [8.6 Application Verifier and page heap](#86-application-verifier-and-page-heap)
- [8.7 Leak hunting, end to end](#87-leak-hunting-end-to-end)
- [8.8 The complete triage flowchart](#88-the-complete-triage-flowchart)
- [8.9 Reference shelf](#89-reference-shelf)
- [8.10 Final self-assessment](#810-final-self-assessment)
- [8.11 The ten rules, consolidated](#811-the-ten-rules-consolidated)

---

## 8.1 What to collect for each symptom

Start with the observed symptom, then choose evidence that can distinguish its possible causes. There is no fixed collection order: a crash or hang may require a dump before the evidence disappears.

**For every case, record:**

- Reproduction steps, frequency, and expected versus actual behavior.
- The failing API or method and exact error output, including the HRESULT when available.
- Client and server machine names, process IDs, account identities, bitness, and relevant Windows/application/component versions.
- Timestamps with time zone so logs and captures from different machines can be correlated.

| Symptom | Data to collect | Tools |
|---|---|---|
| **Object creation fails** | CLSID, registration, startup events, lookup trace | Process Monitor (ProcMon), Event Viewer, `reg query` |
| **Access denied** | Server-side caller identity, correlated events, effective COM security settings | Event Viewer, `dcomcnfg` (registry policy), ProcMon |
| **Crash** | Crash dump, events, matching binaries and symbols | ProcDump, WinDbg, Event Viewer |
| **Hang** | Client/server dumps taken a few seconds apart; operation and duration | ProcDump, WinDbg |
| **Memory or handles grow** | Usage trend, workload, comparative dumps | Performance Monitor, ProcDump, WinDbg |
| **Intermittent or slow calls** | Timestamped logs, COM/RPC trace | `logman`, Windows Performance Analyzer |

**ProcDump** captures dumps; **WinDbg** analyzes them. For ETW traces, **logman** captures and **Windows Performance Analyzer** analyzes.

For crashes involving suspected memory corruption or invalid COM/handle use, **Application Verifier / page heap with WinDbg** provides targeted checks during a controlled reproduction (section 8.6). These checks change runtime behavior and can add substantial overhead.

Capture evidence while the symptom is present, or arrange capture for the next reproduction. The following sections explain the tools; use them to answer a specific question, not as mandatory steps in a sequence.

**Collect safely:** agree on capture impact and disk space before running diagnostics in production; dump capture can pause a process. Dumps and traces may contain credentials or customer data: use restricted storage, approved secure sharing, and the agreed retention policy. Stop traces and restore any verifier/page-heap or crash-capture settings you changed when finished.

---

## 8.2 HRESULT decoding

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

## 8.3 Process Monitor for COM

The single highest-value tool for activation failures.

### Filter recipe

```
Path contains {CLSID-with-braces}          -> Include
Operation is RegOpenKey                    -> Include
Operation is RegQueryValue                 -> Include
Result is NAME NOT FOUND                   -> Highlight
Result is ACCESS DENIED                    -> Highlight
```

Use these filters for registration lookups. Keep **Drop Filtered Events** off so you can change the view later. Remove the CLSID and registry-operation filters to inspect file loads; include the server host for out-of-proc activation, not just the client.

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

Follow the lookup sequence through to the failed operation. Windows often probes optional keys and several search paths before succeeding; a single unsuccessful probe is not a diagnosis.

| ProcMon result | What it establishes | Check next |
|---|---|---|
| `NAME NOT FOUND` on a CLSID key | That lookup did not find the key | Did another applicable registry lookup succeed? |
| `ACCESS DENIED` on a CLSID key | That operation was denied | Process token and effective key permissions |
| `NAME NOT FOUND` on a DLL path | The file was absent at that path | Was it found at another path? |
| `ACCESS DENIED` on a DLL file | That file access was denied | Token, requested access, and file permissions |
| Repeated unresolved dependency probes | Possible missing dependency | Does the load ultimately fail? |
| Probe under `Wow6432Node` | Access to that registry path | Verify the process's bitness separately |

The registry path alone does not prove process bitness: a process can explicitly access another registry view. Confirm its architecture in Process Explorer.

### Dependency failures

A failed load with `ERROR_MOD_NOT_FOUND` (`0x8007007E`) can mean the target DLL or a dependency is missing. Correlate unresolved probes, such as for `MSVCP140.dll`, with that failure. List static dependencies with:

```powershell
dumpbin /dependents C:\Components\Calc.dll
```

or the Dependencies tool. Neither list alone proves runtime resolution; use the trace to see which paths were actually tried and loaded.

---

## 8.4 WinDbg for COM

### Choose the process

- **In-process DLL:** capture the process hosting the DLL, normally the client.
- **EXE server or surrogate:** confirm the PID, executable path, user, session, and bitness in Process Explorer. For `dllhost.exe`, use **Find > Find Handle or DLL** to locate the component DLL, then confirm its full loaded path and host PID.
- **Cross-process hang:** capture the client and server close together on each sampling round, including both machines for remote calls. A server crash needs the server's crash dump.

### Setup

```
.symfix
.sympath+ C:\MySymbols
.reload /f
```

Keep the exact application binaries and matching **PDB (Program Database)** symbol files; replace `C:\MySymbols` with their folder. Microsoft's symbol server supplies Windows symbols, not your own application's PDBs. Check `lmvm <module>` for image and symbol status before trusting source lines or object layouts.

### Capture a usable dump

A small minidump may omit heap and object memory. When that memory is needed, use a full **user-mode** dump (`-ma`) if collection policy permits. This captures process memory, not the whole machine. Prepare the output folder and substitute the actual PID or EXE name:

```powershell
procdump -ma -o <pid> C:\dumps\out.dmp          # full dump, now
procdump -ma -e -w app.exe C:\dumps            # on unhandled exception
procdump -ma -h app.exe C:\dumps\hang.dmp       # on window hang
procdump -ma -s 5 -n 3 <pid> C:\dumps\seq.dmp   # 3 dumps, 5s apart - great for leaks/hangs
```

For a startup crash, start the `-w` command before normal COM activation; it waits for the EXE rather than launching it manually. Avoid name-based capture when multiple processes share that name. If the crash precedes attachment, arrange per-application [Windows Error Reporting (WER) LocalDumps](https://learn.microsoft.com/en-us/windows/win32/wer/collecting-user-mode-dumps) with an administrator before reproducing; applications with custom crash reporting may need their own capture mechanism.

The [ProcDump](https://learn.microsoft.com/en-us/sysinternals/downloads/procdump) `-h` trigger detects an unresponsive window, not every COM hang. Use PID-based timed captures for waits without a hung window.

For a **hang**, compare three dumps about five seconds apart. Repeated stacks suggest a persistent wait or recurring work, not necessarily a deadlock; changing stacks do not guarantee useful progress. Follow wait dependencies and correlate with CPU usage and application logs. Exclude intentional waits, such as the lab client's Enter prompt.

### Command reference for COM work

| Command | Use |
|---|---|
| `!error <hr>` | Decode an HRESULT |
| `~*kb` | All thread stacks — the first command for any hang |
| `!uniqstack` | Deduplicated stacks; much faster to scan in a 200-thread process |
| `!runaway` | CPU time per thread; compare across dumps to identify CPU-consuming threads |
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

**STA thread waiting inside COM's modal loop:**
```
ntdll!NtWaitForMultipleObjects
combase!CCliModalLoop::BlockFn
combase!ModalLoop
combase!ThreadSendReceive
```

**STA thread in a non-pumping wait (investigate the dependency):**
```
ntdll!NtWaitForSingleObject
KERNELBASE!WaitForSingleObjectEx
<YourApp>!SomeFunction              <- no ModalLoop, no CoWaitForMultipleHandles
```

> **Stacks are clues:** a COM modal loop can dispatch incoming calls but does not rule out deadlock. An STA's non-pumping wait is a deadlock risk when completion depends on calls or messages to that STA. Establish that dependency rather than diagnosing from one frame; exact COM function names vary by Windows build.

### Cross-process hang analysis

For an out-of-proc hang you need **both** dumps.

1. Client: find the thread in `SendReceive`.
2. Get the target: the RPC call carries a **causality ID**. In practice, the faster route is to look at what the *server* is doing — dump it and look for a thread in `CStdStubBuffer_Invoke`, then see what *it* is blocked on.
3. Common shape: server thread is blocked calling *back* into the client, and the client's STA isn't pumping. That's a distributed version of Module 3's Pattern 2.

---

## 8.5 ETW tracing

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

## 8.6 Application Verifier and page heap

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

### Step 1 — Check what keeps growing

```
0:000> !heap -s
0:000> !heap -stat -h 0            ; all heaps, allocation stats by size
```

Look for size buckets that keep growing across the same repeated workload. Repeated allocation sizes are a clue, not proof of a COM leak; identify the allocations and their owners before drawing that conclusion.

Also check:
```
0:000> lm                          ; is a server DLL still loaded that shouldn't be?
0:000> !handle 0 0                 ; handle count growing?
```

A DLL remaining loaded does not prove leaked COM references. Unloading may be deferred or never requested, and other loader references can keep the DLL mapped. Correlate object/server-lock counts with the host's unload behavior (Module 2); `DllCanUnloadNow` returning `S_OK` only means the DLL is eligible for unloading.

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

These branches suggest checks, not diagnoses from a single error or stack. For DCOM permissions, transport, and disconnection details, use [Module 7's support triage flow](07-dcom-and-security.md#713-the-dcom-support-triage-flow).

```
COM PROBLEM
│
├── Does it FAIL with an error?
│   │
│   ├── At ACTIVATION (CoCreateInstance)
│   │   ├── 0x80040154  → Module 2: bitness → hive → ProcMon (NAME NOT FOUND vs ACCESS DENIED)
│   │   ├── 0x8007007E  → target DLL or dependency missing; correlate ProcMon with the failed load
│   │   ├── 0x80070005  → Module 7 §7.13: correlated events, effective permissions, authentication
│   │   ├── 0x80080005  → Module 7: RunAs account, server crash at startup, session 0
│   │   ├── 0x800706BA  → Module 7 §7.13: server availability, TCP 135 and actual RPC endpoint
│   │   └── 0x800401F0  → CoInitializeEx missing on this thread
│   │
│   ├── At QUERYINTERFACE
│   │   ├── 0x80004002 in-proc too   → the object genuinely lacks it
│   │   ├── 0x80004002 only out-of-proc → Module 4: marshaling not registered
│   │   └── 0x80040155  → Module 4: HKCR\Interface\{IID}\ProxyStubClsid32
│   │
│   ├── At CALL TIME
│   │   ├── 0x8001010E  → Module 3: wrong apartment, raw pointer shared
│   │   ├── 0x80010108  → object disconnected; check server exit or object/apartment shutdown
│   │   ├── 0x800706F7  → Module 4: mismatched IDL builds; get all three version numbers
│   │   ├── 0x80020009  → Module 5: open EXCEPINFO, the real error is inside
│   │   └── 0x80070005  → Module 7: effective call-access policy; also check application authorization
│   │
│   └── Intermittent / under load only
│       └── Module 3: threading. ThreadingModel + client apartment + shared raw pointers
│
├── Does it HANG?
│   ├── Compare repeated dumps (§8.4); confirm wait dependencies and useful progress.
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
    ├── ThreadingModel + client apartment → is every call marshaled? (Module 3 §3.3, Lab 4.2)
    ├── Late-bound IDispatch instead of vtable? (Module 5 Lab 5.1)
    └── Out-of-proc where in-proc would do? CLSCTX_ALL picking the wrong server?
```

---

## 8.9 Reference shelf

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

## 8.10 Final self-assessment

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

## 8.11 The ten rules, consolidated

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
