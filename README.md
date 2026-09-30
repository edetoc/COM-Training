# COM Training — from first principles to real-world troubleshooting

COM is more than thirty years old, and it's still everywhere in Windows: File Explorer and its extensions, Office automation, WMI, DirectX, .NET interop, and the Windows Runtime all run on it. When one of them breaks, the only clue is often a bare error code such as `0x80040154` — and the fix depends on understanding what COM was doing underneath.

This course takes you from "what is an interface pointer?" to diagnosing activation, threading, and DCOM security problems with confidence, with hands-on labs at every step.

**Who it's for:** developers and support/escalation engineers who are new to COM. You should be comfortable reading C++; [Module 0](modules/00-why-com-exists.md) explains how much you need.

**By the end, you'll be able to:**

- Explain what COM guarantees, and why its rules exist.
- Build COM components by hand and with frameworks, and use them from C++, C#, and scripts.
- Diagnose the classic failures — "class not registered", wrong-thread calls, deadlocks, leaks, access denied — from the evidence.
- Configure and troubleshoot out-of-process and remote (DCOM) servers safely.

## The course

Nine modules, taken in order. Module 0 sets the scene; Modules 1–7 pair concepts with hands-on labs; Module 8 is a diagnostics reference. Modules 0–7 each end with a checkpoint, answers included.

| # | Module | You'll learn to… |
|---|---|---|
| 0 | [Why COM exists](modules/00-why-com-exists.md) | Understand why a C++ class can't safely cross a DLL boundary, how COM's three pillars fix that, and the core vocabulary — then make your first COM call |
| 1 | [`IUnknown`, vtables, and lifetime](modules/01-iunknown-and-lifetime.md) | See what an interface pointer is in memory, implement `IUnknown` by hand, apply the reference-counting rules, read `HRESULT`s, and spot leaks and over-releases |
| 2 | [Activation, registration, and the registry](modules/02-activation-and-registry.md) | Follow `CoCreateInstance` from CLSID to running object — registry keys, class factories, 32/64-bit registry views, registration-free COM — and diagnose "class not registered" with Process Monitor |
| 3 | [Threading and apartments](modules/03-apartments-and-threading.md) | Understand apartments and `ThreadingModel`, pass objects between threads safely with marshaling or the Global Interface Table, and recognize reentrancy and deadlocks |
| 4 | [Interfaces, IDL, MIDL, and marshaling](modules/04-idl-and-marshaling.md) | Describe interfaces in IDL, get memory ownership right, choose proxy/stub or type-library marshaling, version interfaces safely, and fix "works in-process, fails out-of-process" |
| 5 | [Automation, `IDispatch`, and scripting](modules/05-automation-and-idispatch.md) | Make components usable from scripts and .NET with `IDispatch` and dual interfaces, handle `VARIANT` and `BSTR`, and add events, enumerators, and rich errors |
| 6 | [Choosing COM tools: ATL, WRL, WIL, C++/WinRT, .NET](modules/06-frameworks-and-interop.md) | Compare what ATL, WRL, WIL, C++/WinRT, and .NET interop each provide, choose the right one for a task, and avoid the classic .NET interop bugs |
| 7 | [DCOM, security, and out-of-proc](modules/07-dcom-and-security.md) | Host servers out of process or in a surrogate, set identity and launch/access permissions, read Event 10016, handle UAC and DCOM hardening, and call across machines |
| 8 | [Debugging and diagnostics](modules/08-debugging-and-diagnostics.md) | Collect the right evidence for each symptom, analyze it with Process Monitor, WinDbg, ETW, and Application Verifier, and follow a triage flowchart to the root cause |

**Appendices** are reference material. Read them when a module or a ticket sends you there:

| | Appendix | Read it when |
|---|---|---|
| A | [Monikers, persistence, and structured storage](modules/appendix-a-monikers-and-persistence.md) | You meet `GetObject`, WMI connection strings, the Running Object Table, or compound files |
| B | [COM+ and WinRT](modules/appendix-b-com-plus-and-winrt.md) | You support COM+ server applications, or work with the Windows Runtime |

---

## How to work through it

- **Go in order.** Each module builds on the previous one, and COM is unforgiving of skipped fundamentals. A module a week is a comfortable pace.
- **Do the labs.** Reading about COM produces false confidence; breaking things on purpose is how the rules stick.
- **Keep a `COM-Notes.md` file** with every `HRESULT` you hit and what actually caused it. It becomes your personal support cheat sheet.
- **Answer each checkpoint before opening the answers.**

### Labs

The labs grow one `Calculator` component through five stages in [`labs/`](labs/), from a hand-written `IUnknown` to an out-of-process server. Each stage supplies starting code, so you can begin any lab without finishing the previous one; every lab names its starting stage in its **Requirements** block. Build steps are in [labs/README.md](labs/README.md) and each stage's README.

| Stage | Starting code for |
|---|---|
| [1 — manual `IUnknown`](labs/stage-1-manual-iunknown/) | Labs 1.1–1.2 |
| [2 — in-proc DLL server](labs/stage-2-inproc-server/) | Labs 2.1–2.4, 3.1–3.3, 6.1 |
| [3 — IDL and proxy/stub](labs/stage-3-idl-marshaling/) | Labs 4.1–4.3 |
| [4 — ATL rewrite](labs/stage-4-atl-server/) | Labs 5.1–5.2, 6.2–6.3 |
| [5 — out-of-process hosting](labs/stage-5-exe-server/) | Labs 7.1–7.3 (self-contained) |

### Before you start

Install Visual Studio with *Desktop development with C++* (including ATL) and *.NET desktop development*, plus the Windows SDK, WinDbg, the Sysinternals tools (Process Monitor and Process Explorer), and OleView.NET. [Module 0 §0.8](modules/00-why-com-exists.md#08-tooling-setup-do-this-now) walks you through setup and a quick smoke test.

---

## Progress tracker

| # | Module | Read | Labs done | Checkpoint passed | Notes written |
|---|---|---|---|---|---|
| 0 | [Why COM exists](modules/00-why-com-exists.md) | ☐ | — | ☐ | ☐ |
| 1 | [IUnknown & lifetime](modules/01-iunknown-and-lifetime.md) | ☐ | ☐ | ☐ | ☐ |
| 2 | [Activation & registry](modules/02-activation-and-registry.md) | ☐ | ☐ | ☐ | ☐ |
| 3 | [Threading & apartments](modules/03-apartments-and-threading.md) | ☐ | ☐ | ☐ | ☐ |
| 4 | [IDL & marshaling](modules/04-idl-and-marshaling.md) | ☐ | ☐ | ☐ | ☐ |
| 5 | [Automation & IDispatch](modules/05-automation-and-idispatch.md) | ☐ | ☐ | ☐ | ☐ |
| 6 | [Choosing COM tools and interop](modules/06-frameworks-and-interop.md) | ☐ | ☐ | ☐ | ☐ |
| 7 | [DCOM & security](modules/07-dcom-and-security.md) | ☐ | ☐ | ☐ | ☐ |
| 8 | [Debugging & diagnostics](modules/08-debugging-and-diagnostics.md) | ☐ | — | ☐ | ☐ |
| A | [Monikers & persistence](modules/appendix-a-monikers-and-persistence.md) | ☐ | — | ☐ | ☐ |
| B | [COM+ & WinRT](modules/appendix-b-com-plus-and-winrt.md) | ☐ | — | ☐ | ☐ |

---

## Keep going

[Module 8](modules/08-debugging-and-diagnostics.md) ends with a reference shelf of books, online resources, and tools, a final self-assessment, and the ten rules to carry forward. Corrections and suggestions are welcome.

**Ready? Start with [Module 0 — Why COM exists](modules/00-why-com-exists.md).**
