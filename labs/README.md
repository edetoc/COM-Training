# Lab snapshots

The course builds **one** `Calculator` component and grows it across all eight modules. That is
deliberate — you watch a single object acquire an IDL, a proxy/stub, a type library, an ATL
rewrite, and finally its own process.

The cost of that design is that a lab can depend on the one before it. These snapshots remove
that cost: **every lab names a stage below, and you can start from that stage cold.**

```
   Stage 1 ──► Stage 2 ──► Stage 3 ──► Stage 4 ──► Stage 5
   manual      in-proc      IDL +       ATL          out-of-proc
   IUnknown    DLL server   proxy/stub  rewrite      EXE server
```

## The five stages

| Stage | Folder | What you get | Labs that start here |
|---|---|---|---|
| 1 | [`stage-1-manual-iunknown/`](stage-1-manual-iunknown/) | `Calculator` with hand-written `IUnknown`, no COM runtime involved | 1.1, 1.2 |
| 2 | [`stage-2-inproc-server/`](stage-2-inproc-server/) | The same object as a registered in-proc DLL server, plus a client | 2.1, 2.2, 2.3, 2.4, 3.1, 3.2, 3.3, 6.1 |
| 3 | [`stage-3-idl-marshaling/`](stage-3-idl-marshaling/) | `ICalculator` in IDL, MIDL output, and a proxy/stub DLL | 4.1, 4.2, 4.3, 7.2 |
| 4 | [`stage-4-atl-server/`](stage-4-atl-server/) | The dual-interface ATL rewrite with a type library and events | 5.1, 5.2, 6.2, 6.3 |
| 5 | [`stage-5-exe-server/`](stage-5-exe-server/) | An out-of-process EXE server | 7.1, 7.3, 7.4 |

Labs 1.3, 8.1 and 8.2 need no starting code at all.

## How to use a stage

Each stage folder has its own `README.md` with the exact steps, but the shape is always the same:

1. Open a **Developer PowerShell for VS** (x64). Every build script needs `cl.exe` and `midl.exe`
   on `PATH`, and a plain PowerShell window does not have them.
2. `cd` into the stage folder.
3. Run `.\build.ps1`.
4. Follow the stage README's **Verify** section to confirm it works before starting the lab.

Nothing here uses `.vcxproj` files. The builds are single `cl.exe` invocations so you can read
exactly what is happening, and so nothing breaks when Visual Studio versions change.

## Working copies

Build output and any edits you make are ignored by git (see [`.gitignore`](.gitignore)), so you can
break a stage freely. To get back to a clean copy:

```powershell
git checkout -- labs/stage-2-inproc-server
git clean -fd labs/stage-2-inproc-server
```

Several labs ask you to **deliberately break** the code. Do that in a copy if you want to keep the
original around:

```powershell
Copy-Item labs\stage-2-inproc-server work\lab-2-4 -Recurse
```

## A note on GUIDs

Every stage uses the same placeholder GUIDs as the course text, so the module listings and these
files line up exactly:

| Name | GUID |
|---|---|
| `IID_ICalculator` | `{A1B2C3D4-0001-4000-9000-000000000001}` |
| `IID_IAdvancedCalculator` | `{A1B2C3D4-0002-4000-9000-000000000002}` |
| `CLSID_Calculator` | `{A1B2C3D4-1111-4000-9000-000000000001}` |
| `LIBID_TrainingCalcLib` | `{A1B2C3D4-9999-4000-9000-000000000099}` |

**These are for learning only.** Generate your own with `New-Guid` for anything you ship — Module 4
explains why reusing an IID for a changed interface is one of the worst bugs in COM.
