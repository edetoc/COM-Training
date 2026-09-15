# Stage 5 - self-contained out-of-process hosting

**Used by:** Lab 7.1 (hosting), Lab 7.2 (permissions), and Lab 7.3 (remote DCOM).

This folder contains everything needed for Lab 7.1. No binaries, project edits, or registrations from previous modules are required. The main path is **x64 only** and does not require a VM. Administrator rights are needed for registration, not for normal client execution.

## One solution, four projects

| Project | Produces | Purpose |
|---|---|---|
| `CalcSrv` | `CalcSrv.exe` | An EXE server with its own factory and process lifetime |
| `CalcDll` | `CalcDll.dll` | A DLL that can run in-process or inside Windows' `dllhost.exe` |
| `CalcSrvClient` | `CalcSrvClient.exe` | Selects either class and calls `Add` and `Subtract` |
| `Stage5PS` | `Stage5PS.dll` | Matching proxy/stub generated from this folder's `Calculator.idl` |

The generated header is shared by all three native projects. The interface contains **only Add and Subtract**, with a dedicated IID. The EXE and DLL have different CLSIDs and AppIDs, so both can stay registered without switching registry entries. Earlier course components are left alone.

## Build and register

1. Open **Stage5.sln** in Visual Studio with Desktop development with C++ and a Windows SDK installed.
2. Choose **Debug | x64**, then **Build > Build Solution**. Project dependencies build `Stage5PS` first and generate the header. Do not copy generated files from earlier labs. The four binaries appear in this folder's `x64` directory.
3. Open a **64-bit PowerShell window as administrator**, change to this folder, and run:

  ```powershell
  .\Setup.ps1 -Action Register
  ```

  This registers the matching proxy/stub, the DLL and its surrogate AppID, and the EXE. The script waits for each registration and stops on failure. It neither elevates itself nor changes machine-wide DCOM permissions, firewall rules, or server identities. If PowerShell blocks scripts, follow your machine's script policy; do not disable organizational controls.
4. Close the elevated window. Use an **ordinary PowerShell window** in this folder to run the clients below. Keep the registered binaries at their build paths, and close clients before rebuilding.

`Setup.ps1 -Action Register -WhatIf` previews the action without changing registration. If setup partially fails, correct the reported problem and rerun it, or use the cleanup command to remove this lab's partial registration. Do not delete the binaries before cleanup.

## Run and compare

```powershell
.\x64\CalcSrvClient.exe
.\x64\CalcSrvClient.exe --surrogate
.\x64\CalcSrvClient.exe --inproc
```

Run these **one at a time**. Each prints:

```text
Add   -> hr=0x00000000  40 + 2 = 42
Sub   -> hr=0x00000000  44 - 2 = 42
```

It then prints the expected host and waits for Enter. While paused, use Process Explorer to verify the actual host; the printed label alone is not proof.

| Client arguments | Requested class/context | What to inspect |
|---|---|---|
| None | EXE CLSID, `CLSCTX_LOCAL_SERVER` | `CalcSrv.exe` is a separate process |
| `--surrogate` | DLL CLSID, `CLSCTX_LOCAL_SERVER` | `CalcDll.dll` is loaded in `dllhost.exe` |
| `--inproc` | DLL CLSID, `CLSCTX_INPROC_SERVER` | `CalcDll.dll` is loaded in `CalcSrvClient.exe` |

Use **Ctrl+D** to show the selected process's DLLs. Press Enter to release the object. The EXE should exit when its final object/server lock goes away. Windows manages the surrogate's idle shutdown; do not expect the same immediate exit timing or terminate unrelated `dllhost.exe` processes.

For unattended checks, append `--auto`; the client exits nonzero on activation, call, or result failure. To debug the EXE, attach Visual Studio to `CalcSrv.exe` while the client is waiting. Set a breakpoint in `Calculator::Add`, then launch a second client to make another call. Startup breakpoints need a debugger attached before COM launches the server.

## Optional: 64-bit client, 32-bit surrogate

The main comparison above is complete without this exercise.

1. In the same solution, build **Debug | x86**. The output is in `x86` (project platform `Win32`).
2. From elevated **64-bit PowerShell**, run:

  ```powershell
  .\Setup.ps1 -Action Register -Architecture x86
  ```

  This adds only the x86 DLL and proxy/stub. It does not register the x86 EXE. Shared AppID/ProgID settings are identical for both DLL builds.
3. From ordinary PowerShell, run the **x64** client:

  ```powershell
  .\x64\CalcSrvClient.exe --surrogate --x86
  ```

  The client uses `CLSCTX_ACTIVATE_32_BIT_SERVER`. Confirm that the client is 64-bit and the `dllhost.exe` containing `CalcDll.dll` is 32-bit. No code or registry switching is needed.

## Registration reference

| Identifier | Value |
|---|---|
| Interface IID | `{68AE8A82-BD63-4110-9FBF-C649EF09C3D8}` |
| EXE CLSID | `{647E1E3B-42CF-4A82-B91E-6DCADB217D9B}` |
| EXE AppID | `{A57CC9FD-4A41-4579-9D75-97ECC20C504B}` |
| DLL CLSID | `{21C831C0-DA4D-4727-97C4-E52ED08D9F18}` |
| DLL AppID | `{C714BFAA-5711-46A7-91E1-42E2B9A14920}` |

The EXE is displayed as **Stage 5 Calculator EXE** in Component Services and uses `Training.Stage5.Calculator.1`. The DLL uses `Training.Stage5.Surrogate.1`. `Registration.h` writes the class and AppID settings; DLL registration includes an empty `DllSurrogate` string. No custom surrogate or separate type-library registration is needed: this lab uses its generated proxy/stub.

The EXE calls `CoInitializeSecurity` with `EOAC_APPID` before publishing its factory, so Lab 7.2 can configure its policy through the AppID. Back up that AppID before the later permission experiments. Do not use a Windows-owned AppID instead.

## Troubleshooting

| Symptom | Check |
|---|---|
| Missing generated header | Build the solution, including the `Stage5PS` dependency |
| `0x80040154` | Run this folder's setup for the architecture requested by the client |
| `0x80004002` | Confirm this lab's `Stage5PS.dll` is registered; earlier labs' proxies are not a substitute |
| `0x80080005` | Inspect EXE startup/security initialization and class registration; confirm the executable path exists |
| `0x80070005` | Check the relevant AppID and machine policy; do not broadly grant permissions |
| DLL cannot be rebuilt | Release the client objects; check which process still has this specific DLL loaded |

## Cleanup

Keep registration while doing Labs 7.2 and 7.3. When finished, close the lab clients, let the servers finish, and run from elevated 64-bit PowerShell in this folder:

```powershell
.\Setup.ps1 -Action Unregister -Architecture x86
.\Setup.ps1 -Action Unregister
```

**Skip the first command if you did not register the optional x86 build.** Cleanup removes this lab's class, AppID, ProgID, and proxy registrations. The DLL unregistration preserves shared AppID/ProgID keys while the other architecture remains registered. It does not unregister Stage 2, Stage 3, or the ATL server from earlier modules.
