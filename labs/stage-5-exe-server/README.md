# Stage 5 — the out-of-process EXE server

**Used by:** Labs 7.1, 7.3, 7.4

The same `Calculator`, now living in its own process. Activation stops being a `LoadLibrary` and
becomes a `CreateProcess` performed by the SCM, with a security check in front of it.

## Prerequisite that is not optional

**Stage 3's `CalcPS.dll` must be registered first.** Out-of-proc means every call crosses a process
boundary, and a boundary needs marshaling. Without a registered proxy/stub, `CoCreateInstance`
returns `0x80004002 E_NOINTERFACE` and this stage stops at step one.

```powershell
cd ..\stage-3-idl-marshaling
# open CalcPS.vcxproj, build Debug|x64, then from an elevated prompt:
regsvr32 "<full path>\stage-3-idl-marshaling\x64\CalcPS.dll"
```

## Files

| File | What it is |
|---|---|
| `Calculator.h` | `ICalculator`, the CLSID, and an **AppID** — new at this stage |
| `CalcSrv.cpp` | Object, factory, `CoRegisterClassObject`, message loop, self-registration |
| `CalcSrvClient.cpp` | A client that asks specifically for `CLSCTX_LOCAL_SERVER` |
| `Stage5.sln` | The solution to open |

## Steps

1. Open **`Stage5.sln`** — it contains `CalcSrv` (the server) and `CalcSrvClient`.
2. Pick **Debug | x64**, then **Build → Build Solution**.
3. Register from an **elevated** prompt (this writes `LocalServer32` and the AppID):
   ```powershell
   & "<full path>\labs\stage-5-exe-server\x64\CalcSrv.exe" -RegServer
   ```
4. Right-click **CalcSrvClient → Set as Startup Project**, then press **F5**.

To debug the *server*, use **Debug → Attach to Process** and pick `CalcSrv.exe` while the client is
waiting at its prompt. You cannot F5 the server directly: the SCM launches it, not you.

## Verify

```
Add   -> hr=0x00000000  40 + 2 = 42
Sub   -> hr=0x00000000  44 - 2 = 42
```

While the client waits at the prompt, check **Task Manager** — `CalcSrv.exe` is running as a
separate process. Press Enter and it exits within a second or two, because releasing the last
object drives `CoReleaseServerProcess` to zero.

That start-on-demand, exit-when-idle behaviour is the whole point of the stage. Watch it happen.

## Troubleshooting

| HRESULT | Meaning | Fix |
|---|---|---|
| `0x80040154` | not registered | run `CalcSrv.exe -RegServer` **elevated** |
| `0x80004002` | no marshaling | register Stage 3's `CalcPS.dll` |
| `0x80080005` | server failed to start | run `CalcSrv.exe` by hand — it should show a message box; check the Application event log |
| `0x80070005` | access denied | Launch permission (§7.6). Expected during Lab 7.3 |

## Cleaning up

```powershell
& ".\x64\CalcSrv.exe" -UnregServer     # elevated
```

Lab 7.3 edits this AppID's security in `dcomcnfg`. **Export
`HKCR\AppID\{B1B2C3D4-2222-4000-9000-000000000002}` before you start**, and prefer a VM.

## Notes

- `REGCLS_MULTI_SEPARATE | REGCLS_SUSPENDED`, then `CoResumeClassObjects()`. Registering suspended
  and resuming afterwards closes a real race — see §7.9's `REGCLS` table.
- The server is `/SUBSYSTEM:WINDOWS` because it needs a message loop and no console.
- Run by hand without `-Embedding` it shows a message box instead of hanging invisibly. Servers
  that *do* hang invisibly are a classic session-0 support case (§7.3).
