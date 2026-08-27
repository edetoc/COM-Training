# Stage 2 — the in-proc DLL server

**Used by:** Labs 2.1, 2.2, 2.3, 2.4, 3.1, 3.2, 3.3, 6.1

The Stage 1 object, now wrapped in a class factory and the four DLL exports, registered so
`CoCreateInstance` can find it by CLSID. This is the most reused stage in the course.

## Files

| File | What it is |
|---|---|
| `Calculator.h` | `ICalculator` plus the IID and the **CLSID** |
| `Calc.cpp` | Object, class factory, `DllGetClassObject`, `DllCanUnloadNow`, self-registration |
| `Calc.def` | Exports the four entry points — without this `regsvr32` cannot find them |
| `CalcClient.cpp` | A console client that activates by CLSID and by ProgID |

## Steps

1. Open a **Developer PowerShell for VS (x64)**.
2. `cd` into this folder.
3. Build:
   ```powershell
   .\build.ps1
   ```
   Output lands in `.\x64\`.
4. Open a **second, elevated** PowerShell (registration writes to `HKLM`), and register:
   ```powershell
   regsvr32 "<full path>\labs\stage-2-inproc-server\x64\Calc.dll"
   ```
5. Back in the normal shell, run the client:
   ```powershell
   .\x64\CalcClient.exe
   ```

## Verify

```
40 + 2 = 42
CLSIDFromProgID: 0x00000000
```

If you get `0x80040154 REGDB_E_CLASSNOTREG`, the DLL is not registered, or you registered the
32-bit build and are running the 64-bit client (that is Lab 2.2, on purpose).

## Building 32-bit as well (Lab 2.2)

Lab 2.2 needs both bitnesses. Open a **Developer PowerShell for VS (x86)** and run:

```powershell
.\build.ps1 -Arch x86
```

Then register that copy with the **32-bit** `regsvr32`:

```powershell
C:\Windows\SysWOW64\regsvr32.exe "<full path>\x86\Calc.dll"
```

The script refuses to build a bitness that does not match your shell, because a silent mismatch
here makes Lab 2.2 impossible to reason about.

## Cleaning up

Always unregister before rebuilding to a different path, or you leave a stale `InprocServer32`
pointing at a file that no longer exists — which is Lab 2.4, row 3:

```powershell
regsvr32 /u "<full path>\x64\Calc.dll"
```

## Notes

- `ThreadingModel` is registered as `Both`. **Lab 3.1 requires `Apartment` instead** — change the
  string in `DllRegisterServer`, rebuild, and re-register. Lab 3.3 has you register the same DLL
  under five CLSIDs, one per model.
- The factory is a static singleton, so its `AddRef`/`Release` move `g_cLocks` rather than a
  per-object count. That is deliberate and matches §2.5.1.

## Where this goes next

Stage 3 describes this same interface in IDL and generates the proxy/stub that lets it cross an
apartment or process boundary.
