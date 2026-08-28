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
| `Stage2.sln` | The solution to open |

## Steps

1. Open **`Stage2.sln`** — it contains both `Calc` (the DLL) and `CalcClient`.
2. Pick **Debug | x64**, then **Build → Build Solution**.
3. Register the DLL from an **elevated** prompt (registration writes to `HKLM`):
   ```powershell
   regsvr32 "<full path>\labs\stage-2-inproc-server\x64\Calc.dll"
   ```
4. Right-click **CalcClient → Set as Startup Project**, then press **F5**.

The `Calc.def` module-definition file is already wired into the linker settings. §2.5.3 has you set
that by hand and warns what happens when you forget, so it is worth opening
**Calc → Properties → Linker → Input → Module Definition File** to see where it lives.

Output lands in `x64\Calc.dll` and `x64\CalcClient.exe`.

## Verify

```
40 + 2 = 42
CLSIDFromProgID: 0x00000000
```

If you get `0x80040154 REGDB_E_CLASSNOTREG`, the DLL is not registered, or you registered the
32-bit build and are running the 64-bit client (that is Lab 2.2, on purpose).

## Building 32-bit as well (Lab 2.2)

Lab 2.2 needs both bitnesses. Change the platform dropdown to **x86** and build again — that is the
whole step. Output goes to `x86\` alongside the x64 copy.

Then register that copy with the **32-bit** `regsvr32`:

```powershell
C:\Windows\SysWOW64\regsvr32.exe "<full path>\x86\Calc.dll"
```

Note the two different `regsvr32` binaries: `System32` is the **64-bit** one, `SysWOW64` is the
**32-bit** one. Yes, that naming is backwards (§2.3).

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
