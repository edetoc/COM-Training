# Stage 1 — `Calculator` with hand-written `IUnknown`

**Used by:** Lab 1.1, Lab 1.2

No COM runtime, no registry, no `CoCreateInstance`. Just a C++ class that implements `IUnknown`
correctly, and a `CreateCalculator` helper standing in for a class factory. This is the object every
later stage builds on.

## Files

| File | What it is |
|---|---|
| `Calculator.h` | The two interfaces and their IIDs |
| `Calculator.cpp` | The `Calculator` class, with `AddRef`/`Release` tracing, plus `CreateCalculator` |
| `main.cpp` | Four functions that *prove* the `QueryInterface` rules with assertions |

## Steps

1. Open a **Developer PowerShell for VS (x64)**.
2. `cd` into this folder.
3. Build:
   ```powershell
   .\build.ps1
   ```
4. Run:
   ```powershell
   .\Lab01.exe
   ```

## Verify

You should see a trace like this, and the program should exit without an assertion firing:

```
[COM] CREATE   obj=000001... count=1  (identity)
[COM] ADDREF   obj=000001... count=2  (identity)
...
[COM] DESTROY  obj=000001... count=0  (identity)
```

**The one thing to check:** every `CREATE` has a matching `DESTROY`. If a `DESTROY` is missing, an
object leaked. That is the entire skill this stage teaches.

## Notes

- `Calculator.cpp` includes `<initguid.h>` **before** `Calculator.h`. That is what turns the
  `DEFINE_GUID` macros into actual definitions. Every other translation unit gets `extern`
  declarations instead. Do this in exactly one `.cpp` or you will get duplicate-symbol errors at
  link time.
- The build uses `/Zi` and `/DEBUG` so symbols exist. Lab 1.2 and several Module 8 exercises need
  them.

## Where this goes next

Stage 2 takes this same object, wraps it in a class factory and DLL exports, and registers it so
`CoCreateInstance` can find it.
