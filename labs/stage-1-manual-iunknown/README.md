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
| `Lab01.sln` | The solution to open |

## Steps

1. Open **`Lab01.sln`**.
2. Pick **Debug | x64** in the toolbar.
3. Press **F5**.

Use F5 rather than Ctrl+F5. The whole point of this lab is watching the reference count move, and
breakpoints on `AddRef` / `Release` are how you do that.

Output lands in `x64\Lab01.exe`.

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
- The project builds with debug symbols in both configurations. Lab 1.2 and several Module 8
  exercises need them.

## Where this goes next

Stage 2 takes this same object, wraps it in a class factory and DLL exports, and registers it so
`CoCreateInstance` can find it.
