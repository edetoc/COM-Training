# Stage 3 — IDL, MIDL, and the proxy/stub DLL

**Used by:** Labs 4.1, 4.2, 4.3, 7.2

`ICalculator` described properly in IDL, compiled by MIDL, and built into a registered proxy/stub
DLL. This is what lets the interface cross an apartment or process boundary — Stage 2's object
cannot leave its apartment without it.

## Files

| File | What it is |
|---|---|
| `Calculator.idl` | The interface, with the memory-ownership attributes Module 4 dissects |
| `CalcPS.def` | Exports for the proxy/stub DLL |
| `CalcPS.vcxproj` | The project to open |

## Steps

1. Open **`CalcPS.vcxproj`**.
2. Pick **Debug | x64** and build.
3. Register the proxy/stub from an **elevated** prompt:
   ```powershell
   regsvr32 "<full path>\labs\stage-3-idl-marshaling\x64\CalcPS.dll"
   ```

This one is an **NMake project** rather than a normal C++ one, because MSBuild needs its source
list before the build starts and `dlldata.c` does not exist until MIDL has run. Open the `.vcxproj`
in a text editor — the three commands (MIDL, `cl`, `link`) are written out in full, and together
they are the entire recipe for a proxy/stub DLL.

> Lab 4.1 itself has you run MIDL **by hand** from a Developer PowerShell, which is worth doing
> once. This project is the shortcut for later labs that merely need the DLL registered.

## Verify

```powershell
# The interface should now have a ProxyStubClsid32 entry:
Get-Item "Registry::HKEY_CLASSES_ROOT\Interface\{A1B2C3D4-0001-4000-9000-000000000001}\ProxyStubClsid32"
```

If that key exists, marshaling for `ICalculator` is registered and Labs 4.2, 7.1 and 7.2 will work.

## Read the output — this *is* Lab 4.1

The generated files land in `.\x64\`. Open each one:

| File | What to look for |
|---|---|
| `Calculator.h` | Your interface in three forms: C++, C, and `CINTERFACE` macros. Note how `MIDL_INTERFACE` expands. |
| `Calculator_i.c` | The actual IID/CLSID byte definitions — this is what `__uuidof` resolves against |
| `Calculator_p.c` | NDR format strings: your interface compiled into bytecode a marshaler interprets |
| `dlldata.c` | The proxy DLL's own `DllGetClassObject` / `DllRegisterServer` |

## Building 32-bit as well (Lab 4.2)

The surrogate experiment pairs a 32-bit DLL with a 64-bit client, and **each side loads its own
proxy**. Switch the platform dropdown to **x86**, build, and register with the 32-bit `regsvr32`:

```powershell
C:\Windows\SysWOW64\regsvr32.exe "<full path>\x86\CalcPS.dll"
```

## Notes

- `REGISTER_PROXY_DLL` is the define that makes `dlldata.c` emit the registration exports. Without
  it the DLL builds but `regsvr32` reports that the entry point is missing.
- Unregistering this DLL (`regsvr32 /u`) is Lab 4.2 step 2. Expect the failure to appear at the
  first *call*, not at activation — record exactly where.

## Where this goes next

Stage 4 rewrites the server in ATL with a dual interface and a type library; Stage 5 moves it into
its own process, which is where a registered proxy/stub stops being optional.
