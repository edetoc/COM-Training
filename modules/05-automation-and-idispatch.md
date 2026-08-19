# Module 5 — Automation, `IDispatch`, and scripting

**Time: 1 week.**

Everything so far assumed a compiled client that knows the vtable at build time. But VBScript, JScript, PowerShell, VBA, and old VB have no vtables and no headers. They discover methods **at runtime, by name**. This module explains how, and covers the family of technologies built on top: type libraries, events, enumerators, and rich error reporting.

**Contents**

- [5.1 Early binding vs late binding](#51-early-binding-vs-late-binding)
- [5.2 `IDispatch`](#52-idispatch)
- [5.3 Dual interfaces — the best of both worlds](#53-dual-interfaces--the-best-of-both-worlds)
- [5.4 Implementing `IDispatch` the easy way](#54-implementing-idispatch-the-easy-way)
- [5.5 The Automation type system](#55-the-automation-type-system)
- [5.6 Rich errors: `IErrorInfo`](#56-rich-errors-ierrorinfo)
- [5.7 Connection points — COM events](#57-connection-points--com-events)
- [5.8 Enumerators and `For Each`](#58-enumerators-and-for-each)
- [5.9 LAB 5.1 — A dual interface driven from four languages](#59-lab-51--a-dual-interface-driven-from-four-languages)
- [5.10 LAB 5.2 — Events and the `Unadvise` leak](#510-lab-52--events-and-the-unadvise-leak)
- [5.11 Automation error codes](#511-automation-error-codes)
- [5.12 Checkpoint](#512-checkpoint)
- [5.13 Rules to carry forward](#513-rules-to-carry-forward)

---

## 5.1 Early binding vs late binding

**Early binding (vtable binding)** — the compiler knows the interface at build time and emits `call [vtable + offset]`. Fast, type-checked, no runtime lookup.

**Late binding** — the client has only a name string at runtime:

```vbscript
Set obj = CreateObject("Training.Calculator.1")
result = obj.Add(2, 3)          ' "Add"? Never heard of it until this instant.
```

The engine must, at runtime: turn `"Add"` into something callable, pack the arguments, invoke it, and unpack the result. `IDispatch` is the interface that makes this possible.

---

## 5.2 `IDispatch`

```cpp
struct IDispatch : public IUnknown
{
    HRESULT GetTypeInfoCount(UINT* pctinfo);
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                          LCID lcid, DISPID* rgDispId);
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                   DISPPARAMS* pDispParams, VARIANT* pVarResult,
                   EXCEPINFO* pExcepInfo, UINT* puArgErr);
};
```

Two methods do the work:

- **`GetIDsOfNames`** — "what is the DISPID of the member called `Add`?" A **DISPID** is just a `LONG` identifying a member.
- **`Invoke`** — "call member DISPID 1 with these `VARIANT` arguments."

So `obj.Add(2, 3)` becomes:

```
GetIDsOfNames(IID_NULL, ["Add"], 1, LOCALE_USER_DEFAULT, &dispid)   -> dispid = 1
Invoke(1, IID_NULL, lcid, DISPATCH_METHOD, &params, &result, &excep, &argErr)
```

### `IUnknown` vs `IDispatch` — the distinction to keep straight

Beginners frequently treat these as alternatives. They are not: **`IDispatch` derives from `IUnknown`.** Every `IDispatch` is an `IUnknown`; the reverse is false.

```
  IUnknown          slots 0-2   QueryInterface, AddRef, Release
     └─ IDispatch   slots 3-6   GetTypeInfoCount, GetTypeInfo, GetIDsOfNames, Invoke
```

| | `IUnknown` | `IDispatch` |
|---|---|---|
| **Question it answers** | "What else can you do, and when should you die?" | "What is the thing called `Add`, and can you call it for me?" |
| **Responsibility** | Identity, navigation, lifetime | Runtime discovery and invocation by name |
| **Implemented by** | **Every** COM object — mandatory | Only objects that opt into Automation |
| **Binding** | Early — the compiler knows the vtable slot | Late — the name is resolved at runtime |
| **Client needs at build time** | The IDL/header | Nothing; a string suffices |
| **Parameter types** | Anything MIDL can describe | Automation-compatible only (`VARIANT`, `BSTR`, `SAFEARRAY`, …) |
| **Call cost** | One indirect call — nanoseconds | Name lookup + `VARIANT` packing + coercion — microseconds |
| **Error detail** | `HRESULT` (+ `IErrorInfo`) | `HRESULT` + `EXCEPINFO` |
| **Typical caller** | C++, Rust, compiled .NET | VBScript, JScript, VBA, PowerShell, C# `dynamic` |

The two are complementary, not competing:

```cpp
// IUnknown's job: get me the contract I want, and manage lifetime.
pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp);

// IDispatch's job: let me call it without knowing that contract at build time.
pDisp->GetIDsOfNames(...); pDisp->Invoke(...);

pDisp->Release();   // still IUnknown's job
```

Note the last line: even a purely late-bound script client is doing `AddRef`/`Release` underneath — the script engine (or the CLR's RCW) calls them on your behalf. **`IUnknown` never goes away.** §5.3's dual interfaces let one object serve both access paths simultaneously.

### `wFlags`

| Flag | Meaning |
|---|---|
| `DISPATCH_METHOD` | `obj.Foo(1)` |
| `DISPATCH_PROPERTYGET` | `x = obj.Foo` |
| `DISPATCH_PROPERTYPUT` | `obj.Foo = 5` |
| `DISPATCH_PROPERTYPUTREF` | `Set obj.Foo = other` (assign by reference) |

A member can be several at once — hence the flags are a bitmask, and `DISPATCH_METHOD | DISPATCH_PROPERTYGET` is common for members that could be either (a script engine often can't tell).

### `DISPPARAMS` — and its two traps

```cpp
typedef struct tagDISPPARAMS {
    VARIANTARG* rgvarg;             // arguments, IN REVERSE ORDER
    DISPID*     rgdispidNamedArgs;  // DISPIDs for named args
    UINT        cArgs;
    UINT        cNamedArgs;
} DISPPARAMS;
```

1. **Arguments are in reverse order.** `rgvarg[0]` is the *last* argument. This trips up everyone once.
2. **Property puts use a named argument.** For `DISPATCH_PROPERTYPUT`, the value is passed with the special DISPID `DISPID_PROPERTYPUT` (`-3`) in `rgdispidNamedArgs`:

```cpp
VARIANT v; v.vt = VT_I4; v.lVal = 42;
DISPID putid = DISPID_PROPERTYPUT;
DISPPARAMS dp = { &v, &putid, 1, 1 };
pDisp->Invoke(dispidPrecision, IID_NULL, lcid, DISPATCH_PROPERTYPUT,
              &dp, nullptr, &excep, nullptr);
```

Forgetting the named argument gives `DISP_E_PARAMNOTOPTIONAL` or `DISP_E_BADPARAMCOUNT`.

### Well-known DISPIDs

| DISPID | Value | Meaning |
|---|---|---|
| `DISPID_VALUE` | `0` | The **default member**. `obj` alone (no member) invokes this. In VB, `obj` and `obj.Value` are the same. |
| `DISPID_NEWENUM` | `-4` | Returns `IEnumVARIANT` — this is what makes `For Each` work. |
| `DISPID_PROPERTYPUT` | `-3` | The named arg for property puts |
| `DISPID_UNKNOWN` | `-1` | "no member" |
| `DISPID_EVALUATE` | `-5` | For `[ ]` evaluation syntax |
| `DISPID_CONSTRUCTOR` / `_DESTRUCTOR` | `-6` / `-7` | Rare |

---

## 5.3 Dual interfaces — the best of both worlds

A **dual** interface derives from `IDispatch` *and* declares its methods in the vtable after `IDispatch`'s. So:

- Compiled clients call through the vtable — full speed.
- Scripts call `GetIDsOfNames`/`Invoke` — full flexibility.

```
  IUnknown          slots 0-2   QueryInterface, AddRef, Release
  IDispatch         slots 3-6   GetTypeInfoCount, GetTypeInfo, GetIDsOfNames, Invoke
  ICalculator       slots 7+    Add, Subtract, get_Precision, put_Precision
```

```idl
[
    object,
    uuid(A1B2C3D4-0001-4000-9000-000000000001),
    dual,                          // implies oleautomation + IDispatch-derived
    nonextensible,                 // the interface is complete; scripts can't add members
    pointer_default(unique)
]
interface ICalculator : IDispatch
{
    [id(1)] HRESULT Add([in] LONG a, [in] LONG b, [out, retval] LONG* result);
    [id(2)] HRESULT Subtract([in] LONG a, [in] LONG b, [out, retval] LONG* result);
    [id(3), propget] HRESULT Precision([out, retval] LONG* value);
    [id(3), propput] HRESULT Precision([in] LONG value);
    [id(DISPID_VALUE)] HRESULT Describe([out, retval] BSTR* text);   // default member
}
```

- `[id(n)]` assigns the DISPID explicitly. **Do this.** Auto-assigned DISPIDs can shift between builds and break compiled scripts.
- The propget/propput pair **shares** DISPID 3 — that's how a property is expressed.
- `dual` implies `oleautomation`, so **typelib marshaling works** (Module 4) and you need no proxy/stub DLL.

### Restrictions on dual interfaces

- Automation-compatible types only.
- Must derive **directly** from `IDispatch` (not from another dual interface — deriving breaks `GetIDsOfNames` for the base's members in some engines, and confuses tooling).
- All methods return `HRESULT`, with the real result as `[out, retval]`.

---

## 5.4 Implementing `IDispatch` the easy way

You almost never hand-write `GetIDsOfNames`/`Invoke`. Instead, delegate to the type library:

```cpp
class Calculator : public ICalculator
{
    ITypeInfo* m_pTypeInfo = nullptr;

    HRESULT EnsureTypeInfo()
    {
        if (m_pTypeInfo) return S_OK;
        ITypeLib* pTL = nullptr;
        HRESULT hr = LoadRegTypeLib(LIBID_TrainingCalcLib, 1, 0,
                                    LOCALE_SYSTEM_DEFAULT, &pTL);
        if (FAILED(hr)) return hr;
        hr = pTL->GetTypeInfoOfGuid(IID_ICalculator, &m_pTypeInfo);
        pTL->Release();
        return hr;
    }

public:
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override
    {
        if (!pctinfo) return E_POINTER;
        *pctinfo = 1;                       // 1 = "I have type info"
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID, ITypeInfo** ppTI) override
    {
        if (!ppTI) return E_POINTER;
        *ppTI = nullptr;
        if (iTInfo != 0) return DISP_E_BADINDEX;
        HRESULT hr = EnsureTypeInfo();
        if (FAILED(hr)) return hr;
        *ppTI = m_pTypeInfo;
        (*ppTI)->AddRef();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* rgszNames, UINT cNames,
                                            LCID, DISPID* rgDispId) override
    {
        HRESULT hr = EnsureTypeInfo();
        if (FAILED(hr)) return hr;
        return DispGetIDsOfNames(m_pTypeInfo, rgszNames, cNames, rgDispId);
    }

    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispid, REFIID, LCID, WORD wFlags,
                                     DISPPARAMS* pDispParams, VARIANT* pVarResult,
                                     EXCEPINFO* pExcepInfo, UINT* puArgErr) override
    {
        HRESULT hr = EnsureTypeInfo();
        if (FAILED(hr)) return hr;
        // ITypeInfo::Invoke reads the type library, converts the VARIANTs to the
        // declared parameter types, and calls the right vtable slot for us.
        return m_pTypeInfo->Invoke(static_cast<ICalculator*>(this), dispid, wFlags,
                                   pDispParams, pVarResult, pExcepInfo, puArgErr);
    }

    // ... ICalculator methods, plain vtable implementations ...
};
```

`ITypeInfo::Invoke` does the entire job: name lookup, type coercion, argument reordering, calling the vtable slot, and packaging the result. **In ATL this is `IDispatchImpl<ICalculator, &IID_ICalculator, &LIBID_TrainingCalcLib>` — one line.**

---

## 5.5 The Automation type system

### `VARIANT`

A tagged union. `vt` says what's in it.

| `vt` | Contents | Field |
|---|---|---|
| `VT_EMPTY` | nothing (uninitialized) | — |
| `VT_NULL` | SQL-style null | — |
| `VT_I4` / `VT_I2` / `VT_I1` | signed int | `lVal` / `iVal` / `cVal` |
| `VT_UI4` etc. | unsigned | `ulVal` … |
| `VT_R8` / `VT_R4` | double / float | `dblVal` / `fltVal` |
| `VT_BSTR` | string | `bstrVal` |
| `VT_BOOL` | `VARIANT_TRUE` (-1) / `VARIANT_FALSE` (0) | `boolVal` |
| `VT_DATE` | double, days since 1899-12-30 | `date` |
| `VT_CY` | currency, scaled by 10,000 | `cyVal` |
| `VT_DISPATCH` | `IDispatch*` | `pdispVal` |
| `VT_UNKNOWN` | `IUnknown*` | `punkVal` |
| `VT_ERROR` | `SCODE`; `DISP_E_PARAMNOTFOUND` = "argument omitted" | `scode` |
| `VT_ARRAY \| VT_x` | `SAFEARRAY*` of x | `parray` |
| `VT_BYREF \| VT_x` | pointer to x (for `[in,out]`) | `pl`, `pbstrVal`, … |
| `VT_VARIANT \| VT_BYREF` | `VARIANT*` — very common for by-ref script args | `pvarVal` |

Two gotchas that cause real bugs:

- **`VARIANT_TRUE` is `-1`, not `1`.** Testing `boolVal == 1` fails. Test `!= VARIANT_FALSE`.
- **`VT_BYREF|VT_VARIANT` is everywhere** in script calls. Always resolve it before reading:

```cpp
VARIANT* pv = &arg;
while (pv->vt == (VT_BYREF | VT_VARIANT)) pv = pv->pvarVal;   // unwrap
```

Conversion helper:

```cpp
VARIANT dst; VariantInit(&dst);
HRESULT hr = VariantChangeType(&dst, &src, 0, VT_I4);   // coerce to LONG
// ...
VariantClear(&dst);
```

Use `CComVariant`, which is a `VARIANT` with a constructor, destructor (`VariantClear`), and `ChangeType`.

### `SAFEARRAY`

A self-describing array: element type, dimensions, bounds, and (for `VT_BSTR`/`VT_VARIANT`/`VT_DISPATCH`) proper element cleanup.

```cpp
#include <atlsafe.h>

CComSafeArray<LONG> sa(5);              // 5 elements, lower bound 0
for (LONG i = 0; i < 5; ++i) sa[i] = i * i;

CComVariant v;
v.vt = VT_ARRAY | VT_I4;
v.parray = sa.Detach();                 // v now owns it; VariantClear will destroy it
```

Raw form, for when you must:

```cpp
SAFEARRAYBOUND bound = { 5, 0 };        // cElements, lLbound
SAFEARRAY* psa = SafeArrayCreate(VT_I4, 1, &bound);
LONG* pData = nullptr;
SafeArrayAccessData(psa, (void**)&pData);   // lock
for (LONG i = 0; i < 5; ++i) pData[i] = i * i;
SafeArrayUnaccessData(psa);                 // unlock — MUST pair
// ...
SafeArrayDestroy(psa);
```

> **Beware the lower bound.** VB defaults to 1, C++ to 0. Always call `SafeArrayGetLBound`/`GetUBound` rather than assuming.

---

## 5.6 Rich errors: `IErrorInfo`

An `HRESULT` is a number. Scripts want a message. `IErrorInfo` carries a description, source, and help context back to the caller — across process boundaries.

### Server side

```cpp
#include <atlbase.h>
#include <atlcom.h>

class ATL_NO_VTABLE Calculator :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<Calculator, &CLSID_Calculator>,
    public ISupportErrorInfoImpl<&IID_ICalculator>,   // "yes, I report rich errors"
    public IDispatchImpl<ICalculator, &IID_ICalculator, &LIBID_TrainingCalcLib>
{
public:
    STDMETHOD(Divide)(LONG a, LONG b, LONG* result)
    {
        if (!result) return E_POINTER;
        *result = 0;
        if (b == 0)
            return Error(L"Cannot divide by zero.", IID_ICalculator, E_INVALIDARG);
        *result = a / b;
        return S_OK;
    }
};
```

`CComCoClass::Error` wraps `ICreateErrorInfo` + `SetErrorInfo`. The raw form:

```cpp
HRESULT SetRichError(REFIID iid, PCWSTR desc, PCWSTR source, HRESULT hr)
{
    ICreateErrorInfo* pCEI = nullptr;
    if (SUCCEEDED(CreateErrorInfo(&pCEI)))
    {
        pCEI->SetGUID(iid);
        pCEI->SetDescription(const_cast<LPOLESTR>(desc));
        pCEI->SetSource(const_cast<LPOLESTR>(source));

        IErrorInfo* pEI = nullptr;
        if (SUCCEEDED(pCEI->QueryInterface(IID_IErrorInfo, (void**)&pEI)))
        {
            SetErrorInfo(0, pEI);        // attaches to the current LOGICAL THREAD
            pEI->Release();
        }
        pCEI->Release();
    }
    return hr;
}
```

`ISupportErrorInfo::InterfaceSupportsErrorInfo(riid)` must return `S_OK` for the interface, or clients won't even look.

### Client side

```cpp
HRESULT hr = pCalc->Divide(10, 0, &r);
if (FAILED(hr))
{
    IErrorInfo* pEI = nullptr;
    if (GetErrorInfo(0, &pEI) == S_OK && pEI)       // note: S_FALSE means "no info"
    {
        CComBSTR desc, source;
        pEI->GetDescription(&desc);
        pEI->GetSource(&source);
        wprintf(L"0x%08X: %s (%s)\n", hr, (BSTR)desc, (BSTR)source);
        pEI->Release();
    }
}
```

Two important details:

- `GetErrorInfo` **returns `S_FALSE`** (not a failure) when there's no error info. Test `== S_OK`.
- `GetErrorInfo` **clears** the stored info. Call it once, and call it immediately — any intervening COM call can overwrite it.
- In .NET, this machinery is what populates `COMException.Message`. In VBScript it's `Err.Description`. In PowerShell it surfaces in the exception message.

### `EXCEPINFO`

For late-bound calls, `Invoke`'s `pExcepInfo` carries the same information. If you use `ITypeInfo::Invoke`, it populates `EXCEPINFO` from the `IErrorInfo` you set — so implementing `IErrorInfo` correctly gets you both paths.

---

## 5.7 Connection points — COM events

COM's callback/event mechanism. The **source** (your object) notifies **sinks** (client objects).

### The shape

```
   Client                                 Server (source object)
 ┌────────────────┐                     ┌──────────────────────────────┐
 │  Sink object   │                     │  IConnectionPointContainer   │
 │ (_ICalcEvents) │◄────────────────────┤    └─ IConnectionPoint       │
 └────────────────┘   source calls it   │         └─ sink list         │
        │                               └──────────────────────────────┘
        └── Advise() registers it ─────────────────►
```

### IDL

```idl
[uuid(A1B2C3D4-0003-4000-9000-000000000003)]
dispinterface _ICalculatorEvents        // 'dispinterface' = late-bound only
{
    properties:
    methods:
        [id(1)] void OnCalculated([in] LONG result);
        [id(2)] void OnError([in] BSTR message);
};

[uuid(A1B2C3D4-1111-4000-9000-000000000001)]
coclass Calculator
{
    [default] interface ICalculator;
    [default, source] dispinterface _ICalculatorEvents;   // <- 'source' marks the event set
};
```

`[default, source]` is what tells VB/VBA/C# "this is the event interface" so `WithEvents` / `+=` work automatically.

### Server: `IConnectionPointContainer` / `IConnectionPoint`

```cpp
struct IConnectionPointContainer : IUnknown
{
    HRESULT EnumConnectionPoints(IEnumConnectionPoints** ppEnum);
    HRESULT FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP);
};

struct IConnectionPoint : IUnknown
{
    HRESULT GetConnectionInterface(IID* pIID);
    HRESULT GetConnectionPointContainer(IConnectionPointContainer** ppCPC);
    HRESULT Advise(IUnknown* pUnkSink, DWORD* pdwCookie);      // subscribe
    HRESULT Unadvise(DWORD dwCookie);                          // unsubscribe
    HRESULT EnumConnections(IEnumConnections** ppEnum);
};
```

ATL makes this nearly free:

```cpp
class ATL_NO_VTABLE Calculator :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<Calculator, &CLSID_Calculator>,
    public IConnectionPointContainerImpl<Calculator>,
    public CProxy_ICalculatorEvents<Calculator>,     // generated by the wizard
    public IDispatchImpl<ICalculator, &IID_ICalculator, &LIBID_TrainingCalcLib>
{
public:
    BEGIN_COM_MAP(Calculator)
        COM_INTERFACE_ENTRY(ICalculator)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY(IConnectionPointContainer)
    END_COM_MAP()

    BEGIN_CONNECTION_POINT_MAP(Calculator)
        CONNECTION_POINT_ENTRY(DIID__ICalculatorEvents)
    END_CONNECTION_POINT_MAP()

    STDMETHOD(Add)(LONG a, LONG b, LONG* r)
    {
        *r = a + b;
        Fire_OnCalculated(*r);      // generated helper: calls every subscribed sink
        return S_OK;
    }
};
```

### Client: implementing a sink

```cpp
class CalcSink : public IDispatch
{
    LONG m_cRef = 1;
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        // MUST answer to the event dispinterface IID, not just IDispatch.
        if (riid == IID_IUnknown || riid == IID_IDispatch || riid == DIID__ICalculatorEvents)
            *ppv = static_cast<IDispatch*>(this);
        else return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef()  override { return InterlockedIncrement(&m_cRef); }
    ULONG STDMETHODCALLTYPE Release() override
    { ULONG n = InterlockedDecrement(&m_cRef); if (!n) delete this; return n; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* p) override { *p = 0; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override
    { return E_NOTIMPL; }

    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispid, REFIID, LCID, WORD,
                                     DISPPARAMS* pDP, VARIANT*, EXCEPINFO*, UINT*) override
    {
        switch (dispid)
        {
        case 1:   // OnCalculated(LONG result)
            if (pDP && pDP->cArgs == 1)
                wprintf(L"[event] result = %ld\n", pDP->rgvarg[0].lVal);
            return S_OK;
        case 2:   // OnError(BSTR message)
            if (pDP && pDP->cArgs == 1)
                wprintf(L"[event] error: %s\n", pDP->rgvarg[0].bstrVal);
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }
};
```

Subscribe and — critically — unsubscribe:

```cpp
CComPtr<IConnectionPointContainer> spCPC;
spCalc->QueryInterface(IID_IConnectionPointContainer, (void**)&spCPC);

CComPtr<IConnectionPoint> spCP;
spCPC->FindConnectionPoint(DIID__ICalculatorEvents, &spCP);

CalcSink* pSink = new CalcSink();
DWORD cookie = 0;
spCP->Advise(pSink, &cookie);        // source now holds a reference to pSink
pSink->Release();                    // we drop ours; the source keeps it alive

long r = 0;
spCalc->Add(2, 3, &r);               // fires OnCalculated

spCP->Unadvise(cookie);              // <<<<<< MANDATORY
```

### The classic leak

`Advise` makes the source hold a **strong** reference to the sink. The client typically holds a strong reference to the source. That is a **cycle**:

```
   Client ──strong──► Source ──strong──► Sink ──(often)──► Client
```

Forget `Unadvise` and nothing is ever destroyed. This is the single most common COM leak in the wild.

**Symptoms:** memory grows with every dialog opened / document loaded / connection made; a server process never exits; a DLL never unloads.

**RAII fix:**

```cpp
class ConnectionCookie
{
    CComPtr<IConnectionPoint> m_cp;
    DWORD m_cookie = 0;
public:
    HRESULT Advise(IUnknown* pSource, REFIID iid, IUnknown* pSink)
    {
        CComPtr<IConnectionPointContainer> cpc;
        RETURN_IF_FAILED(pSource->QueryInterface(IID_PPV_ARGS(&cpc)));
        RETURN_IF_FAILED(cpc->FindConnectionPoint(iid, &m_cp));
        return m_cp->Advise(pSink, &m_cookie);
    }
    ~ConnectionCookie() { if (m_cp && m_cookie) m_cp->Unadvise(m_cookie); }
};
```

ATL provides `CComPtr`-based `AtlAdvise`/`AtlUnadvise`; WIL provides `wil::com_ptr` helpers. Use them.

---

## 5.8 Enumerators and `For Each`

The COM enumerator pattern:

```cpp
struct IEnumVARIANT : IUnknown
{
    HRESULT Next(ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched);
    HRESULT Skip(ULONG celt);
    HRESULT Reset();
    HRESULT Clone(IEnumVARIANT** ppEnum);
};
```

`Next` returns `S_OK` if it filled all `celt` slots, `S_FALSE` if it filled fewer (including zero). **Both are success.**

```cpp
CComPtr<IEnumVARIANT> spEnum;
// ... obtain it ...

CComVariant v;
ULONG fetched = 0;
while (spEnum->Next(1, &v, &fetched) == S_OK && fetched == 1)
{
    // use v
    v.Clear();                 // you own each fetched element
}
```

You owe cleanup for exactly `fetched` elements — `VariantClear` for `VARIANT`s, `Release` for `IUnknown*`s from `IEnumUnknown`.

### Making `For Each` work

A collection object exposes `DISPID_NEWENUM` (`-4`) returning an `IEnumVARIANT`:

```idl
[id(DISPID_NEWENUM), propget, restricted]
HRESULT _NewEnum([out, retval] IUnknown** ppEnum);
```

`[restricted]` hides it from object browsers; the name `_NewEnum` is conventional. VB's `For Each`, C#'s `foreach` over a COM collection, and PowerShell's pipeline all call DISPID -4.

Standard collection shape (follow it — tools expect it):

```idl
[id(DISPID_VALUE), propget] HRESULT Item([in] VARIANT index, [out, retval] VARIANT* value);
[id(1), propget]            HRESULT Count([out, retval] LONG* count);
[id(DISPID_NEWENUM), propget, restricted] HRESULT _NewEnum([out, retval] IUnknown** ppEnum);
```

ATL's `CComEnumOnSTL` / `IEnumOnSTLImpl` implement `IEnumVARIANT` over an STL container in a couple of lines.

---

## 5.9 LAB 5.1 — A dual interface driven from four languages

> **Requirements**
> - **Tools:** Visual Studio C++ (for `#import`); **Windows PowerShell 5.1 *and* PowerShell 7** — run the lab in both, their COM behaviour differs; `cscript.exe` for the VBScript client; the **.NET SDK** for the C# client.
> - **VBScript:** on Windows 11 24H2 and later VBScript is an **optional feature on demand**, not installed by default. If `cscript test.vbs` fails, add it under *Settings → System → Optional features → VBSCRIPT*. It is deprecated — you learn it to support the customers still running it, not to write new code.
> - **Elevation:** required, to register the server **and its type library** (`regsvr32` on a server with an embedded TLB, or `RegisterTypeLib`). Late binding by ProgID needs the CLSID; `#import` and early-bound C# need the TLB.
> - **Bitness:** register x64 and use the 64-bit hosts — `%SystemRoot%\System32\cscript.exe` is 64-bit, `%SystemRoot%\SysWOW64\cscript.exe` is 32-bit. Picking the wrong one reproduces Lab 2.2's error, which is a useful accident.
> - **Depends on:** a dual-interface `Calculator` with a registered type library — easiest via the ATL server in Lab 6.1, or the Module 4 IDL plus a hand-written `IDispatch`.
> - **Time:** ~3 h.

Build a `Calculator` with a dual interface, then call it from C++ (early and late bound), PowerShell, VBScript, and C#.

### C++ early bound (`#import`)

```cpp
#import "Calculator.tlb" no_namespace named_guids
// generates Calculator.tlh / .tli with smart pointers and wrapper methods

int main()
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    {
        ICalculatorPtr calc;                     // _com_ptr_t from #import
        calc.CreateInstance(CLSID_Calculator);
        long r = calc->Add(2, 3);                // [retval] becomes the return value!
        wprintf(L"Add -> %ld\n", r);
        wprintf(L"Describe -> %s\n", (LPCWSTR)calc->Describe());
    }
    CoUninitialize();
}
```

Note how `#import` turns `HRESULT Add([in] LONG, [in] LONG, [out,retval] LONG*)` into `long Add(long, long)` that throws `_com_error` on failure. Look at the generated `.tlh` to see exactly how.

### C++ late bound (raw `IDispatch`)

Write this by hand once. It's tedious, and that's the lesson.

```cpp
HRESULT CallAddLateBound(IDispatch* pDisp, long a, long b, long* pResult)
{
    OLECHAR* name = const_cast<OLECHAR*>(L"Add");
    DISPID dispid = 0;
    RETURN_IF_FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1,
                                          LOCALE_USER_DEFAULT, &dispid));

    VARIANTARG args[2];
    VariantInit(&args[0]); args[0].vt = VT_I4; args[0].lVal = b;   // REVERSE ORDER
    VariantInit(&args[1]); args[1].vt = VT_I4; args[1].lVal = a;

    DISPPARAMS dp = { args, nullptr, 2, 0 };
    CComVariant result;
    EXCEPINFO excep = {};
    UINT argErr = 0;

    HRESULT hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD, &dp, &result, &excep, &argErr);
    if (FAILED(hr))
    {
        if (hr == DISP_E_EXCEPTION)
            wprintf(L"server error: %s\n", excep.bstrDescription ? excep.bstrDescription : L"");
        else if (hr == DISP_E_TYPEMISMATCH)
            wprintf(L"bad type for argument %u\n", argErr);
        return hr;
    }
    *pResult = result.lVal;
    return S_OK;
}
```

### PowerShell

```powershell
$calc = New-Object -ComObject Training.Calculator.1
$calc.Add(2, 3)
$calc.Precision = 4          # property put
$calc.Precision              # property get
"$calc"                      # invokes DISPID_VALUE (Describe)

try { $calc.Divide(10, 0) } catch { $_.Exception.Message }   # IErrorInfo surfaces here

[Runtime.InteropServices.Marshal]::ReleaseComObject($calc) | Out-Null
```

### VBScript

```vbscript
Set calc = CreateObject("Training.Calculator.1")
WScript.Echo calc.Add(2, 3)
calc.Precision = 4
WScript.Echo calc                        ' default member

On Error Resume Next
calc.Divide 10, 0
If Err.Number <> 0 Then WScript.Echo "Error: " & Err.Description
```

### C#

```csharp
// Late bound - no reference needed
Type t = Type.GetTypeFromProgID("Training.Calculator.1");
dynamic calc = Activator.CreateInstance(t);
Console.WriteLine(calc.Add(2, 3));

// Early bound - add a COM reference to the type library
var calc2 = new TrainingCalcLib.Calculator();
Console.WriteLine(calc2.Add(2, 3));
```

### What to compare

Put a breakpoint (or `OutputDebugString`) in **both** `Invoke` and `Add`. For each client, record which path was taken:

| Client | `GetIDsOfNames`? | `Invoke`? | Direct vtable `Add`? |
|---|---|---|---|
| C++ `#import` | | | |
| C++ raw `IDispatch` | | | |
| PowerShell | | | |
| VBScript | | | |
| C# `dynamic` | | | |
| C# early bound | | | |

Then time 100,000 calls each. Early-bound C++ vs VBScript typically differs by **two to three orders of magnitude**. That measurement is the answer to "why is the script version so slow?"

---

## 5.10 LAB 5.2 — Events and the `Unadvise` leak

> **Requirements**
> - **Tools:** Visual Studio C++ with **ATL** (connection-point implementation); Windows PowerShell for `Register-ObjectEvent`.
> - **Elevation:** required, to register the server and TLB.
> - **Bitness:** x64, matching the PowerShell host you use.
> - **Depends on:** the Lab 5.1 server, plus the ref-count tracing from Module 1 — without the trace the leak is invisible, which is the lesson.
> - **Time:** ~2 h.

1. Add `_ICalculatorEvents` with `OnCalculated`, implement the connection point, and wire up the C++ sink from §5.7. Confirm the callback fires.
2. Add the Module 1 ref-count tracing to both the source and the sink.
3. **Remove the `Unadvise`.** Run. Observe in the trace: neither the sink nor the source is ever destroyed. Note that nothing *fails* — the program runs correctly and just leaks.
4. Restore `Unadvise` via the `ConnectionCookie` RAII wrapper.
5. Subscribe from **PowerShell** and confirm the same mechanism serves scripts:

```powershell
$calc = New-Object -ComObject Training.Calculator.1
Register-ObjectEvent -InputObject $calc -EventName OnCalculated -Action {
    Write-Host "event: $($EventArgs)"
}
$calc.Add(2, 3)
Get-EventSubscriber | Unregister-Event      # the PowerShell equivalent of Unadvise
```

6. **Support drill:** deliberately fire an event from the source while the sink's `Invoke` calls back into the source. On an STA, this is reentrancy (Module 3). Observe what happens, and note the fix (queue the notification instead of firing synchronously).

---

## 5.11 Automation error codes

| HRESULT | Symbol | Meaning |
|---|---|---|
| `0x80020003` | `DISP_E_MEMBERNOTFOUND` | The DISPID isn't a member; often a typo, or the object doesn't implement what the script expects |
| `0x80020006` | `DISP_E_UNKNOWNNAME` | `GetIDsOfNames` couldn't resolve the name |
| `0x80020005` | `DISP_E_TYPEMISMATCH` | An argument couldn't be coerced; `puArgErr` says which |
| `0x8002000E` | `DISP_E_BADPARAMCOUNT` | Wrong number of arguments |
| `0x8002000F` | `DISP_E_PARAMNOTOPTIONAL` | A required argument was omitted |
| `0x80020004` | `DISP_E_PARAMNOTFOUND` | A named argument doesn't exist |
| `0x80020009` | `DISP_E_EXCEPTION` | The **method itself failed**; read `EXCEPINFO` for the real error |
| `0x80020008` | `DISP_E_BADVARTYPE` | Invalid `VARIANT` type |
| `0x80028018` | `TYPE_E_INVDATAREAD` | Corrupt / wrong-version type library |
| `0x80029C4A` | `TYPE_E_CANTLOADLIBRARY` | Type library missing, unregistered, or **wrong bitness** |
| `0x800288C5` | `TYPE_E_LIBNOTREGISTERED` | `HKCR\TypeLib\{LIBID}` missing |

> **`DISP_E_EXCEPTION` is not the error.** It means "the call reached the method and the method failed." Always dig into `EXCEPINFO.scode`/`wCode` and `bstrDescription`. Support tickets that report `0x80020009` and stop there are incomplete.

### Support scenario: Office automation from a service

A recurring ticket. Symptoms: `CO_E_SERVER_EXEC_FAILURE` (`0x80080005`), `E_ACCESSDENIED`, or hangs, when Office automation runs under a service or IIS app pool.

**The answer is that it is not supported** — Microsoft explicitly does not support server-side Office automation. Reasons: Office assumes an interactive desktop and user profile; session 0 isolation blocks UI; modal dialogs (a "do you want to save?" prompt) block forever with no one to click them; licensing; and Office isn't reentrant/thread-safe for concurrent use.

Recognize it fast: check whether the process is a service, then look for `Interactive User` in the AppID's `RunAs`, and for `dllhost`/`Excel.exe` instances piling up in session 0. Recommend the Open XML SDK or a server-supported library instead.

---

## 5.12 Checkpoint

1. Trace what `obj.Add(2, 3)` in VBScript does at the COM level, call by call.
2. A colleague says "our object uses `IDispatch` instead of `IUnknown`." Correct them precisely, and state what each interface is responsible for.
3. Why are `DISPPARAMS::rgvarg` arguments in reverse order, and what's the second surprise in `DISPPARAMS` for property puts?
4. What is a dual interface, and what does it give you that a pure `dispinterface` doesn't — and vice versa?
5. `VARIANT_TRUE` is `-1`. Why, and what bug does assuming `1` cause?
6. A component `Advise`s a sink and the client forgets `Unadvise`. Draw the reference graph and state what leaks.
7. `Invoke` returns `0x80020009`. What have you actually learned, and what's your next step?
8. `IEnumVARIANT::Next(10, rgVar, &fetched)` returns `S_FALSE` with `fetched == 3`. Is that an error? What cleanup do you owe?
9. A customer's VBScript works on their old server and fails on the new one with `TYPE_E_CANTLOADLIBRARY`. Name three things to check.

<details>
<summary>Answers</summary>

1. `CreateObject` → `CLSIDFromProgID` → `CoCreateInstance(clsid, …, IID_IDispatch, …)`. Then for the call: `GetIDsOfNames(IID_NULL, ["Add"], 1, lcid, &dispid)`; build `DISPPARAMS` with two `VT_I4` args in reverse order; `Invoke(dispid, IID_NULL, lcid, DISPATCH_METHOD|DISPATCH_PROPERTYGET, &dp, &result, &excep, &argErr)`; read `result`; `VariantClear` it.

2. They're not alternatives — `IDispatch` **derives from** `IUnknown`, so an object exposing `IDispatch` necessarily exposes `IUnknown` too (slots 0–2 of the `IDispatch` vtable *are* `IUnknown`). `IUnknown` is mandatory for every COM object and handles identity (`QueryInterface`), navigation, and lifetime (`AddRef`/`Release`). `IDispatch` is optional and adds runtime discovery: name→DISPID resolution and `VARIANT`-based invocation, so clients with no compile-time knowledge of the vtable can call the object. Even a pure VBScript client is having `AddRef`/`Release` called on its behalf.

3. Historical, from VB's calling convention and the way arguments were pushed on the stack — `rgvarg[0]` is the last declared parameter. The second surprise: for `DISPATCH_PROPERTYPUT`, the new value must be passed as a **named argument** with DISPID `DISPID_PROPERTYPUT` (-3) in `rgdispidNamedArgs`, and `cNamedArgs` must be 1.

4. A dual interface derives from `IDispatch` and also exposes its methods in the vtable after `IDispatch`'s four slots. Compiled clients get fast vtable calls; scripts get `Invoke`. A pure `dispinterface` supports **only** late binding (no vtable access) but has no restriction on being derivable and is what event sets use. Dual costs you: Automation-compatible types only, and it must derive directly from `IDispatch`.

5. `VARIANT_TRUE` is `0xFFFF` (-1 as a `short`) because VB's `True` is all-bits-set, which makes bitwise and logical operators coincide. Assuming `1` means `if (v.boolVal == 1)` is false for a true value — a silent logic inversion. Test `!= VARIANT_FALSE`.

6. Client → Source (strong), Source → Sink (strong, from `Advise`), Sink → Client or Source (strong, typically). Nothing reaches zero: the sink leaks, the source leaks, and by extension the server DLL never unloads / server process never exits.

7. Only that the call **reached the method and the method reported failure** — `DISP_E_EXCEPTION` is a wrapper. Next step: read `EXCEPINFO` — `scode` (or `wCode`), `bstrDescription`, `bstrSource`. That's the real error.

8. Not an error — `S_FALSE` means "fewer than requested," which is the normal way an enumerator signals the end. You owe `VariantClear` on `rgVar[0..2]` only; elements 3..9 were never written.

9. (a) Is the type library registered on the new box (`HKCR\TypeLib\{LIBID}`)? (b) **Bitness** — is the script host 32- or 64-bit, and is the typelib registered in the matching view (`win32` vs `win64` subkey, `Wow6432Node`)? (c) Does the path in the TypeLib key point at a file that actually exists — and is it the right version subkey (`1.0` vs `2.0`)?

</details>

---

## 5.13 Rules to carry forward

1. Assign DISPIDs explicitly with `[id(n)]`; never let them be auto-generated in shipping interfaces.
2. Implement `IDispatch` by delegating to `ITypeInfo::Invoke` (or use `IDispatchImpl`). Never hand-roll it.
3. `VariantInit` before use, `VariantClear` after — or use `CComVariant`.
4. `VARIANT_TRUE` is `-1`. Unwrap `VT_BYREF|VT_VARIANT`.
5. `GetErrorInfo` returns `S_FALSE` when there's nothing; it also *clears* the info. Read it immediately.
6. Every `Advise` needs an `Unadvise`. Wrap it in RAII.
7. `Next` returning `S_FALSE` is success; clean up exactly `fetched` elements.
8. `DISP_E_EXCEPTION` is a wrapper — always open `EXCEPINFO`.
9. Dual interfaces get typelib marshaling for free; keep to Automation types.
10. Server-side Office automation is unsupported. Recognize the signature and redirect the customer.

---

**Next: [Module 6 — Frameworks: ATL, WRL, WIL, and .NET interop](06-frameworks-and-interop.md)**
