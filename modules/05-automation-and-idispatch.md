# Module 5 — Automation, `IDispatch`, and scripting

Everything so far assumed a compiled client that knows the vtable at build time. But VBScript, JScript, PowerShell, VBA, and old VB have no vtables and no headers. They discover methods **at runtime, by name**. This module explains how, and covers the family of technologies built on top.

**What this module covers**

How a language with no headers and no vtables calls a COM object at all: `IDispatch`, binding by name at runtime, and dual interfaces that serve both kinds of caller from one object. It covers the Automation type system — `VARIANT`, `BSTR`, `SAFEARRAY` — type libraries and what reads them, rich errors that reach a script as a message rather than a number, connection points and the leak built into them, and how `For Each` actually works.

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
Set obj = CreateObject("TrainingCalc.Calculator.1")
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

This is the fourth parameter of `Invoke`, and the **caller** supplies it — the script engine, PowerShell, or the CLR.

Its job is to say what the caller is trying to *do*. `dispIdMember` names **which member**; `wFlags` says **read it, write it, or call it**. Both are needed, because `GetIDsOfNames` maps a name to a DISPID and one name can be used several ways: `obj.Value`, `obj.Value = 5` and `obj.Value(2)` all arrive with the *same* DISPID.

| Flag | The caller wrote | What your `Invoke` should do |
|---|---|---|
| `DISPATCH_METHOD` | `obj.Foo(1)` | Call the method |
| `DISPATCH_PROPERTYGET` | `x = obj.Foo` | Return the property's value |
| `DISPATCH_PROPERTYPUT` | `obj.Foo = 5` | Assign the value |
| `DISPATCH_PROPERTYPUTREF` | `Set obj.Foo = other` | Assign an *interface pointer*, not a copied value |

A member can be several of these at once, so the flags are a bitmask. `DISPATCH_METHOD | DISPATCH_PROPERTYGET` is common: some languages cannot tell a no-argument method call from a property read, so they ask for either and let the object decide.

If you implement `Invoke` by hand you must branch on this. If you delegate to `ITypeInfo::Invoke` — as §5.4 does — it reads the type library and does the branching for you, which is one of the better reasons not to write `Invoke` by hand.

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
2. **Property puts use a named argument.**

   > A **named argument** is one the caller identifies by name rather than by position — `cells.Item(Row:=1, Column:=2)` in VBA, instead of `cells.Item(1, 2)`. Since `Invoke` deals in DISPIDs rather than names, each named argument carries a DISPID: `rgdispidNamedArgs[i]` gives the DISPID for `rgvarg[i]`, and named arguments occupy the **first** `cNamedArgs` slots of `rgvarg`. Everything after them is positional.

   A property assignment reuses that mechanism. The value being assigned is passed as a named argument with the reserved DISPID `DISPID_PROPERTYPUT` (`-3`), which is how the object tells "the new value" apart from any ordinary arguments the property may also take:

```cpp
VARIANT v; v.vt = VT_I4; v.lVal = 42;
DISPID putid = DISPID_PROPERTYPUT;
DISPPARAMS dp = { &v, &putid, 1, 1 };      // 1 argument, and it is a named one
pDisp->Invoke(dispidPrecision, IID_NULL, lcid, DISPATCH_PROPERTYPUT,
              &dp, nullptr, &excep, nullptr);
```

Forgetting the named argument — passing `{ &v, nullptr, 1, 0 }` — gives `DISP_E_PARAMNOTOPTIONAL` or `DISP_E_BADPARAMCOUNT`.

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
  ICalculator       slots 7+    Add, Subtract, get_Precision, put_Precision, SumTo, Describe
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
    [id(4)] HRESULT SumTo([in] LONG n, [out, retval] LONG* total);   // slow on purpose (§5.7)
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

You almost never hand-write `GetIDsOfNames`/`Invoke`.

Consider what a hand-written `Invoke` would have to do. Given DISPID 1, `DISPATCH_METHOD`, and a `DISPPARAMS` holding two `VARIANT`s, it must: check the DISPID is one you recognise, check `wFlags` is a use that member supports, check the argument count, coerce each `VARIANT` to the type the method actually takes (the script may have passed a `VT_BSTR` where you want a `LONG`), remember the arguments are in reverse order, call the method, wrap the result back into a `VARIANT`, and fill in `puArgErr` if any of it failed. Then repeat that for every member, forever, keeping it in step with the interface.

Every fact needed to do that is already recorded in your **type library**: the member names, their DISPIDs, which are methods and which are properties, the parameter types, and the vtable slot each one occupies. MIDL put it there when it compiled your IDL.

So instead of writing that code, hand the call to the type library and let it do the work. `ITypeInfo::Invoke` looks up the DISPID, coerces the arguments to the declared types, and calls the correct vtable slot on your object. Two calls do the whole job:

| You implement | You call | What it does |
|---|---|---|
| `GetIDsOfNames` | `DispGetIDsOfNames` | Looks each name up in the type library and returns its DISPID |
| `Invoke` | `ITypeInfo::Invoke` | Coerces the arguments and dispatches to the real method |

The rest of the code below is just obtaining the `ITypeInfo` once and caching it:

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

**The problem it solves.** A scripting language has no types at compile time. When VBScript runs `x = calc.Add(2, 3)`, the engine must hand COM "a value" without knowing whether it is a number, a string, a date, or an object — and receive one back the same way. C++ has no such type, so Automation defines one.

**What it physically is.** A fixed-size struct with two parts: a tag called `vt` saying *which* type is stored, and a union holding the value. You set both, and you read `vt` before ever touching the union.

```cpp
VARIANT v;
VariantInit(&v);          // vt = VT_EMPTY

v.vt   = VT_I4;           // 1. declare what it holds
v.lVal = 42;              // 2. store it in the matching union field

if (v.vt == VT_I4)        // reading: ALWAYS check the tag first
    printf("%ld\n", v.lVal);
```

**A `VARIANT` can own what it holds.** If it contains a `BSTR`, a `SAFEARRAY`, or an interface pointer, that resource belongs to the variant — so it has to be freed. `VariantClear` reads `vt`, does whichever of `SysFreeString` / `SafeArrayDestroy` / `Release` applies, and resets the variant to `VT_EMPTY`:

```cpp
VariantClear(&v);         // the only correct way to dispose of a VARIANT
```

Forgetting it leaks; guessing the wrong free function corrupts the heap (§4.3). In practice use **`CComVariant`** (ATL) or **`_variant_t`**, which call `VariantInit` and `VariantClear` for you.

### The `vt` values

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

**How it relates to `VARIANT`.** They are two separate types that are designed to compose:

| | Holds |
|---|---|
| `VARIANT` | **one** value, of **any** type |
| `SAFEARRAY` | **many** values, all of **one** type |

They *have* to compose, because `IDispatch::Invoke` carries every argument and every return value as a `VARIANT` (§5.2). An array can therefore only reach a script by riding inside one. That is what `VT_ARRAY` means: set `vt` to `VT_ARRAY | <element type>` and put the `SAFEARRAY*` in the `parray` field.

```cpp
#include <atlsafe.h>

CComSafeArray<LONG> sa(5);              // 5 elements, lower bound 0
for (LONG i = 0; i < 5; ++i) sa[i] = i * i;

CComVariant v;
v.vt = VT_ARRAY | VT_I4;                // "this variant holds an array of LONG"
v.parray = sa.Detach();                 // v now owns it; VariantClear will destroy it
```

**Ownership follows containment.** Once the `SAFEARRAY*` is inside the variant, `VariantClear` calls `SafeArrayDestroy` for you. That is why the code says `Detach()` — it hands ownership over. Assigning without detaching leaves two owners and a double free.

**And the nesting works the other way too.** A `SAFEARRAY` whose *elements* are `VARIANT`s — `VT_ARRAY | VT_VARIANT` — gives you many values each of any type. That is exactly what VBScript's `Array(1, "two", #2020-03-04#)` produces, and it is the shape you will meet most often coming from scripts.

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

Every call so far has gone one direction: the client calls the object. This section is about the other direction.

### 1. The problem

An object often needs to tell its client something **without being asked**: a long operation finished, a value changed, a document was saved, a device was plugged in.

It cannot simply call the client back, because **it does not know who its clients are.** `CoCreateInstance` hands interface pointers out; it never tells the object who received them.

```
   client A ──┐
   client B ──┼──►  Calculator        "Something just changed.
   client C ──┘                        Who do I tell?"
```

**Why not just add a callback method?** You could put `SetCallback(IProgress*)` on your interface and it would work. But it takes one subscriber, for one kind of notification, through a convention *you* invented. No scripting host, form designer or IDE can discover it, because there is nothing standard to look for — the same problem Module 0 says COM exists to remove.

### 2. The idea

The client builds a small COM object of its own and hands it to the server. The server now has something concrete to call.

That is the whole mechanism. Everything else is vocabulary and plumbing.

The running example from here on is a method that takes a while — `SumTo(n)`, which adds up every number from 1 to *n* — and reports how far it has got as it goes.

```
        the client calls the server - every call in Modules 1-4
   ┌──────────┐  ──────────────────────►  ┌──────────────┐
   │  CLIENT  │     SumTo(1000000, &t)    │    SERVER    │
   │          │                           │              │
   │  + SINK  │  ◄──────────────────────  │  = SOURCE    │
   └──────────┘        OnProgress(10)     └──────────────┘
   └──────────┘        OnProgress(20)     └──────────────┘
        the server calls the client back - these are the events
```

Note the shape of it: **one call in, many calls back, before the first one has returned.** No return value can do that. That is the first of the two reasons events exist; the second one appears in §5.

Both arrows are ordinary COM calls: same vtables, same `HRESULT`s, same apartment rules. **The only thing that makes the lower ones "events" is their direction.** So the roles swap depending on which call you mean:

| Call | Caller | Callee | Whose code runs |
|---|---|---|---|
| `SumTo(1000000, &t)` | client | server | the server's `Calculator::SumTo` |
| `OnProgress(10)` | server | client | the client's sink class |

### 3. The three words

| Term | Meaning |
|---|---|
| **Outgoing interface** | An interface the object **calls** instead of implementing. An event set is an outgoing interface. |
| **Source** | The object that raises the events — the server you wrote. |
| **Sink** | The object that receives them — **written by the client**. |

A sink is **a full COM object, not a function pointer.** It implements `IUnknown` plus the outgoing interface, so the source can call it exactly like any other COM interface. In VB, C# and VBA the tooling writes the sink class for you behind `WithEvents` and `+=`; in C++ you write it yourself.

### 4. What COM standardizes

A "connectable object" is one that supports four things — and these four are precisely what the documentation lists:

| The need | The API |
|---|---|
| Say which outgoing interfaces I support | `IConnectionPointContainer::EnumConnectionPoints` — one connection point per outgoing IID |
| Find the connection point for one event set | `IConnectionPointContainer::FindConnectionPoint(IID)` |
| Connect and disconnect sinks | `IConnectionPoint::Advise` / `Unadvise` |
| List the connections that exist | `IConnectionPoint::EnumConnections` — a connection point holds a **list** of sinks, not one pointer |

Because the type library also describes the outgoing interface, tools can generate the wiring: VB's `WithEvents`, C#'s `+=`, VBA's `Private Sub obj_Event()`. That is the payoff — an event handler costs one line and no plumbing.

### 5. When the event fires

Two questions get confused here, and they have different answers.

**Is the call synchronous?** Always. Raising an event means calling a method on the sink, and that is an ordinary COM call: the source blocks until your handler returns an `HRESULT`. Connection points offer no post-and-forget path.

**When does the source raise it?** That is the **server author's choice** — nothing in the standard fixes it. Most sources raise an event at the moment the thing happens, which puts it inside the method that caused it. That is what this section's implementation does, calling `Fire_OnProgress` from inside the loop:

```
   CLIENT  (implements the sink)      ║      SERVER  (is the source)
   ═══════════════════════════════════╬═══════════════════════════════════
                                      ║
   1. SumTo(1000000, &t) ─────────────╫───► SumTo begins looping
                                      ║        │
   3. OnProgress(10) runs here ◄──────╫────────┤  2. Fire_OnProgress(10)
      └── returns S_OK ───────────────╫───────►│     ...loop continues
                                      ║        │
   5. OnProgress(20) runs here ◄──────╫────────┤  4. Fire_OnProgress(20)
      └── returns S_OK ───────────────╫───────►│     ...and so on
                                      ║        │
   7. SumTo returns, t is set ◄───────╫────────┘  6. *total = sum; return S_OK
                                      ║
```

Follow the numbers: everything from step 2 onwards happens **inside** step 1. `SumTo` has not returned when your handler runs, and it is all one thread.

#### The other reason events exist: one caller, many observers

Section 2 gave the first reason — a progress report cannot be a return value, because there are many of them and the call has not finished. Here is the second: a connection point keeps a **set of connections**, so firing an event reaches every sink attached to it, not just whoever made the call.

```
   client A ── SumTo(...) ──►  Calculator  ── OnProgress ──► client A   (it asked)
                              (ONE instance)  │
                                    ├──────── OnProgress ──► client B   (it did not,
                                    │                                    and could not
                                    └──────── OnProgress ──► client C    otherwise know)
```

A progress bar in one window, a log file, and an audit component can all watch the same operation without any of them being the caller. B and C never invoked anything, so no return value is coming their way — polling would be the only alternative.

Two conditions on that picture, and both matter in practice:

- **It is one object, not one class.** A connection point belongs to a specific *instance*. B and C must be holding the **same** `Calculator` object — obtained from the Running Object Table, or handed to them by A. If each of them called `CoCreateInstance` they would each own a separate object with its own sink list, and no event would ever cross between them. This is exactly why Office automation examples fetch the *running* Excel rather than creating a new one.
- **The number of sinks may be capped.** Nothing obliges a component to accept many subscribers; some accept only one, and the second attempt to subscribe simply fails. The subscription call and its error codes come later in this section.

The same shape covers events that follow no call at all: `OnDeviceArrived`, `OnDocumentSaved` when the *user* pressed Ctrl+S, `OnCellChanged` when another client wrote to the sheet. Nobody invoked a method, so there is no return value to attach the news to.

Three consequences, each of which shows up in real tickets:

- **Your handler blocks the server.** `SumTo` is waiting for `OnProgress` to return before it resumes the loop. A handler that opens a dialog or waits on a lock stops the calculation dead. This holds no matter when the author chose to fire.
- **Every subscriber pays, on every event.** The source calls each sink in turn. Ten progress events and three slow handlers means thirty blocking calls inside one method — the classic "adding a listener made the app crawl" ticket.
- **This is reentrancy** (Module 3 §3.6). Your code runs while `SumTo` is unfinished, so calling back into that same object re-enters a method already in progress.

A server may instead queue the event and raise it later from another thread. That is a legitimate design — but the call then crosses an apartment boundary to reach your sink, and Module 3's rules apply in full. **When diagnosing someone else's component, do not assume either pattern: check which thread the handler runs on.**

### 6. The catch

A connection is a **reference cycle by construction**: the source holds a reference to the sink so it can call it, and the sink usually holds a reference to the source so it can use it. Neither can reach zero on its own.

That is what `Unadvise` is for, and forgetting it is the classic leak at the end of this section.

### 7. Putting the objects on a diagram

The four APIs above become three objects. The container answers "which event sets do you have?", each connection point owns the list of sinks for one event set, and the sink is the client's own object:

```
   Client                                 Server (source object)
 ┌────────────────┐                     ┌──────────────────────────────┐
 │  Sink object   │                     │  IConnectionPointContainer   │
 │ (_ICalcEvents) │◄────────────────────┤    └─ IConnectionPoint       │
 └────────────────┘   source calls it   │         └─ sink list         │
        │                               └──────────────────────────────┘
        └── Advise() registers it ─────────────────►
```

The rest of this section builds exactly that: the IDL that declares the outgoing interface, the server side, the client's sink, and what `Advise` really does.

### IDL

```idl
[uuid(A1B2C3D4-0003-4000-9000-000000000003)]
dispinterface _ICalculatorEvents        // 'dispinterface' = late-bound only
{
    properties:
    methods:
        [id(1)] void OnProgress([in] LONG percent);
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

    STDMETHOD(SumTo)(LONG n, LONG* total)
    {
        LONG sum = 0;
        for (LONG i = 1; i <= n; ++i)
        {
            sum += i;
            if (n >= 100 && i % (n / 10) == 0)
                Fire_OnProgress(i / (n / 100));   // generated helper: calls every sink
        }
        *total = sum;
        return S_OK;
    }
};
```

`Fire_OnProgress` is generated by ATL from the connection point map. It walks the sink list and calls each one in turn — which is why the loop pauses at every tenth of the work, for as long as the slowest handler takes.

### Client: implementing a sink

**Where is `OnProgress`?** You are about to read a sink class that has no method by that name, which looks wrong until you remember how the event set was declared:

```idl
dispinterface _ICalculatorEvents        // 'dispinterface' = late-bound only
```

A `dispinterface` has **no vtable**. There is nothing for the source to call directly, so `Fire_OnProgress` does not invoke a method called `OnProgress` — it calls **`IDispatch::Invoke` on the sink, passing DISPID 1**, which is the ID the IDL assigned to `OnProgress`. Firing an event is exactly the late binding from §5.2, running in the opposite direction.

That is why the sink below implements `IDispatch` and switches on the DISPID, rather than overriding named methods:

| In the IDL | On the wire | In the sink |
|---|---|---|
| `[id(1)] void OnProgress(LONG)` | `Invoke(1, ..., DISPPARAMS{percent})` | `case 1:` |
| `[id(2)] void OnError(BSTR)` | `Invoke(2, ..., DISPPARAMS{message})` | `case 2:` |

Had the event set been declared as a **vtable interface** instead of a `dispinterface`, the sink *would* implement `OnProgress` as a real method and the source would call it directly. Dispinterfaces are used for event sets because scripting hosts can only consume that form — it is what lets the same events reach VBScript and PowerShell.

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
        case 1:   // OnProgress(LONG percent)
            if (pDP && pDP->cArgs == 1)
                wprintf(L"[event] %ld%% done\n", pDP->rgvarg[0].lVal);
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

### Subscribing: who calls `Advise`, and what it does

**The client calls it** — and it is an ordinary client → server call, the last one that runs in that direction. It means: *"here is my sink object; add it to your list of subscribers."*

Getting there takes four calls, and only the third is `Advise` itself:

| Step | Call | What it asks for |
|---|---|---|
| 1 | `QueryInterface(IID_IConnectionPointContainer)` | "do you raise events at all?" |
| 2 | `FindConnectionPoint(DIID__ICalculatorEvents)` | "give me the connection point for **this particular** event set" — one object can offer several |
| 3 | `Advise(pSink, &cookie)` | "here is my sink object; subscribe it" |
| 4 | `Unadvise(cookie)` | later: "cancel that subscription" |

Inside `Advise`, the **source** — that is, the server object you called — does three things:

1. calls `QueryInterface` on your sink for the event interface — if the sink does not implement it, `Advise` fails with `CONNECT_E_CANNOTCONNECT`;
2. **`AddRef`s the sink** and stores the pointer in its subscriber list;
3. returns a **cookie** — a token identifying *this one* subscription, which is why `Unadvise` takes it and why many sinks can subscribe independently.

Note the reversal in step 1: for the whole course so far, *you* have been the one calling `QueryInterface` on someone else's object. Here the server calls it on **yours**, which is the first sign that the roles are about to swap.

Step 2 is the source half of the reference cycle. And from the moment `Advise` returns, the direction reverses for good: everything the source sends afterwards is a server → client call.

Subscribe and — critically — unsubscribe:

```cpp
// spCalc is the Calculator object, already created with CoCreateInstance.
// (HRESULTs are omitted here for room - check every one of them in real code.)

// --- Step 1: does this object raise events at all? ------------------------
CComPtr<IConnectionPointContainer> spCPC;
spCalc->QueryInterface(IID_IConnectionPointContainer, (void**)&spCPC);
// E_NOINTERFACE here would mean "this component offers no events".

// --- Step 2: get the connection point for THIS event set ------------------
CComPtr<IConnectionPoint> spCP;
spCPC->FindConnectionPoint(DIID__ICalculatorEvents, &spCP);
// DIID_ is the IID of a dispinterface. One object can expose several
// connection points, so you must say which event set you mean.

// --- Step 3: create our sink object and subscribe it ----------------------
CalcSink* pSink = new CalcSink();    // CalcSink starts at m_cRef = 1  -> count 1
DWORD cookie = 0;                    // receives the subscription's token

spCP->Advise(pSink, &cookie);        // the source AddRefs the sink    -> count 2

pSink->Release();                    // we drop OUR reference          -> count 1
                                     // The source still holds one, so the sink lives
                                     // exactly as long as the subscription does.
                                     // Only correct because Advise succeeded - if it
                                     // failed, this Release is the last one and the
                                     // sink dies while we still expect events.

// --- Events now travel server -> client -----------------------------------
long total = 0;
spCalc->SumTo(1000000, &total);      // inside SumTo the source calls Fire_OnProgress
                                     // ten times, each landing in CalcSink::Invoke
                                     // BEFORE SumTo returns

// --- Step 4: unsubscribe --------------------------------------------------
spCP->Unadvise(cookie);              // the source Releases the sink    -> count 0
                                     // -> CalcSink deletes itself.
                                     // MANDATORY. Omit it and nothing is ever
                                     // destroyed - see the next section.

// spCP and spCPC are CComPtr, so they Release themselves at end of scope.
```

### The classic leak

> **Strong and weak, recalled from Module 1 §1.8.** A **strong** reference is one you `AddRef`'d: it keeps the target alive and you owe it a `Release`. Every interface pointer in this course so far has been strong. A **weak** reference is a pointer you hold *without* `AddRef` — it does not keep anything alive, and it is your problem to be sure the target still exists before you use it.
>
> | | Strong | Weak |
> |---|---|---|
> | `AddRef`'d? | yes | no |
> | Keeps the target alive? | **yes** | no |
> | Can it be used safely at any time? | yes | only while something *else* holds a strong reference |
> | Participates in a cycle? | **yes** | no — which is exactly why weak references break cycles |

`Advise` makes the source hold a **strong** reference to the sink. The client typically holds a strong reference to the source. That is a **cycle** — but only once the third arrow exists:

```
   Client ──strong──► Source ──strong──► Sink ──(usually)──► Client or Source
```

That third arrow is worth being precise about, because it is what the leak actually depends on. With only the first two, forgetting `Unadvise` costs you nothing: the client releases the source, the source's count reaches zero, and as it is destroyed the connection point releases every sink still registered. Cleanup happens late, but it happens.

The sink almost always does hold that reference, though — it needs the source to read a property while handling an event, or to `Unadvise` itself later, or because the sink *is* the client object. Once it does, neither side can reach zero: forget `Unadvise` and nothing is ever destroyed. Lab 5.2 has you build both versions and watch the second one stop cleaning up.

**Why not just make one of them weak?** That is a legitimate cycle-breaker in general — Module 1 lists it alongside `IWeakReference` — but it cannot work here. The source may need to call the sink at any moment, possibly across a process boundary, so it must be certain the sink is still alive: that requires a strong reference. **Connection points therefore break the cycle by hand rather than by weakness**, and `Unadvise` is that break. Nothing else will do it for you.

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
> - **Tools:** Visual Studio C++ (for `#import`); **either** Windows PowerShell 5.1 **or** PowerShell 7 — late binding behaves the same in both, so one is enough; `cscript.exe` for the VBScript client; the **.NET SDK** for the C# client, built from Visual Studio rather than `dotnet build` (see the C# section for why).
> - **VBScript:** on Windows 11 24H2 and later VBScript is an **optional feature on demand**, not installed by default. If `cscript test.vbs` fails, add it under *Settings → System → Optional features → VBSCRIPT*. It is deprecated — you learn it to support the customers still running it, not to write new code.
> - **Elevation:** yes — **run Visual Studio as administrator**. An ATL DLL project has **Register Output** switched on by default, so every successful build runs `regsvr32` on the result, and that writes the CLSID, ProgID and type library to `HKLM\Software\Classes`. Without elevation the build itself succeeds and the registration step fails. (To work without admin, see the per-user option in the Stage 4 README.) Late binding by ProgID needs the CLSID; `#import` and early-bound C# need the TLB.
> - **Bitness:** register x64 and use the 64-bit hosts — `%SystemRoot%\System32\cscript.exe` is 64-bit, `%SystemRoot%\SysWOW64\cscript.exe` is 32-bit. Picking the wrong one reproduces Lab 2.2's error, which is a useful accident.
> - **Depends on:** a dual-interface `Calculator` with a registered type library.
> - **Starting point — do this first:** work through the whole of [`labs/stage-4-atl-server/README.md`](../labs/stage-4-atl-server/README.md), steps 1 to 6 (**15–20 min**). It produces exactly the server this lab drives, and **none of the clients below will compile or run until it is finished and registered.** Stop when `regsvr32` reports success.
> - **Time:** ~2 h for the five clients and the comparison, plus the 15–20 min for the starting point. The raw `IDispatch` client alone is about a third of it — which is the point of writing it once.

**Step 0 — build the server.** This lab has nothing of its own to run against: the component comes from [`labs/stage-4-atl-server/README.md`](../labs/stage-4-atl-server/README.md). Work through its steps 1 to 6 now, and come back when `regsvr32` has reported success.

Then confirm the server is really there — this is the same check every client below depends on:

```powershell
$calc = New-Object -ComObject TrainingCalc.Calculator.1
$calc.Add(2, 3)          # 5
```

If that fails with `0x80040154`, the component is not registered; if it succeeds but `$calc.Describe()` cannot be found, the **type library** is not registered. Fix that before writing any client code, or you will be debugging the wrong layer all afternoon.

One component, five callers, no changes to the component. That is the claim Automation makes, and this lab is where you check it.

You build a `Calculator` with a **dual** interface, then call it from C++ (early *and* late bound), PowerShell, VBScript, and C#. Two of those callers use the vtable; three go through `IDispatch` and have never seen your header.

What to take away is not the code — it is which capability each caller depends on. When a customer says "it works in C# but not in VBScript," this lab is how you already know where to look.

### Step 1 — create the client projects

The server is a DLL, so every client is a separate program of its own. Three of them need a project; two do not.

| Client | Create | Name it |
|---|---|---|
| C++ `#import` | **Console App (C++)** | `Client_Import` |
| C++ raw `IDispatch` | **Console App (C++)** | `Client_Dispatch` |
| PowerShell | a `.ps1` file — no project | `client.ps1` |
| VBScript | a `.vbs` file — no project | `client.vbs` |
| C# | **Console App (C#)** | `Client_CSharp` |

Add the three projects to the **same solution as `TrainingCalc`** (right-click the solution → **Add → New Project**). You then have one place to set breakpoints on both sides of the call, which is exactly what *What to compare* at the end of this lab asks for. Set each C++ client to **Debug | x64**: a 32-bit client cannot load the x64 server in-proc (Module 2 §2.3).

Nothing else is needed — no extra linker input, no include directories, no references.

> **Which C++ toolset?** Visual Studio 2026 asks you to choose between the **Latest** toolset (v14.51) and the **LTS** toolset (v14.50). **Take either — and it does not have to match the toolset the server was built with.**
>
> That is worth a pause rather than a shrug, because it is the whole point of COM. Client and server meet at a vtable of `HRESULT`-returning functions and nothing else: no C++ name mangling in common, no shared CRT, no shared allocator, no shared header. A toolset mismatch *cannot* break it — which is precisely why the PowerShell and VBScript clients below can call the same object without a compiler at all. Pass a `std::string` across that boundary instead and every one of those guarantees disappears at once (Module 0 §0.1).
>
> If you want to see it rather than take it on trust, build one C++ client on Latest and the other on LTS. Both work.

### C++ early bound (`#import`)

`#import` reads a type library and generates two headers next to your object files — `TrainingCalc.tlh` (declarations) and `TrainingCalc.tli` (inline wrapper bodies) — then includes them. You do not write a single interface declaration yourself.

Point it at the **DLL**, not at a `.tlb` file. ATL embeds the type library inside the DLL as a resource, which is how `DllRegisterServer` was able to register it in the first place, so there is no separate file to go hunting for:

```cpp
#include <windows.h>
#include <stdio.h>

// Use the path to the DLL you registered. Double the backslashes, or use forward slashes.
#import "C:\\Users\\you\\source\\repos\\TrainingCalc\\x64\\Debug\\TrainingCalc.dll" \
    no_namespace named_guids

int main()
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    {
        ICalculatorPtr calc;                        // _com_ptr_t, generated by #import
        HRESULT hr = calc.CreateInstance(CLSID_Calculator);   // named_guids gives you this name
        if (FAILED(hr))
        {
            wprintf(L"CreateInstance failed 0x%08lX\n", hr);
        }
        else
        {
            long r = calc->Add(2, 3);               // [retval] becomes the return value!
            wprintf(L"Add -> %ld\n", r);
            wprintf(L"Describe -> %s\n", (LPCWSTR)calc->Describe());

            try { calc->Divide(10, 0); }
            catch (const _com_error& e)             // a failed HRESULT is thrown, not returned
            {
                wprintf(L"Divide -> %s\n", (LPCWSTR)e.Description());
            }
        }
    }
    CoUninitialize();
    return 0;
}
```

Three things that surprise people the first time:

- `#import` turns `HRESULT Add([in] LONG, [in] LONG, [out, retval] LONG*)` into `long Add(long, long)`. The `HRESULT` does not vanish — a failure is turned into a thrown `_com_error`, which is why the `Divide` call above is wrapped in `try`. (The console app template already compiles with `/EHsc`, so this needs no project change.)
- `named_guids` is what gives you the names `CLSID_Calculator` and `IID_ICalculator`. Without it you write `__uuidof(Calculator)`.
- The braces around the body matter: `ICalculatorPtr` must release before `CoUninitialize` runs. That is the same bug you hit in Lab 3.1.

Open the generated `TrainingCalc.tlh` and read it. It is the clearest possible answer to "what does a type library actually contain" — and it is generated, so it cannot be out of date with the server.

### C++ late bound (raw `IDispatch`)

Write this by hand once. It's tedious, and that's the lesson.

This is the `Client_Dispatch` project, and it is the whole program — there is no `#import`, no type library, and no generated header anywhere in it. It knows two things about the server: a ProgID string and a method name. That is exactly what PowerShell and VBScript know, which is why this client is worth writing: **everything the scripting hosts do for you is in here, spelled out.**

It needs `#include <atlbase.h>` for `CComVariant` (installed with ATL in Stage 4) and the `RETURN_IF_FAILED` macro from Module 1 §1.5.

```cpp
#include <windows.h>
#include <atlbase.h>
#include <stdio.h>

#define RETURN_IF_FAILED(x) do { HRESULT _hr = (x); if (FAILED(_hr)) return _hr; } while (0)

// Calls any method taking two LONGs and returning a LONG - by name, with no header.
HRESULT CallLateBound(IDispatch* pDisp, LPCOLESTR method, long a, long b, long* pResult)
{
    OLECHAR* name = const_cast<OLECHAR*>(method);
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
        {
            wprintf(L"  server error: %s (scode 0x%08lX)\n",
                    excep.bstrDescription ? excep.bstrDescription : L"(none given)",
                    excep.scode);
            SysFreeString(excep.bstrSource);        // EXCEPINFO strings are YOURS to free
            SysFreeString(excep.bstrDescription);
            SysFreeString(excep.bstrHelpFile);
        }
        else if (hr == DISP_E_TYPEMISMATCH)
        {
            wprintf(L"  bad type for argument %u\n", argErr);
        }
        return hr;
    }

    *pResult = result.lVal;
    return S_OK;
}

int main()
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    {
        CLSID clsid;
        HRESULT hr = CLSIDFromProgID(L"TrainingCalc.Calculator.1", &clsid);
        if (FAILED(hr))
        {
            wprintf(L"ProgID not registered: 0x%08lX\n", hr);
        }
        else
        {
            CComPtr<IDispatch> disp;
            hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&disp));
            if (FAILED(hr))
            {
                wprintf(L"CoCreateInstance failed: 0x%08lX\n", hr);
            }
            else
            {
                long r = 0;

                hr = CallLateBound(disp, L"Add", 2, 3, &r);
                wprintf(L"Add(2,3)       hr=0x%08lX  r=%ld\n", hr, r);

                hr = CallLateBound(disp, L"Divide", 10, 0, &r);   // the server raises
                wprintf(L"Divide(10,0)   hr=0x%08lX\n", hr);

                hr = CallLateBound(disp, L"Multiply", 6, 7, &r);  // no such member
                wprintf(L"Multiply(6,7)  hr=0x%08lX\n", hr);
            }
        }
    }
    CoUninitialize();
    return 0;
}
```

The last two calls are there to give you the two failures you will spend your career reading:

```text
Add(2,3)       hr=0x00000000  r=5
  server error: Cannot divide by zero. (scode 0x80070057)
Divide(10,0)   hr=0x80020009
Multiply(6,7)  hr=0x80020006
```

| Call | HRESULT | Where it came from |
|---|---|---|
| `Divide(10, 0)` | `0x80020009` `DISP_E_EXCEPTION` | The **method ran** and failed. The detail is in `EXCEPINFO`, not in the HRESULT. |
| `Multiply(6, 7)` | `0x80020006` `DISP_E_UNKNOWNNAME` | `GetIDsOfNames` failed. The method never ran, and the server never saw the call. |

That distinction is the whole reason late binding has two steps. `DISP_E_UNKNOWNNAME` means the *name* is wrong — a typo, or a client written against a newer version of the server. `DISP_E_EXCEPTION` means the name was fine and the object objected. A customer reporting "it can't find the method" and a customer reporting "the method threw" are in completely different places.

Follow the description in that transcript back to its source. `Cannot divide by zero.` is the literal string from the `Error(...)` call you wrote in the Stage 4 server, and `scode 0x80070057` is the `E_INVALIDARG` you passed alongside it. `ISupportErrorInfo` carried both across to a caller that has never seen your header — which is the same route they take to reach `$_.Exception.Message` in PowerShell and `Err.Description` in VBScript, a few sections below.

> **`bstrDescription` can be empty.** You get it here only because the server went to the trouble; an object that just returns a failed HRESULT leaves you `scode` and nothing else. Print both and never assume the description is there.

Now count what you wrote. Roughly sixty lines to call `Add(2, 3)` — and PowerShell does the same thing in one. That is not PowerShell being clever; it is PowerShell running this code on your behalf, every single call.

### PowerShell

```powershell
$calc = New-Object -ComObject TrainingCalc.Calculator.1
$calc.Add(2, 3)
$calc.Precision = 4          # property put
$calc.Precision              # property get
$calc.Describe()
$calc.SumTo(10)

try { $calc.Divide(10, 0) } catch { $_.Exception.Message }   # IErrorInfo surfaces here

[Runtime.InteropServices.Marshal]::ReleaseComObject($calc)
```

```text
5
4
Training Calculator 1.0
55
Cannot divide by zero.
0
```

Six lines of output for six lines of script, and not one line of type information anywhere — no header, no type library reference, no `Add-Type`. `New-Object -ComObject` took a ProgID string; everything after it went through `GetIDsOfNames` and `Invoke`, exactly as your `Client_Dispatch` project does by hand.

Two of those lines are worth stopping on:

- **`Cannot divide by zero.`** is the string you passed to `Error(...)` in the C++ server, arriving as a first-class PowerShell exception message. Nothing in between translated it; `ISupportErrorInfo` carried it.
- **The trailing `0`** is what `ReleaseComObject` returns: the reference count *remaining* after it released. Zero means the object is gone. If you see anything else, something is still holding a reference — which is how Lab 5.2's leak announces itself.

> **`"$calc"` does not call `Describe`.** It prints `System.__ComObject`. Interpolation asks .NET's runtime-callable wrapper for a string, and that wrapper answers with its own type name without consulting the object at all. In PowerShell, call `$calc.Describe()` and mean it. The default member is a *scripting* convenience, and the next section is where you can actually see it.

### VBScript

```vbscript
Set calc = CreateObject("TrainingCalc.Calculator.1")
WScript.Echo "Add(2,3)   = " & calc.Add(2, 3)
calc.Precision = 4
WScript.Echo "Precision  = " & calc.Precision
WScript.Echo "Describe() = " & calc.Describe()

On Error Resume Next
calc.Divide 10, 0
If Err.Number <> 0 Then
    WScript.Echo "Divide     -> " & Hex(Err.Number) & " " & Err.Description
End If
```

Run it with the **64-bit** host to match the server (Module 2 §2.3):

```text
C:\> %SystemRoot%\System32\cscript.exe //nologo client.vbs
Add(2,3)   = 5
Precision  = 4
Describe() = Training Calculator 1.0
Divide     -> 80070057 Cannot divide by zero.
```

Compare that last line with what `Client_Dispatch` printed for the same call. C++ got `DISP_E_EXCEPTION` (`0x80020009`) and had to dig `scode` and `bstrDescription` out of `EXCEPINFO` itself. VBScript reports `Err.Number = 80070057` and `Err.Description = Cannot divide by zero.` — it has thrown away the `DISP_E_EXCEPTION` wrapper and handed you the two `EXCEPINFO` fields directly. Same bytes, unpacked for you.

> **The default member is missing here, on purpose.** Try adding `WScript.Echo calc` on its own and you get:
>
> ```text
> Object doesn't support this property or method
> ```
>
> — error `1B6`, because the Stage 4 wizard gave `Describe` the DISPID `5`, and a default member must be DISPID `0`. Change it to `[id(DISPID_VALUE)]` in `TrainingCalc.idl`, rebuild, and the bare `WScript.Echo calc` starts working. That is §5.3's `DISPID_VALUE` discussion, in the one language where you can watch it switch on and off.

### C#

Two clients in one project. Set the target framework to **`net10.0-windows`** and **Platform target = x64**:

```xml
<PropertyGroup>
  <OutputType>Exe</OutputType>
  <TargetFramework>net10.0-windows</TargetFramework>
  <PlatformTarget>x64</PlatformTarget>
</PropertyGroup>
```

The `-windows` suffix is not strictly required — a plain `net10.0` target still builds and runs — but it declares the truth, and without it every COM call raises `CA1416: 'Type.GetTypeFromProgID(string)' is only supported on: 'windows'`. Any supported .NET works here; pick the current LTS.

```csharp
// Late bound - no reference needed
Type t = Type.GetTypeFromProgID("TrainingCalc.Calculator.1");
dynamic calc = Activator.CreateInstance(t);
Console.WriteLine(calc.Add(2, 3));

// Early bound - add a COM reference to the type library
var calc2 = new TrainingCalcLib.Calculator();
Console.WriteLine(calc2.Add(2, 3));
```

For the early-bound half, right-click the project → **Add → COM Reference** → *TrainingCalc 1.0 Type Library*. It appears in that list only because `regsvr32` registered the type library, so if it is not there, go back to the Stage 4 README step 6.

> **Build the early-bound client from Visual Studio, not from `dotnet build`.** Resolving a COM reference means running `tlbimp`, and that step exists only in the .NET Framework MSBuild. `dotnet build` and `dotnet run` fail with:
>
> ```
> error MSB4803: The task "ResolveComReference" is not supported on the .NET Core
> version of MSBuild. Please use the .NET Framework version of MSBuild.
> ```
>
> Build from the IDE, or from a Developer PowerShell with `msbuild Client_CSharp.csproj`. The **late-bound** half has no such restriction — it needs no reference at all, so `dotnet run` is fine for it.

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
> - **Tools:** Visual Studio C++ with **ATL** (connection-point implementation); **either** Windows PowerShell 5.1 **or** PowerShell 7 — they behave identically here, including the failure in step 5a, so one is enough; `cscript.exe` for the script sink.
> - **Elevation:** none, if the server is already registered from Lab 5.1. Rebuilding it re-runs registration, which needs admin unless you took the per-user redirection option in the Stage 4 README.
> - **Bitness:** x64 — the sink, the scripting host and the server must all agree.
> - **Depends on:** Lab 5.1, complete and registered, plus the ref-count tracing from Module 1 — without the trace the leak is invisible, which is the lesson.
> - **Starting point:** [`labs/stage-4-atl-server/`](../labs/stage-4-atl-server/), built with connection points enabled (step 4 of its README).
> - **Time:** ~2 h.

Connection points build a reference cycle **by construction**: the source holds the sink so it can raise events, and the sink holds the source so it can unsubscribe. Nothing is wrong with either half.

Here you wire one up, remove the `Unadvise`, and watch the trace show that *nothing is ever destroyed* — while the program keeps working perfectly. Then you subscribe from a script, where the same cycle exists under a different name and the tooling to find it is thinner.

1. **Write the sink.** The server half already exists: Stage 4 gave you `_ICalculatorEvents`, the connection point, the `Fire_OnProgress` helper and a `SumTo` that calls it, and Lab 5.1 had you build and register the lot.

   What has never existed is anything on the **receiving** end. `SumTo(1000)` calls `Fire_OnProgress` ten times *today* — but `Fire_OnProgress` walks the connection point's list of subscribers, that list is empty, and so the ten notifications go nowhere at all. Every client you have written so far only ever called *into* the object. This lab is where you supply the other direction.

   Add one more **Console App (C++)** to the Lab 5.1 solution, named `Client_Sink`, **Debug | x64**. It needs three things, in this order:

   | | What | Where it comes from |
   |---|---|---|
   | a | the same `#import` line as `Client_Import` | Lab 5.1 |
   | b | the `CalcSink` class, verbatim | §5.7 |
   | c | the `main` below | here |

   `#import` is what makes (b) compile: `named_guids` generates **`DIID__ICalculatorEvents`**, the IID of the event dispinterface, which `CalcSink::QueryInterface` must answer to and which you pass to `FindConnectionPoint`. Without it you would be writing that GUID out by hand.

   ```cpp
   #include <windows.h>
   #include <atlbase.h>
   #include <stdio.h>

   #import "C:\\...\\TrainingCalc\\x64\\Debug\\TrainingCalc.dll" no_namespace named_guids

   // ... paste CalcSink from §5.7 here ...

   int main()
   {
       CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
       {
           ICalculatorPtr calc;
           HRESULT hr = calc.CreateInstance(CLSID_Calculator);
           if (FAILED(hr)) { wprintf(L"CreateInstance 0x%08lX\n", hr); }
           else
           {
               CComPtr<IConnectionPointContainer> cpc;
               hr = calc->QueryInterface(IID_PPV_ARGS(&cpc));   // "do you raise events?"
               wprintf(L"QI IConnectionPointContainer   0x%08lX\n", hr);

               CComPtr<IConnectionPoint> cp;
               hr = cpc->FindConnectionPoint(DIID__ICalculatorEvents, &cp);
               wprintf(L"FindConnectionPoint            0x%08lX\n", hr);

               CalcSink* sink = new CalcSink();                 // starts at m_cRef = 1
               DWORD cookie = 0;
               hr = cp->Advise(sink, &cookie);                  // source AddRefs -> 2
               wprintf(L"Advise                         0x%08lX cookie=%lu\n", hr, cookie);
               sink->Release();                                 // drop ours     -> 1

               wprintf(L"--- before SumTo\n");
               long total = calc->SumTo(1000);
               wprintf(L"--- after SumTo, total=%ld\n", total);

               hr = cp->Unadvise(cookie);                       // source Releases -> 0
               wprintf(L"Unadvise                       0x%08lX\n", hr);
           }
       }
       CoUninitialize();
       return 0;
   }
   ```

   ```text
   QI IConnectionPointContainer   0x00000000
   FindConnectionPoint            0x00000000
   Advise                         0x00000000 cookie=1
   --- before SumTo
     [event] 10% done
     [event] 20% done
     ...
     [event] 100% done
   --- after SumTo, total=500500
   Unadvise                       0x00000000
   ```

   Every `OnProgress` lands **between** the two `---` lines. Events are delivered synchronously, on your thread, inside the call you made — the source is still sitting in the middle of `SumTo` while your handler runs. That is the fact the rest of this lab depends on, and step 6's reentrancy drill is what happens when you abuse it.

   > Call `SumTo(1000)`, not `SumTo(10)`. The Stage 4 implementation reports every 10% and skips the notification entirely for `n < 100`, so a small argument looks exactly like a sink that was never connected.
2. **Make both lifetimes visible.** The leak in step 3 is silent — the program produces correct answers whether or not it leaks — so you need to *see* construction and destruction before you can see them stop happening. The two sides need different techniques, because you own one and ATL owns the other.

   **The sink is yours**, so trace it directly. `CalcSink` already has `AddRef` and `Release` as one-liners — **delete those two and paste these in their place**, then add a destructor:

   ```cpp
   ~CalcSink() { wprintf(L"  [sink] destroyed\n"); }

   ULONG STDMETHODCALLTYPE AddRef() override
   {
       LONG n = InterlockedIncrement(&m_cRef);
       wprintf(L"  [sink] AddRef  -> %ld\n", n);
       return n;
   }
   ULONG STDMETHODCALLTYPE Release() override
   {
       LONG n = InterlockedDecrement(&m_cRef);
       wprintf(L"  [sink] Release -> %ld\n", n);
       if (!n) delete this;                  // read n BEFORE this line, not after
       return n;
   }
   ```

   > Pasting these *alongside* the originals gives you `error C2535: member function already defined or declared`, pointing at the one-liners you meant to remove.

   **The source is ATL's.** `CComObject` implements `AddRef`/`Release` for you and there is no clean place to put a print inside them — but you do not need one. ATL calls **`FinalConstruct`** when the object is created and **`FinalRelease`** when its reference count reaches zero. Birth and death, which is exactly what this lab is about. Both already exist in the generated `Calculator.h`, empty:

   ```cpp
   // Calculator.h, inside CCalculator - fill in the two the wizard left blank
   HRESULT FinalConstruct() { wprintf(L"[source] constructed\n"); return S_OK; }
   void    FinalRelease()   { wprintf(L"[source] destroyed\n"); }
   ```

   You do not need to add `#include <stdio.h>` — `framework.h` already includes `<atlbase.h>`, which pulls it in. If you ever see `C3861: 'wprintf': identifier not found` in a project that is not using ATL, that is the include you are missing.

   > **The server's output appears on your client's console.** The DLL is loaded *into* your client process, and both link the same shared UCRT, so `[source]` and `[sink]` lines interleave with the client's own in the right order. No DebugView, no `OutputDebugString`, nothing to attach.

   > **No, you do not need to register anything again.** A rebuild writes the same DLL to the same path under the same CLSID, so every registry entry still points at it. (ATL's *Register Output* re-runs `regsvr32` on each build anyway, harmlessly.) What *will* catch you is the opposite problem: **the build fails while a client is still running**, because the DLL is loaded and locked —
   >
   > ```text
   > LNK1168: cannot open ...\TrainingCalc.dll for writing
   > ```
   >
   > Close `Client_Sink` and any PowerShell session that ran `New-Object`, then build again. A PowerShell window holds the DLL until the object is released *and* the process exits.

   **Now build and run.** Step 1's HRESULT prints have done their job; replace them with plain phase markers — `--- new CalcSink`, `--- Advise`, `--- our Release`, `--- Unadvise`, `--- end of scope` — so that every reference count below can be attributed to the call that caused it.

   Everything still works. This is the healthy case, and it is the baseline step 3 will be a diff against:

   ```text
   [source] constructed
   --- new CalcSink
   --- Advise
     [sink] AddRef  -> 2
   --- our Release
     [sink] Release -> 1
   --- before SumTo
     [sink] AddRef  -> 2
     [sink] AddRef  -> 3
     [event] 10% done
     [sink] Release -> 2
     [sink] Release -> 1
     ...                       (the same five lines, eight more times)
     [sink] AddRef  -> 2
     [sink] AddRef  -> 3
     [event] 100% done
     [sink] Release -> 2
     [sink] Release -> 1
   --- after SumTo, total=500500
   --- Unadvise
     [sink] Release -> 0
     [sink] destroyed
   --- end of scope
   [source] destroyed
   ```

   Read it line by line before moving on — every claim §5.7 made is sitting in this transcript:

   | In the trace | What it proves |
   |---|---|
   | `AddRef -> 2` during `Advise` | the source took a **strong** reference to your sink. This is the half of the cycle you cannot see from the client. |
   | `Release -> 1` immediately after | you dropped yours, and the sink did **not** die — the subscription alone is keeping it alive. |
   | the count returns to **1** between every event | each callback is bracketed by a matched `AddRef`/`Release`, so the sink cannot be destroyed halfway through its own handler. The exact number of pairs is an ATL implementation detail; what matters is that it balances and never reaches zero. |
   | `Release -> 0` at `Unadvise` | `Unadvise` is the **only** thing that destroys the sink. Nothing else was ever going to. |
   | `[source] destroyed` last | the `ICalculatorPtr` released at end of scope, and the server object went with it. |

   Both objects were created and both were destroyed. Keep this output somewhere you can compare against.
3. **Remove the `Unadvise`.**

   **3a. Watch it *not* leak.** Comment out `cp->Unadvise(cookie)` and run. Both objects are still destroyed:

   ```text
   --- Unadvise SKIPPED
   --- end of scope
   [source] destroyed
     [sink] Release -> 0
     [sink] destroyed
   ```

   That is not the lab going wrong. It is the most useful thing in it, and it is worth understanding before you go any further: **there is no cycle yet.** Every reference you have created so far points one way.

   ```text
   Client ──strong──► Source ──strong──► Sink
   ```

   When `calc` leaves scope the source's count reaches zero, and as the source is destroyed ATL's connection point releases every sink still sitting in its list. Forgetting `Unadvise` cost you nothing, because there was always something able to reach zero on its own.

   Remember this when you read a bug report. "We never call `Unadvise`" is not, by itself, a leak.

   **3b. Close the loop.** Give the sink a strong reference back to the source. This is not a contrivance to make the lab work — it is what real sinks look like. A sink almost always needs the object it is listening to: to read a property while handling the event, to `Unadvise` itself later, or simply because the sink *is* the client object that owns the source in the first place.

   ```cpp
   class CalcSink : public IDispatch
   {
       LONG m_cRef = 1;
       CComPtr<ICalculator> m_source;              // add this
   public:
       void SetSource(ICalculator* p) { m_source = p; }    // CComPtr AddRefs
       // ... rest unchanged ...
   ```

   and in `main`, immediately after `new CalcSink()`:

   ```cpp
   sink->SetSource(calc);
   ```

   Run again, still with no `Unadvise`:

   ```text
   [source] constructed
   --- before SumTo
     ... ten events, exactly as before ...
   --- after SumTo, total=500500
   --- Unadvise SKIPPED
   --- end of scope
   ```

   And then nothing. **Neither `[source] destroyed` nor `[sink] destroyed` ever prints.** The program produced the right answer, exited cleanly, reported no error, and destroyed nothing:

   ```text
   Client ──strong──► Source ──strong──► Sink
                        ▲                 │
                        └─────strong──────┘
   ```

   The client dropped its reference and the source's count fell to one — the sink's. The sink's count is one — the source's. Each is keeping the other alive and neither can ever reach zero. One added line turned a program that cleaned up perfectly into one that cannot clean up at all, and **the only visible difference is two lines of output that no longer appear.**

   This is the single most common leak in COM, and that is why: nothing fails.
4. **Break the cycle, and make it impossible to forget.** Put the `Unadvise` back and both destructors return — `[sink] destroyed` first, because the sink's count reaches zero immediately, and `[source] destroyed` right after, once the sink's `m_source` has released too.

   Then delete the `cp->Unadvise(cookie)` line again — this time replacing it with a destructor rather than with a leak. `ConnectionCookie` owns the whole subscription: it finds the connection point, `Advise`s, and `Unadvise`s when it goes out of scope. Paste it above `main` — this is §5.7's wrapper with the error handling written out, so it needs no `RETURN_IF_FAILED`:

   ```cpp
   class ConnectionCookie
   {
       CComPtr<IConnectionPoint> m_cp;
       DWORD m_cookie = 0;
   public:
       HRESULT Advise(IUnknown* pSource, REFIID iid, IUnknown* pSink)
       {
           CComPtr<IConnectionPointContainer> cpc;
           HRESULT hr = pSource->QueryInterface(IID_PPV_ARGS(&cpc));
           if (FAILED(hr)) return hr;
           hr = cpc->FindConnectionPoint(iid, &m_cp);
           if (FAILED(hr)) return hr;
           return m_cp->Advise(pSink, &m_cookie);
       }
       ~ConnectionCookie()
       {
           if (m_cp && m_cookie) m_cp->Unadvise(m_cookie);
       }
   };
   ```

   The `cpc`, `cp`, `Advise` and `Unadvise` lines all disappear from `main`, and four steps collapse into one:

   ```cpp
   int main()
   {
       CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
       {
           ICalculatorPtr calc;
           calc.CreateInstance(CLSID_Calculator);

           CalcSink* sink = new CalcSink();
           sink->SetSource(calc);                  // step 3b's back-reference - the cycle

           ConnectionCookie cookie;                // declared AFTER calc
           HRESULT hr = cookie.Advise(calc, DIID__ICalculatorEvents, sink);
           wprintf(L"Advise 0x%08lX\n", hr);
           sink->Release();

           wprintf(L"--- before SumTo\n");
           long total = calc->SumTo(1000);
           wprintf(L"--- after SumTo, total=%ld\n", total);

           wprintf(L"--- end of scope\n");
       }                                           // ~ConnectionCookie -> ~CalcSink -> ~CCalculator
       wprintf(L"=== done\n");
       CoUninitialize();
       return 0;
   }
   ```

   ```text
   [source] constructed
     [sink] AddRef  -> 2
   Advise 0x00000000
     [sink] Release -> 1
   --- before SumTo
     ... ten events ...
   --- after SumTo, total=500500
   --- end of scope
     [sink] Release -> 0
     [sink] destroyed
   [source] destroyed
   === done
   ```

   Both objects destroyed again, with the cycle still in place — the wrapper broke it. Note that everything after `--- end of scope` happens with no code of yours running at all.

   Declare `cookie` **after** `calc`, as above: destructors run in reverse order, so `Unadvise` fires while the source is still alive and the output reads in the natural order.

   The point is not tidiness. A manual `Unadvise` is skipped by every early `return`, every thrown exception, and every future edit that adds one. A destructor is not. ATL offers `AtlAdvise`/`AtlUnadvise` and WIL has equivalents — but write this one once, because the wrapper is the only thing standing between step 3b and production.
5. **Subscribe from a script.** Calling a method needed nothing but a ProgID string. Receiving an event is not the same problem, and this step is where that stops being an abstract claim.

   **5a. Try the obvious thing, and watch it fail.**

   ```powershell
   $calc = New-Object -ComObject TrainingCalc.Calculator.1
   $calc | Get-Member -MemberType Event
   Register-ObjectEvent -InputObject $calc -EventName OnProgress -Action {
       Write-Host "progress: $($EventArgs)%"
   }
   ```

   ```text
   Cannot register for the specified event. An event with the name 'OnProgress' does not exist.
   ```

   `Get-Member` lists no events at all. `New-Object -ComObject` returns a `System.__ComObject` — a wrapper holding an `IDispatch` pointer and **no type information**. `GetIDsOfNames` can resolve a method name at the moment you call it, but there is no equivalent for "tell me what events you raise", so there is nothing for `Register-ObjectEvent` to bind to. Run it in Windows PowerShell 5.1 as well if you like; the message is identical.

   **5b. VBScript needs no extra machinery at all.**

   ```vbscript
   Set calc = WScript.CreateObject("TrainingCalc.Calculator.1", "calc_")
   WScript.Echo "SumTo(1000) = " & calc.SumTo(1000)

   Sub calc_OnProgress(percent)
       WScript.Echo "  OnProgress " & percent
   End Sub
   ```

   ```text
     OnProgress 10
     OnProgress 20
     ...
     OnProgress 100
   SumTo(1000) = 500500
   ```

   The second argument to `WScript.CreateObject` is a **prefix**. WSH reads the coclass's `[default, source]` dispinterface out of the type library, implements a sink for it, calls `Advise`, and routes each incoming DISPID to a sub named `<prefix><EventName>`. Everything you wrote by hand in §5.7, done by the host.

   Note the ordering: all ten events print **before** `SumTo` returns. That is step 1's lesson again, now with no C++ in sight.

   **5c. PowerShell can do it — once it has a type.**

   Rebuild Lab 5.1's `Client_CSharp` project with `<EmbedInteropTypes>false</EmbedInteropTypes>` on the COM reference. That leaves an **`Interop.TrainingCalcLib.dll`** in its output folder: a real .NET assembly, generated from the type library, in which the source dispinterface has become ordinary .NET events.

   ```powershell
   Add-Type -Path .\Interop.TrainingCalcLib.dll
   $calc = New-Object TrainingCalcLib.CalculatorClass
   $calc | Get-Member -MemberType Event | ForEach-Object Name     # OnError, OnProgress

   Register-ObjectEvent -InputObject $calc -EventName OnProgress -SourceIdentifier p1
   $calc.SumTo(1000)
   Get-Event -SourceIdentifier p1 | ForEach-Object { "  OnProgress " + $_.SourceArgs[0] }

   Get-EventSubscriber | Unregister-Event      # the PowerShell equivalent of Unadvise
   ```

   Ten events, in both hosts. Nothing about the *server* changed between 5a and 5c — only whether the caller had type information. Late binding was enough to **call** this object; it was never enough to **listen** to it.

   Leave out the `Unregister-Event` and you have rebuilt this lab's leak in a scripting host: the subscription holds a reference to the object, the object holds the sink, and neither is collected while the session lives. `Get-EventSubscriber` is your `Unadvise` audit.

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

**Next: [Module 6 — Choosing COM tools and interop](06-frameworks-and-interop.md)**
