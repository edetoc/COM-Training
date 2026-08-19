# Module 1 — `IUnknown`, vtables, and lifetime

**Time: 1 week. This is the most important module in the course.**

If you understand this module completely, you can reason about 70% of real COM bugs. If you skim it, everything later will feel like arbitrary rules.

---

## 1.1 What an interface pointer actually is

Module 0 gave the one-paragraph version. Here is the precise one, because everything in this course rests on it.

> **Definition.** A **vtable** is a compile-time-generated array of function pointers — one entry per virtual method, in declaration order. There is exactly **one vtable per class**, not per object; it lives in the binary's read-only data section and is shared by every instance.
>
> A **vptr** is a hidden pointer *inside each object* that points to its class's vtable. Under MSVC (Microsoft Visual C++) it sits at offset 0 — its placement is a compiler choice, not a language rule. An object with two independent interface bases has **two** vptrs.
>
> A **slot** is an index into the vtable. The caller hardcodes the slot number; the callee's vtable supplies the address.

So an "interface pointer" is not a pointer to code, and not really a pointer to an object either — it is **a pointer to a vptr**. That indirection is the entire foundation of COM.

Start with plain C++. When a class has virtual functions, the compiler puts a hidden pointer at the start of each object. That pointer points to a **vtable** — a static array of function pointers.

```cpp
struct ICalculator
{
    virtual HRESULT __stdcall Add(long a, long b, long* result) = 0;
    virtual HRESULT __stdcall Subtract(long a, long b, long* result) = 0;
};
```

An object implementing this looks like:

```
  pInterface ──►  ┌──────────────┐
                  │  vptr        │────► ┌─────────────────────┐  vtable
                  ├──────────────┤      │ &Impl::Add          │  slot 0
                  │  m_data...   │      ├─────────────────────┤
                  │              │      │ &Impl::Subtract     │  slot 1
                  └──────────────┘      └─────────────────────┘
```

When the client writes `pCalc->Add(2, 3, &r)`, the compiler emits roughly:

```asm
mov  rax, [pCalc]        ; load the vptr
call qword ptr [rax + 0] ; call slot 0, with pCalc as the hidden 'this'
```

**That's the whole magic.** Two machine instructions. There is nothing compiler-specific about "load a pointer, index an array, call it, pass `this` in the first register." Every compiler, every language that can make an indirect call, can do this.

This also explains, physically, three things:

- **Why you can't add a method in the middle.** The client compiled `call [rax + 8]` for `Subtract`. Insert a method before it and slot 1 is now something else. The client will call the wrong function with the wrong arguments.
- **Why interfaces can't have data members.** The client never knows the object's size, so any layout the client assumed would be a lie.
- **Why `dps <interface-ptr>` in WinDbg is so useful.** Dereference the pointer, dump the vtable, and you see the symbol names of the actual implementation — which tells you *which component you really got*.

### Verify it yourself

```cpp
// Prove the vptr is at offset 0 and the vtable is an array of function pointers.
ICalculator* p = /* ... */;
void** vptr = *reinterpret_cast<void***>(p);
printf("vtable at %p\n", vptr);
printf("slot 0 (Add)      = %p\n", vptr[0]);
printf("slot 1 (Subtract) = %p\n", vptr[1]);
```

Run that under a debugger and compare `vptr[0]` with the address of your `Add` implementation. They match.

---

## 1.2 `IUnknown` — the interface every interface inherits

```cpp
struct IUnknown
{
    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) = 0;
    virtual ULONG   __stdcall AddRef() = 0;
    virtual ULONG   __stdcall Release() = 0;
};
```

Because every COM interface derives from `IUnknown`, **slots 0, 1, 2 of every COM vtable are always these three**. Your own methods start at slot 3.

```
  ICalculator vtable:
    slot 0  QueryInterface     ─┐
    slot 1  AddRef              ├─ IUnknown, always
    slot 2  Release            ─┘
    slot 3  Add
    slot 4  Subtract
```

`QueryInterface` gives you **navigation** (what else can this object do?). `AddRef`/`Release` give you **lifetime**.

---

## 1.3 `QueryInterface` — the four rules

`QueryInterface` answers: *"Do you support interface X? If so, give me a pointer to it."*

```cpp
HRESULT QueryInterface(REFIID riid, void** ppv);
// S_OK          -> *ppv holds a valid, AddRef'd pointer
// E_NOINTERFACE -> *ppv is NULL, object does not support riid
// E_POINTER     -> ppv was NULL
```

The rules exist so that clients can reason about objects they've never seen.

### Rule 1 — Reflexive / identity

`QI(IID_IUnknown)` on **any** interface of an object must return **the exact same pointer value**.

```cpp
IUnknown* pUnk1 = nullptr;
IUnknown* pUnk2 = nullptr;
pCalc->QueryInterface(IID_IUnknown, (void**)&pUnk1);
pPersist->QueryInterface(IID_IUnknown, (void**)&pUnk2);
// If pUnk1 == pUnk2, pCalc and pPersist are the SAME OBJECT.
```

**This is the only legal way to test COM object identity.** Comparing `pCalc == pPersist` is meaningless: with multiple inheritance those are different addresses within the same object; with tear-offs or proxies they're different objects entirely.

Also: `QI(IID_IUnknown)` must succeed. Always.

### Rule 2 — Symmetric

If you can get from `A` to `B`, you must be able to get from `B` back to `A`.

### Rule 3 — Transitive

If `A`→`B` succeeds and `B`→`C` succeeds, then `A`→`C` must succeed.

### Rule 4 — Static

The set of interfaces an object supports must **not change over its lifetime**. If `QI(IID_IFoo)` succeeds once, it must succeed every time. If it fails once, it must fail every time.

> This rule is why "the interface appears after initialization" is illegal, and why intermittent `E_NOINTERFACE` is always a bug — usually an aggregation error, a thread/apartment problem, or an object that isn't really the object you think it is.

### Rule 5 (the one everyone forgets) — Always `AddRef` on success

`QueryInterface` returns a **new reference**. Even when it returns `this`. Even when you ask for the same interface you already have.

```cpp
IFoo* p2 = nullptr;
pFoo->QueryInterface(IID_IFoo, (void**)&p2);   // ref count is now 2
// You must Release BOTH p and p2.
```

### And on failure

Set `*ppv = nullptr` **before** returning a failure. Clients rely on this. Also null it on entry, so an early failure path can't leave garbage.

---

## 1.4 Reference counting — the rules that actually matter

The mental model that makes everything click:

> **You do not count objects. You count *references*.**
> Every *variable* holding an interface pointer that outlives the current statement owns one count.

### When the count is incremented

1. **Activation.** `CoCreateInstance`, `CoCreateInstanceEx`, `CoGetClassObject`, `IClassFactory::CreateInstance`. The pointer arrives already `AddRef`ed — usually count 1.
2. **`QueryInterface`,** always, on success.
3. **Any interface pointer returned via `[out]` or `[out, retval]`.** Enumerators (`IEnumXxx::Next`), property getters, factory methods, `IServiceProvider::QueryService`. The callee `AddRef`s; the caller owns it and must `Release`.
4. **Copying into storage that outlives the current scope** — a member variable, a global, an array, a `std::vector`, a lambda capture, a queued work item.
5. **A callee that keeps an `[in]` pointer past the call.** If you save it, you `AddRef` it.
6. **Marshaling.** `CoMarshalInterface`, `CoMarshalInterThreadInterfaceInStream`, `IGlobalInterfaceTable::RegisterInterfaceInGlobal`/`GetInterfaceFromGlobal`, proxy creation, `CoCreateInstanceEx` MULTI_QI entries.
7. **Smart pointer copy/assign** — `CComPtr`, `_com_ptr_t`, `winrt::com_ptr`, `Microsoft::WRL::ComPtr` all `AddRef` implicitly.
8. **Defensively, when lifetime is ambiguous** — e.g. before calling out to code that might release the owner, or before releasing the object that gave you the pointer.

### When it is NOT incremented

- Passing an `[in]` interface pointer for the duration of a call only. The caller's reference already guarantees the object lives.
- Local aliases in a scope clearly dominated by an existing reference.
- `Attach`/`Detach` on smart pointers — ownership *transfer*, no count change.
- `[in, out]`: the callee `Release`s the old value if it replaces it — net zero.
- Deliberate weak references (WinRT `IWeakReference`, or the classic parent/child back-pointer).

### The symmetry law

> Every `AddRef` — explicit or implicit — has exactly one matching `Release`.
> Miss one → **leak**. Do one too many → **use-after-free**.

### The pessimistic rule for beginners

When in doubt, `AddRef`. A leak is a slow, findable, non-crashing bug. An over-release is a random crash in unrelated code, often in a different thread, hours later.

### Watching it happen — a worked example

The rules above are a list. Here they are as one continuous story.

First, the three helpers the walkthrough calls. Each one shows the *other* side of a rule — where the reference is created, or deliberately isn't. (`Calculator` is the class you'll write in §1.6; `CreateCalculator` is a stand-in for `CoCreateInstance`, which Module 2 covers properly.)

```cpp
// ---- helper 1: creation --------------------------------------------------
// Hands back an object with ONE reference already taken, which the CALLER owns.
// This is the shape every real class factory uses.
HRESULT CreateCalculator(REFIID riid, void** ppv, const char* tag)
{
    if (!ppv) return E_POINTER;
    *ppv = nullptr;

    Calculator* p = new (std::nothrow) Calculator(tag);   // born at count = 1
    if (!p) return E_OUTOFMEMORY;

    HRESULT hr = p->QueryInterface(riid, ppv);            // count = 2
    p->Release();                                         // count = 1 — the caller's
    return hr;                                            // (on E_NOINTERFACE this
}                                                         //  correctly destroys it)

// ---- helper 2: borrowing via an [in] parameter ---------------------------
// NO AddRef here. The caller's reference already guarantees the object
// outlives this call, and we don't keep the pointer afterwards.
void PrintDescription(ICalculator* p)
{
    long r = 0;
    p->Add(1, 1, &r);
    printf("calculator says 1 + 1 = %ld\n", r);
}   // nothing to release — we never owned anything

// ---- helper 3: handing one back via an [out] parameter -------------------
// The AddRef that gives the caller its new reference happens HERE, in the callee.
// That is why the caller owes a Release for a pointer it never AddRef'd itself.
HRESULT GetCalculator(ICalculator** ppOut)
{
    if (!ppOut) return E_POINTER;
    *ppOut = nullptr;                    // null the [out] param before anything can fail
    if (!g_pCached) return E_FAIL;

    *ppOut = g_pCached;
    (*ppOut)->AddRef();                  // <-- step (6)'s increment, seen from this side
    return S_OK;
}
```

> **Wait — `new`? Module 0 said COM objects are never created with `new`.**
>
> Both are true, because they describe opposite sides of the boundary.
>
> | | Who | What they do |
> |---|---|---|
> | **Client side** | Anyone *using* the object | Never `new`, never `delete`. Asks for an interface pointer; calls `Release` when finished. |
> | **Implementation side** | The component's *author* | Uses `new` exactly once, in one place, and `delete this` in `Release`. |
>
> `CreateCalculator` above is **implementation-side code**. It lives in the same binary as `Calculator`, compiled by the same compiler at the same moment — so every objection from Module 0 §0.1 (unknown size, mismatched heaps, differing layout) simply doesn't apply. The object's own author is the one person who legitimately knows its size and owns its heap.
>
> The client never sees any of it. It receives an `ICalculator*` and cannot tell how the object came into existence.
>
> In Module 2 this exact function moves into `IClassFactory::CreateInstance`, and the full chain becomes:
>
> ```
> client:  CoCreateInstance(CLSID_Calculator, ..., IID_ICalculator, &p)
>              │
>              ▼   (COM finds the server via the registry)
> server:  IClassFactory::CreateInstance(...)
>              │
>              ▼
>          new Calculator()      <-- the new is HERE, inside the component
> ```
>
> So `CoCreateInstance` doesn't replace `new` — it's the client-side doorway that eventually *reaches* somebody's `new`, on the far side of the boundary. Module 1 uses a plain function instead so you can study lifetime without activation getting in the way; Module 2 adds the activation half.

Now the walkthrough. **The right-hand column is the object's reference count *after* that line runs.**

```cpp
ICalculator* g_pCached = nullptr;        // a global that outlives the function below

void Walkthrough()
{
    ICalculator* pCalc = nullptr;

    // (1) ACTIVATION — the object is born with your reference already taken.
    CreateCalculator(IID_ICalculator, (void**)&pCalc, "demo");      // count = 1

    // (2) A LOCAL ALIAS in the same scope. No new reference is needed: pCalc's
    //     reference already guarantees the object outlives this variable.
    ICalculator* pAlias = pCalc;                                    // count = 1  unchanged

    // (3) QUERYINTERFACE — always AddRefs on success. This is a second,
    //     independent reference with its own lifetime.
    IAdvancedCalculator* pAdv = nullptr;
    pCalc->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);  // count = 2

    // (4) PASSING AS AN [in] PARAMETER — the callee only borrows it for the
    //     duration of the call, so no AddRef. Your reference keeps it alive.
    PrintDescription(pCalc);                                        // count = 2  unchanged

    // (5) STORING IT SOMEWHERE THAT OUTLIVES THIS SCOPE — a new owner, so a
    //     new reference. Forgetting this AddRef is the classic dangling global.
    g_pCached = pCalc;
    g_pCached->AddRef();                                            // count = 3

    // (6) RECEIVING AN INTERFACE VIA AN [out] PARAMETER — the callee AddRef'd
    //     it on your behalf. You own it, even though you never wrote AddRef.
    ICalculator* pFromOut = nullptr;
    GetCalculator(&pFromOut);          // hands back the SAME object // count = 4

    // ---- hand every reference back: one Release per reference held ----

    pFromOut->Release();                                            // count = 3
    pAdv->Release();                                                // count = 2
    pCalc->Release();                                               // count = 1
    // pAlias gets NO Release — it never owned a reference of its own.

}   // The object is STILL ALIVE and that is correct: g_pCached owns the last one.

void Shutdown()
{
    g_pCached->Release();                                           // count = 0
    g_pCached = nullptr;                 // -> the object deletes itself here
}
```

With the tracing from Lab 1.1 switched on, that produces:

```
[COM] CREATE   obj=000001F2A3B4C5D0 count=1     <- (1) activation
[COM] ADDREF   obj=000001F2A3B4C5D0 count=2     <- (3) QueryInterface
[COM] ADDREF   obj=000001F2A3B4C5D0 count=3     <- (5) stored in a global
[COM] ADDREF   obj=000001F2A3B4C5D0 count=4     <- (6) [out] parameter
[COM] RELEASE  obj=000001F2A3B4C5D0 count=3
[COM] RELEASE  obj=000001F2A3B4C5D0 count=2
[COM] RELEASE  obj=000001F2A3B4C5D0 count=1
[COM] RELEASE  obj=000001F2A3B4C5D0 count=0
[COM] DESTROY  obj=000001F2A3B4C5D0 count=0
```

**Four increments, four `Release` calls, ends at zero.** That balance is the entire discipline.

Note what did *not* appear in the trace: steps (2) and (4). No line was emitted for them, because no reference was created. Getting those two wrong in the other direction — `AddRef`ing an alias or an `[in]` parameter — produces a leak just as surely as forgetting a `Release`. **Lab 1.4** (§1.11) is eight snippets built from exactly these mistakes.

#### The same code, two ways to break it

```cpp
// LEAK — the QI reference is never returned.
pCalc->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);   // count = 2
pCalc->Release();                                                // count = 1
// pAdv->Release() missing  ->  count stays at 1, DESTROY never prints.
// Symptom: nothing. The program runs correctly and quietly grows.
```

```cpp
// OVER-RELEASE — one Release too many.
pCalc->Release();          // count = 0  -> object deleted here
pCalc->Release();          // count = ?? -> reads freed memory. Undefined behaviour.
// Symptom: a crash somewhere, sometime. Possibly not here, possibly not now.
```

Both bugs are one line. One is invisible and harmless-looking; the other is a crash you'll be handed in a dump with no obvious cause. **This is why the rules are absolute rather than advisory** — you cannot rely on testing to find either.

You'll reproduce both deliberately in Lab 1.1, and §1.7 shows how smart pointers make the whole example collapse to a handful of lines with no explicit `AddRef` or `Release` at all.

---

## 1.5 `HRESULT` — read it properly

```
 3 3 2 2 2 2 2 2 2 2 2 2 1 1 1 1 1 1 1 1 1 1
 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0 9 8 7 6 5 4 3 2 1 0
┌─┬─┬─┬─┬─┬───────────────────────┬───────────────────────────────┐
│S│R│C│N│X│        Facility       │             Code              │
└─┴─┴─┴─┴─┴───────────────────────┴───────────────────────────────┘
 S = Severity: 0 = success, 1 = failure
```

```cpp
#define SUCCEEDED(hr)  (((HRESULT)(hr)) >= 0)
#define FAILED(hr)     (((HRESULT)(hr)) <  0)
```

**Never write `if (hr == S_OK)`.** There are legitimate success codes other than `S_OK`:

| Code | Value | Meaning |
|---|---|---|
| `S_OK` | `0x00000000` | Success, "true" |
| `S_FALSE` | `0x00000001` | **Success**, but "false"/"nothing to do"/"fewer items than requested" |
| `E_NOINTERFACE` | `0x80004002` | Interface not supported |
| `E_POINTER` | `0x80004003` | Null pointer argument |
| `E_FAIL` | `0x80004005` | Unspecified failure |
| `E_UNEXPECTED` | `0x8000FFFF` | Catastrophic |
| `E_NOTIMPL` | `0x80004001` | Not implemented |
| `E_INVALIDARG` | `0x80070057` | Bad argument |
| `E_OUTOFMEMORY` | `0x8007000E` | OOM |
| `E_ACCESSDENIED` | `0x80070005` | Access denied |

`S_FALSE` is returned by, among others: `CoInitializeEx` (already initialized on this thread), `IEnumXxx::Next` (fewer than requested), `IPersistStream::IsDirty`, `IStream::Clone` in some impls, `CreateStreamOnHGlobal` variants. Treating it as failure is a classic bug; treating it as `S_OK` is also a classic bug.

**Win32 errors as HRESULTs**: `HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)` = `0x80070002`. Facility `0x7` = FACILITY_WIN32, so any `0x8007xxxx` HRESULT's low word is a Win32 error code. You will use this constantly:

```
0x80070005 -> Win32 5   -> ERROR_ACCESS_DENIED
0x8007007E -> Win32 126 -> ERROR_MOD_NOT_FOUND
0x800706BA -> Win32 1722-> RPC_S_SERVER_UNAVAILABLE
```

In WinDbg: `!error 0x8007007e`. On the command line: `certutil -error 0x8007007e`.

### A discipline that saves hours

```cpp
#define RETURN_IF_FAILED(x) do { HRESULT _hr = (x); if (FAILED(_hr)) return _hr; } while(0)
```

Or don't write your own — use **WIL**.

> **WIL** — the **Windows Implementation Library**. Microsoft's open-source, header-only C++ helper library for Windows programming: https://github.com/microsoft/wil. It's not a COM framework; it's a collection of RAII wrappers and error-handling macros that make Windows and COM code far harder to get wrong. Nothing to build or link — just add the headers.
>
> Its `RETURN_IF_FAILED` / `THROW_IF_FAILED` do what the macro above does **and** record the originating file, line, and HRESULT. That matters more than it sounds: when a failure bubbles up through ten layers, WIL tells you where it *started* instead of handing you a bare `E_FAIL` from somewhere unknown. Module 6 §6.4 covers it properly; Module 8 uses its logging for diagnostics.

You'll see it used from Module 6 onward. For the hand-written labs in this module, the three-line macro above is enough.

### One more term: RAII

You'll meet this word constantly, in this course and everywhere else in C++.

> **RAII** — "Resource Acquisition Is Initialization." A famously bad name for a simple idea: **tie a resource's lifetime to a variable's scope.** Acquire the resource in an object's constructor; release it in that object's destructor. C++ guarantees the destructor runs when the variable goes out of scope — on a normal return, an *early* return, a `break`, or a thrown exception. So the cleanup cannot be forgotten or skipped.

For COM, the resource is a reference and the release is `Release()`. Compare:

```cpp
// WITHOUT RAII — every exit path needs its own cleanup, and eventually one won't get it.
HRESULT Work()
{
    ICalculator* p = nullptr;
    HRESULT hr = CreateCalculator(IID_ICalculator, (void**)&p, "x");
    if (FAILED(hr)) return hr;

    long r = 0;
    hr = p->Add(2, 3, &r);
    if (FAILED(hr)) { p->Release(); return hr; }    // remembered here...
    if (r != 5)     { return E_UNEXPECTED; }        // ...and FORGOTTEN here. Leak.

    p->Release();
    return S_OK;
}
```

```cpp
// WITH RAII — the wrapper's destructor releases on every exit path, automatically.
HRESULT Work()
{
    CComPtr<ICalculator> p;                          // acquires nothing yet
    RETURN_IF_FAILED(CreateCalculator(IID_ICalculator, (void**)&p, "x"));

    long r = 0;
    RETURN_IF_FAILED(p->Add(2, 3, &r));
    if (r != 5) return E_UNEXPECTED;                 // still released — impossible to forget

    return S_OK;
}                                                    // destructor runs here, whatever happened
```

The second version has no `Release` call at all, yet leaks on none of its four exit paths. Here's why that works.

#### Why there's no leak — the mechanism

`CComPtr<T>` isn't magic. It's an ordinary C++ class wrapping one raw pointer, and the only part that matters is its destructor:

```cpp
template <class T>
class CComPtr                        // radically simplified
{
    T* p = nullptr;
public:
    ~CComPtr() { if (p) p->Release(); }     // <-- the entire trick

    T** operator&()     { return &p; }      // lets &sp receive an [out] parameter
    T*  operator->()    { return p; }       // lets sp->Add(...) work
    operator T*() const { return p; }
};
```

Two things to notice before the walkthrough:

- **How `p` acquires the reference.** `(void**)&p` invokes `operator&`, which hands out the address of the *internal* pointer. So `CreateCalculator` writes its already-`AddRef`'d pointer straight into the member. From that moment the `CComPtr` owns that reference.
- **How it gives it back.** Only in the destructor. Nowhere else.

Now the C++ guarantee this depends on:

> **When a local variable goes out of scope, its destructor runs.** The compiler emits that call at *every* point where control leaves the scope. You cannot opt out, and you cannot forget one.

So the compiler effectively turns the RAII version into this:

```cpp
HRESULT Work()
{
    CComPtr<ICalculator> p;                              // p.p == nullptr

    HRESULT hr = CreateCalculator(IID_ICalculator, (void**)&p, "x");
    if (FAILED(hr)) { /* ~CComPtr() */ return hr; }              // EXIT 1  ← compiler-inserted

    long r = 0;
    hr = p->Add(2, 3, &r);
    if (FAILED(hr)) { /* ~CComPtr() */ return hr; }              // EXIT 2  ← compiler-inserted

    if (r != 5)     { /* ~CComPtr() */ return E_UNEXPECTED; }    // EXIT 3  ← compiler-inserted

                      /* ~CComPtr() */ return S_OK;              // EXIT 4  ← compiler-inserted
}
```

You wrote none of those four destructor calls. The compiler wrote all of them — **including at exit 3, the one a human forgot in the raw version.**

What each exit actually does:

| Exit | Reached when | State of `p` | Destructor does |
|---|---|---|---|
| 1 | creation failed | still `nullptr` | `if (p)` is false → nothing. Correct: there was no reference to give back. |
| 2 | `Add` failed | holds 1 reference | `Release()` → count 0 → object destroyed |
| 3 | wrong answer | holds 1 reference | `Release()` → object destroyed |
| 4 | success | holds 1 reference | `Release()` → object destroyed |

And there's a fifth path you never wrote at all: if any call **throws an exception**, C++ unwinds the stack and runs destructors on the way out — so the reference is still returned. The raw-pointer version leaks there too, silently.

#### The contrast, stated exactly

In the raw version, `p->Release()` exists **only where a human typed it.** Exit 3 has no such line, so that reference is never given back and the object lives until the process dies.

In the RAII version, the release lives in **one place** — the destructor — and the *compiler* decides where that runs. It runs at all four exits, plus every exception path, without anyone deciding anything.

> That's the real point. Not that RAII cleans up "automatically" in some vague sense, but that **cleanup stops being a thing you remember at N call sites and becomes a thing you write once at one site**, which the compiler then applies everywhere. N chances to be wrong become zero.

This same reasoning is why Module 3 wraps `Unadvise` in a destructor, Module 5 wraps connection-point cookies, and Module 7 wraps `CoRevertToSelf`. In every case the bug being prevented is "somebody added an early return six months later and didn't notice the cleanup line above it."

You can recognize RAII types by name — anything like `*_ptr`, `unique_*`, `scope_exit`, or `*Lock` is one:

| Type | Acquires / releases |
|---|---|
| `CComPtr<T>`, `wil::com_ptr<T>`, `winrt::com_ptr<T>` | a COM reference / `Release()` |
| `wil::unique_bstr`, `CComBSTR` | a `BSTR` / `SysFreeString()` |
| `CComVariant` | a `VARIANT` / `VariantClear()` |
| `std::lock_guard`, `CComCritSecLock` | a lock / unlock |
| `wil::scope_exit` | nothing / runs an arbitrary lambda on scope exit |

§1.7 puts COM smart pointers to work; Module 3 uses RAII to guarantee an `Unadvise`, and Module 7 to guarantee a `CoRevertToSelf`. In each case the pattern is identical — and in each case the bug it prevents is one this course will have shown you first.

### The SDK macros you're about to see

COM code is written in a dialect of C++ built from Windows SDK macros. They obscure more than they reveal until someone decodes them, so:

| Macro | Expands to | Why it exists |
|---|---|---|
| `STDMETHODCALLTYPE` | `__stdcall` | The calling convention COM mandates (Module 0 §0.2). Callee cleans the stack; stable across compilers. |
| `STDMETHOD(name)` | `virtual HRESULT __stdcall name` | Declaring a method **in an interface** |
| `STDMETHOD_(type, name)` | `virtual type __stdcall name` | Same, for the rare non-`HRESULT` method (`AddRef`, `Release`) |
| `STDMETHODIMP` | `HRESULT __stdcall` | **Implementing** a method in a class |
| `STDMETHODIMP_(type)` | `type __stdcall` | Implementing `AddRef`/`Release` |
| `STDAPI` | `extern "C" HRESULT __stdcall` | A **DLL export** like `DllGetClassObject` — note `extern "C"`, which suppresses name mangling so `GetProcAddress` can find it |
| `REFIID` / `REFCLSID` | `const IID&` in C++, `const IID*` in C | One declaration that compiles in both languages |
| `DEFINE_GUID(name, ...)` | a `GUID` declaration | Defines the constant **only** in the file that first `#include <initguid.h>`; elsewhere it's a declaration. Getting this wrong yields `LNK2001: unresolved external symbol IID_IFoo`. |
| `__uuidof(IFoo)` | the GUID attached by `__declspec(uuid(...))` | MSVC-specific; avoids a separate `_i.c` file. MIDL emits it via `MIDL_INTERFACE`. |
| `IID_PPV_ARGS(&p)` | `__uuidof(T), (void**)&p` | Derives the IID **from the pointer type**, so the two can't disagree. Use it everywhere. |

The last one is worth adopting immediately, because it eliminates an entire bug class:

```cpp
// Bug: the IID and the pointer type disagree. Compiles fine. Corrupts at runtime,
// because the returned vtable is IBar's but the compiler calls it as IFoo.
IFoo* p = nullptr;
pUnk->QueryInterface(IID_IBar, (void**)&p);

// Cannot go wrong: the IID is derived from the type of p.
IFoo* p = nullptr;
pUnk->QueryInterface(IID_PPV_ARGS(&p));
```

One more you'll meet in the lab: `__declspec(novtable)` on an interface declaration tells the compiler not to emit a vtable pointer initializer for that abstract type. It's a size/speed optimization that's safe on pure interfaces because they're never instantiated directly. Module 6 covers the version ATL applies to *implementation* classes, where the rules are subtler.

---

## 1.6 LAB 1.1 — Implement `IUnknown` by hand

> **Requirements**
> - **Tools:** Visual Studio with *Desktop development with C++*. Nothing else.
> - **Elevation:** not required — no registry, no registration.
> - **Bitness:** x64.
> - **Depends on:** nothing. This is the first lab.
> - **Time:** ~2 h.

**No ATL. No WRL. No smart pointers.** You need to feel the machinery once.

Create a Visual Studio **Console App (C++)** named `Lab01_Unknown`.

### `Calculator.h`

```cpp
#pragma once
#include <windows.h>
#include <unknwn.h>

// {A1B2C3D4-0001-4000-9000-000000000001}
DEFINE_GUID(IID_ICalculator,
    0xa1b2c3d4, 0x0001, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01);

// {A1B2C3D4-0002-4000-9000-000000000002}
DEFINE_GUID(IID_IAdvancedCalculator,
    0xa1b2c3d4, 0x0002, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02);

struct __declspec(novtable) ICalculator : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Add(long a, long b, long* result) = 0;
    virtual HRESULT STDMETHODCALLTYPE Subtract(long a, long b, long* result) = 0;
};

struct __declspec(novtable) IAdvancedCalculator : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Divide(long a, long b, long* result) = 0;
};
```

> Generate your own GUIDs with `New-Guid` in PowerShell. Never copy GUIDs from tutorials into shipping code.

### `Calculator.cpp`

```cpp
#include "Calculator.h"
#include <cstdio>

// Trace helper so you can SEE the reference count move.
static void Trace(const char* what, const void* obj, ULONG count, const char* who)
{
    char buf[256];
    sprintf_s(buf, "[COM] %-8s obj=%p count=%lu  (%s)\n", what, obj, count, who);
    OutputDebugStringA(buf);
    printf("%s", buf);
}

class Calculator : public ICalculator, public IAdvancedCalculator
{
    LONG m_cRef = 1;               // Born with one reference: the creator's.
    const char* m_tag;
public:
    explicit Calculator(const char* tag) : m_tag(tag)
    {
        Trace("CREATE", this, 1, m_tag);
    }
    ~Calculator()
    {
        Trace("DESTROY", this, 0, m_tag);
    }

    // ---------- IUnknown ----------
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;                       // ALWAYS null out first.

        if (riid == IID_IUnknown)
            // Ambiguous base: pick ONE branch and be consistent forever.
            *ppv = static_cast<ICalculator*>(this);
        else if (riid == IID_ICalculator)
            *ppv = static_cast<ICalculator*>(this);
        else if (riid == IID_IAdvancedCalculator)
            *ppv = static_cast<IAdvancedCalculator*>(this);
        else
            return E_NOINTERFACE;             // *ppv already NULL.

        static_cast<IUnknown*>(static_cast<ICalculator*>(this))->AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        ULONG n = InterlockedIncrement(&m_cRef);
        Trace("ADDREF", this, n, m_tag);
        return n;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG n = InterlockedDecrement(&m_cRef);
        Trace("RELEASE", this, n, m_tag);
        if (n == 0) delete this;              // Suicide at zero. Do NOT touch members after.
        return n;                             // Never touch 'this' after delete.
    }

    // ---------- ICalculator ----------
    HRESULT STDMETHODCALLTYPE Add(long a, long b, long* result) override
    {
        if (!result) return E_POINTER;
        *result = a + b;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Subtract(long a, long b, long* result) override
    {
        if (!result) return E_POINTER;
        *result = a - b;
        return S_OK;
    }

    // ---------- IAdvancedCalculator ----------
    HRESULT STDMETHODCALLTYPE Divide(long a, long b, long* result) override
    {
        if (!result) return E_POINTER;
        *result = 0;                          // Null out [out] params on failure paths.
        if (b == 0) return E_INVALIDARG;
        *result = a / b;
        return S_OK;
    }
};

// A factory function standing in for CoCreateInstance (Module 2 does the real thing).
HRESULT CreateCalculator(REFIID riid, void** ppv, const char* tag)
{
    if (!ppv) return E_POINTER;
    *ppv = nullptr;
    Calculator* p = new (std::nothrow) Calculator(tag);
    if (!p) return E_OUTOFMEMORY;
    HRESULT hr = p->QueryInterface(riid, ppv);   // count -> 2
    p->Release();                                // count -> 1 : the caller's
    return hr;
}
```

Two details worth pausing on:

- **`new` then `QI` then `Release`.** Why not just cast? Because if the caller asks for an interface the object doesn't support, this pattern correctly returns `E_NOINTERFACE` *and* destroys the object. It's the exact pattern every real class factory uses.
- **`return n;` after `delete this;`** — returning a *local* is fine. Reading `m_cRef` after `delete` would be use-after-free. This is why `Release` caches the count in a local.

### `main.cpp` — prove the rules

```cpp
#include "Calculator.h"
#include <cstdio>
#include <cassert>

HRESULT CreateCalculator(REFIID riid, void** ppv, const char* tag);

static void ProveIdentityRule()
{
    printf("\n=== Rule 1: reflexive / identity ===\n");
    ICalculator* pCalc = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&pCalc, "identity");

    IAdvancedCalculator* pAdv = nullptr;
    pCalc->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);

    printf("pCalc = %p, pAdv = %p  -> DIFFERENT addresses, same object!\n", pCalc, pAdv);

    IUnknown* u1 = nullptr; IUnknown* u2 = nullptr;
    pCalc->QueryInterface(IID_IUnknown, (void**)&u1);
    pAdv ->QueryInterface(IID_IUnknown, (void**)&u2);
    printf("u1 = %p, u2 = %p  -> MUST be equal: %s\n", u1, u2, (u1 == u2) ? "YES" : "NO!!!");
    assert(u1 == u2);

    u2->Release(); u1->Release(); pAdv->Release(); pCalc->Release();
}

static void ProveSymmetryAndTransitivity()
{
    printf("\n=== Rules 2 & 3: symmetric, transitive ===\n");
    ICalculator* a = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&a, "sym");

    IAdvancedCalculator* b = nullptr;
    assert(SUCCEEDED(a->QueryInterface(IID_IAdvancedCalculator, (void**)&b)));   // A -> B

    ICalculator* backToA = nullptr;
    assert(SUCCEEDED(b->QueryInterface(IID_ICalculator, (void**)&backToA)));     // B -> A
    printf("symmetry OK\n");

    backToA->Release(); b->Release(); a->Release();
}

static void ProveQIAddRefs()
{
    printf("\n=== Rule 5: QI always AddRefs, even for the SAME iid ===\n");
    ICalculator* p1 = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&p1, "qi-addref");   // count 1

    ICalculator* p2 = nullptr;
    p1->QueryInterface(IID_ICalculator, (void**)&p2);              // count 2  <-- watch the trace
    printf("p1 == p2 : %s, but the count went to 2.\n", (p1 == p2) ? "yes" : "no");

    p2->Release();   // 1
    p1->Release();   // 0 -> DESTROY
}

static void ShowFailureBehaviour()
{
    printf("\n=== E_NOINTERFACE and [out] param hygiene ===\n");
    ICalculator* p = nullptr;
    CreateCalculator(IID_ICalculator, (void**)&p, "fail");

    // 1. A failed QI must NULL the [out] pointer, even if it arrived holding garbage.
    IClassFactory* pCF = reinterpret_cast<IClassFactory*>(0xDEADBEEF);
    HRESULT hr = p->QueryInterface(IID_IClassFactory, (void**)&pCF);
    printf("QI  hr = 0x%08X (E_NOINTERFACE), pCF = %p (MUST be null)\n", hr, pCF);
    assert(hr == E_NOINTERFACE && pCF == nullptr);

    // 2. A FAILING METHOD must null its [out] params too - not just QueryInterface.
    IAdvancedCalculator* pAdv = nullptr;
    hr = p->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);
    assert(SUCCEEDED(hr));

    long r = 999;                        // garbage the callee is obliged to overwrite
    hr = pAdv->Divide(10, 0, &r);
    printf("Div hr = 0x%08X (E_INVALIDARG), r = %ld (MUST be 0)\n", hr, r);
    assert(hr == E_INVALIDARG && r == 0);

    pAdv->Release();
    p->Release();
}

int main()
{
    ProveIdentityRule();
    ProveSymmetryAndTransitivity();
    ProveQIAddRefs();
    ShowFailureBehaviour();
    printf("\nDone. Every CREATE above must have a matching DESTROY.\n");
    return 0;
}
```

### Exercises for Lab 1.1

1. **Run it and read the trace.** Confirm each `CREATE` has exactly one `DESTROY`.
2. **Make a leak.** Comment out `p2->Release()` in `ProveQIAddRefs`. Run. Observe: no `DESTROY`. This is what a leak looks like — *nothing happens*. That's what makes leaks hard.
3. **Make a crash.** Add a second `p1->Release()` at the end of `ProveQIAddRefs`. Run under the debugger. Note the callstack: you'll be inside `Release` reading `m_cRef` of freed memory, or crashing in `delete`. **Memorize this stack shape** — you'll see it in customer dumps.
4. **Break Rule 1.** In `QueryInterface`, change the `IID_IUnknown` branch to `static_cast<IAdvancedCalculator*>(this)`. Run `ProveIdentityRule`. Watch the assert fire. This bug is real: it happens whenever someone uses multiple inheritance and casts inconsistently.
5. **Add `IAdvancedCalculator2`** deriving from `IAdvancedCalculator` with a `Modulo` method. Note you did *not* modify `IAdvancedCalculator`. That's COM versioning.
6. **Verify the vtable.** Print `(*(void***)pCalc)[3]` and compare it in the debugger with `&Calculator::Add`.

---

## 1.7 LAB 1.2 — Smart pointers

> **Requirements**
> - **Tools:** Visual Studio C++ plus the optional component **C++ ATL for latest build tools** (for `CComPtr`). WRL (`wrl/client.h`) ships with the Windows SDK — no extra install.
> - **Elevation:** not required.
> - **Bitness:** x64.
> - **Depends on:** Lab 1.1 — you reuse `Calculator.h` and the `Calculator` implementation.
> - **Time:** ~1 h.

Manual `Release` is error-prone the moment you have early returns or exceptions. Compare four styles.

### Style A — raw pointers with a single exit

```cpp
HRESULT DoWorkRaw()
{
    HRESULT hr = S_OK;
    ICalculator* pCalc = nullptr;
    IAdvancedCalculator* pAdv = nullptr;
    long r = 0;

    hr = CreateCalculator(IID_ICalculator, (void**)&pCalc, "raw");
    if (FAILED(hr)) goto Cleanup;

    hr = pCalc->Add(2, 3, &r);
    if (FAILED(hr)) goto Cleanup;

    hr = pCalc->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);
    if (FAILED(hr)) goto Cleanup;

    hr = pAdv->Divide(10, 2, &r);
    if (FAILED(hr)) goto Cleanup;

Cleanup:
    if (pAdv)  pAdv->Release();
    if (pCalc) pCalc->Release();
    return hr;
}
```

This is the classic Windows-SDK style. `goto Cleanup` is *idiomatic and correct* here — it guarantees a single release path.

### Style B — ATL `CComPtr`

```cpp
#include <atlbase.h>

HRESULT DoWorkATL()
{
    CComPtr<ICalculator> spCalc;
    RETURN_IF_FAILED(CreateCalculator(IID_ICalculator, (void**)&spCalc, "atl"));

    long r = 0;
    RETURN_IF_FAILED(spCalc->Add(2, 3, &r));

    CComPtr<IAdvancedCalculator> spAdv;
    RETURN_IF_FAILED(spCalc->QueryInterface(IID_IAdvancedCalculator, (void**)&spAdv));
    RETURN_IF_FAILED(spAdv->Divide(10, 2, &r));

    return S_OK;   // both released automatically
}
```

**Trap:** `CComPtr::operator&` asserts if the pointer is already non-null (to prevent silent leaks). So this leaks/asserts:

```cpp
CComPtr<IFoo> sp;
GetFoo(&sp);
GetFoo(&sp);      // ASSERT / leak — sp already holds a reference
```

Use `sp.Release()` first, or a fresh variable.

### Style C — `_com_ptr_t` via `_COM_SMARTPTR_TYPEDEF`

```cpp
#include <comdef.h>
_COM_SMARTPTR_TYPEDEF(ICalculator, __uuidof(ICalculator));
// gives you ICalculatorPtr

ICalculatorPtr calc;
calc.CreateInstance(CLSID_Calculator);   // throws _com_error on failure
IAdvancedCalculatorPtr adv = calc;       // implicit QueryInterface via operator=
```

Compact, but throws `_com_error`, which surprises people. Common in `#import`-based Office automation code.

### Style D — `winrt::com_ptr` / `wil::com_ptr`

```cpp
#include <winrt/base.h>

winrt::com_ptr<ICalculator> calc;
winrt::check_hresult(CreateCalculator(IID_ICalculator, calc.put_void(), "cppwinrt"));

auto adv = calc.as<IAdvancedCalculator>();   // throws hresult_error on failure
// or:
auto advMaybe = calc.try_as<IAdvancedCalculator>();   // returns null com_ptr instead
```

Note the explicit `put()` / `put_void()` — no `operator&` ambiguity, which eliminates the `CComPtr` trap by design.

### Exercises for Lab 1.2

1. Implement all four styles against your `Calculator`. Add the ref-count trace and confirm all four produce identical `ADDREF`/`RELEASE` sequences.
2. Introduce an early `return` in the middle of Style A **without** a `goto`. Watch the leak in the trace. Then fix it.
3. Trigger the `CComPtr::operator&` assert deliberately so you recognize it.
4. Write a function that returns an interface to its caller as an `[out]` parameter, using `CComPtr` internally. Get the "transfer of ownership at the boundary" right:

```cpp
HRESULT GetCalculator(ICalculator** ppOut)
{
    if (!ppOut) return E_POINTER;
    *ppOut = nullptr;
    CComPtr<ICalculator> sp;
    RETURN_IF_FAILED(CreateCalculator(IID_ICalculator, (void**)&sp, "out"));
    *ppOut = sp.Detach();     // Detach: transfer ownership, no AddRef/Release
    return S_OK;
}
```

Using `*ppOut = sp;` instead would `AddRef` and then `sp`'s destructor would `Release` — net correct, but `Detach` states the intent. Using `*ppOut = sp.p;` without `Detach` is a **bug** (returns a pointer that's about to be released).

---

## 1.8 Reference cycles

Reference counting cannot collect cycles.

```
   Parent ──strong──► Child
      ▲                 │
      └────strong───────┘        Neither ever reaches 0.
```

Real examples:

- A document holds its views; each view holds the document.
- **Connection points**: a source `Advise`s a sink, holding a strong reference to it; the sink holds the source so it can `Unadvise`. If you forget `Unadvise`, both leak forever. (Module 5.)
- .NET RCW/CCW cycles across the interop boundary — the GC cannot see through COM ref counts.

**Solutions:**

| Technique | How |
|---|---|
| Weak back-pointer | Child stores a raw `Parent*` without `AddRef`. Parent must outlive child by construction. |
| Explicit teardown | Parent calls `child->SetParent(nullptr)` before releasing. `IOleObject::SetClientSite(nullptr)` and `IObjectWithSite::SetSite(nullptr)` exist for exactly this. |
| `IWeakReference` | WinRT: `IWeakReferenceSource::GetWeakReference`, then `Resolve` when you need it. |
| Always `Unadvise` | Pair every `Advise` with an `Unadvise`, ideally with RAII. |

**Support signal:** a process whose private bytes grow monotonically, with a matching growth in a single object type, and which *never* shrinks even at idle → suspect a cycle or a missing `Release`, not a heap fragmentation issue.

---

## 1.9 Debugging reference counts (support skills)

### Technique 1 — instrument

The trace in Lab 1.1 is the technique. In production code, do it conditionally:

```cpp
ULONG Release()
{
    ULONG n = InterlockedDecrement(&m_cRef);
#ifdef _DEBUG
    WCHAR buf[128];
    swprintf_s(buf, L"Obj %p -> %lu\n", this, n);
    OutputDebugStringW(buf);
#endif
    if (n == 0) delete this;
    return n;
}
```

View with **DebugView** (Sysinternals) or the VS Output window.

### Technique 2 — WinDbg breakpoint logging

When you don't have source, break on the vtable slot and log stacks:

```
0:000> dps @rcx L4               ; dump the vtable of an interface in rcx
0:000> bp <Release-address> "kb 8; gc"
0:000> g
```

Every `Release` now prints its callstack and continues. Compare the count of `AddRef` stacks with `Release` stacks; the surplus stack is your leak site.

More targeted — break only when the count reaches a specific value:

```
0:000> bp mymodule!CFoo::Release ".if (poi(@rcx+8) == 3) { kb } .else { gc }"
```

### Technique 3 — Application Verifier

Enable **Basics → COM** in Application Verifier for your EXE. It catches:

- `CoInitialize` mismatch (init/uninit imbalance),
- interface pointers used on the wrong apartment,
- releasing an interface more times than acquired (in many cases),
- `CoUninitialize` while objects are still alive.

Combine with **Basics → Heaps (page heap)** so a use-after-free faults *at the moment of use*, not later.

### Technique 4 — recognize the two symptom families

| Symptom | Almost certainly |
|---|---|
| Private bytes / handle count grow monotonically; process never shrinks; server DLL never unloads | **Leak** — missing `Release`, or a cycle |
| `0xC0000005` in a call through an interface pointer; vtable pointer is garbage or points to freed memory; crash is non-deterministic and moves around | **Over-release** — use-after-free |
| `0x80004002 E_NOINTERFACE` intermittently on an object that worked before | Wrong apartment / stale proxy / identity bug |
| Crash inside `Release` itself | Double-release, or releasing a pointer that was never AddRef'd |

In a dump, for the crash family: `!heap -p -a <ptr>` will often show the block as **freed** and give you the *allocation* and *free* stacks with page heap enabled. That free stack is the culprit's over-release.

---

## 1.10 LAB 1.3 — Leak hunt (support drill)

> **Requirements**
> - **Tools:** WinDbg with the Microsoft symbol server configured (`_NT_SYMBOL_PATH=srv*C:\Symbols*https://msdl.microsoft.com/download/symbols`); **Application Verifier** (Windows SDK, *Debugging Tools for Windows*); a colleague or a coin flip to inject the bug for you.
> - **Elevation:** required — Application Verifier writes machine-wide image-execution options.
> - **Bitness:** x64. Build **Release with symbols** (`/Zi` and linker `/DEBUG`).
> - **Depends on:** Lab 1.1.
> - **Time:** ~45 min, timed. Stop the clock and read the answer rather than running over.

Do this as a timed exercise. It simulates a real ticket.

1. Take your Lab 1.1 code. Have a colleague (or a coin flip) insert **one** of: a missing `Release`, an extra `Release`, or an `AddRef` without a matching `Release` — somewhere in `main.cpp`. Don't look.
2. Build in Release with symbols (`/Zi /DEBUG`).
3. Find it using only:
   - the `OutputDebugString` trace,
   - a WinDbg `bp` on `AddRef`/`Release` with `kb`,
   - Application Verifier.
4. Write down, in your `COM-Notes.md`: the symptom you saw first, the tool that localized it, and the fix.

Target: under 15 minutes. Repeat until it's boring.

---

## 1.11 LAB 1.4 — Spot the bug

> **Requirements**
> - **Tools:** none to begin with. Read these with the compiler closed; open Visual Studio afterwards only to check yourself.
> - **Elevation:** not required.
> - **Depends on:** §1.3 and §1.4. No code from the earlier labs.
> - **Time:** ~30 min. Give yourself 60 seconds per snippet.

Eight snippets of plausible, review-passing code. For each one, write down three things **before** you open the answer key:

| # | Verdict | Offending line | Fix |
|---|---|---|---|
| 1 | leak / crash / correct | | |
| 2 | … | | |

Assume `p`, `pFirst`, `pSecond`, and `pEnum` are valid pointers already owned by the enclosing scope, and that every call succeeds unless the code shows otherwise.

**Not all eight are broken.** One is correct as written, and deciding which is part of the exercise.

### Snippet 1

```cpp
IAdvancedCalculator* pAdv = nullptr;
p->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);
pAdv->AddRef();

long r = 0;
pAdv->Divide(10, 2, &r);

pAdv->Release();
```

### Snippet 2

```cpp
void Report(ICalculator* pCalc)          // [in] parameter
{
    long r = 0;
    pCalc->Add(2, 3, &r);
    printf("%ld\n", r);
    pCalc->Release();
}
```

### Snippet 3

```cpp
HRESULT Compute(ICalculator* p, long* out)
{
    IAdvancedCalculator* pAdv = nullptr;
    HRESULT hr = p->QueryInterface(IID_IAdvancedCalculator, (void**)&pAdv);
    if (FAILED(hr)) return hr;

    hr = pAdv->Divide(10, 0, out);
    if (FAILED(hr)) return hr;

    pAdv->Release();
    return S_OK;
}
```

### Snippet 4

```cpp
IUnknown* pUnk = nullptr;
pFirst ->QueryInterface(IID_IUnknown, (void**)&pUnk);
pSecond->QueryInterface(IID_IUnknown, (void**)&pUnk);

bool same = (pUnk == pFirst);
pUnk->Release();
```

### Snippet 5

```cpp
class Dashboard
{
    ICalculator* m_pCalc = nullptr;
public:
    void Attach(ICalculator* p) { m_pCalc = p; }
    void Refresh()              { long r = 0; m_pCalc->Add(1, 1, &r); }
    ~Dashboard()                { if (m_pCalc) m_pCalc->Release(); }
};
```

### Snippet 6

```cpp
void LogOnce(ICalculator* pCalc)         // [in] parameter
{
    long r = 0;
    pCalc->Add(0, 0, &r);
    printf("calc is alive: %ld\n", r);
}
```

### Snippet 7

```cpp
ICalculator* p1 = nullptr;
CreateCalculator(IID_ICalculator, (void**)&p1, "alias");

ICalculator* p2 = p1;

p2->Release();
p1->Release();
```

### Snippet 8

```cpp
IUnknown* rg[10] = {};
ULONG fetched = 0;

pEnum->Next(10, rg, &fetched);           // returns S_FALSE, fetched == 4

for (int i = 0; i < 10; ++i)
    rg[i]->Release();
```

### Answer key

<details>
<summary>Check yourself</summary>

| # | Verdict | Line | Why |
|---|---|---|---|
| 1 | **Leak** | `pAdv->AddRef();` | `QueryInterface` already incremented on your behalf. Two increments, one `Release`, count never reaches 0. Note that *every variable looks balanced* — one acquire, one release each — which is exactly why reviews wave it through. |
| 2 | **Crash** | `pCalc->Release();` | An `[in]` parameter is **borrowed**. Releasing it destroys an object the caller still holds, so the caller's next call runs on freed memory. The crash surfaces in the *caller*, later, with nothing pointing back here. |
| 3 | **Leak** | the second `if (FAILED(hr)) return hr;` | The error path returns without releasing. The happy path is correct, so this leaks only when something else has already gone wrong — i.e. precisely when the customer is already having a bad day. The real fix is `CComPtr`, not a third `Release` (§1.4). |
| 4 | **Leak** | the second `QueryInterface` | It overwrites `pUnk` and destroys the only pointer to the first reference: two references, one `Release`. There is a second bug too — `pUnk == pFirst` is **not** the identity test. §1.3 Rule 1 requires `QI(IID_IUnknown)` on *both* sides and a comparison of the two results. |
| 5 | **Crash** | `Attach` — the missing `AddRef` | Two bugs in one class. It stores the pointer for later use without taking a reference, so `Refresh` can run on a dead object; and the destructor releases a reference it never owned. Keeping a pointer beyond the current call is the defining case where you **must** `AddRef`. |
| 6 | **Correct** | — | Borrowed pointer, used only for the duration of the call: no `AddRef`, no `Release`. This is right. It is also the shape that well-meaning people "fix" into Snippet 1 or Snippet 2. Leaving a borrowed pointer alone is correct code, not an oversight. |
| 7 | **Crash** | `ICalculator* p2 = p1;` | Pointer assignment creates an **alias**, not a reference — it is not on §1.4's list of things that increment. Two `Release` calls against one reference; the second is an over-release. |
| 8 | **Crash** | `rg[i]->Release()` once `i >= 4` | `Next` returned `S_FALSE` with `fetched == 4`, so only `rg[0..3]` hold references and `rg[4..9]` are still null. Loop to `fetched`, never to the count you asked for. This is the concrete cost of testing `hr == S_OK` instead of `FAILED(hr)` (§1.5) — and the reason `[out]` arrays are zero-initialized, so the mistake is a clean null-dereference instead of a jump through garbage. |

**Scoring:** 6 or better and you're ready for Module 2. Below that, re-read §1.4's list of what increments the count, then come back tomorrow — not in ten minutes.

</details>

### The single question behind all eight

Every one of these is the same misunderstanding:

> Did this hand me a reference to **own**, or am I **borrowing** one?

**You own it** (and owe exactly one `Release`) after: `QueryInterface`, `CoCreateInstance`, a successful `[out]` interface pointer, or an explicit `AddRef` you wrote yourself.

**You own nothing** from: an `[in]` parameter, a plain pointer assignment, or a raw copy out of a member — and releasing any of them is theft from whoever does own it.

Copy those two lists into your `COM-Notes.md` now. Practically every ref-counting ticket you will ever be handed is one of these eight shapes.

---

## 1.12 Checkpoint

1. Why must `QueryInterface` `AddRef` before returning, even when it returns `this` for the same IID you already hold?
2. An object implements `IFoo` and `IBar` via multiple inheritance. Why is `pFoo == pBar` false, and how do you *correctly* test whether they're the same object?
3. What does `S_FALSE` mean? Name two APIs that return it, and describe the bug caused by testing `hr == S_OK`.
4. `Release` does `delete this` and then `return n;`. Why is returning `n` safe but returning `m_cRef` a bug?
5. You call `IEnumUnknown::Next(10, rgUnk, &fetched)` and it returns `S_FALSE` with `fetched == 4`. How many `Release` calls do you owe, and on what?
6. Your process leaks 4 KB/second. `Release` is called correctly everywhere you can see. What's the next hypothesis and how do you test it?
7. Decode `0x8007000E` and `0x80004002` from memory.

<details>
<summary>Answers</summary>

1. Because the *caller* now holds an independent reference in a new variable with its own lifetime, and will call `Release` on it. Without the `AddRef` the count would be one short and the object would die early. The rule is uniform precisely so callers never have to special-case anything.

2. With multiple inheritance the object contains two vptrs at different offsets; `static_cast<IFoo*>(this)` and `static_cast<IBar*>(this)` yield different addresses. Correct test: `QI(IID_IUnknown)` on both and compare the results.

3. `S_FALSE` is a **success** code meaning "no"/"nothing to do"/"partial." Examples: `CoInitializeEx` when the apartment is already initialized on that thread; `IEnumXxx::Next` when it returns fewer items than requested. Testing `hr == S_OK` treats these as failures — e.g. you'd call `CoUninitialize` incorrectly, or discard the 4 valid items the enumerator did return.

4. `n` is a local copy on the stack. `m_cRef` lives inside the object, which `delete this` has already freed — reading it is a use-after-free.

5. Four. `Next` `AddRef`s each element it actually returns, so you owe `Release` on `rgUnk[0..3]`. Elements 4..9 were never written and must not be touched.

6. Hypothesis: a **reference cycle**, or a leak in a component you don't own (a callback sink, a proxy, an event subscription). Test: instrument `AddRef` with stack capture (`CaptureStackBackTrace` or WinDbg `bp ... "kb;gc"`), aggregate the stacks, and diff `AddRef` stacks against `Release` stacks. Also check for `Advise` without `Unadvise`.

7. `0x8007000E` = FACILITY_WIN32 (0x7), Win32 code `0x0E` = 14 = `ERROR_OUTOFMEMORY` → `E_OUTOFMEMORY`. `0x80004002` = `E_NOINTERFACE`.

</details>

---

## 1.13 Rules to carry forward

1. Every `AddRef` gets exactly one `Release`.
2. `QueryInterface` always `AddRef`s on success; always nulls `*ppv` on failure.
3. `QI(IID_IUnknown)` is the only identity test.
4. The interface set is fixed for the object's lifetime.
5. Check `FAILED(hr)`, never `hr == S_OK`.
6. Null every `[out]` parameter on entry and on every failure path.
7. Cache the count in a local in `Release`; never touch members after `delete this`.
8. Prefer smart pointers everywhere except when you're learning.
9. Break cycles deliberately — weak back-pointers, explicit teardown, or paired `Advise`/`Unadvise`.
10. When in doubt, `AddRef`. Leaks are findable; use-after-free is not.

---

**Next: [Module 2 — Activation, registration, and the registry](02-activation-and-registry.md)**
