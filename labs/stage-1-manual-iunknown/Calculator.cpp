// initguid.h must come first, and in exactly ONE translation unit: it turns the
// DEFINE_GUID macros in Calculator.h into real definitions rather than externs.
#include <initguid.h>
#include "Calculator.h"

#include <cstdio>
#include <new>

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
