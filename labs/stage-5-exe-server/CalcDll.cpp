#include <initguid.h>
#include "Calculator.h"
#include "Registration.h"
#include <new>

static HMODULE g_module = nullptr;
static LONG g_moduleReferences = 0;

class DllCalculator : public ICalculator
{
    LONG m_references = 1;
public:
    DllCalculator() { InterlockedIncrement(&g_moduleReferences); }
    ~DllCalculator() { InterlockedDecrement(&g_moduleReferences); }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (!result) return E_POINTER;
        *result = nullptr;
        if (iid != IID_IUnknown && iid != IID_ICalculator) return E_NOINTERFACE;
        *result = static_cast<ICalculator*>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&m_references); }
    ULONG STDMETHODCALLTYPE Release() override
    {
        const ULONG remaining = InterlockedDecrement(&m_references);
        if (!remaining) delete this;
        return remaining;
    }
    HRESULT STDMETHODCALLTYPE Add(LONG left, LONG right, LONG* result) override
    { if (!result) return E_POINTER; *result = left + right; return S_OK; }
    HRESULT STDMETHODCALLTYPE Subtract(LONG left, LONG right, LONG* result) override
    { if (!result) return E_POINTER; *result = left - right; return S_OK; }
};

class DllCalculatorFactory : public IClassFactory
{
    LONG m_references = 1;
public:
    DllCalculatorFactory() { InterlockedIncrement(&g_moduleReferences); }
    ~DllCalculatorFactory() { InterlockedDecrement(&g_moduleReferences); }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override
    {
        if (!result) return E_POINTER;
        *result = nullptr;
        if (iid != IID_IUnknown && iid != IID_IClassFactory) return E_NOINTERFACE;
        *result = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&m_references); }
    ULONG STDMETHODCALLTYPE Release() override
    {
        const ULONG remaining = InterlockedDecrement(&m_references);
        if (!remaining) delete this;
        return remaining;
    }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID iid, void** result) override
    {
        if (!result) return E_POINTER;
        *result = nullptr;
        if (outer) return CLASS_E_NOAGGREGATION;
        auto* calculator = new (std::nothrow) DllCalculator();
        if (!calculator) return E_OUTOFMEMORY;
        const HRESULT status = calculator->QueryInterface(iid, result);
        calculator->Release();
        return status;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override
    {
        if (lock) InterlockedIncrement(&g_moduleReferences);
        else InterlockedDecrement(&g_moduleReferences);
        return S_OK;
    }
};

STDAPI DllGetClassObject(REFCLSID classId, REFIID iid, void** result)
{
    if (!result) return E_POINTER;
    *result = nullptr;
    if (classId != CLSID_SurrogateCalculator) return CLASS_E_CLASSNOTAVAILABLE;
    auto* factory = new (std::nothrow) DllCalculatorFactory();
    if (!factory) return E_OUTOFMEMORY;
    const HRESULT status = factory->QueryInterface(iid, result);
    factory->Release();
    return status;
}

STDAPI DllCanUnloadNow()
{
    return InterlockedCompareExchange(&g_moduleReferences, 0, 0) == 0 ? S_OK : S_FALSE;
}

STDAPI DllRegisterServer() { return Stage5Registration::Register(true, g_module); }
STDAPI DllUnregisterServer() { return Stage5Registration::Unregister(true); }

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_module = instance;
        DisableThreadLibraryCalls(instance);
    }
    return TRUE;
}