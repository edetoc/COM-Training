#include <initguid.h>     // must precede Calculator.h: defines the GUIDs in this TU
#include "Calculator.h"

#include <objbase.h>
#include <olectl.h>       // SELFREG_E_CLASS
#include <new>
#include <strsafe.h>

static LONG    g_cObjects = 0;      // live objects
static LONG    g_cLocks   = 0;      // LockServer count
static HMODULE g_hModule  = nullptr;

// ------------------------------------------------------------------ object
class Calculator : public ICalculator
{
    LONG m_cRef = 1;
public:
    Calculator()  { InterlockedIncrement(&g_cObjects); }
    ~Calculator() { InterlockedDecrement(&g_cObjects); }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == IID_ICalculator)
            *ppv = static_cast<ICalculator*>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&m_cRef); }
    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG n = InterlockedDecrement(&m_cRef);
        if (!n) delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE Add(long a, long b, long* r) override
    { if (!r) return E_POINTER; *r = a + b; return S_OK; }
    HRESULT STDMETHODCALLTYPE Subtract(long a, long b, long* r) override
    { if (!r) return E_POINTER; *r = a - b; return S_OK; }
};

// ----------------------------------------------------------------- factory
class CalculatorFactory : public IClassFactory
{
public:
    // The factory is a singleton with an artificially high ref count; it is
    // never destroyed, so QI/AddRef/Release are trivial.
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == IID_IClassFactory)
            *ppv = static_cast<IClassFactory*>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef()  override { return InterlockedIncrement(&g_cLocks); }
    ULONG STDMETHODCALLTYPE Release() override { return InterlockedDecrement(&g_cLocks); }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (pUnkOuter) return CLASS_E_NOAGGREGATION;   // we don't support aggregation

        Calculator* p = new (std::nothrow) Calculator();
        if (!p) return E_OUTOFMEMORY;
        HRESULT hr = p->QueryInterface(riid, ppv);
        p->Release();
        return hr;
    }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override
    {
        if (fLock) InterlockedIncrement(&g_cLocks);
        else       InterlockedDecrement(&g_cLocks);
        return S_OK;
    }
};

static CalculatorFactory g_factory;   // static instance; never freed

// ------------------------------------------------------------- DLL exports
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (!ppv) return E_POINTER;
    *ppv = nullptr;
    if (rclsid != CLSID_Calculator) return CLASS_E_CLASSNOTAVAILABLE;
    return g_factory.QueryInterface(riid, ppv);
}

STDAPI DllCanUnloadNow(void)
{
    return (g_cObjects == 0 && g_cLocks == 0) ? S_OK : S_FALSE;
}

// ------------------------------------------------------------ registration
static HRESULT SetKeyValue(HKEY root, PCWSTR subkey, PCWSTR name, PCWSTR value)
{
    HKEY hKey = nullptr;
    LONG rc = RegCreateKeyExW(root, subkey, 0, nullptr, REG_OPTION_NON_VOLATILE,
                              KEY_WRITE, nullptr, &hKey, nullptr);
    if (rc != ERROR_SUCCESS) return HRESULT_FROM_WIN32(rc);
    rc = RegSetValueExW(hKey, name, 0, REG_SZ,
                        reinterpret_cast<const BYTE*>(value),
                        static_cast<DWORD>((wcslen(value) + 1) * sizeof(WCHAR)));
    RegCloseKey(hKey);
    return HRESULT_FROM_WIN32(rc);
}

STDAPI DllRegisterServer(void)
{
    WCHAR modulePath[MAX_PATH];
    if (!GetModuleFileNameW(g_hModule, modulePath, ARRAYSIZE(modulePath)))
        return HRESULT_FROM_WIN32(GetLastError());

    WCHAR clsidStr[64];
    StringFromGUID2(CLSID_Calculator, clsidStr, ARRAYSIZE(clsidStr));

    WCHAR key[256];
    // HKCR\CLSID\{...}
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s", clsidStr);
    if (FAILED(SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, L"Calculator Component")))
        return SELFREG_E_CLASS;

    // HKCR\CLSID\{...}\InprocServer32
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s\\InprocServer32", clsidStr);
    if (FAILED(SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, modulePath)))
        return SELFREG_E_CLASS;
    if (FAILED(SetKeyValue(HKEY_CLASSES_ROOT, key, L"ThreadingModel", L"Both")))
        return SELFREG_E_CLASS;

    // HKCR\CLSID\{...}\ProgID
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s\\ProgID", clsidStr);
    SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, L"Training.Calculator.1");

    // HKCR\Training.Calculator.1\CLSID
    SetKeyValue(HKEY_CLASSES_ROOT, L"Training.Calculator.1", nullptr, L"Calculator Component");
    SetKeyValue(HKEY_CLASSES_ROOT, L"Training.Calculator.1\\CLSID", nullptr, clsidStr);

    return S_OK;
}

STDAPI DllUnregisterServer(void)
{
    WCHAR clsidStr[64];
    StringFromGUID2(CLSID_Calculator, clsidStr, ARRAYSIZE(clsidStr));
    WCHAR key[256];

    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s", clsidStr);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, key);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, L"Training.Calculator.1");
    return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hInst, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hModule = hInst;
        DisableThreadLibraryCalls(hInst);   // we don't need thread notifications
    }
    return TRUE;
}
