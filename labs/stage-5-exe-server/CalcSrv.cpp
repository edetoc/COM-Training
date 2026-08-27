#include <initguid.h>     // must precede Calculator.h: defines the GUIDs in this TU
#include "Calculator.h"

#include <objbase.h>
#include <olectl.h>
#include <new>
#include <strsafe.h>

static DWORD g_dwRegister   = 0;
static DWORD g_mainThreadId = 0;

// Server-process lifetime. When the last object and lock go away, the process
// must exit - otherwise it lingers forever and the next activation reuses a
// server nobody wanted.
static void ServerLock()
{
    CoAddRefServerProcess();
}

static void ServerUnlock()
{
    if (CoReleaseServerProcess() == 0)
    {
        // Stop accepting new activations, then break the message loop.
        CoSuspendClassObjects();
        PostThreadMessageW(g_mainThreadId, WM_QUIT, 0, 0);
    }
}

// ------------------------------------------------------------------ object
class Calculator : public ICalculator
{
    LONG m_cRef = 1;
public:
    Calculator()  { ServerLock(); }
    ~Calculator() { ServerUnlock(); }

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
    ULONG STDMETHODCALLTYPE AddRef()  override { return 2; }   // static singleton
    ULONG STDMETHODCALLTYPE Release() override { return 1; }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (pUnkOuter) return CLASS_E_NOAGGREGATION;

        auto* p = new (std::nothrow) Calculator();
        if (!p) return E_OUTOFMEMORY;
        HRESULT hr = p->QueryInterface(riid, ppv);
        p->Release();
        return hr;
    }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override
    {
        if (fLock) ServerLock();
        else       ServerUnlock();
        return S_OK;
    }
};

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

static HRESULT RegisterServer()
{
    WCHAR exePath[MAX_PATH];
    if (!GetModuleFileNameW(nullptr, exePath, ARRAYSIZE(exePath)))
        return HRESULT_FROM_WIN32(GetLastError());

    WCHAR clsidStr[64], appidStr[64], key[256];
    StringFromGUID2(CLSID_Calculator, clsidStr, ARRAYSIZE(clsidStr));
    StringFromGUID2(APPID_CalcSrv,    appidStr, ARRAYSIZE(appidStr));

    // HKCR\CLSID\{clsid}
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s", clsidStr);
    SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, L"Calculator Server");
    SetKeyValue(HKEY_CLASSES_ROOT, key, L"AppID", appidStr);      // links class -> process

    // HKCR\CLSID\{clsid}\LocalServer32   <- an EXE, not a DLL
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s\\LocalServer32", clsidStr);
    SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, exePath);

    // HKCR\CLSID\{clsid}\ProgID
    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s\\ProgID", clsidStr);
    SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, L"Training.CalcSrv.1");
    SetKeyValue(HKEY_CLASSES_ROOT, L"Training.CalcSrv.1", nullptr, L"Calculator Server");
    SetKeyValue(HKEY_CLASSES_ROOT, L"Training.CalcSrv.1\\CLSID", nullptr, clsidStr);

    // HKCR\AppID\{appid}  - process-wide settings live here (Module 7)
    StringCchPrintfW(key, ARRAYSIZE(key), L"AppID\\%s", appidStr);
    SetKeyValue(HKEY_CLASSES_ROOT, key, nullptr, L"Calculator Server");

    // HKCR\AppID\CalcSrv.exe  - lets the SCM map the EXE name back to the AppID
    SetKeyValue(HKEY_CLASSES_ROOT, L"AppID\\CalcSrv.exe", L"AppID", appidStr);

    return S_OK;
}

static HRESULT UnregisterServer()
{
    WCHAR clsidStr[64], appidStr[64], key[256];
    StringFromGUID2(CLSID_Calculator, clsidStr, ARRAYSIZE(clsidStr));
    StringFromGUID2(APPID_CalcSrv,    appidStr, ARRAYSIZE(appidStr));

    StringCchPrintfW(key, ARRAYSIZE(key), L"CLSID\\%s", clsidStr);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, key);
    StringCchPrintfW(key, ARRAYSIZE(key), L"AppID\\%s", appidStr);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, key);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, L"AppID\\CalcSrv.exe");
    RegDeleteTreeW(HKEY_CLASSES_ROOT, L"Training.CalcSrv.1");
    return S_OK;
}

// ------------------------------------------------------------------- entry
int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR pCmdLine, int)
{
    g_mainThreadId = GetCurrentThreadId();

    if (wcsstr(pCmdLine, L"-RegServer")   || wcsstr(pCmdLine, L"/RegServer"))
        return SUCCEEDED(RegisterServer())   ? 0 : 1;
    if (wcsstr(pCmdLine, L"-UnregServer") || wcsstr(pCmdLine, L"/UnregServer"))
        return SUCCEEDED(UnregisterServer()) ? 0 : 1;

    // The SCM always launches us with "-Embedding". Started any other way, the
    // user ran us by hand - do nothing rather than sit there invisibly.
    const bool embedding = wcsstr(pCmdLine, L"-Embedding") || wcsstr(pCmdLine, L"/Embedding");
    if (!embedding)
    {
        MessageBoxW(nullptr,
                    L"This is a COM server.\n\n"
                    L"Register:   CalcSrv.exe -RegServer\n"
                    L"Unregister: CalcSrv.exe -UnregServer",
                    L"CalcSrv", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) return 1;

    static CalculatorFactory factory;
    hr = CoRegisterClassObject(
        CLSID_Calculator, &factory,
        CLSCTX_LOCAL_SERVER,
        REGCLS_MULTI_SEPARATE | REGCLS_SUSPENDED,   // suspended: don't serve calls yet
        &g_dwRegister);

    if (SUCCEEDED(hr))
    {
        CoResumeClassObjects();     // NOW start accepting activations - closes a race

        MSG msg;
        while (GetMessageW(&msg, nullptr, 0, 0))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        CoRevokeClassObject(g_dwRegister);
    }

    CoUninitialize();
    return SUCCEEDED(hr) ? 0 : 1;
}
