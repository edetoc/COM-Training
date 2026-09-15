#include <initguid.h>     // must precede Calculator.h: defines the GUIDs in this TU
#include "Calculator.h"
#include "Registration.h"

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

// ------------------------------------------------------------------- entry
int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR pCmdLine, int)
{
    g_mainThreadId = GetCurrentThreadId();

    if (wcsstr(pCmdLine, L"-RegServer")   || wcsstr(pCmdLine, L"/RegServer"))
        return SUCCEEDED(Stage5Registration::Register(false, nullptr)) ? 0 : 1;
    if (wcsstr(pCmdLine, L"-UnregServer") || wcsstr(pCmdLine, L"/UnregServer"))
        return SUCCEEDED(Stage5Registration::Unregister(false)) ? 0 : 1;

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

    GUID securityAppId = APPID_CalcSrv;
    hr = CoInitializeSecurity(&securityAppId, -1, nullptr, nullptr,
                              RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IDENTIFY,
                              nullptr, EOAC_APPID, nullptr);
    if (FAILED(hr)) { CoUninitialize(); return 1; }

    MSG msg = {};
    PeekMessageW(&msg, nullptr, WM_USER, WM_USER, PM_NOREMOVE);

    static CalculatorFactory factory;
    hr = CoRegisterClassObject(
        CLSID_Calculator, &factory,
        CLSCTX_LOCAL_SERVER,
        REGCLS_MULTI_SEPARATE | REGCLS_SUSPENDED,   // suspended: don't serve calls yet
        &g_dwRegister);

    if (SUCCEEDED(hr))
    {
        hr = CoResumeClassObjects();

        if (SUCCEEDED(hr))
        {
            BOOL status;
            while ((status = GetMessageW(&msg, nullptr, 0, 0)) > 0)
            {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            if (status == -1) hr = HRESULT_FROM_WIN32(GetLastError());
        }

        CoRevokeClassObject(g_dwRegister);
    }

    CoUninitialize();
    return SUCCEEDED(hr) ? 0 : 1;
}
