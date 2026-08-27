#include <initguid.h>
#include "Calculator.h"

#include <objbase.h>
#include <cstdio>

// A client for the out-of-proc server. Note CLSCTX_LOCAL_SERVER: we are
// explicitly refusing an in-proc answer, so a mis-registration fails loudly
// instead of silently giving us the wrong kind of server (§2.4, Support fact #3).
int wmain()
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) { wprintf(L"CoInitializeEx: 0x%08X\n", hr); return 1; }

    ICalculator* pCalc = nullptr;
    hr = CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_LOCAL_SERVER,
                          IID_ICalculator, reinterpret_cast<void**>(&pCalc));
    if (FAILED(hr))
    {
        wprintf(L"CoCreateInstance failed: 0x%08X\n", hr);
        wprintf(L"  0x80040154 = not registered   (run CalcSrv.exe -RegServer, elevated)\n");
        wprintf(L"  0x80004002 = no marshaling    (register Stage 3's CalcPS.dll)\n");
        wprintf(L"  0x80080005 = server failed to start\n");
        CoUninitialize();
        return 1;
    }

    long r = 0;
    hr = pCalc->Add(40, 2, &r);
    wprintf(L"Add   -> hr=0x%08X  40 + 2 = %ld\n", hr, r);

    hr = pCalc->Subtract(44, 2, &r);
    wprintf(L"Sub   -> hr=0x%08X  44 - 2 = %ld\n", hr, r);

    wprintf(L"\nCalcSrv.exe should be visible in Task Manager right now.\n");
    wprintf(L"Press Enter to release the object and let the server exit...\n");
    (void)getchar();

    pCalc->Release();
    CoUninitialize();
    return 0;
}
