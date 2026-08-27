#include <initguid.h>     // must precede Calculator.h: defines the GUIDs in this TU
#include "Calculator.h"

#include <objbase.h>
#include <cstdio>

int wmain()
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) { wprintf(L"CoInitializeEx: 0x%08X\n", hr); return 1; }

    ICalculator* pCalc = nullptr;
    hr = CoCreateInstance(CLSID_Calculator, nullptr, CLSCTX_INPROC_SERVER,
                          IID_ICalculator, reinterpret_cast<void**>(&pCalc));
    if (SUCCEEDED(hr))
    {
        long r = 0;
        pCalc->Add(40, 2, &r);
        wprintf(L"40 + 2 = %ld\n", r);
        pCalc->Release();
    }
    else
    {
        wprintf(L"CoCreateInstance failed: 0x%08X\n", hr);
    }

    // Also demonstrate ProgID -> CLSID resolution.
    CLSID fromProgID{};
    hr = CLSIDFromProgID(L"Training.Calculator.1", &fromProgID);
    wprintf(L"CLSIDFromProgID: 0x%08X\n", hr);

    CoUninitialize();
    return (SUCCEEDED(hr)) ? 0 : 1;
}
