#include <initguid.h>
#include "Calculator.h"

#include <objbase.h>
#include <cstdio>

int wmain(int argumentCount, wchar_t* arguments[])
{
    bool surrogate = false;
    bool inProcess = false;
    bool requestX86 = false;
    bool automatic = false;
    for (int index = 1; index < argumentCount; ++index)
    {
        if (wcscmp(arguments[index], L"--surrogate") == 0)
            surrogate = true;
        else if (wcscmp(arguments[index], L"--inproc") == 0)
            inProcess = true;
        else if (wcscmp(arguments[index], L"--x86") == 0)
            requestX86 = true;
        else if (wcscmp(arguments[index], L"--auto") == 0)
            automatic = true;
        else
        {
            wprintf(L"Usage: CalcSrvClient [--surrogate | --inproc] [--x86] [--auto]\n");
            return 2;
        }
    }
    if (surrogate && inProcess)
    { wprintf(L"Choose either --surrogate or --inproc.\n"); return 2; }
    if (inProcess && requestX86)
    { wprintf(L"--x86 selects an out-of-process server, not an in-process DLL.\n"); return 2; }
    if (requestX86 && !surrogate)
    { wprintf(L"This lab registers only the x86 surrogate. Use --surrogate --x86.\n"); return 2; }

    const CLSID classId = surrogate || inProcess ? CLSID_SurrogateCalculator : CLSID_Calculator;
    DWORD context = inProcess ? CLSCTX_INPROC_SERVER : CLSCTX_LOCAL_SERVER;
    if (requestX86) context |= CLSCTX_ACTIVATE_32_BIT_SERVER;
    PCWSTR host = inProcess ? L"CalcSrvClient.exe (in-process)" : surrogate ? L"dllhost.exe" : L"CalcSrv.exe";

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) { wprintf(L"CoInitializeEx: 0x%08X\n", hr); return 1; }

    ICalculator* calculator = nullptr;
    hr = CoCreateInstance(classId, nullptr, context,
                          IID_ICalculator, reinterpret_cast<void**>(&calculator));
    if (FAILED(hr))
    {
        wprintf(L"CoCreateInstance failed: 0x%08X\n", hr);
        wprintf(L"  0x80040154 = not registered   (run this lab's Setup.ps1 -Action Register)\n");
        wprintf(L"  0x80004002 = no marshaling    (register this lab's Stage5PS.dll)\n");
        wprintf(L"  0x80080005 = server failed to start\n");
        CoUninitialize();
        return 1;
    }

    long result = 0;
    hr = calculator->Add(40, 2, &result);
    wprintf(L"Add   -> hr=0x%08X  40 + 2 = %ld\n", hr, result);
    bool passed = SUCCEEDED(hr) && result == 42;

    hr = calculator->Subtract(44, 2, &result);
    wprintf(L"Sub   -> hr=0x%08X  44 - 2 = %ld\n", hr, result);
    passed = passed && SUCCEEDED(hr) && result == 42;

    wprintf(L"\nHost: %ls\n", host);
    if (!automatic)
    {
        wprintf(L"Press Enter to release the object...\n");
        (void)getchar();
    }

    calculator->Release();
    CoUninitialize();
    return passed ? 0 : 1;
}
