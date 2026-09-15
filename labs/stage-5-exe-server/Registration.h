#pragma once
#include "Calculator.h"
#include <string>

namespace Stage5Registration
{
inline std::wstring GuidText(REFGUID guid)
{
    wchar_t text[40] = {};
    StringFromGUID2(guid, text, ARRAYSIZE(text));
    return text;
}

inline HRESULT SetString(const std::wstring& subkey, PCWSTR name, const std::wstring& value)
{
    HKEY key = nullptr;
    const std::wstring path = L"Software\\Classes\\" + subkey;
    LONG status = RegCreateKeyExW(HKEY_LOCAL_MACHINE, path.c_str(), 0, nullptr,
                                 REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, nullptr, &key, nullptr);
    if (status != ERROR_SUCCESS) return HRESULT_FROM_WIN32(status);
    status = RegSetValueExW(key, name, 0, REG_SZ,
                           reinterpret_cast<const BYTE*>(value.c_str()),
                           static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t)));
    RegCloseKey(key);
    return HRESULT_FROM_WIN32(status);
}

inline HRESULT RemoveKey(const std::wstring& subkey)
{
    HKEY classes = nullptr;
    LONG status = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"Software\\Classes", 0,
                               KEY_READ | KEY_WRITE, &classes);
    if (status != ERROR_SUCCESS) return HRESULT_FROM_WIN32(status);
    status = RegDeleteTreeW(classes, subkey.c_str());
    RegCloseKey(classes);
    return status == ERROR_FILE_NOT_FOUND ? S_OK : HRESULT_FROM_WIN32(status);
}

inline HRESULT Register(bool surrogate, HMODULE module)
{
    wchar_t modulePath[MAX_PATH] = {};
    const DWORD length = GetModuleFileNameW(module, modulePath, ARRAYSIZE(modulePath));
    if (length == 0) return HRESULT_FROM_WIN32(GetLastError());
    if (length >= ARRAYSIZE(modulePath)) return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);

    const std::wstring clsid = GuidText(surrogate ? CLSID_SurrogateCalculator : CLSID_Calculator);
    const std::wstring appid = GuidText(surrogate ? APPID_CalcSurrogate : APPID_CalcSrv);
    const std::wstring classKey = L"CLSID\\" + clsid;
    const std::wstring appKey = L"AppID\\" + appid;
    const std::wstring progId = surrogate ? L"Training.Stage5.Surrogate.1" : L"Training.Stage5.Calculator.1";
    const std::wstring title = surrogate ? L"Stage 5 Calculator Surrogate" : L"Stage 5 Calculator EXE";

    struct Value { std::wstring key; PCWSTR name; std::wstring data; };
    const Value values[] = {
        {classKey, nullptr, title},
        {classKey, L"AppID", appid},
        {classKey + L"\\ProgID", nullptr, progId},
        {progId, nullptr, title},
        {progId + L"\\CLSID", nullptr, clsid},
        {appKey, nullptr, title},
        {classKey + (surrogate ? L"\\InprocServer32" : L"\\LocalServer32"), nullptr,
         surrogate ? std::wstring(modulePath) : L"\"" + std::wstring(modulePath) + L"\""}
    };
    for (const auto& value : values)
    {
        const HRESULT result = SetString(value.key, value.name, value.data);
        if (FAILED(result)) return result;
    }
    if (surrogate)
    {
        const HRESULT result = SetString(classKey + L"\\InprocServer32", L"ThreadingModel", L"Both");
        if (FAILED(result)) return result;
        return SetString(appKey, L"DllSurrogate", L"");
    }
    return S_OK;
}

inline HRESULT Unregister(bool surrogate)
{
    const std::wstring classKey = L"CLSID\\" + GuidText(surrogate ? CLSID_SurrogateCalculator : CLSID_Calculator);
    HKEY otherClass = nullptr;
    const REGSAM otherView = sizeof(void*) == 8 ? KEY_WOW64_32KEY : KEY_WOW64_64KEY;
    const LONG otherStatus = RegOpenKeyExW(HKEY_LOCAL_MACHINE,
        (L"Software\\Classes\\" + classKey).c_str(), 0, KEY_READ | otherView, &otherClass);
    if (otherStatus == ERROR_SUCCESS) RegCloseKey(otherClass);
    else if (otherStatus != ERROR_FILE_NOT_FOUND) return HRESULT_FROM_WIN32(otherStatus);
    if (otherStatus == ERROR_SUCCESS) return RemoveKey(classKey);

    const std::wstring keys[] = {
        classKey,
        L"AppID\\" + GuidText(surrogate ? APPID_CalcSurrogate : APPID_CalcSrv),
        surrogate ? L"Training.Stage5.Surrogate.1" : L"Training.Stage5.Calculator.1"
    };
    HRESULT firstFailure = S_OK;
    for (const auto& key : keys)
    {
        const HRESULT result = RemoveKey(key);
        if (FAILED(result) && SUCCEEDED(firstFailure)) firstFailure = result;
    }
    return firstFailure;
}
}