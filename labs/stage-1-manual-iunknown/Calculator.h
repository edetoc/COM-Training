#pragma once
#include <windows.h>
#include <unknwn.h>

// {A1B2C3D4-0001-4000-9000-000000000001}
DEFINE_GUID(IID_ICalculator,
    0xa1b2c3d4, 0x0001, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01);

// {A1B2C3D4-0002-4000-9000-000000000002}
DEFINE_GUID(IID_IAdvancedCalculator,
    0xa1b2c3d4, 0x0002, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02);

// __declspec(uuid(...)) attaches the GUID to the type, which is what makes
// __uuidof(ICalculator) work. MIDL does this for you via MIDL_INTERFACE.
struct __declspec(uuid("A1B2C3D4-0001-4000-9000-000000000001"))
       __declspec(novtable) ICalculator : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Add(long a, long b, long* result) = 0;
    virtual HRESULT STDMETHODCALLTYPE Subtract(long a, long b, long* result) = 0;
};

struct __declspec(uuid("A1B2C3D4-0002-4000-9000-000000000002"))
       __declspec(novtable) IAdvancedCalculator : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Divide(long a, long b, long* result) = 0;
};

// Stands in for CoCreateInstance until Module 2 introduces the real thing.
HRESULT CreateCalculator(REFIID riid, void** ppv, const char* tag);
