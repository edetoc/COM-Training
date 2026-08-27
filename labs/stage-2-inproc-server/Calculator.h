#pragma once
#include <windows.h>
#include <unknwn.h>

// {A1B2C3D4-0001-4000-9000-000000000001}
DEFINE_GUID(IID_ICalculator,
    0xa1b2c3d4, 0x0001, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01);

// {A1B2C3D4-1111-4000-9000-000000000001}
DEFINE_GUID(CLSID_Calculator,
    0xa1b2c3d4, 0x1111, 0x4000, 0x90, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01);

struct __declspec(novtable) ICalculator : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Add(long a, long b, long* result) = 0;
    virtual HRESULT STDMETHODCALLTYPE Subtract(long a, long b, long* result) = 0;
};
