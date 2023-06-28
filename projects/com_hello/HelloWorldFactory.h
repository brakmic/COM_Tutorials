#pragma once
#include <Windows.h>

class HelloWorldFactory : public IClassFactory
{
    long m_cRef;

public:
    HelloWorldFactory();

    HRESULT __stdcall QueryInterface(const IID& riid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();
    HRESULT __stdcall CreateInstance(IUnknown* pUnkOuter, const IID& riid, void** ppv);
    HRESULT __stdcall LockServer(BOOL fLock);
};
