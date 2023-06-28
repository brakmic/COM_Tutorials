#pragma once
#include "./midl/IHelloWorld.h"

class HelloWorld : public IHelloWorld
{
    long m_cRef;

public:
    HelloWorld();

    // IUnknown methods
    HRESULT __stdcall QueryInterface(const IID& riid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IDispatch methods
    HRESULT __stdcall GetTypeInfoCount(UINT* pctinfo);
    HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
    HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
    HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

    // IHelloWorld methods
	  HRESULT __stdcall SayHello();
    HRESULT __stdcall SayHelloStr(BSTR* greeting);
    HRESULT __stdcall SayHelloTo(BSTR name, BSTR* greeting);
};
