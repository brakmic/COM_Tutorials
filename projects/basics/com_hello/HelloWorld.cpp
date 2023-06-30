#include "HelloWorld.h"
#include <iostream>
#include <string>

// Constructor to initialize the reference count
HelloWorld::HelloWorld() : m_cRef(1) {}

// QueryInterface allows a client to obtain pointers to other interfaces on a given object
HRESULT __stdcall HelloWorld::QueryInterface(const IID& riid, void** ppv)
{
    // If the requested interface is IUnknown, IDispatch, or IHelloWorld
    // we increment the ref count and return a pointer to it
    if (riid == IID_IUnknown || riid == IID_IDispatch || riid == IID_IHelloWorld)
    {
        *ppv = static_cast<IHelloWorld*>(this);
    }
    else
    {
        // If the requested interface does not exist, return an error
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    // Increment the reference count before the pointer is returned
    reinterpret_cast<IUnknown*>(*ppv)->AddRef();
    return S_OK;
}

// AddRef method increments the reference count for an object
ULONG __stdcall HelloWorld::AddRef()
{
    // Use interlocked increment for thread safety
    return InterlockedIncrement(&m_cRef);
}

// Release method decrements the reference count for an object
ULONG __stdcall HelloWorld::Release()
{
    // Use interlocked decrement for thread safety
    ULONG ulRefCount = InterlockedDecrement(&m_cRef);
    // If reference count is 0, delete the object
    if (0 == m_cRef)
    {
        delete this;
    }
    return ulRefCount;
}

// GetTypeInfoCount method retrieves the number of type information interfaces that an object provides
HRESULT __stdcall HelloWorld::GetTypeInfoCount(UINT* pctinfo)
{
    // We're not providing any type info, so set the count to 0
    *pctinfo = 0;
    return S_OK;
}

// GetTypeInfo retrieves the type information for an object
HRESULT __stdcall HelloWorld::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    // We're not providing any type info, so set the out parameter to NULL and return an error
    *ppTInfo = NULL;
    return DISP_E_BADINDEX;
}

// GetIDsOfNames method maps a set of names to a corresponding set of dispatch identifiers
HRESULT __stdcall HelloWorld::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
    // Map the method names to dispatch IDs
    // If the name matches, set the ID and return S_OK
    if (_wcsicmp(*rgszNames, L"SayHello") == 0)
    {
        *rgDispId = 1;
        return S_OK;
    }
    else if (_wcsicmp(*rgszNames, L"SayHelloStr") == 0)
    {
        *rgDispId = 2;
        return S_OK;
    }
    else if (_wcsicmp(*rgszNames, L"SayHelloTo") == 0)
    {
        *rgDispId = 3;
        return S_OK;
    }
    else
    {
        // If the name does not match, set the ID to unknown and return an error
        *rgDispId = DISPID_UNKNOWN;
        return DISP_E_UNKNOWNNAME;
    }
}

// Invoke provides access to properties and methods exposed by an object
HRESULT __stdcall HelloWorld::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    // Dispatch the call to the correct method based on the dispatch ID
    switch (dispIdMember)
    {
        case 1: // SayHello
        {
            return SayHello();
        }
        case 2: // SayHelloStr
        {
            BSTR greeting;
            HRESULT hr = SayHelloStr(&greeting);

            // If successful, store the return value
            if (SUCCEEDED(hr))
            {
                pVarResult->vt = VT_BSTR;
                pVarResult->bstrVal = greeting;
            }

            return hr;
        }
        case 3: // SayHelloTo
        {
            // Check for correct number and type of arguments
            if (pDispParams->cArgs != 1 || pDispParams->rgvarg[0].vt != VT_BSTR)
                return DISP_E_TYPEMISMATCH;

            BSTR greeting;
            HRESULT hr = SayHelloTo(pDispParams->rgvarg[0].bstrVal, &greeting);

            // If successful, store the return value
            if (SUCCEEDED(hr))
            {
                pVarResult->vt = VT_BSTR;
                pVarResult->bstrVal = greeting;
            }

            return hr;
        }
        default: // Unknown dispatch ID
        {
            return DISP_E_MEMBERNOTFOUND;
        }
    }
}

HRESULT __stdcall HelloWorld::SayHello()
{
    std::cout << "Hello, World!\n";
    return S_OK;
}

HRESULT __stdcall HelloWorld::SayHelloStr(BSTR* greeting)
{
    *greeting = SysAllocString(L"Hello, World!\n");
    if (*greeting == NULL)
    {
        return E_OUTOFMEMORY;
    }
    return S_OK;
}

HRESULT __stdcall HelloWorld::SayHelloTo(BSTR name, BSTR* greeting)
{
    try
    {
        // Prepare the greeting
        std::wstring nameWStr(name);
        std::wstring greetingWStr = L"Hello, " + nameWStr + L"!\n";

        // Allocate a BSTR for the greeting
        *greeting = SysAllocStringLen(greetingWStr.c_str(), greetingWStr.length());

        if (*greeting == NULL)
        {
            return E_OUTOFMEMORY;
        }

        return S_OK;
    }
    catch(...)
    {
        return E_FAIL;
    }
}
