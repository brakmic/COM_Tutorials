#include "HelloWorld.h"
#include "HelloWorldFactory.h"


HelloWorldFactory::HelloWorldFactory() : m_cRef(1) {}

HRESULT __stdcall HelloWorldFactory::QueryInterface(const IID& riid, void** ppv)
{
    // Check if the requested interface ID either belongs to IUnknown or IClassFactory.
    if (riid == IID_IUnknown || riid == IID_IClassFactory)
    {
        // If it does, we cast 'this' to an IClassFactory pointer, meaning we consider this instance as an IClassFactory.
        *ppv = static_cast<IClassFactory*>(this);
    }
    else
    {
        // If it doesn't, then we don't know about this interface. Return an appropriate error code.
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    
    // Here, we cast the pointer to IUnknown before calling AddRef.
    // This is to ensure that we increase the reference count of the IUnknown interface of the object.
    // It's a standard practice because every COM object must at least support the IUnknown interface,
    // and IUnknown->AddRef is used to manage the lifecycle of the COM object.
    // Even though we could call AddRef directly on the IClassFactory pointer,
    // we do it on the IUnknown pointer to follow this standard practice.
    // This becomes particularly important with COM objects implementing multiple interfaces,
    // as each interface implementation has its own IUnknown with reference counting.
    reinterpret_cast<IUnknown*>(*ppv)->AddRef();
    
    // Return S_OK to indicate success.
    return S_OK;
}

ULONG __stdcall HelloWorldFactory::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

ULONG __stdcall HelloWorldFactory::Release()
{
    ULONG ulRefCount = InterlockedDecrement(&m_cRef);
    if (0 == m_cRef)
    {
        delete this;
    }
    return ulRefCount;
}

HRESULT __stdcall HelloWorldFactory::CreateInstance(IUnknown* pUnkOuter, const IID& riid, void** ppv)
{
    // Ensure the outer unknown (used for aggregation) is NULL. Aggregation is not supported in this example.
    if (pUnkOuter != NULL)
    {
        return CLASS_E_NOAGGREGATION;
    }

    // Create a new instance of HelloWorld
    HelloWorld* pHelloWorld = new HelloWorld;
    if (pHelloWorld == NULL) // Check if memory allocation was successful
    {
        return E_OUTOFMEMORY;
    }

    // Attempt to obtain a pointer to the requested interface by calling the object's QueryInterface()
    HRESULT hr = pHelloWorld->QueryInterface(riid, ppv);
    if (FAILED(hr))
    {
        // QueryInterface() failed, delete the HelloWorld object because no one else has a reference to clean it up
        delete pHelloWorld;
    }
    else
    {
        // QueryInterface() succeeded, meaning the object's reference count has been incremented.
        // Since CreateInstance() itself does not need a reference to the object beyond this point,
        // it must release its reference. Not doing so would lead to a memory leak, as the object
        // would never be deleted even when all other references are released by the client(s).
        pHelloWorld->Release();
    }
    return hr;
}

HRESULT __stdcall HelloWorldFactory::LockServer(BOOL fLock)
{
    // Do nothing. We are not implementing server locking in this simple example.
    return S_OK;
}
