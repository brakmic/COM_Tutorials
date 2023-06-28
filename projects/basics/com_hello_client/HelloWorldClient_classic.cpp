#include <iostream>
#include <objbase.h>
#include <Windows.h>
#include "../com_hello/midl/IHelloWorld.h"

int main()
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
    {
        std::cout << "Failed to initialize COM library. Error code = 0x" 
                  << std::hex << hr << std::endl;
        return hr;
    }

    CLSID clsid;
    hr = CLSIDFromProgID(L"HelloWorldLib.HelloWorld", &clsid);
    if (FAILED(hr))
    {
        std::cout << "CLSIDFromProgID() failed. Error code = 0x" 
                  << std::hex << hr << std::endl;
        return hr;
    }

    IHelloWorld* pHelloWorld;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IHelloWorld, (LPVOID*)&pHelloWorld);
    if (FAILED(hr))
    {
        std::cout << "CoCreateInstance() failed. Error code = 0x" 
                  << std::hex << hr << std::endl;
        return hr;
    }

    BSTR greeting;
    hr = pHelloWorld->SayHelloStr(&greeting);

    if (SUCCEEDED(hr)) {
        std::wcout << greeting;
        SysFreeString(greeting); // free the BSTR when you're done with it
    } else {
        std::cout << "SayHello() failed. Error code = 0x" 
                  << std::hex << hr << std::endl;
    }
	
    pHelloWorld->SayHello();

    pHelloWorld->Release();

    CoUninitialize();

    return 0;
}
