#include <windows.h>
#include <iostream>
#include "../com_hello/midl/IHelloWorld.h"

int main() {
    HRESULT hr;
    IHelloWorld *pHelloWorld = NULL;

    hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        std::cerr << "Failed to initialize COM library. Error code = " << hr;
        return hr;
    }

    CLSID clsid;
    hr = CLSIDFromProgID(L"HelloWorldLib.HelloWorld", &clsid);
    if (FAILED(hr)) {
        std::cerr << "CLSIDFromProgID error: " << hr;
        CoUninitialize();
        return hr;
    }

    hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, __uuidof(IHelloWorld), (void**)&pHelloWorld);
    if (FAILED(hr)) {
        std::cerr << "Failed to create HelloWorld instance. Error code = " << hr;
        CoUninitialize();
        return hr;
    }

    BSTR name = SysAllocString(L"John Doe");
    BSTR greeting = SysAllocString(L"");

    hr = pHelloWorld->SayHelloTo(name, &greeting);

    if (SUCCEEDED(hr)) {
        std::wcout << greeting;
    }
    else {
        std::cerr << "Failed to call SayHelloTo method. Error code = " << hr;
    }

    // Call SayHello and output the greeting
    BSTR genericGreeting = SysAllocString(L"");

    hr = pHelloWorld->SayHelloStr(&genericGreeting);

    if (SUCCEEDED(hr)) {
        std::wcout << genericGreeting;
    }
    else {
        std::cerr << "Failed to call SayHello method. Error code = " << hr;
    }
	
    pHelloWorld->SayHello();

    SysFreeString(name);
    SysFreeString(greeting);
    SysFreeString(genericGreeting); // don't forget to free the BSTR allocated for genericGreeting

    pHelloWorld->Release();
    CoUninitialize();

    return 0;
}
