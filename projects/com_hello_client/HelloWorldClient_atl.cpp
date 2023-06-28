#include <windows.h>
#include <iostream>
#include <atlbase.h>
#include <atlstr.h>
#include "../com_hello/midl/IHelloWorld.h"

int main() {
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        std::cerr << "Failed to initialize COM library. Error code = " << hr;
        return hr;
    }

    CComPtr<IHelloWorld> pHelloWorld;
    hr = pHelloWorld.CoCreateInstance(__uuidof(HelloWorld));

    if (FAILED(hr)) {
        std::cerr << "Failed to create HelloWorld instance. Error code = " << hr;
        CoUninitialize();
        return hr;
    }

    CComBSTR name(L"John Doe");
    CComBSTR greeting;

    hr = pHelloWorld->SayHelloTo(name, &greeting);

    if (SUCCEEDED(hr)) {
        std::wcout << greeting.m_str;
    }
    else {
        std::cerr << "Failed to call SayHelloTo method. Error code = " << hr;
    }

    // Remember to uninitialize when you're done.
    CoUninitialize();
    return 0;
}
