#include <Windows.h>
#include <winerror.h>
#include <shlwapi.h>
#include "./midl/IHelloWorld.h"
#include "HelloWorldFactory.h"

LONG dllRefCount = 0;
EXTERN_C IMAGE_DOS_HEADER __ImageBase;

#define SELFREG_E_CLASS HRESULT_FROM_WIN32(ERROR_CANNOT_MAKE)

extern "C" BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    return TRUE;
}

extern "C" HRESULT __stdcall DllGetClassObject(const CLSID &clsid, const IID &iid, void **ppv)
{
    if (clsid != CLSID_HelloWorld) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    
    HelloWorldFactory *factory = new HelloWorldFactory();
    if (factory == NULL) {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = factory->QueryInterface(iid, ppv);
    if (SUCCEEDED(hr)) {
        InterlockedIncrement(&dllRefCount);
    }

    factory->Release();
    return hr;
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    return (dllRefCount == 0) ? S_OK : S_FALSE;
}

extern "C" HRESULT __stdcall DllRegisterServer()
{
    HKEY hKey;
    LONG lResult;

    // Register CLSID
    lResult = RegCreateKeyExW(HKEY_CLASSES_ROOT, L"CLSID\\{DC0F3891-93F3-42E9-A117-729B4F3C775A}", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        return SELFREG_E_CLASS;
    }
    
    lResult = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)L"HelloWorld", (wcslen(L"HelloWorld") + 1) * sizeof(WCHAR));
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    // Register ProgID under CLSID
    HKEY hSubKey;
    lResult = RegCreateKeyExW(hKey, L"ProgID", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    lResult = RegSetValueExW(hSubKey, NULL, 0, REG_SZ, (const BYTE*)L"HelloWorldLib.HelloWorld", (wcslen(L"HelloWorldLib.HelloWorld") + 1) * sizeof(WCHAR));
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hSubKey);
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    RegCloseKey(hSubKey);

    // Register InprocServer32
    lResult = RegCreateKeyExW(hKey, L"InprocServer32", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    // Set path
    WCHAR path[MAX_PATH];
    GetModuleFileNameW((HMODULE)&__ImageBase, path, MAX_PATH);
    lResult = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)path, (wcslen(path) + 1) * sizeof(WCHAR));
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    // Set ThreadingModel
    lResult = RegSetValueExW(hKey, L"ThreadingModel", 0, REG_SZ, (const BYTE*)L"Apartment", (wcslen(L"Apartment") + 1) * sizeof(WCHAR));
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    RegCloseKey(hKey);

    // Register ProgID
    lResult = RegCreateKeyExW(HKEY_CLASSES_ROOT, L"HelloWorldLib.HelloWorld", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        return SELFREG_E_CLASS;
    }

    lResult = RegCreateKeyExW(hKey, L"CLSID", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    lResult = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)L"{DC0F3891-93F3-42E9-A117-729B4F3C775A}", (wcslen(L"{DC0F3891-93F3-42E9-A117-729B4F3C775A}") + 1) * sizeof(WCHAR));
    if (lResult != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return SELFREG_E_CLASS;
    }

    RegCloseKey(hKey);
    return S_OK;
}

extern "C" HRESULT __stdcall DllUnregisterServer()
{
    LONG lResult;

    // Unregister ProgID under CLSID
    lResult = SHDeleteKeyW(HKEY_CLASSES_ROOT, L"CLSID\\{DC0F3891-93F3-42E9-A117-729B4F3C775A}\\ProgID");
    if (lResult != ERROR_SUCCESS) {
        return SELFREG_E_CLASS;
    }

    // Unregister InprocServer32 under CLSID
    lResult = SHDeleteKeyW(HKEY_CLASSES_ROOT, L"CLSID\\{DC0F3891-93F3-42E9-A117-729B4F3C775A}\\InprocServer32");
    if (lResult != ERROR_SUCCESS) {
        return SELFREG_E_CLASS;
    }

    // Unregister CLSID
    lResult = SHDeleteKeyW(HKEY_CLASSES_ROOT, L"CLSID\\{DC0F3891-93F3-42E9-A117-729B4F3C775A}");
    if (lResult != ERROR_SUCCESS) {
        return SELFREG_E_CLASS;
    }

    // Unregister CLSID under ProgID
    lResult = SHDeleteKeyW(HKEY_CLASSES_ROOT, L"HelloWorldLib.HelloWorld\\CLSID");
    if (lResult != ERROR_SUCCESS) {
        return SELFREG_E_CLASS;
    }

    // Unregister ProgID
    lResult = SHDeleteKeyW(HKEY_CLASSES_ROOT, L"HelloWorldLib.HelloWorld");
    if (lResult != ERROR_SUCCESS) {
        return SELFREG_E_CLASS;
    }

    return S_OK;
}


