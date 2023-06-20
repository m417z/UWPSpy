#include "stdafx.h"

#include "MainDlg.h"
#include "tap.hpp"

using PFN_INITIALIZE_XAML_DIAGNOSTICS_EX =
    decltype(&InitializeXamlDiagnosticsEx);

CAppModule _Module;

namespace {

ATOM RegisterDialogClass(LPCTSTR lpszClassName, HINSTANCE hInstance) {
    WNDCLASS wndcls;
    ::GetClassInfo(hInstance, MAKEINTRESOURCE(32770), &wndcls);

    // Set our own class name.
    wndcls.lpszClassName = lpszClassName;

    // Just register the class.
    return ::RegisterClass(&wndcls);
}

}  // namespace

BOOL WINAPI DllMain(HINSTANCE hinstDLL,  // handle to DLL module
                    DWORD fdwReason,     // reason for calling function
                    LPVOID lpvReserved)  // reserved
{
    switch (fdwReason) {
        case DLL_PROCESS_ATTACH: {
            HRESULT hRes = _Module.Init(nullptr, hinstDLL);
            ATLASSERT(SUCCEEDED(hRes));

            RegisterDialogClass(L"UWPSpy", hinstDLL);
        } break;

        case DLL_THREAD_ATTACH:
        case DLL_THREAD_DETACH:
            break;

        case DLL_PROCESS_DETACH:
            UnregisterClass(L"UWPSpy", hinstDLL);

            _Module.Term();
            break;
    }

    return TRUE;
}

HRESULT WINAPI start(DWORD pid) {
    // Calling InitializeXamlDiagnosticsEx the second time will reset the
    // existing element tree callbacks. Therefore, first check for an existing
    // window for the target process.
    auto windowTitle = std::format(L"UWPSpy - PID: {}", pid);
    HWND existingWindow = FindWindow(L"UWPSpy", windowTitle.c_str());
    if (existingWindow) {
        PostMessage(existingWindow, CMainDlg::UWM_ACTIVATE_WINDOW, 0, 0);
        return S_OK;
    }

    WCHAR location[MAX_PATH];
    switch (GetModuleFileName(_Module.m_hInst, location, ARRAYSIZE(location))) {
        case 0:
        case ARRAYSIZE(location):
            return HRESULT_FROM_WIN32(GetLastError());
    }

    const HMODULE wux(LoadLibraryEx(L"Windows.UI.Xaml.dll", nullptr,
                                    LOAD_LIBRARY_SEARCH_SYSTEM32));
    if (!wux) [[unlikely]] {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const auto ixde = reinterpret_cast<PFN_INITIALIZE_XAML_DIAGNOSTICS_EX>(
        GetProcAddress(wux, "InitializeXamlDiagnosticsEx"));
    if (!ixde) [[unlikely]] {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const HRESULT hr2 = ixde(L"VisualDiagConnection1", pid, nullptr, location,
                             CLSID_ExplorerTAP, nullptr);
    if (FAILED(hr2)) [[unlikely]] {
        return hr2;
    }

    return S_OK;
}
