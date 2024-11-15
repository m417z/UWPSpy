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

bool GetLoadedDllPath(DWORD processId,
                      PCWSTR dllName,
                      WCHAR resultDllPath[MAX_PATH]) {
    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, processId);
    if (snapshot == INVALID_HANDLE_VALUE) {
        return false;
    }

    bool succeeded = false;

    MODULEENTRY32 entry{
        .dwSize = sizeof(entry),
    };
    if (Module32First(snapshot, &entry)) {
        do {
            if (_wcsicmp(entry.szModule, dllName) == 0) {
                wcscpy_s(resultDllPath, MAX_PATH, entry.szExePath);
                succeeded = true;
                break;
            }
        } while (Module32Next(snapshot, &entry));
    }

    CloseHandle(snapshot);
    return succeeded;
}

HRESULT UwpInitializeXamlDiagnostics(DWORD pid, PCWSTR dllLocation) {
    const HMODULE wux(LoadLibraryEx(L"Windows.UI.Xaml.dll", nullptr,
                                    LOAD_LIBRARY_SEARCH_SYSTEM32));
    if (!wux) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const auto ixde = reinterpret_cast<PFN_INITIALIZE_XAML_DIAGNOSTICS_EX>(
        GetProcAddress(wux, "InitializeXamlDiagnosticsEx"));
    if (!ixde) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // I didn't find a better way than trying many connections until one works.
    // Reference:
    // https://github.com/microsoft/microsoft-ui-xaml/blob/d74a0332cf0d5e58f12eddce1070fa7a79b4c2db/src/dxaml/xcp/dxaml/lib/DXamlCore.cpp#L2782
    HRESULT hr;
    for (int i = 0; i < 10000; i++) {
        WCHAR connectionName[256];
        wsprintf(connectionName, L"VisualDiagConnection%d", i + 1);

        hr = ixde(connectionName, pid, L"", dllLocation, CLSID_UWPSpyTAP,
                  nullptr);
        if (hr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) {
            break;
        }
    }

    return hr;
}

HRESULT WinUIInitializeXamlDiagnostics(DWORD pid, PCWSTR dllLocation) {
    WCHAR microsoftInternalFrameworkUdkPath[MAX_PATH];
    if (!GetLoadedDllPath(pid, L"Microsoft.Internal.FrameworkUdk.dll",
                          microsoftInternalFrameworkUdkPath)) {
        return HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    }

    const HMODULE mux(LoadLibraryEx(microsoftInternalFrameworkUdkPath, nullptr,
                                    LOAD_WITH_ALTERED_SEARCH_PATH));
    if (!mux) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const auto ixde = reinterpret_cast<PFN_INITIALIZE_XAML_DIAGNOSTICS_EX>(
        GetProcAddress(mux, "InitializeXamlDiagnosticsEx"));
    if (!ixde) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // In Windows 10 24H2, Explorer creates a new connection identifier for each
    // new window, while previous connections are eventually closed. I didn't
    // find a better way than trying many connections until one works.
    // Reference:
    // https://github.com/microsoft/microsoft-ui-xaml/blob/d74a0332cf0d5e58f12eddce1070fa7a79b4c2db/src/dxaml/xcp/dxaml/lib/DXamlCore.cpp#L2782
    HRESULT hr;
    for (int i = 0; i < 10000; i++) {
        WCHAR connectionName[256];
        wsprintf(connectionName, L"WinUIVisualDiagConnection%d", i + 1);

        hr = ixde(connectionName, pid, L"", dllLocation, CLSID_UWPSpyTAP,
                  nullptr);
        if (hr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) {
            break;
        }
    }

    return hr;
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

HRESULT WINAPI start(DWORD pid, DWORD framework) {
    AllowSetForegroundWindow(pid);

    // Calling InitializeXamlDiagnosticsEx the second time will reset the
    // existing element tree callbacks. Therefore, first check for existing
    // windows for the target process.

    struct EnumWindowsData {
        DWORD pid;
        BOOL found;
    };

    EnumWindowsData data = {pid, FALSE};

    EnumWindows(
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            EnumWindowsData& data = *reinterpret_cast<EnumWindowsData*>(lParam);

            DWORD windowPid;
            if (!GetWindowThreadProcessId(hWnd, &windowPid) ||
                windowPid != data.pid) {
                return TRUE;
            }

            WCHAR windowClass[16];
            if (!GetClassName(hWnd, windowClass, ARRAYSIZE(windowClass)) ||
                _wcsicmp(windowClass, L"UWPSpy") != 0) {
                return TRUE;
            }

            data.found = TRUE;
            PostMessage(hWnd, CMainDlg::UWM_ACTIVATE_WINDOW, 0, 0);
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&data));

    if (data.found) {
        return S_OK;
    }

    WCHAR location[MAX_PATH];
    switch (GetModuleFileName(_Module.GetModuleInstance(), location,
                              ARRAYSIZE(location))) {
        case 0:
        case ARRAYSIZE(location):
            return HRESULT_FROM_WIN32(GetLastError());
    }

    enum Framework : DWORD {
        kFrameworkUWP = 1,
        kFrameworkWinUI,
    };

    switch (framework) {
        case kFrameworkUWP:
            return UwpInitializeXamlDiagnostics(pid, location);

        case kFrameworkWinUI:
            return WinUIInitializeXamlDiagnostics(pid, location);
    }

    return E_INVALIDARG;
}

BOOL WINAPI isDebugging(DWORD pid) {
    struct EnumWindowsData {
        DWORD pid;
        BOOL found;
    };

    EnumWindowsData data = {pid, FALSE};

    EnumWindows(
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            EnumWindowsData& data = *reinterpret_cast<EnumWindowsData*>(lParam);

            DWORD windowPid;
            if (!GetWindowThreadProcessId(hWnd, &windowPid) ||
                windowPid != data.pid) {
                return TRUE;
            }

            WCHAR windowClass[16];
            if (!GetClassName(hWnd, windowClass, ARRAYSIZE(windowClass)) ||
                _wcsicmp(windowClass, L"UWPSpy") != 0) {
                return TRUE;
            }

            data.found = TRUE;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&data));

    return data.found;
}
