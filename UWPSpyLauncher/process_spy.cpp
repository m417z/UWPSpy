#include "stdafx.h"

#include "process_spy.h"

namespace {

bool AllowAppContainerAccess(PCWSTR path) {
    PSECURITY_DESCRIPTOR sd = nullptr;
    ULONG sd_length = 0;
    // Set read and execute access to AppContainers and Low Privilege
    // AppContainers.
    BOOL success = ConvertStringSecurityDescriptorToSecurityDescriptor(
        L"D:(A;;GRGX;;;S-1-15-2-1)(A;;GRGX;;;S-1-15-2-2)", SDDL_REVISION_1, &sd,
        &sd_length);
    if (success) {
        PACL dacl = nullptr;
        BOOL present = FALSE, defaulted = FALSE;
        success = GetSecurityDescriptorDacl(sd, &present, &dacl, &defaulted);
        if (success) {
            DWORD result = SetNamedSecurityInfo(
                (PWSTR)path, SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr,
                nullptr, dacl, nullptr);
            if (result == ERROR_SUCCESS) {
                success = true;
            }
        }

        LocalFree(sd);
    }

    return success;
}

}  // namespace

bool ProcessSpy(HWND hWnd, DWORD pid) {
    WCHAR path[MAX_PATH];
    switch (GetModuleFileName(nullptr, path, ARRAYSIZE(path))) {
        case 0:
        case ARRAYSIZE(path):
            MessageBox(hWnd, L"Failed to get module path", L"Error",
                       MB_ICONERROR);
            return false;
    }

    PWSTR filename = PathFindFileName(path);

    wcscpy_s(filename, ARRAYSIZE(path) - (filename - path), L"UWPSpy.dll");

    if (GetFileAttributes(path) == INVALID_FILE_ATTRIBUTES) {
        MessageBox(hWnd, L"UWPSpy.dll is missing", L"Error", MB_ICONERROR);
        return false;
    }

    if (!AllowAppContainerAccess(path)) {
        PCWSTR warningMsg =
            L"Failed to adjust the file permissions for UWPSpy.dll. A UWP "
            L"application running in a sandboxed AppContainer environment "
            L"might not be able to load the module.\n"
            L"\n"
            L"If UWPSpy is running from a USB flash drive or a network drive, "
            L"consider copying it to the main system drive and running it from "
            L"there.\n"
            L"\n"
            L"Proceed anyway?";
        if (MessageBox(hWnd, warningMsg, L"Warning",
                       MB_ICONWARNING | MB_YESNO) == IDNO) {
            return false;
        }
    }

    HMODULE lib = LoadLibrary(path);
    if (!lib) {
        MessageBox(hWnd, L"Failed to load UWPSpy.dll", L"Error", MB_ICONERROR);
        return false;
    }

    using start_proc_t = HRESULT(WINAPI*)(DWORD pid);

    start_proc_t start = (start_proc_t)GetProcAddress(lib, "start");
    if (!start) {
        MessageBox(hWnd, L"Failed to find spy function", L"Error",
                   MB_ICONERROR);
        return false;
    }

    HRESULT hr = start(pid);
    if (FAILED(hr)) {
        CString str = AtlGetErrorDescription(hr);
        MessageBox(hWnd, L"Failed to start spying:\n" + str, L"Error",
                   MB_ICONERROR);
        return false;
    }

    return true;
}
