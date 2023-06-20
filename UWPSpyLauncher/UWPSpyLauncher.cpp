#include "stdafx.h"

#include "resource.h"

#include <aclapi.h>
#include <sddl.h>
#include <shlwapi.h>
#include <tlhelp32.h>

using start_proc_t = HRESULT(WINAPI*)(DWORD pid);

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
                success = TRUE;
            }
        }

        LocalFree(sd);
    }

    return success;
}

DWORD GetMsPaintPid() {
    DWORD pid = 0;

    PROCESSENTRY32 entry;
    entry.dwSize = sizeof(PROCESSENTRY32);

    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

    if (Process32First(snapshot, &entry)) {
        do {
             if (_wcsicmp(entry.szExeFile, L"mspaint.exe") == 0)
             //if (_wcsicmp(entry.szExeFile, L"CalculatorApp.exe") == 0)
             //if (_wcsicmp(entry.szExeFile, L"explorer.exe") == 0)
            {
                pid = entry.th32ProcessID;
                break;
            }
        } while (Process32Next(snapshot, &entry));
    }

    CloseHandle(snapshot);

    return pid;
}

int WINAPI wWinMain(_In_ HINSTANCE hInstance,
                    _In_opt_ HINSTANCE hPrevInstance,
                    _In_ LPWSTR lpCmdLine,
                    _In_ int nCmdShow) {
    WCHAR path[MAX_PATH];
    GetModuleFileName(nullptr, path, ARRAYSIZE(path));

    PWSTR filename = PathFindFileName(path);

    wcscpy_s(filename, ARRAYSIZE(path) - (filename - path), L"UWPSpy.dll");

    AllowAppContainerAccess(path);

    HMODULE lib = LoadLibrary(path);
    if (!lib) {
        return 1;
    }

    start_proc_t start = (start_proc_t)GetProcAddress(lib, "start");
    if (!start) {
        return 1;
    }

    DWORD pid = GetMsPaintPid();
    if (!pid) {
        return E_FAIL;
    }

    return start(pid);
}
