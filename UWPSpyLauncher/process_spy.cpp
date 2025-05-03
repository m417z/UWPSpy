#include "stdafx.h"

#include "process_spy.h"

namespace {

constexpr USHORT kCurrentArch =
#if defined(_M_IX86)
    IMAGE_FILE_MACHINE_I386;
#elif defined(_M_X64)
    IMAGE_FILE_MACHINE_AMD64;
#elif defined(_M_ARM64)
    IMAGE_FILE_MACHINE_ARM64;
#else
#error "Unsupported architecture"
#endif

USHORT GetProcessArch(HANDLE hProcess) {
    PROCESS_MACHINE_INFORMATION pmi;
    if (GetProcessInformation(hProcess, ProcessMachineTypeInfo, &pmi,
                              sizeof(pmi))) {
        return pmi.ProcessMachine;
    }

    // If GetProcessInformation(ProcessMachineTypeInfo) isn't supported, assume
    // only IMAGE_FILE_MACHINE_I386 and IMAGE_FILE_MACHINE_AMD64.

#ifndef _WIN64
    SYSTEM_INFO siSystemInfo;
    GetNativeSystemInfo(&siSystemInfo);
    if (siSystemInfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_INTEL) {
        // 32-bit machine, only one option.
        return IMAGE_FILE_MACHINE_I386;
    }
#endif  // _WIN64

    BOOL bIsWow64Process;
    if (IsWow64Process(hProcess, &bIsWow64Process) && bIsWow64Process) {
        return IMAGE_FILE_MACHINE_I386;
    }

    return IMAGE_FILE_MACHINE_AMD64;
}

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

bool ProcessSpy(HWND hWnd, DWORD pid, ProcessSpyFramework framework) {
    HANDLE hProcess =
        OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (hProcess) {
        USHORT arch = GetProcessArch(hProcess);
        CloseHandle(hProcess);
        if (arch != kCurrentArch) {
            CString message;
            message.Format(
                L"UWPSpy is not compatible with the target process "
                L"architecture.\n"
                L"\n"
                L"Target process architecture: %X\n"
                L"UWPSpy architecture: %X\n"
                L"\n"
                L"Please use the correct version of UWPSpy.\n"
                L"\n"
                L"Proceed anyway?",
                arch, kCurrentArch);
            if (MessageBox(hWnd, message, L"Warning",
                           MB_ICONWARNING | MB_YESNO) == IDNO) {
                return false;
            }
        }
    }

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

    using isDebugging_proc_t = BOOL(WINAPI*)(DWORD pid);

    isDebugging_proc_t isDebugging =
        (isDebugging_proc_t)GetProcAddress(lib, "isDebugging");
    if (isDebugging && isDebugging(pid)) {
        PCWSTR warningMsg =
            L"UWPSpy is already inspecting the target process. To start a new "
            L"inspection session, close all existing UWPSpy windows and try "
            L"again.\n"
            L"\n"
            L"Resume the existing inspection session?";
        if (MessageBox(hWnd, warningMsg, L"Warning",
                       MB_ICONWARNING | MB_YESNO) == IDNO) {
            return false;
        }
    }

    using start_proc_t = HRESULT(WINAPI*)(DWORD pid, DWORD framework);

    start_proc_t start = (start_proc_t)GetProcAddress(lib, "start");
    if (!start) {
        MessageBox(hWnd, L"Failed to find spy function", L"Error",
                   MB_ICONERROR);
        return false;
    }

    HRESULT hr = start(pid, framework);
    if (FAILED(hr)) {
        CString message =
            L"Failed to start spying:\n" + AtlGetErrorDescription(hr);

        switch (framework) {
            case kFrameworkUWP:
                if (hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) {
                    message +=
                        L"\n\nMake sure that the target process is a UWP "
                        L"application.";
                }
                break;

            case kFrameworkWinUI:
                if (hr == HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND) ||
                    hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) {
                    message +=
                        L"\n\nMake sure that the target process is a WinUI 3 "
                        L"application.";
                }
                break;
        }

        MessageBox(hWnd, message, L"Error", MB_ICONERROR);
        return false;
    }

    return true;
}
