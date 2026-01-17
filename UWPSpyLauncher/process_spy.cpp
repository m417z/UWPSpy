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
                L"UWPSpy 与目标进程架构不兼容。\n"
                L"\n"
                L"目标进程架构：%X\n"
                L"UWPSpy 架构：%X\n"
                L"\n"
                L"请使用正确版本的 UWPSpy。\n"
                L"\n"
                L"仍要继续吗？",
                arch, kCurrentArch);
            if (MessageBox(hWnd, message, L"警告",
                           MB_ICONWARNING | MB_YESNO) == IDNO) {
                return false;
            }
        }
    }

    WCHAR path[MAX_PATH];
    switch (GetModuleFileName(nullptr, path, ARRAYSIZE(path))) {
        case 0:
        case ARRAYSIZE(path):
            MessageBox(hWnd, L"无法获取模块路径", L"错误", MB_ICONERROR);
            return false;
    }

    PWSTR filename = PathFindFileName(path);

    wcscpy_s(filename, ARRAYSIZE(path) - (filename - path), L"UWPSpy.dll");

    if (GetFileAttributes(path) == INVALID_FILE_ATTRIBUTES) {
        MessageBox(hWnd, L"缺少 UWPSpy.dll", L"错误", MB_ICONERROR);
        return false;
    }

    if (!AllowAppContainerAccess(path)) {
        PCWSTR warningMsg =
            L"无法调整 UWPSpy.dll 的文件权限。运行在沙盒 AppContainer 环境中的 UWP "
            L"应用可能无法加载该模块。\n"
            L"\n"
            L"如果 UWPSpy 在 U 盘或网络驱动器上运行，建议将其复制到主系统盘后再运行。\n"
            L"\n"
            L"仍要继续吗？";
        if (MessageBox(hWnd, warningMsg, L"警告",
                       MB_ICONWARNING | MB_YESNO) == IDNO) {
            return false;
        }
    }

    HMODULE lib = LoadLibrary(path);
    if (!lib) {
        MessageBox(hWnd, L"无法加载 UWPSpy.dll", L"错误", MB_ICONERROR);
        return false;
    }

    using isDebugging_proc_t = BOOL(WINAPI*)(DWORD pid);

    isDebugging_proc_t isDebugging =
        (isDebugging_proc_t)GetProcAddress(lib, "isDebugging");
    if (isDebugging && isDebugging(pid)) {
        PCWSTR warningMsg =
            L"UWPSpy 已在检查该目标进程。若要开始新的检查会话，请关闭所有已打开的 "
            L"UWPSpy 窗口后重试。\n"
            L"\n"
            L"是否继续现有的检查会话？";
        if (MessageBox(hWnd, warningMsg, L"警告",
                       MB_ICONWARNING | MB_YESNO) == IDNO) {
            return false;
        }
    }

    using start_proc_t = HRESULT(WINAPI*)(DWORD pid, DWORD framework);

    start_proc_t start = (start_proc_t)GetProcAddress(lib, "start");
    if (!start) {
        MessageBox(hWnd, L"未找到监视函数", L"错误",
                   MB_ICONERROR);
        return false;
    }

    HRESULT hr = start(pid, framework);
    if (FAILED(hr)) {
        CString message =
            L"启动监视失败：\n" + AtlGetErrorDescription(hr);

        switch (framework) {
            case kFrameworkUWP:
                if (hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) {
                    message +=
                        L"\n\n请确认目标进程是 UWP 应用。";
                }
                break;

            case kFrameworkWinUI:
                if (hr == HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND) ||
                    hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) {
                    message +=
                        L"\n\n请确认目标进程是 WinUI 3 应用。";
                }
                break;
        }

        MessageBox(hWnd, message, L"错误", MB_ICONERROR);
        return false;
    }

    return true;
}
