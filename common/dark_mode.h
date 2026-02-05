#pragma once

#include <dwmapi.h>

namespace dark_mode {

constexpr COLORREF kDarkBg = RGB(32, 32, 32);
constexpr COLORREF kDarkText = RGB(240, 240, 240);

// Requires Win11 24H2 (26100.6899+) or 25H2 (26200.6899+) for
// DarkMode_DarkTheme support.
// https://github.com/ozone10/darkmodelib/issues/8
inline bool IsSupportedBuild() {
    static int cached = -1;
    if (cached >= 0) {
        return cached != 0;
    }

    using FnRtlGetNtVersionNumbers = void(WINAPI*)(DWORD*, DWORD*, DWORD*);
    auto fn = reinterpret_cast<FnRtlGetNtVersionNumbers>(::GetProcAddress(
        ::GetModuleHandleW(L"ntdll.dll"), "RtlGetNtVersionNumbers"));
    if (!fn) {
        cached = 0;
        return false;
    }

    DWORD major = 0, minor = 0, build = 0;
    fn(&major, &minor, &build);
    build &= ~0x80000000UL;

    // Future OS versions: assume supported.
    if (major > 10 || (major == 10 && minor > 0) ||
        (major == 10 && minor == 0 && build > 26200)) {
        cached = 1;
        return true;
    }

    // Need at least Win11 24H2 (build 26100).
    if (major < 10 || (major == 10 && minor == 0 && build < 26100)) {
        cached = 0;
        return false;
    }

    // Build is 26100 (24H2) or 26200 (25H2), check UBR.
    DWORD ubr = 0;
    DWORD ubrSize = sizeof(ubr);
    ::RegGetValue(HKEY_LOCAL_MACHINE,
                  L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", L"UBR",
                  RRF_RT_REG_DWORD, nullptr, &ubr, &ubrSize);
    cached = (ubr >= 6899) ? 1 : 0;
    return cached != 0;
}

inline const VS_FIXEDFILEINFO* GetModuleVersionInfo(HMODULE hModule) {
    HRSRC hResource =
        FindResource(hModule, MAKEINTRESOURCE(VS_VERSION_INFO), RT_VERSION);
    if (!hResource) {
        return nullptr;
    }
    HGLOBAL hGlobal = LoadResource(hModule, hResource);
    if (!hGlobal) {
        return nullptr;
    }
    void* pData = LockResource(hGlobal);
    if (!pData) {
        return nullptr;
    }
    void* pFixedFileInfo = nullptr;
    UINT uPtrLen = 0;
    if (!VerQueryValue(pData, L"\\", &pFixedFileInfo, &uPtrLen) ||
        uPtrLen == 0) {
        return nullptr;
    }
    return static_cast<const VS_FIXEDFILEINFO*>(pFixedFileInfo);
}

inline bool IsComCtlV6() {
    static int cached = -1;
    if (cached >= 0) {
        return cached != 0;
    }
    HMODULE hComCtl = ::GetModuleHandleW(L"comctl32.dll");
    if (!hComCtl) {
        cached = 0;
        return false;
    }
    auto* vi = GetModuleVersionInfo(hComCtl);
    cached = (vi && HIWORD(vi->dwFileVersionMS) >= 6) ? 1 : 0;
    return cached != 0;
}

inline bool IsSystemDarkMode() {
    if (!IsSupportedBuild() || !IsComCtlV6()) {
        return false;
    }

    HMODULE hUxtheme =
        ::LoadLibraryExW(L"uxtheme.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!hUxtheme) {
        return false;
    }
    using FnShouldAppsUseDarkMode = bool(WINAPI*)();
    auto fn = reinterpret_cast<FnShouldAppsUseDarkMode>(
        ::GetProcAddress(hUxtheme, MAKEINTRESOURCEA(132)));
    bool result = fn && fn();
    ::FreeLibrary(hUxtheme);
    return result;
}

enum class AppMode { Default, AllowDark, ForceDark, ForceLight, Max };

inline void SetPreferredAppMode(AppMode mode) {
    HMODULE hUxtheme =
        ::LoadLibraryExW(L"uxtheme.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!hUxtheme) {
        return;
    }
    using FnSetPreferredAppMode = AppMode(WINAPI*)(AppMode);
    auto fn = reinterpret_cast<FnSetPreferredAppMode>(
        ::GetProcAddress(hUxtheme, MAKEINTRESOURCEA(135)));
    if (fn) {
        fn(mode);
    }
    ::FreeLibrary(hUxtheme);
}

inline void SetDarkTitleBar(HWND hWnd, bool dark) {
    BOOL useDark = dark ? TRUE : FALSE;
    ::DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDark,
                            sizeof(useDark));
}

inline void SetControlTheme(HWND hWnd, bool dark) {
    ::SetWindowTheme(hWnd, dark ? L"DarkMode_DarkTheme" : L"Explorer", nullptr);
}

}  // namespace dark_mode
