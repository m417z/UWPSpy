#include "stdafx.h"

#include "aboutdlg.h"

#include "../common/version.h"

namespace {

void OpenUrl(HWND hWnd, PCWSTR url) {
    if ((INT_PTR)ShellExecute(hWnd, L"open", url, nullptr, nullptr,
                              SW_SHOWNORMAL) <= 32) {
        CString errorMsg;
        errorMsg.Format(
            L"Failed to open link, you can copy it with Ctrl+C and open it in "
            L"a browser manually:\n\n%s",
            url);
        MessageBox(hWnd, errorMsg, nullptr, MB_ICONHAND);
    }
}

}  // namespace

BOOL CAboutDlg::OnInitDialog(CWindow wndFocus, LPARAM lInitParam) {
    // Center the dialog on the screen.
    CenterWindow();

    // Set icons.
    m_icon = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR,
                              ::GetSystemMetrics(SM_CXICON),
                              ::GetSystemMetrics(SM_CYICON));
    SetIcon(m_icon, TRUE);
    m_smallIcon = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR,
                                   ::GetSystemMetrics(SM_CXSMICON),
                                   ::GetSystemMetrics(SM_CYSMICON));
    SetIcon(m_smallIcon, FALSE);

    CString titleStr;
    GetDlgItemText(IDC_ABOUT_TITLE, titleStr);
    titleStr.Replace(L"%s", VER_FILE_VERSION_WSTR);
    SetDlgItemText(IDC_ABOUT_TITLE, titleStr);

    ApplyDarkMode();

    return TRUE;
}

void CAboutDlg::ApplyDarkMode() {
    m_darkMode = dark_mode::IsSystemDarkMode();

    dark_mode::SetDarkTitleBar(m_hWnd, m_darkMode);

    if (m_darkMode && m_darkBgBrush.IsNull()) {
        m_darkBgBrush.CreateSolidBrush(dark_mode::kDarkBg);
    }

    // Apply theme to all child controls.
    for (HWND hChild = ::GetWindow(m_hWnd, GW_CHILD); hChild;
         hChild = ::GetWindow(hChild, GW_HWNDNEXT)) {
        dark_mode::SetControlTheme(hChild, m_darkMode);
    }

    InvalidateRect(nullptr, TRUE);
}

void CAboutDlg::OnSettingChange(UINT uFlags, LPCTSTR lpszSection) {
    if (lpszSection && _wcsicmp(lpszSection, L"ImmersiveColorSet") == 0) {
        ApplyDarkMode();
    }
}

HBRUSH CAboutDlg::OnCtlColorDlg(CDCHandle dc, CWindow wnd) {
    if (m_darkMode) {
        return m_darkBgBrush;
    }
    SetMsgHandled(FALSE);
    return nullptr;
}

HBRUSH CAboutDlg::OnCtlColorStatic(CDCHandle dc, CStatic wndStatic) {
    if (m_darkMode) {
        dc.SetTextColor(dark_mode::kDarkText);
        dc.SetBkColor(dark_mode::kDarkBg);
        return m_darkBgBrush;
    }
    SetMsgHandled(FALSE);
    return nullptr;
}

HBRUSH CAboutDlg::OnCtlColorBtn(CDCHandle dc, CButton button) {
    if (m_darkMode) {
        dc.SetTextColor(dark_mode::kDarkText);
        dc.SetBkColor(dark_mode::kDarkBg);
        return m_darkBgBrush;
    }
    SetMsgHandled(FALSE);
    return nullptr;
}

void CAboutDlg::OnButtonRamenSoftware(UINT uNotifyCode,
                                      int nID,
                                      CWindow wndCtl) {
    OpenUrl(m_hWnd, L"https://ramensoftware.com/");
}

void CAboutDlg::OnButtonHomepage(UINT uNotifyCode, int nID, CWindow wndCtl) {
    OpenUrl(m_hWnd, L"https://ramensoftware.com/uwpspy");
}

void CAboutDlg::OnButtonSourceCode(UINT uNotifyCode, int nID, CWindow wndCtl) {
    OpenUrl(m_hWnd, L"https://github.com/m417z/UWPSpy");
}

void CAboutDlg::OnOK(UINT uNotifyCode, int nID, CWindow wndCtl) {
    EndDialog(nID);
}

void CAboutDlg::OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl) {
    EndDialog(nID);
}
