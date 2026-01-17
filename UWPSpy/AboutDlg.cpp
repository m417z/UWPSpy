#include "stdafx.h"

#include "aboutdlg.h"

#include "../common/version.h"

namespace {

void OpenUrl(HWND hWnd, PCWSTR url) {
    if ((INT_PTR)ShellExecute(hWnd, L"open", url, nullptr, nullptr,
                              SW_SHOWNORMAL) <= 32) {
        CString errorMsg;
        errorMsg.Format(
            L"无法打开链接，你可以用 Ctrl+C 复制并在浏览器中手动打开：\n\n%s",
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

    return TRUE;
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
