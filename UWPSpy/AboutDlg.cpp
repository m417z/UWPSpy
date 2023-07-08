#include "stdafx.h"

#include "aboutdlg.h"

#include "../common/version.h"

namespace {

void OpenUrl(HWND hWnd, PCWSTR url) {
    if ((INT_PTR)ShellExecute(hWnd, L"open", url, nullptr, nullptr,
                              SW_SHOWNORMAL) <= 32) {
        MessageBox(hWnd, L"Failed to open link", nullptr, MB_ICONHAND);
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

LRESULT CAboutDlg::OnNotify(int idCtrl, LPNMHDR pnmh) {
    switch (pnmh->idFrom) {
        case IDC_ABOUT_CONTENT_SYSLINK:
            switch (pnmh->code) {
                case NM_CLICK:
                case NM_RETURN:
                    OpenUrl(m_hWnd, ((PNMLINK)pnmh)->item.szUrl);
                    break;
            }
            break;
    }

    SetMsgHandled(FALSE);
    return 0;
}

void CAboutDlg::OnOK(UINT uNotifyCode, int nID, CWindow wndCtl) {
    EndDialog(nID);
}

void CAboutDlg::OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl) {
    EndDialog(nID);
}
