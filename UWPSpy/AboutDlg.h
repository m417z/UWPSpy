#pragma once

#include "resource.h"

#include "../common/dark_mode.h"

class CAboutDlg : public CDialogImpl<CAboutDlg> {
   public:
    enum { IDD = IDD_ABOUT };

   private:
    BEGIN_MSG_MAP_EX(CAboutDlg)
        MSG_WM_INITDIALOG(OnInitDialog)
        MSG_WM_CTLCOLORDLG(OnCtlColorDlg)
        MSG_WM_CTLCOLORSTATIC(OnCtlColorStatic)
        MSG_WM_CTLCOLORBTN(OnCtlColorBtn)
        MSG_WM_SETTINGCHANGE(OnSettingChange)
        COMMAND_HANDLER_EX(IDC_ABOUT_BUTTON_RAMEN_SOFTWARE, BN_CLICKED,
                           OnButtonRamenSoftware)
        COMMAND_HANDLER_EX(IDC_ABOUT_BUTTON_HOMEPAGE, BN_CLICKED,
                           OnButtonHomepage)
        COMMAND_HANDLER_EX(IDC_ABOUT_BUTTON_SOURCE_CODE, BN_CLICKED,
                           OnButtonSourceCode)
        COMMAND_ID_HANDLER_EX(IDOK, OnOK)
        COMMAND_ID_HANDLER_EX(IDCANCEL, OnCancel)
    END_MSG_MAP()

    BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam);
    HBRUSH OnCtlColorDlg(CDCHandle dc, CWindow wnd);
    HBRUSH OnCtlColorStatic(CDCHandle dc, CStatic wndStatic);
    HBRUSH OnCtlColorBtn(CDCHandle dc, CButton button);
    void OnSettingChange(UINT uFlags, LPCTSTR lpszSection);
    void ApplyDarkMode();
    void OnButtonRamenSoftware(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnButtonHomepage(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnButtonSourceCode(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnOK(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl);

    CIcon m_icon, m_smallIcon;
    bool m_darkMode = false;
    CBrush m_darkBgBrush;
};
