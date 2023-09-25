#pragma once

#include "resource.h"

class CAboutDlg : public CDialogImpl<CAboutDlg> {
   public:
    enum { IDD = IDD_ABOUT };

   private:
    BEGIN_MSG_MAP_EX(CAboutDlg)
        MSG_WM_INITDIALOG(OnInitDialog)
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
    void OnButtonRamenSoftware(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnButtonHomepage(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnButtonSourceCode(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnOK(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl);

    CIcon m_icon, m_smallIcon;
};
