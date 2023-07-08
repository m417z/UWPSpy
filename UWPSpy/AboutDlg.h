#pragma once

#include "resource.h"

class CAboutDlg : public CDialogImpl<CAboutDlg> {
   public:
    enum { IDD = IDD_ABOUT };

   private:
    BEGIN_MSG_MAP_EX(CAboutDlg)
        MSG_WM_INITDIALOG(OnInitDialog)
        MSG_WM_NOTIFY(OnNotify)
        COMMAND_ID_HANDLER_EX(IDOK, OnOK)
        COMMAND_ID_HANDLER_EX(IDCANCEL, OnCancel)
    END_MSG_MAP()

    BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam);
    LRESULT OnNotify(int idCtrl, LPNMHDR pnmh);
    void OnOK(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl);

    CIcon m_icon, m_smallIcon;
};
