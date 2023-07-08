#pragma once

#include "resource.h"

class CMainDlg : public CDialogImpl<CMainDlg>, public CDialogResize<CMainDlg> {
   public:
    enum { IDD = IDD_MAINDLG };

    BEGIN_DLGRESIZE_MAP(CMainDlg)
        DLGRESIZE_CONTROL(IDC_PROCESS_LIST, DLSZ_SIZE_X | DLSZ_SIZE_Y)
        DLGRESIZE_CONTROL(IDOK, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDC_REFRESH_BUTTON, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X | DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDCANCEL, DLSZ_MOVE_X | DLSZ_MOVE_Y)
    END_DLGRESIZE_MAP()

   private:
    BEGIN_MSG_MAP(CMainDlg)
        CHAIN_MSG_MAP(CDialogResize<CMainDlg>)
        MSG_WM_INITDIALOG(OnInitDialog)
        COMMAND_ID_HANDLER_EX(ID_APP_ABOUT, OnAppAbout)
        COMMAND_ID_HANDLER_EX(IDOK, OnOK)
        COMMAND_ID_HANDLER_EX(IDC_REFRESH_BUTTON, OnRefresh)
        COMMAND_ID_HANDLER_EX(ID_APP_ABOUT, OnAppAbout)
        COMMAND_ID_HANDLER_EX(IDCANCEL, OnCancel)
        NOTIFY_HANDLER_EX(IDC_PROCESS_LIST, NM_DBLCLK, OnListDblClk)
    END_MSG_MAP()

    BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam);
    void OnOK(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnRefresh(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnAppAbout(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl);
    LRESULT OnListDblClk(LPNMHDR pnmh);

    void InitProcessList();
    void LoadProcessList();

    CIcon m_icon, m_smallIcon;
    CSortListViewCtrl m_processListSort;
};
