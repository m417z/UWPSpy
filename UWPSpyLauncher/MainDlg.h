#pragma once

#include "resource.h"

class CSortListViewCtrlCustom : public CSortListViewCtrl {
    BEGIN_MSG_MAP(CSortListViewCtrlCustom)
        MSG_WM_CHAR(ProcessListOnChar)
        CHAIN_MSG_MAP(CSortListViewCtrl)
    END_MSG_MAP()

    void ProcessListOnChar(TCHAR chChar, UINT nRepCnt, UINT nFlags);
};

class CMainDlg : public CDialogImpl<CMainDlg>, public CDialogResize<CMainDlg> {
   public:
    enum { IDD = IDD_MAINDLG };

    enum {
        TIMER_ID_END_DIALOG = 1,
    };

    enum {
        UWM_SELECT_PROCESS_IN_LIST_FROM_CURSOR = WM_APP,
    };

    BEGIN_DLGRESIZE_MAP(CMainDlg)
        DLGRESIZE_CONTROL(IDC_PROCESS_LIST, DLSZ_SIZE_X | DLSZ_SIZE_Y)
        DLGRESIZE_CONTROL(IDC_STATIC_FRAMEWORK, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDC_RADIO_UWP, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDC_RADIO_WINUI, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDC_STATIC_TIP, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDOK, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDC_REFRESH_BUTTON, DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X | DLSZ_MOVE_Y)
        DLGRESIZE_CONTROL(IDCANCEL, DLSZ_MOVE_X | DLSZ_MOVE_Y)
    END_DLGRESIZE_MAP()

   private:
    BEGIN_MSG_MAP(CMainDlg)
        CHAIN_MSG_MAP(CDialogResize<CMainDlg>)
        MSG_WM_INITDIALOG(OnInitDialog)
        MSG_WM_TIMER(OnTimer)
        COMMAND_ID_HANDLER_EX(ID_APP_ABOUT, OnAppAbout)
        COMMAND_ID_HANDLER_EX(IDOK, OnOK)
        COMMAND_ID_HANDLER_EX(IDC_REFRESH_BUTTON, OnRefresh)
        COMMAND_ID_HANDLER_EX(ID_APP_ABOUT, OnAppAbout)
        COMMAND_ID_HANDLER_EX(IDCANCEL, OnCancel)
        NOTIFY_HANDLER_EX(IDC_PROCESS_LIST, NM_DBLCLK, OnListDblClk)
        MESSAGE_HANDLER_EX(UWM_SELECT_PROCESS_IN_LIST_FROM_CURSOR,
                           SelectProcessInListFromCursor)
    END_MSG_MAP()

    BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam);
    void OnTimer(UINT_PTR nIDEvent);
    void OnOK(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnRefresh(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnAppAbout(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl);
    LRESULT OnListDblClk(LPNMHDR pnmh);
    LRESULT SelectProcessInListFromCursor(UINT uMsg,
                                          WPARAM wParam,
                                          LPARAM lParam);

    void InitProcessList();
    void LoadProcessList();
    void ProcessSpyFromList(int index);

    CIcon m_icon, m_smallIcon;
    CSortListViewCtrlCustom m_processListSort;
};
