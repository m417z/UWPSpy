#include "stdafx.h"

#include "MainDlg.h"

#include "../common/version.h"
#include "process_spy.h"
#include "resource.h"

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

void CSortListViewCtrlCustom::ProcessListOnChar(TCHAR chChar,
                                                UINT nRepCnt,
                                                UINT nFlags) {
    if (chChar == 4) {
        // Ctrl+D.
        GetParent().SendMessage(
            CMainDlg::UWM_SELECT_PROCESS_IN_LIST_FROM_CURSOR);
    } else {
        SetMsgHandled(FALSE);
    }
}

BOOL CMainDlg::OnInitDialog(CWindow wndFocus, LPARAM lInitParam) {
    CenterWindow();

    m_icon.LoadIconMetric(IDR_MAINFRAME, LIM_LARGE);
    SetIcon(m_icon, TRUE);
    m_smallIcon.LoadIconMetric(IDR_MAINFRAME, LIM_SMALL);
    SetIcon(m_smallIcon, FALSE);

    DlgResize_Init();

    InitProcessList();
    LoadProcessList();

    CheckDlgButton(IDC_RADIO_UWP, BST_CHECKED);

    ::SetWindowSubclass(GetDlgItem(IDC_PROCESS_LIST), ListViewSubclassProc, 0,
                        (DWORD_PTR)this);

    HWND hGrip = GetDlgItem(ATL_IDW_STATUS_BAR);
    if (hGrip) {
        ::SetWindowSubclass(hGrip, GripSubclassProc, 0, (DWORD_PTR)this);
    }

    ApplyDarkMode();

    return TRUE;
}

void CMainDlg::OnTimer(UINT_PTR nIDEvent) {
    switch (nIDEvent) {
        case TIMER_ID_END_DIALOG:
            KillTimer(nIDEvent);
            EndDialog(0);
            break;
    }
}

void CMainDlg::OnOK(UINT uNotifyCode, int nID, CWindow wndCtl) {
    int selectedIndex = m_processListSort.GetSelectedIndex();
    if (selectedIndex == -1) {
        MessageBox(L"Select a process from the list to spy on",
                   L"No process selected", MB_ICONWARNING);
        return;
    }

    ProcessSpyFromList(selectedIndex);
}

void CMainDlg::OnRefresh(UINT uNotifyCode, int nID, CWindow wndCtl) {
    LoadProcessList();
}

void CMainDlg::OnAppAbout(UINT uNotifyCode, int nID, CWindow wndCtl) {
    PCWSTR content =
        L"An inspection tool for UWP and WinUI 3 applications. Seamlessly view "
        L"and manipulate UI elements and their properties in real time.\n"
        L"\n"
        L"By <A HREF=\"https://ramensoftware.com/\">Ramen Software</A>\n"
        L"\n"
        L"<A HREF=\"https://ramensoftware.com/uwpspy\">Homepage</a>\n"
        L"<A HREF=\"https://github.com/m417z/UWPSpy\">Source code</a>";

    TASKDIALOGCONFIG taskDialogConfig{
        .cbSize = sizeof(taskDialogConfig),
        .hwndParent = m_hWnd,
        .hInstance = _Module.GetModuleInstance(),
        .dwFlags = TDF_ENABLE_HYPERLINKS | TDF_ALLOW_DIALOG_CANCELLATION,
        .pszWindowTitle = L"About",
        .pszMainIcon = MAKEINTRESOURCE(IDR_MAINFRAME),
        .pszMainInstruction = L"UWPSpy v" VER_FILE_VERSION_WSTR,
        .pszContent = content,
        .pfCallback = [](HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                         LONG_PTR lpRefData) -> HRESULT {
            switch (msg) {
                case TDN_HYPERLINK_CLICKED:
                    OpenUrl(hwnd, (PCWSTR)lParam);
                    break;
            }

            return S_OK;
        },
    };

    ::TaskDialogIndirect(&taskDialogConfig, nullptr, nullptr, nullptr);
}

void CMainDlg::OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl) {
    EndDialog(nID);
}

LRESULT CMainDlg::OnListDblClk(LPNMHDR pnmh) {
    LPNMITEMACTIVATE lpnmitem = (LPNMITEMACTIVATE)pnmh;
    if (lpnmitem->iItem == -1) {
        return 0;
    }

    ProcessSpyFromList(lpnmitem->iItem);

    return 0;
}

LRESULT CMainDlg::SelectProcessInListFromCursor(UINT uMsg,
                                                WPARAM wParam,
                                                LPARAM lParam) {
    POINT pt;
    ::GetCursorPos(&pt);
    CWindow wnd = ::WindowFromPoint(pt);
    if (!wnd) {
        return 0;
    }

    DWORD pid = wnd.GetWindowProcessID();

    CListViewCtrl list(GetDlgItem(IDC_PROCESS_LIST));
    int itemCount = list.GetItemCount();
    for (int i = 0; i < itemCount; i++) {
        DWORD itemPid = static_cast<DWORD>(list.GetItemData(i));
        if (itemPid == pid) {
            list.SetItemState(i, LVIS_FOCUSED | LVIS_SELECTED,
                              LVIS_FOCUSED | LVIS_SELECTED);
            list.EnsureVisible(i, FALSE);
            break;
        }
    }

    return 0;
}

void CMainDlg::InitProcessList() {
    CListViewCtrl list(GetDlgItem(IDC_PROCESS_LIST));

    m_processListSort.SubclassWindow(list);

    list.SetExtendedListViewStyle(LVS_EX_HEADERDRAGDROP | LVS_EX_FULLROWSELECT |
                                  LVS_EX_LABELTIP | LVS_EX_DOUBLEBUFFER);

    using GetDpiForWindow_t = UINT(WINAPI*)(HWND hwnd);
    static GetDpiForWindow_t pGetDpiForWindow = []() {
        HMODULE hUser32 = GetModuleHandle(L"user32.dll");
        return (GetDpiForWindow_t)(hUser32 ? GetProcAddress(hUser32,
                                                            "GetDpiForWindow")
                                           : nullptr);
    }();

    UINT windowDpi = pGetDpiForWindow ? pGetDpiForWindow(m_hWnd) : 96;

    // Sort PIDs as decimals, signed 128-bit (16-byte) values representing
    // 96-bit (12-byte) integer numbers:
    // https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/decimal-data-type
    // A bit of an overkill, but LVCOLSORT_LONG is signed 32-bit while PIDs are
    // unsigned 32-bit, so it doesn't fit. And a custom type isn't worth the
    // effort.
    struct {
        PCWSTR name;
        int width;
        WORD sort;
    } columns[] = {
        {L"Process name", -60, LVCOLSORT_TEXT},
        {L"PID", 60, LVCOLSORT_DECIMAL},
    };

    for (int i = 0; i < ARRAYSIZE(columns); i++) {
        list.InsertColumn(i, columns[i].name);
        int width = columns[i].width;
        width = MulDiv(width, windowDpi, 96);
        if (width < 0) {
            CRect listRect;
            list.GetClientRect(&listRect);
            int scrollbarWidth = ::GetSystemMetrics(SM_CXVSCROLL);
            width += listRect.Width() - scrollbarWidth;
            if (width < scrollbarWidth) {
                width = scrollbarWidth;
            }
        }

        list.SetColumnWidth(i, width);
        m_processListSort.SetColumnSortType(i, columns[i].sort);
    }

    m_processListSort.SetSortColumn(0);
}

void CMainDlg::LoadProcessList() {
    m_processListSort.SetRedraw(FALSE);

    m_processListSort.DeleteAllItems();

    int itemIndex = 0;

    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snapshot != INVALID_HANDLE_VALUE) {
        PROCESSENTRY32 entry{
            .dwSize = sizeof(entry),
        };
        if (Process32First(snapshot, &entry)) {
            do {
                if (entry.th32ProcessID == 0) {
                    continue;
                }

                WCHAR szPid[16];
                swprintf_s(szPid, L"%u", entry.th32ProcessID);

                m_processListSort.AddItem(itemIndex, 0, entry.szExeFile);
                m_processListSort.AddItem(itemIndex, 1, szPid);
                m_processListSort.SetItemData(itemIndex, entry.th32ProcessID);
                itemIndex++;
            } while (Process32Next(snapshot, &entry));
        }

        CloseHandle(snapshot);
    }

    m_processListSort.DoSortItems(m_processListSort.GetSortColumn(),
                                  m_processListSort.IsSortDescending());

    m_processListSort.SetRedraw(TRUE);
    m_processListSort.RedrawWindow(
        nullptr, nullptr,
        RDW_ERASE | RDW_FRAME | RDW_INVALIDATE | RDW_ALLCHILDREN);
}

void CMainDlg::ApplyDarkMode() {
    m_darkMode = dark_mode::IsSystemDarkMode();

    dark_mode::SetPreferredAppMode(dark_mode::AppMode::AllowDark);

    dark_mode::SetDarkTitleBar(m_hWnd, m_darkMode);

    if (m_darkMode && m_darkBgBrush.IsNull()) {
        m_darkBgBrush.CreateSolidBrush(dark_mode::kDarkBg);
    }

    // Apply theme to all child controls.
    for (HWND hChild = ::GetWindow(m_hWnd, GW_CHILD); hChild;
         hChild = ::GetWindow(hChild, GW_HWNDNEXT)) {
        dark_mode::SetControlTheme(hChild, m_darkMode);
    }

    // List view header.
    CHeaderCtrl header =
        CListViewCtrl(GetDlgItem(IDC_PROCESS_LIST)).GetHeader();
    if (header) {
        dark_mode::SetControlTheme(header, m_darkMode);
    }

    // Manual list view colors as fallback.
    {
        auto lv = CListViewCtrl(GetDlgItem(IDC_PROCESS_LIST));
        lv.SetBkColor(m_darkMode ? dark_mode::kDarkBg
                                 : ::GetSysColor(COLOR_WINDOW));
        lv.SetTextBkColor(m_darkMode ? dark_mode::kDarkBg
                                     : ::GetSysColor(COLOR_WINDOW));
        lv.SetTextColor(m_darkMode ? dark_mode::kDarkText
                                   : ::GetSysColor(COLOR_WINDOWTEXT));
    }

    // Resize grip.
    HWND hGrip = GetDlgItem(ATL_IDW_STATUS_BAR);
    if (hGrip) {
        dark_mode::SetControlTheme(hGrip, m_darkMode);
    }

    InvalidateRect(nullptr, TRUE);
}

HBRUSH CMainDlg::OnCtlColorDlg(CDCHandle dc, CWindow wnd) {
    if (m_darkMode) {
        return m_darkBgBrush;
    }
    SetMsgHandled(FALSE);
    return nullptr;
}

HBRUSH CMainDlg::OnCtlColorStatic(CDCHandle dc, CStatic wndStatic) {
    if (m_darkMode) {
        dc.SetTextColor(dark_mode::kDarkText);
        dc.SetBkColor(dark_mode::kDarkBg);
        return m_darkBgBrush;
    }
    SetMsgHandled(FALSE);
    return nullptr;
}

HBRUSH CMainDlg::OnCtlColorBtn(CDCHandle dc, CButton button) {
    if (m_darkMode) {
        dc.SetTextColor(dark_mode::kDarkText);
        dc.SetBkColor(dark_mode::kDarkBg);
        return m_darkBgBrush;
    }
    SetMsgHandled(FALSE);
    return nullptr;
}

LRESULT CALLBACK CMainDlg::ListViewSubclassProc(HWND hWnd,
                                                UINT uMsg,
                                                WPARAM wParam,
                                                LPARAM lParam,
                                                UINT_PTR uIdSubclass,
                                                DWORD_PTR dwRefData) {
    auto* pThis = reinterpret_cast<CMainDlg*>(dwRefData);
    if (uMsg == WM_NOTIFY && pThis->m_darkMode) {
        auto* nmhdr = reinterpret_cast<LPNMHDR>(lParam);
        if (nmhdr->code == NM_CUSTOMDRAW) {
            auto* nmcd = reinterpret_cast<LPNMCUSTOMDRAW>(lParam);
            if (nmcd->dwDrawStage == CDDS_PREPAINT) {
                return CDRF_NOTIFYITEMDRAW;
            }
            if (nmcd->dwDrawStage == CDDS_ITEMPREPAINT) {
                ::SetTextColor(nmcd->hdc, dark_mode::kDarkText);
                return CDRF_DODEFAULT;
            }
        }
    }
    return ::DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK CMainDlg::GripSubclassProc(HWND hWnd,
                                            UINT uMsg,
                                            WPARAM wParam,
                                            LPARAM lParam,
                                            UINT_PTR uIdSubclass,
                                            DWORD_PTR dwRefData) {
    auto* pThis = reinterpret_cast<CMainDlg*>(dwRefData);
    if (pThis->m_darkMode) {
        if (uMsg == WM_ERASEBKGND) {
            return 1;
        }
        if (uMsg == WM_PAINT) {
            PAINTSTRUCT ps;
            HDC hdc = ::BeginPaint(hWnd, &ps);
            RECT rc;
            ::GetClientRect(hWnd, &rc);
            ::FillRect(hdc, &rc, pThis->m_darkBgBrush);

            // Draw grip dots.
            HBRUSH hDotBrush = ::CreateSolidBrush(RGB(96, 96, 96));
            int d = 2;
            int s = 4;
            int x0 = rc.right - 3;
            int y0 = rc.bottom - 3;
            for (int diag = 0; diag < 3; diag++) {
                for (int i = 0; i <= diag; i++) {
                    int x = x0 - (diag - i) * s;
                    int y = y0 - i * s;
                    RECT dot = {x - d, y - d, x, y};
                    ::FillRect(hdc, &dot, hDotBrush);
                }
            }
            ::DeleteObject(hDotBrush);

            ::EndPaint(hWnd, &ps);
            return 0;
        }
    }
    return ::DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void CMainDlg::OnSettingChange(UINT uFlags, LPCTSTR lpszSection) {
    if (lpszSection && _wcsicmp(lpszSection, L"ImmersiveColorSet") == 0) {
        ApplyDarkMode();
    }
}

void CMainDlg::ProcessSpyFromList(int index) {
    DWORD pid = static_cast<DWORD>(m_processListSort.GetItemData(index));

    ProcessSpyFramework framework = kFrameworkUWP;
    if (IsDlgButtonChecked(IDC_RADIO_WINUI) != BST_UNCHECKED) {
        framework = kFrameworkWinUI;
    }

    if (ProcessSpy(m_hWnd, pid, framework)) {
        // EndDialog(0);

        // Starting with Windows 11 build 22621.2506, closing the window right
        // away makes the inspection fail. Use a timer as a workaround.
        SetTimer(TIMER_ID_END_DIALOG, 0);
    }
}
