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

    return TRUE;
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

void CMainDlg::InitProcessList() {
    CListViewCtrl list(GetDlgItem(IDC_PROCESS_LIST));

    m_processListSort.SubclassWindow(list);

    list.SetExtendedListViewStyle(LVS_EX_HEADERDRAGDROP | LVS_EX_FULLROWSELECT |
                                  LVS_EX_LABELTIP | LVS_EX_DOUBLEBUFFER);
    ::SetWindowTheme(list, L"Explorer", nullptr);

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

void CMainDlg::ProcessSpyFromList(int index) {
    DWORD pid = static_cast<DWORD>(m_processListSort.GetItemData(index));

    ProcessSpyFramework framework = kFrameworkUWP;
    if (IsDlgButtonChecked(IDC_RADIO_WINUI) != BST_UNCHECKED) {
        framework = kFrameworkWinUI;
    }

    if (ProcessSpy(m_hWnd, pid, framework)) {
        EndDialog(0);
    }
}
