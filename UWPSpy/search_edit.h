#pragma once

// A single-line edit control for search and filter boxes. Adds a clear button
// which is shown when the box isn't empty, and Ctrl+Backspace/Ctrl+Delete word
// deletion, which plain edit controls don't support.
//
// Since UWPSpy dialogs run in the target's message loop without
// IsDialogMessage, keys such as Enter and Escape are passed to a handler set
// with SetKeyDownHandler.
class CSearchEdit : public CWindowImpl<CSearchEdit, CEdit> {
   public:
    // Return true if the key was handled.
    using KeyDownHandler_t = std::function<bool(UINT nChar)>;

    void Init(HWND hWnd, PCWSTR cueBannerText) {
        SubclassWindow(hWnd);
        ModifyStyle(0, WS_CLIPCHILDREN);
        SetCueBannerText(cueBannerText);

        // Use the "Cancel" icon of the system icon font if available.
        bool iconFont = CreateIconFont();
        m_clearButton.Create(m_hWnd, rcDefault,
                             iconFont ? L"\uE711" : L"\u00D7",
                             WS_CHILD | SS_CENTER | SS_CENTERIMAGE | SS_NOTIFY,
                             0, kClearButtonId);
        m_clearButton.SetFont(iconFont ? m_iconFont.m_hFont : GetFont());

        LayoutClearButton();
        UpdateClearButton();
    }

    void SetKeyDownHandler(KeyDownHandler_t handler) {
        m_keyDownHandler = std::move(handler);
    }

    // Must be called when the text changes (EN_CHANGE).
    void UpdateClearButton() {
        m_clearButton.ShowWindow(GetWindowTextLength() > 0 ? SW_SHOWNA
                                                           : SW_HIDE);
    }

   private:
    enum { kClearButtonId = 1 };

    BEGIN_MSG_MAP_EX(CSearchEdit)
        MSG_WM_KEYDOWN(OnKeyDown)
        MSG_WM_CHAR(OnChar)
        MSG_WM_SIZE(OnSize)
        MSG_WM_SETCURSOR(OnSetCursor)
        MSG_WM_CTLCOLORSTATIC(OnCtlColorStatic)
        COMMAND_HANDLER_EX(kClearButtonId, STN_CLICKED, OnClearButtonClicked)
    END_MSG_MAP()

    void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) {
        if ((nChar == VK_BACK || nChar == VK_DELETE) &&
            ::GetKeyState(VK_CONTROL) < 0) {
            DeleteWord(/*forward=*/nChar == VK_DELETE);
            return;
        }

        if (m_keyDownHandler && m_keyDownHandler(nChar)) {
            return;
        }

        SetMsgHandled(FALSE);
    }

    void OnChar(TCHAR chChar, UINT nRepCnt, UINT nFlags) {
        switch (chChar) {
            case 6:
                // Ctrl+F, which is used to focus the box, selects all text if
                // the box is already focused.
                SetSelAll();
                break;

            case VK_RETURN:
            case VK_ESCAPE:
            case 0x7F:
                // Return, Escape and Ctrl+Backspace. Handled in OnKeyDown,
                // don't pass them to the edit control to avoid a beep or an
                // inserted character.
                break;

            default:
                SetMsgHandled(FALSE);
                break;
        }
    }

    void OnSize(UINT nType, CSize size) {
        LayoutClearButton();
        SetMsgHandled(FALSE);
    }

    BOOL OnSetCursor(CWindow wnd, UINT nHitTest, UINT message) {
        if (wnd == m_clearButton) {
            ::SetCursor(::LoadCursor(nullptr, IDC_HAND));
            return TRUE;
        }

        SetMsgHandled(FALSE);
        return FALSE;
    }

    HBRUSH OnCtlColorStatic(CDCHandle dc, CStatic wndStatic) {
        if (wndStatic != m_clearButton) {
            SetMsgHandled(FALSE);
            return nullptr;
        }

        // Use the same colors as the edit control, as provided by the parent
        // (which can be dark mode colors), with a dimmed text color.
        auto brush = reinterpret_cast<HBRUSH>(GetParent().SendMessage(
            WM_CTLCOLOREDIT, reinterpret_cast<WPARAM>(dc.m_hDC),
            reinterpret_cast<LPARAM>(m_hWnd)));

        COLORREF text = dc.GetTextColor();
        COLORREF bk = dc.GetBkColor();
        dc.SetTextColor(RGB((GetRValue(text) + GetRValue(bk)) / 2,
                            (GetGValue(text) + GetGValue(bk)) / 2,
                            (GetBValue(text) + GetBValue(bk)) / 2));

        return brush;
    }

    void OnClearButtonClicked(UINT uNotifyCode, int nID, CWindow wndCtl) {
        SetWindowText(L"");
        SetFocus();
    }

    bool CreateIconFont() {
        constexpr WCHAR kIconFontName[] = L"Segoe MDL2 Assets";

        if (!m_iconFont.CreatePointFont(80, kIconFontName)) {
            return false;
        }

        // Make sure that the font exists and wasn't substituted.
        CClientDC dc(m_hWnd);
        HFONT oldFont = dc.SelectFont(m_iconFont);
        WCHAR faceName[LF_FACESIZE];
        dc.GetTextFace(faceName, ARRAYSIZE(faceName));
        dc.SelectFont(oldFont);

        if (_wcsicmp(faceName, kIconFontName) != 0) {
            m_iconFont.DeleteObject();
            return false;
        }

        return true;
    }

    void LayoutClearButton() {
        CRect rect;
        GetClientRect(&rect);
        int width = rect.Height();
        m_clearButton.SetWindowPos(nullptr, rect.right - width, 0, width,
                                   rect.Height(),
                                   SWP_NOZORDER | SWP_NOACTIVATE);
        SetMargins(0, width, EC_RIGHTMARGIN);
    }

    // Deletes the selection, or the word before or after the caret if there's
    // no selection. Words are runs of alphanumeric characters or runs of
    // punctuation, so that e.g. "Windows.UI.Xaml" is deleted part by part.
    void DeleteWord(bool forward) {
        int start, end;
        GetSel(start, end);

        if (start == end) {
            CString text;
            GetWindowText(text);
            int length = text.GetLength();

            auto charClass = [&text](int i) {
                if (iswspace(text[i])) {
                    return 0;
                }
                if (iswalnum(text[i]) || text[i] == L'_') {
                    return 1;
                }
                return 2;
            };

            if (forward) {
                if (end < length) {
                    int startClass = charClass(end);
                    while (end < length && charClass(end) == startClass) {
                        end++;
                    }
                }
                while (end < length && charClass(end) == 0) {
                    end++;
                }
            } else {
                while (start > 0 && charClass(start - 1) == 0) {
                    start--;
                }
                if (start > 0) {
                    int endClass = charClass(start - 1);
                    while (start > 0 && charClass(start - 1) == endClass) {
                        start--;
                    }
                }
            }

            if (start == end) {
                return;
            }

            SetSel(start, end);
        }

        ReplaceSel(L"", TRUE);
    }

    CStatic m_clearButton;
    CFont m_iconFont;
    KeyDownHandler_t m_keyDownHandler;
};
