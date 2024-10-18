#pragma once

#include "resource.h"
#include "winrt.hpp"

class CMainDlg : public CDialogImpl<CMainDlg>, public CDialogResize<CMainDlg> {
   public:
    enum { IDD = IDD_MAINDLG };

    enum {
        TIMER_ID_REDRAW_TREE = 1,
        TIMER_ID_SET_SELECTED_ELEMENT_INFORMATION,
        TIMER_ID_REFRESH_SELECTED_ELEMENT_INFORMATION,
    };

    enum {
        UWM_ACTIVATE_WINDOW = WM_APP,
        UWM_DESTROY_WINDOW,
    };

    enum class EventId {
        Shown,
        Hidden,
        FinalMessage,
    };

    using OnEventCallback_t = std::function<void(HWND hWnd, EventId eventId)>;

    CMainDlg(winrt::com_ptr<IXamlDiagnostics> diagnostics,
             OnEventCallback_t eventCallback);
    void Hide();
    void Show();
    void ElementAdded(const ParentChildRelation& parentChildRelation,
                      const VisualElement& element);
    void ElementRemoved(InstanceHandle handle);

   private:
    BEGIN_MSG_MAP_EX(CMainDlg)
        CHAIN_MSG_MAP(CDialogResize<CMainDlg>)
        MSG_WM_INITDIALOG(OnInitDialog)
        MSG_WM_DESTROY(OnDestroy)
        MSG_WM_TIMER(OnTimer)
        MSG_WM_CONTEXTMENU(OnContextMenu)
        NOTIFY_HANDLER_EX(IDC_ELEMENT_TREE, TVN_SELCHANGED,
                          OnElementTreeSelChanged)
        COMMAND_ID_HANDLER_EX(IDC_SPLIT_TOGGLE, OnSplitToggle)
        NOTIFY_HANDLER_EX(IDC_DETAILS_TABS, TCN_SELCHANGE,
                          OnDetailsTabsSelChange)
        NOTIFY_HANDLER_EX(IDC_ATTRIBUTE_LIST, NM_DBLCLK, OnAttributeListDblClk)
        COMMAND_HANDLER_EX(IDC_PROPERTY_NAME, CBN_SELCHANGE,
                           OnPropertyNameSelChange)
        COMMAND_ID_HANDLER_EX(IDC_PROPERTY_IS_XAML, OnPropertyIsXaml)
        COMMAND_ID_HANDLER_EX(IDC_PROPERTY_REMOVE, OnPropertyRemove)
        COMMAND_ID_HANDLER_EX(IDC_PROPERTY_SET, OnPropertySet)
        COMMAND_ID_HANDLER_EX(IDC_COLLAPSE_ALL, OnCollapseAll)
        COMMAND_HANDLER_EX(IDC_HIGHLIGHT_SELECTION, BN_CLICKED,
                           OnHighlightSelection)
        COMMAND_HANDLER_EX(IDC_DETAILED_PROPERTIES, BN_CLICKED,
                           OnDetailedProperties)
        COMMAND_ID_HANDLER_EX(ID_APP_ABOUT, OnAppAbout)
        COMMAND_ID_HANDLER_EX(IDCANCEL, OnCancel)
        MESSAGE_HANDLER_EX(UWM_ACTIVATE_WINDOW, OnActivateWindow)
        MESSAGE_HANDLER_EX(UWM_DESTROY_WINDOW, OnDestroyWindow)
        // ----------
        ALT_MSG_MAP(1)
        MSG_WM_CHAR(ElementTreeOnChar)
    END_MSG_MAP()

    struct ResizeData : public CDialogResize<ResizeData> {
        BEGIN_DLGRESIZE_MAP(ResizeData)
            DLGRESIZE_CONTROL(IDC_ELEMENT_TREE, DLSZ_SIZE_X | DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_SPLIT_TOGGLE, DLSZ_MOVE_X | DLSZ_CENTER_Y)
            DLGRESIZE_CONTROL(IDC_CLASS_STATIC, DLSZ_MOVE_X)
            DLGRESIZE_CONTROL(IDC_CLASS_EDIT, DLSZ_MOVE_X | DLSZ_SIZE_X)
            DLGRESIZE_CONTROL(IDC_NAME_STATIC, DLSZ_MOVE_X)
            DLGRESIZE_CONTROL(IDC_NAME_EDIT, DLSZ_MOVE_X | DLSZ_SIZE_X)
            DLGRESIZE_CONTROL(IDC_RECT_STATIC, DLSZ_MOVE_X)
            DLGRESIZE_CONTROL(IDC_RECT_EDIT, DLSZ_MOVE_X)
            DLGRESIZE_CONTROL(IDC_DETAILS_TABS, DLSZ_MOVE_X)
            DLGRESIZE_CONTROL(IDC_ATTRIBUTE_LIST, DLSZ_MOVE_X | DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_VISUAL_STATE_TREE, DLSZ_MOVE_X | DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_NAME, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_VALUE, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_VALUE_XAML,
                              DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_IS_XAML, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_REMOVE, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_SET, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_COLLAPSE_ALL, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_HIGHLIGHT_SELECTION, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_DETAILED_PROPERTIES, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X | DLSZ_MOVE_Y)
        END_DLGRESIZE_MAP()
    };

    struct ResizeDataAttributesExpanded
        : public CDialogResize<ResizeDataAttributesExpanded> {
        BEGIN_DLGRESIZE_MAP(ResizeDataAttributesExpanded)
            DLGRESIZE_CONTROL(IDC_ELEMENT_TREE, DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_SPLIT_TOGGLE, DLSZ_CENTER_Y)
            DLGRESIZE_CONTROL(IDC_CLASS_STATIC, 0)
            DLGRESIZE_CONTROL(IDC_CLASS_EDIT, DLSZ_SIZE_X)
            DLGRESIZE_CONTROL(IDC_NAME_STATIC, 0)
            DLGRESIZE_CONTROL(IDC_NAME_EDIT, DLSZ_SIZE_X)
            DLGRESIZE_CONTROL(IDC_RECT_STATIC, 0)
            DLGRESIZE_CONTROL(IDC_RECT_EDIT, DLSZ_SIZE_X)
            DLGRESIZE_CONTROL(IDC_DETAILS_TABS, DLSZ_SIZE_X)
            DLGRESIZE_CONTROL(IDC_ATTRIBUTE_LIST, DLSZ_SIZE_X | DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_VISUAL_STATE_TREE, DLSZ_SIZE_X | DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_NAME, DLSZ_SIZE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_VALUE, DLSZ_SIZE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_VALUE_XAML,
                              DLSZ_SIZE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_IS_XAML, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_REMOVE, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_SET, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_COLLAPSE_ALL, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_HIGHLIGHT_SELECTION, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_DETAILED_PROPERTIES, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X | DLSZ_MOVE_Y)
        END_DLGRESIZE_MAP()
    };

    const _AtlDlgResizeMap* GetDlgResizeMap() {
        return m_splitModeAttributesExpanded
                   ? ResizeDataAttributesExpanded::GetDlgResizeMap()
                   : ResizeData::GetDlgResizeMap();
    }

    struct ElementItem {
        InstanceHandle parentHandle;
        std::wstring itemTitle;
        HTREEITEM treeItem;
    };

    BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam);
    void OnDestroy();
    void OnTimer(UINT_PTR nIDEvent);
    void OnContextMenu(CWindow wnd, CPoint point);
    LRESULT OnElementTreeSelChanged(LPNMHDR pnmh);
    void OnSplitToggle(UINT uNotifyCode, int nID, CWindow wndCtl);
    LRESULT OnDetailsTabsSelChange(LPNMHDR pnmh);
    LRESULT OnAttributeListDblClk(LPNMHDR pnmh);
    void OnPropertyNameSelChange(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnPropertyIsXaml(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnPropertyRemove(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnPropertySet(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCollapseAll(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnHighlightSelection(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnDetailedProperties(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnAppAbout(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl);
    LRESULT OnActivateWindow(UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT OnDestroyWindow(UINT uMsg, WPARAM wParam, LPARAM lParam);

    void OnFinalMessage(HWND hWnd) override;

    void ElementTreeOnChar(TCHAR chChar, UINT nRepCnt, UINT nFlags);

    void RedrawTreeQueue();
    bool SetSelectedElementInformation();
    bool RefreshSelectedElementInformation(UINT delay = 20);
    void ResetAttributesListColumns();
    void PopulateAttributesList(InstanceHandle handle);
    void PopulateVisualStatesTree(InstanceHandle handle);
    void AddItemToTree(HTREEITEM parentTreeItem,
                       HTREEITEM insertAfter,
                       InstanceHandle handle,
                       ElementItem* elementItem);
    InstanceHandle ElementFromPoint(CPoint pt);
    InstanceHandle ElementFromPointInSubtree(wux::UIElement subtree, CPoint pt);
    InstanceHandle ElementFromPointInSubtree(mux::UIElement subtree, CPoint pt);
    bool CreateFlashArea(InstanceHandle handle);
    void DestroyFlashArea();
    bool SelectElementFromCursor();

    CIcon m_icon, m_smallIcon;
    CContainedWindowT<CTreeViewCtrlEx> m_elementTree;
    bool m_listCollapsed = false;
    bool m_highlightSelection = true;
    bool m_detailedProperties = false;
    bool m_splitModeAttributesExpanded = false;
    bool m_redrawTreeQueued = false;
    bool m_redrawTreeQueuedEnsureSelectionVisible = false;

    winrt::com_ptr<IVisualTreeService3> m_visualTreeService;
    winrt::com_ptr<IXamlDiagnostics> m_xamlDiagnostics;
    OnEventCallback_t m_eventCallback;

    std::unordered_map<InstanceHandle, ElementItem> m_elementItems;

    // Note: A parent might be in m_parentToChildren but not in m_elementItems
    // at some point if the child is added before its parent.
    std::unordered_map<InstanceHandle, std::vector<InstanceHandle>>
        m_parentToChildren;

    CString m_lastPropertySelection;

    CWindow m_flashAreaWindow;
};
