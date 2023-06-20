#pragma once

#include "resource.h"
#include "winrt.hpp"

class CMainDlg : public CDialogImpl<CMainDlg>, public CDialogResize<CMainDlg> {
   public:
    enum { IDD = IDD_MAINDLG };

    enum {
        HOTKEY_SELECT_ELEMENT_FROM_CURSOR = 1,
    };

    enum {
        UWM_ACTIVATE_WINDOW = WM_APP,
    };

    CMainDlg(winrt::com_ptr<IVisualTreeService3> service,
             winrt::com_ptr<IXamlDiagnostics> diagnostics);
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
        MSG_WM_HOTKEY(OnHotKey)
        NOTIFY_HANDLER_EX(IDC_ELEMENT_TREE, TVN_SELCHANGED,
                          OnElementTreeSelChaneged)
        NOTIFY_HANDLER_EX(IDC_ELEMENT_TREE, TVN_KEYDOWN, OnElementTreeKeyDown)
        COMMAND_ID_HANDLER_EX(IDC_SPLIT_TOGGLE, OnSplitToggle)
        COMMAND_HANDLER_EX(IDC_PROPERTY_NAME, CBN_SELCHANGE,
                           OnPropertyNameSelChange)
        COMMAND_ID_HANDLER_EX(IDC_PROPERTY_REMOVE, OnPropertyRemove)
        COMMAND_ID_HANDLER_EX(IDC_PROPERTY_SET, OnPropertySet)
        COMMAND_ID_HANDLER_EX(IDC_COLLAPSE_ALL, OnCollapseAll)
        COMMAND_ID_HANDLER_EX(IDC_EXPAND_ALL, OnExpandAll)
        COMMAND_HANDLER_EX(IDC_HIGHLIGHT_SELECTION, BN_CLICKED,
                           OnHighlightSelection)
        COMMAND_HANDLER_EX(IDC_DETAILED_PROPERTIES, BN_CLICKED,
                           OnDetailedProperties)
        COMMAND_ID_HANDLER_EX(ID_APP_ABOUT, OnAppAbout)
        COMMAND_ID_HANDLER_EX(IDCANCEL, OnCancel)
        MESSAGE_HANDLER_EX(UWM_ACTIVATE_WINDOW, OnActivateWindow)
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
            DLGRESIZE_CONTROL(IDC_ATTRIBUTE_LIST, DLSZ_MOVE_X | DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_NAME, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_VALUE, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_REMOVE, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_SET, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_COLLAPSE_ALL, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_EXPAND_ALL, DLSZ_MOVE_Y)
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
            DLGRESIZE_CONTROL(IDC_ATTRIBUTE_LIST, DLSZ_SIZE_X | DLSZ_SIZE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_NAME, DLSZ_SIZE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_VALUE, DLSZ_SIZE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_REMOVE, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_PROPERTY_SET, DLSZ_MOVE_X | DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_COLLAPSE_ALL, DLSZ_MOVE_Y)
            DLGRESIZE_CONTROL(IDC_EXPAND_ALL, DLSZ_MOVE_Y)
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
    void OnHotKey(int nHotKeyID, UINT uModifiers, UINT uVirtKey);
    LRESULT OnElementTreeSelChaneged(LPNMHDR pnmh);
    LRESULT OnElementTreeKeyDown(LPNMHDR pnmh);
    void OnSplitToggle(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnPropertyNameSelChange(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnPropertyRemove(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnPropertySet(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCollapseAll(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnExpandAll(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnHighlightSelection(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnDetailedProperties(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnAppAbout(UINT uNotifyCode, int nID, CWindow wndCtl);
    void OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl);
    LRESULT OnActivateWindow(UINT uMsg, WPARAM wParam, LPARAM lParam);

    bool SetSelectedElementInformation();
    void ResetAttributesListColumns();
    void PopulateAttributesList(InstanceHandle handle);
    void AddItemToTree(HTREEITEM parentTreeItem,
                       HTREEITEM insertAfter,
                       InstanceHandle handle,
                       ElementItem* elementItem);
    InstanceHandle ElementFromPoint(CPoint pt);
    InstanceHandle ElementFromPointInSubtree(wux::UIElement subtree, CPoint pt);
    bool CreateFlashArea(InstanceHandle handle);
    void DestroyFlashArea();
    bool SelectElementFromCursor();

    bool m_highlightSelection = true;
    bool m_detailedProperties = false;
    bool m_splitModeAttributesExpanded = false;
    bool m_registeredHotkeySelectElementFromCursor = false;

    winrt::com_ptr<IVisualTreeService3> m_visualTreeService;
    winrt::com_ptr<IXamlDiagnostics> m_xamlDiagnostics;

    std::unordered_map<InstanceHandle, ElementItem> m_elementItems;

    // Note: A parent might be in m_parentToChildren but not in m_elementItems
    // at some point if the child is added before its parent.
    std::unordered_map<InstanceHandle, std::vector<InstanceHandle>>
        m_parentToChildren;

    CString m_lastPropertySelection;

    CWindow m_flashAreaWindow;
};
