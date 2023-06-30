#include "stdafx.h"

#include "MainDlg.h"

#include "flash_area.h"
#include "version.h"

namespace {

// Delay each tree redraw, redrawing at most once every kRedrawTreeDelay ms.
// Otherwise, multiple redraw operations can make the UI very slow.
constexpr UINT kRedrawTreeDelay = 200;

// https://github.com/sumatrapdfreader/sumatrapdf/blob/9a2183db3c3db5cbf242ac9d8f8576750f581096/src/utils/WinUtil.cpp#L2863
void TreeViewExpandRecursively(HWND hTree, HTREEITEM hItem, DWORD flag) {
    while (hItem) {
        TreeView_Expand(hTree, hItem, flag);
        HTREEITEM child = TreeView_GetChild(hTree, hItem);
        if (child) {
            TreeViewExpandRecursively(hTree, child, flag);
        }
        hItem = TreeView_GetNextSibling(hTree, hItem);
    }
}

bool IsTreeItemInView(CTreeItem treeItem) {
    CRect rect;
    if (!treeItem.GetRect(rect, FALSE)) {
        return false;
    }

    CRect clientRect;
    treeItem.GetTreeView()->GetClientRect(clientRect);

    CRect intercectionRect;
    return intercectionRect.IntersectRect(rect, clientRect);
}

// https://stackoverflow.com/a/35277401
void ListViewSetTopIndex(CListViewCtrl listView, int index) {
    CSize itemSpacing;
    listView.GetItemSpacing(itemSpacing, TRUE);

    CSize scrollsize(0, (index - listView.GetTopIndex()) * itemSpacing.cy);
    listView.Scroll(scrollsize);
}

// https://stackoverflow.com/a/41088729
int GetRequiredComboDroppedWidth(CComboBox rCombo) {
    CString str;
    CSize sz;
    int dx = 0;
    TEXTMETRIC tm;
    CDCHandle pDC = rCombo.GetDC();
    CFontHandle pFont = rCombo.GetFont();

    // Select the listbox font, save the old font
    CFontHandle pOldFont = pDC.SelectFont(pFont);
    // Get the text metrics for avg char width
    pDC.GetTextMetrics(&tm);

    for (int i = 0; i < rCombo.GetCount(); i++) {
        rCombo.GetLBText(i, str);
        pDC.GetTextExtent(str, -1, &sz);

        // Add the avg width to prevent clipping
        sz.cx += tm.tmAveCharWidth;

        if (sz.cx > dx)
            dx = sz.cx;
    }
    // Select the old font back into the DC
    pDC.SelectFont(pOldFont);
    rCombo.ReleaseDC(pDC);

    // Adjust the width for the vertical scroll bar and the left and right
    // border.
    dx += ::GetSystemMetrics(SM_CXVSCROLL) + 2 * ::GetSystemMetrics(SM_CXEDGE);

    return dx;
}

std::optional<CRect> GetRelativeElementRect(wf::IInspectable element) {
    if (auto uiElement = element.try_as<wux::UIElement>()) {
        const auto offset = uiElement.TransformToVisual(nullptr).TransformPoint(
            wf::Point(0, 0));
        const auto size = uiElement.ActualSize();

        CRect rect;
        rect.left = std::lroundf(offset.X);
        rect.top = std::lroundf(offset.Y);
        rect.right = rect.left + std::lroundf(size.x);
        rect.bottom = rect.top + std::lroundf(size.y);

        return rect;
    }

    return std::nullopt;
}

std::optional<CRect> GetRootElementRect(wf::IInspectable element,
                                        HWND* outWnd = nullptr) {
    if (auto window = element.try_as<winrt::Windows::UI::Xaml::Window>()) {
        const auto bounds = window.Bounds();

        CRect rect;
        rect.left = std::lroundf(bounds.X);
        rect.top = std::lroundf(bounds.Y);
        rect.right = rect.left + std::lroundf(bounds.Width);
        rect.bottom = rect.top + std::lroundf(bounds.Height);

        if (outWnd) {
            if (auto coreWindow = window.CoreWindow()) {
                if (const auto coreInterop =
                        coreWindow.try_as<ICoreWindowInterop>()) {
                    coreInterop->get_WindowHandle(outWnd);
                }
            }
        }

        return rect;
    }

    if (auto desktopWindowXamlSource =
            element.try_as<wux::Hosting::DesktopWindowXamlSource>()) {
        const auto nativeSource =
            desktopWindowXamlSource.as<IDesktopWindowXamlSourceNative>();

        // A window which is no longer valid might be returned if the element is
        // being destroyed.
        CWindow wnd;
        if (FAILED(nativeSource->get_WindowHandle(&wnd.m_hWnd)) || !wnd ||
            !wnd.IsWindow()) {
            return std::nullopt;
        }

        CRect rectWithDpi;
        wnd.GetWindowRect(rectWithDpi);

        UINT dpi = ::GetDpiForWindow(wnd);
        CRect rect;
        rect.left = MulDiv(rectWithDpi.left, 96, dpi);
        rect.top = MulDiv(rectWithDpi.top, 96, dpi);
        // Round width and height, not right and bottom, for consistent results.
        rect.right =
            rect.left + MulDiv(rectWithDpi.right - rectWithDpi.left, 96, dpi);
        rect.bottom =
            rect.top + MulDiv(rectWithDpi.bottom - rectWithDpi.top, 96, dpi);

        if (outWnd) {
            *outWnd = wnd;
        }

        return rect;
    }

    return std::nullopt;
}

void OpenUrl(HWND hWnd, PCWSTR url) {
    if ((INT_PTR)ShellExecute(hWnd, L"open", url, nullptr, nullptr,
                              SW_SHOWNORMAL) <= 32) {
        MessageBox(hWnd, L"Failed to open link", nullptr, MB_ICONHAND);
    }
}

}  // namespace

CMainDlg::CMainDlg(winrt::com_ptr<IVisualTreeService3> service,
                   winrt::com_ptr<IXamlDiagnostics> diagnostics)
    : m_visualTreeService(std::move(service)),
      m_xamlDiagnostics(std::move(diagnostics)) {}

void CMainDlg::Hide() {
    ShowWindow(SW_HIDE);

    if (m_flashAreaWindow) {
        m_flashAreaWindow.ShowWindow(SW_HIDE);
    }
}

void CMainDlg::Show() {
    ShowWindow(SW_SHOW);

    if (m_flashAreaWindow) {
        m_flashAreaWindow.ShowWindow(SW_SHOW);
    }
}

void CMainDlg::ElementAdded(const ParentChildRelation& parentChildRelation,
                            const VisualElement& element) {
    ATLASSERT(parentChildRelation.Parent
                  ? parentChildRelation.Child &&
                        parentChildRelation.Child == element.Handle
                  : !parentChildRelation.Child);

    const std::wstring_view elementType{
        element.Type, element.Type ? SysStringLen(element.Type) : 0};
    const std::wstring_view elementName{
        element.Name, element.Name ? SysStringLen(element.Name) : 0};

    std::wstring itemTitle(elementType);
    if (!elementName.empty()) {
        itemTitle += L" - ";
        itemTitle += elementName;
    }

    ElementItem elementItem{
        .parentHandle = parentChildRelation.Parent,
        .itemTitle = itemTitle,
        .treeItem = nullptr,
    };

    auto [itElementItem, inserted] =
        m_elementItems.try_emplace(element.Handle, std::move(elementItem));
    if (!inserted) {
        // Element already exists, I'm not sure what that means but let's remove
        // the existing element from the tree and hope that it works.
        ElementRemoved(element.Handle);
        const auto [itElementItem2, inserted2] =
            m_elementItems.try_emplace(element.Handle, std::move(elementItem));
        ATLASSERT(inserted2);
        itElementItem = itElementItem2;
    }

    HTREEITEM parentItem = nullptr;
    HTREEITEM insertAfter = TVI_LAST;

    if (parentChildRelation.Parent) {
        auto& children = m_parentToChildren[parentChildRelation.Parent];

        if (parentChildRelation.ChildIndex <= children.size()) {
            if (parentChildRelation.ChildIndex == 0) {
                insertAfter = TVI_FIRST;
            } else if (parentChildRelation.ChildIndex < children.size()) {
                auto it = m_elementItems.find(
                    children[parentChildRelation.ChildIndex - 1]);
                if (it != m_elementItems.end()) {
                    if (it->second.treeItem) {
                        insertAfter = it->second.treeItem;
                    }
                } else {
                    ATLASSERT(FALSE);
                }
            }

            children.insert(children.begin() + parentChildRelation.ChildIndex,
                            element.Handle);
        } else {
            // I've seen this happen, for example with mspaint if you open the
            // color picker. Not sure why or what to do about it, for now at
            // least avoid inserting out of bounds.
            // ATLASSERT(FALSE);
            children.push_back(element.Handle);
        }

        auto it = m_elementItems.find(parentChildRelation.Parent);
        if (it == m_elementItems.end()) {
            return;
        }

        parentItem = it->second.treeItem;
        if (!parentItem) {
            return;
        }
    }

    RedrawTreeQueue();

    AddItemToTree(parentItem, insertAfter, element.Handle,
                  &itElementItem->second);
}

void CMainDlg::ElementRemoved(InstanceHandle handle) {
    auto it = m_elementItems.find(handle);
    if (it == m_elementItems.end()) {
        // I've seen this happen, for example with mspaint if you open the color
        // picker and then close it. Not sure why or what to do about it, for
        // now just return.
        // ATLASSERT(FALSE);
        return;
    }

    if (it->second.treeItem) {
        RedrawTreeQueue();

        auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));

        bool deleted = treeView.DeleteItem(it->second.treeItem);
        ATLASSERT(deleted);

        std::function<void(InstanceHandle, ElementItem*)>
            clearTreeItemRecursive;
        clearTreeItemRecursive = [this, &clearTreeItemRecursive](
                                     InstanceHandle handle,
                                     ElementItem* elementItem) {
            elementItem->treeItem = nullptr;

            if (auto it = m_parentToChildren.find(handle);
                it != m_parentToChildren.end()) {
                for (const auto& childHandle : it->second) {
                    auto childElementItem = m_elementItems.find(childHandle);
                    if (childElementItem == m_elementItems.end()) {
                        ATLASSERT(FALSE);
                        continue;
                    }

                    clearTreeItemRecursive(childHandle,
                                           &childElementItem->second);
                }
            }
        };

        clearTreeItemRecursive(handle, &it->second);
    }

    auto itChildrenOfParent = m_parentToChildren.find(it->second.parentHandle);
    if (itChildrenOfParent != m_parentToChildren.end()) {
        auto& children = itChildrenOfParent->second;
        children.erase(std::remove(children.begin(), children.end(), handle),
                       children.end());
    }

    m_elementItems.erase(it);
}

BOOL CMainDlg::OnInitDialog(CWindow wndFocus, LPARAM lInitParam) {
    // Center the dialog on the screen.
    CenterWindow();

    // Set icons.
    HICON hIcon = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR,
                                   ::GetSystemMetrics(SM_CXICON),
                                   ::GetSystemMetrics(SM_CYICON));
    SetIcon(hIcon, TRUE);
    HICON hIconSmall = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR,
                                        ::GetSystemMetrics(SM_CXSMICON),
                                        ::GetSystemMetrics(SM_CYSMICON));
    SetIcon(hIconSmall, FALSE);

    // Init resizing.
    DlgResize_Init();

    // Init UI elements.
    SetWindowText(
        std::format(L"UWPSpy - PID: {}", GetCurrentProcessId()).c_str());

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    treeView.SetExtendedStyle(TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);
    ::SetWindowTheme(treeView, L"Explorer", nullptr);

    auto attributesList = CListViewCtrl(GetDlgItem(IDC_ATTRIBUTE_LIST));
    attributesList.SetExtendedListViewStyle(
        LVS_EX_FULLROWSELECT | LVS_EX_LABELTIP | LVS_EX_DOUBLEBUFFER);
    ::SetWindowTheme(attributesList, L"Explorer", nullptr);
    ResetAttributesListColumns();

    CButton(GetDlgItem(IDC_HIGHLIGHT_SELECTION))
        .SetCheck(m_highlightSelection ? BST_CHECKED : BST_UNCHECKED);

    CButton(GetDlgItem(IDC_DETAILED_PROPERTIES))
        .SetCheck(m_detailedProperties ? BST_CHECKED : BST_UNCHECKED);

    // Register hotkeys.
    m_registeredHotkeySelectElementFromCursor =
        ::RegisterHotKey(m_hWnd, HOTKEY_SELECT_ELEMENT_FROM_CURSOR,
                         MOD_CONTROL | MOD_SHIFT, 'C');

    return TRUE;
}

void CMainDlg::OnDestroy() {
    if (m_registeredHotkeySelectElementFromCursor) {
        ::UnregisterHotKey(m_hWnd, HOTKEY_SELECT_ELEMENT_FROM_CURSOR);
        m_registeredHotkeySelectElementFromCursor = false;
    }
}

void CMainDlg::OnHotKey(int nHotKeyID, UINT uModifiers, UINT uVirtKey) {
    switch (nHotKeyID) {
        case HOTKEY_SELECT_ELEMENT_FROM_CURSOR:
            SelectElementFromCursor();
            break;
    }
}

void CMainDlg::OnTimer(UINT_PTR nIDEvent) {
    switch (nIDEvent) {
        case TIMER_ID_REDRAW_TREE: {
            KillTimer(nIDEvent);
            m_redrawTreeQueued = false;

            auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
            treeView.SetRedraw(TRUE);

            // Make sure the selected item remains visible.
            if (m_redrawTreeQueuedEnsureSelectionVisible) {
                auto selectedItem = treeView.GetSelectedItem();
                if (selectedItem) {
                    selectedItem.EnsureVisible();
                }

                m_redrawTreeQueuedEnsureSelectionVisible = false;
            }
            break;
        }

        case TIMER_ID_REFRESH_SELECTED_ELEMENT_INFORMATION:
            KillTimer(nIDEvent);
            RefreshSelectedElementInformation(0);
            break;
    }
}

void CMainDlg::OnContextMenu(CWindow wnd, CPoint point) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    if (wnd != treeView) {
        return;
    }

    CTreeItem targetItem;

    if (point.x == -1 && point.y == -1) {
        // Keyboard context menu.
        targetItem = treeView.GetSelectedItem();
    } else {
        CPoint mappedPoint = point;
        treeView.ScreenToClient(&mappedPoint);

        targetItem = treeView.HitTest(mappedPoint, nullptr);
    }

    if (!targetItem) {
        return;
    }

    CPoint menuPoint = point;
    if (menuPoint.x == -1 && menuPoint.y == -1) {
        // Keyboard context menu.
        CRect rect;
        if (targetItem.GetRect(rect, FALSE)) {
            menuPoint = rect.CenterPoint();
            treeView.ClientToScreen(&menuPoint);
        } else {
            ::GetCursorPos(&menuPoint);
        }
    }

    CMenu menu;
    menu.CreatePopupMenu();

    enum {
        MENU_ID_VISIBLE = 1,
    };

    auto handle = static_cast<InstanceHandle>(targetItem.GetData());

    wf::IInspectable element;
    HRESULT hr = m_xamlDiagnostics->GetIInspectableFromHandle(
        handle, reinterpret_cast<::IInspectable**>(winrt::put_abi(element)));
    if (FAILED(hr) || !element) {
        return;
    }

    auto uiElement = element.try_as<wux::UIElement>();
    if (!uiElement) {
        return;
    }

    bool visible = uiElement.Visibility() == wux::Visibility::Visible;

    menu.AppendMenu(MF_STRING | (visible ? MF_CHECKED : 0), MENU_ID_VISIBLE,
                    L"Visible");

    int nCmd = menu.TrackPopupMenu(TPM_RIGHTBUTTON | TPM_RETURNCMD, menuPoint.x,
                                   menuPoint.y, m_hWnd);
    switch (nCmd) {
        case MENU_ID_VISIBLE:
            uiElement.Visibility(visible ? wux::Visibility::Collapsed
                                         : wux::Visibility::Visible);

            RefreshSelectedElementInformation();
            break;
    }
}

InstanceHandle CMainDlg::ElementFromPoint(CPoint pt) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));

    for (auto item = treeView.GetRootItem(); item;
         item = item.GetNextSibling()) {
        auto handle = static_cast<InstanceHandle>(item.GetData());

        wf::IInspectable rootElement;
        HRESULT hr = m_xamlDiagnostics->GetIInspectableFromHandle(
            handle,
            reinterpret_cast<::IInspectable**>(winrt::put_abi(rootElement)));
        if (FAILED(hr) || !rootElement) {
            continue;
        }

        CWindow rootWnd;
        auto rootElementRect = GetRootElementRect(rootElement, &rootWnd.m_hWnd);
        if (!rootElementRect) {
            continue;
        }

        wux::UIElement subtree = nullptr;

        if (auto window =
                rootElement.try_as<winrt::Windows::UI::Xaml::Window>()) {
            subtree = window.Content();
        } else if (auto desktopWindowXamlSource =
                       rootElement
                           .try_as<wux::Hosting::DesktopWindowXamlSource>()) {
            subtree = desktopWindowXamlSource.Content();
        }

        if (!subtree) {
            continue;
        }

        CPoint ptWithoutDpi;

        if (rootWnd) {
            UINT dpi = ::GetDpiForWindow(rootWnd);
            ptWithoutDpi.x = MulDiv(pt.x, 96, dpi);
            ptWithoutDpi.y = MulDiv(pt.y, 96, dpi);
        } else {
            ptWithoutDpi = pt;
        }

        if (!rootElementRect->PtInRect(ptWithoutDpi)) {
            continue;
        }

        CPoint ptWithoutDpiRelative = ptWithoutDpi;
        ptWithoutDpiRelative.Offset(-rootElementRect->TopLeft());

        InstanceHandle foundHandle =
            ElementFromPointInSubtree(subtree, ptWithoutDpiRelative);
        if (foundHandle) {
            return foundHandle;
        }
    }

    return 0;
}

InstanceHandle CMainDlg::ElementFromPointInSubtree(wux::UIElement subtree,
                                                   CPoint pt) {
    wf::Rect rect(wf::Point{static_cast<float>(pt.x), static_cast<float>(pt.y)},
                  wf::Size{1, 1});

    auto elements = winrt::Windows::UI::Xaml::Media::VisualTreeHelper::
        FindElementsInHostCoordinates(rect, subtree);

    for (auto element : elements) {
        InstanceHandle handle;
        HRESULT hr = m_xamlDiagnostics->GetHandleFromIInspectable(
            reinterpret_cast<::IInspectable*>(winrt::get_abi(element)),
            &handle);
        if (FAILED(hr)) {
            continue;
        }

        return handle;
    }

    return 0;
}

bool CMainDlg::CreateFlashArea(InstanceHandle handle) {
    wf::IInspectable element;
    wf::IInspectable rootElement;

    InstanceHandle iterHandle = handle;
    while (true) {
        auto it = m_elementItems.find(iterHandle);
        if (it == m_elementItems.end()) {
            ATLASSERT(FALSE);
            return false;
        }

        if (!it->second.parentHandle) {
            HRESULT hr = m_xamlDiagnostics->GetIInspectableFromHandle(
                it->first, reinterpret_cast<::IInspectable**>(
                               winrt::put_abi(rootElement)));
            if (FAILED(hr) || !rootElement) {
                return false;
            }

            break;
        }

        if (!element) {
            HRESULT hr = m_xamlDiagnostics->GetIInspectableFromHandle(
                it->first,
                reinterpret_cast<::IInspectable**>(winrt::put_abi(element)));
            if (FAILED(hr) || !element) {
                return false;
            }
        }

        iterHandle = it->second.parentHandle;
    }

    CWindow rootWnd;
    auto rootElementRect = GetRootElementRect(rootElement, &rootWnd.m_hWnd);
    if (!rootElementRect) {
        return false;
    }

    CRect rect;
    if (element) {
        auto elementRect = GetRelativeElementRect(element);
        if (!elementRect) {
            return false;
        }

        rect = *elementRect;
        rect.OffsetRect(rootElementRect->TopLeft());
    } else {
        rect = *rootElementRect;
    }

    if (rootWnd) {
        UINT dpi = ::GetDpiForWindow(rootWnd);
        CRect rectWithDpi;
        rectWithDpi.left = MulDiv(rect.left, dpi, 96);
        rectWithDpi.top = MulDiv(rect.top, dpi, 96);
        // Round width and height, not right and bottom, for consistent results.
        rectWithDpi.right =
            rectWithDpi.left + MulDiv(rect.right - rect.left, dpi, 96);
        rectWithDpi.bottom =
            rectWithDpi.top + MulDiv(rect.bottom - rect.top, dpi, 96);

        rect = rectWithDpi;
    }

    DestroyFlashArea();

    if (rect.IsRectEmpty()) {
        return true;
    }

    m_flashAreaWindow = FlashArea(m_hWnd, _Module.GetModuleInstance(), rect);
    if (!m_flashAreaWindow) {
        return false;
    }

    return true;
}

void CMainDlg::DestroyFlashArea() {
    if (m_flashAreaWindow) {
        m_flashAreaWindow.DestroyWindow();
    }
}

LRESULT CMainDlg::OnElementTreeSelChaneged(LPNMHDR pnmh) {
    SetSelectedElementInformation();

    return 0;
}

LRESULT CMainDlg::OnElementTreeKeyDown(LPNMHDR pnmh) {
    auto wVKey = reinterpret_cast<LPNMTVKEYDOWN>(pnmh)->wVKey;

    // Ctrl+Shift+C is registered as a hotkey, this is a fallback in case that
    // fails. It seems that UWP sandboxed apps are not allowed to register
    // hotkeys.
    if (wVKey == 'C' && GetKeyState(VK_CONTROL) < 0 &&
        GetKeyState(VK_SHIFT) < 0) {
        SelectElementFromCursor();
    }

    return 0;
}

void CMainDlg::OnSplitToggle(UINT uNotifyCode, int nID, CWindow wndCtl) {
    m_splitModeAttributesExpanded = !m_splitModeAttributesExpanded;
    wndCtl.SetWindowText(m_splitModeAttributesExpanded ? L">" : L"<");

    // Below is a very ugly hack to switch to another DlgResize layout.

    RECT rectDlg = {};
    GetClientRect(&rectDlg);

    SetWindowPos(nullptr, 0, 0, m_sizeDialog.cx, m_sizeDialog.cy,
                 SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

    DlgResize_Init();

    ResizeClient(rectDlg.right, rectDlg.bottom);
}

LRESULT CMainDlg::OnAttributeListDblClk(LPNMHDR pnmh) {
    auto itemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pnmh);
    if (itemActivate->iItem == -1 ||
        (itemActivate->iSubItem != 0 && itemActivate->iSubItem != 1)) {
        return 0;
    }

    auto attributesList = CListViewCtrl(pnmh->hwndFrom);

    if (itemActivate->iSubItem == 0) {
        unsigned int clickedPropertyIndex = static_cast<unsigned int>(
            attributesList.GetItemData(itemActivate->iItem));

        auto propertiesComboBox = CComboBox(GetDlgItem(IDC_PROPERTY_NAME));
        int index = CB_ERR;

        int propertiesCount = propertiesComboBox.GetCount();
        for (int i = 0; i < propertiesCount; i++) {
            unsigned int propertyIndex =
                static_cast<unsigned int>(propertiesComboBox.GetItemData(i));
            if (propertyIndex == clickedPropertyIndex) {
                index = i;
                break;
            }
        }

        if (index != CB_ERR) {
            propertiesComboBox.SetCurSel(index);
            propertiesComboBox.GetLBText(index, m_lastPropertySelection);
        }
    } else if (itemActivate->iSubItem == 1) {
        bool valueShownAsIs =
            (attributesList.GetItemData(itemActivate->iItem) >> 32) & 1;
        if (valueShownAsIs) {
            CString itemValue;
            attributesList.GetItemText(itemActivate->iItem,
                                       itemActivate->iSubItem, itemValue);

            auto propertyValueEdit = CEdit(GetDlgItem(IDC_PROPERTY_VALUE));
            propertyValueEdit.SetWindowText(itemValue);
        }
    }

    return 0;
}

void CMainDlg::OnPropertyNameSelChange(UINT uNotifyCode,
                                       int nID,
                                       CWindow wndCtl) {
    auto propertiesComboBox = CComboBox(wndCtl);
    int index = propertiesComboBox.GetCurSel();
    if (index != CB_ERR) {
        propertiesComboBox.GetLBText(index, m_lastPropertySelection);
    }
}

void CMainDlg::OnPropertyRemove(UINT uNotifyCode, int nID, CWindow wndCtl) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto selectedItem = treeView.GetSelectedItem();
    if (!selectedItem) {
        return;
    }

    auto handle = static_cast<InstanceHandle>(selectedItem.GetData());

    auto propertiesComboBox = CComboBox(GetDlgItem(IDC_PROPERTY_NAME));

    int propertiesComboBoxIndex = propertiesComboBox.GetCurSel();
    if (propertiesComboBoxIndex == CB_ERR) {
        return;
    }

    unsigned int propertyIndex = static_cast<unsigned int>(
        propertiesComboBox.GetItemData(propertiesComboBoxIndex));

    HRESULT hr = m_visualTreeService->ClearProperty(handle, propertyIndex);
    if (FAILED(hr)) {
        auto errorMsg = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        MessageBox(errorMsg.c_str(), L"Error");
        return;
    }

    RefreshSelectedElementInformation();
}

void CMainDlg::OnPropertySet(UINT uNotifyCode, int nID, CWindow wndCtl) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto selectedItem = treeView.GetSelectedItem();
    if (!selectedItem) {
        return;
    }

    auto handle = static_cast<InstanceHandle>(selectedItem.GetData());

    auto propertiesComboBox = CComboBox(GetDlgItem(IDC_PROPERTY_NAME));

    int propertiesComboBoxIndex = propertiesComboBox.GetCurSel();
    if (propertiesComboBoxIndex == CB_ERR) {
        return;
    }

    CString propertyNameAndType;
    propertiesComboBox.GetLBText(propertiesComboBoxIndex, propertyNameAndType);
    unsigned int propertyIndex = static_cast<unsigned int>(
        propertiesComboBox.GetItemData(propertiesComboBoxIndex));

    CString propertyName;
    CString propertyType;
    if (int pos = propertyNameAndType.ReverseFind(L'('); pos != -1) {
        if (pos >= 2 && propertyNameAndType.GetLength() - pos >= 3 &&
            propertyNameAndType[pos - 1] == L' ' &&
            propertyNameAndType[propertyNameAndType.GetLength() - 1] == L')') {
            propertyName = propertyNameAndType.Left(pos - 1);
            propertyType = propertyNameAndType.Mid(
                pos + 1, propertyNameAndType.GetLength() - pos - 2);
        }
    }

    if (propertyName.IsEmpty() || propertyType.IsEmpty()) {
        MessageBox(L"Something went wrong", L"Error");
        return;
    }

    auto propertyValueEdit = CEdit(GetDlgItem(IDC_PROPERTY_VALUE));

    CString propertyValue;
    propertyValueEdit.GetWindowText(propertyValue);

    InstanceHandle newValueHandle;
    HRESULT hr = m_visualTreeService->CreateInstance(
        _bstr_t(propertyType), _bstr_t(propertyValue), &newValueHandle);

    if (SUCCEEDED(hr)) {
        hr = m_visualTreeService->SetProperty(handle, newValueHandle,
                                              propertyIndex);
    }

    if (FAILED(hr)) {
        auto errorMsg = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        MessageBox(errorMsg.c_str(), L"Error");
        return;
    }

    RefreshSelectedElementInformation();
}

void CMainDlg::OnCollapseAll(UINT uNotifyCode, int nID, CWindow wndCtl) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto root = treeView.GetRootItem();
    if (root) {
        if (!m_redrawTreeQueued) {
            treeView.SetRedraw(FALSE);
        }

        treeView.SelectItem(root);
        TreeViewExpandRecursively(treeView, root, TVE_COLLAPSE);
        treeView.EnsureVisible(root);

        if (!m_redrawTreeQueued) {
            treeView.SetRedraw(TRUE);
        }
    }
}

void CMainDlg::OnExpandAll(UINT uNotifyCode, int nID, CWindow wndCtl) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto root = treeView.GetRootItem();
    if (root) {
        if (!m_redrawTreeQueued) {
            treeView.SetRedraw(FALSE);
        }

        TreeViewExpandRecursively(treeView, root, TVE_EXPAND);
        auto selected = treeView.GetSelectedItem();
        treeView.EnsureVisible(selected ? selected : root);

        if (!m_redrawTreeQueued) {
            treeView.SetRedraw(TRUE);
        }
    }
}

void CMainDlg::OnHighlightSelection(UINT uNotifyCode, int nID, CWindow wndCtl) {
    m_highlightSelection = CButton(wndCtl).GetCheck() != BST_UNCHECKED;

    DestroyFlashArea();

    if (m_highlightSelection) {
        auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
        auto selectedItem = treeView.GetSelectedItem();
        if (selectedItem) {
            auto handle = static_cast<InstanceHandle>(selectedItem.GetData());
            CreateFlashArea(handle);
        }
    }
}

void CMainDlg::OnDetailedProperties(UINT uNotifyCode, int nID, CWindow wndCtl) {
    m_detailedProperties = CButton(wndCtl).GetCheck() != BST_UNCHECKED;

    ResetAttributesListColumns();

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto selectedItem = treeView.GetSelectedItem();
    if (selectedItem && selectedItem.GetParent()) {
        auto handle = static_cast<InstanceHandle>(selectedItem.GetData());
        PopulateAttributesList(handle);
    }
}

void CMainDlg::OnAppAbout(UINT uNotifyCode, int nID, CWindow wndCtl) {
    PCWSTR content =
        L"By <A HREF=\"https://ramensoftware.com/\">Ramen Software</A>\n"
        L"\n"
        L"<A HREF=\"https://ramensoftware.com/uwpspy\">Web page</a>\n"
        L"<A HREF=\"https://github.com/m417z/UWPSpy\">Source code</a>\n"
        L"\n"
        L"Tips:\n"
        L"• When focused on the element tree, place your mouse cursor over a "
        L"UWP element and press Ctrl+Shift+C to select it in the tree.\n"
        L"• To have a good overview of an element's surroundings, collapse all "
        L"tree elements, select the desired element with the previous tip, and "
        L"go up the element tree.\n"
        L"• Double-click property list entries to copy the data to the form "
        L"below.";

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
    Hide();
}

LRESULT CMainDlg::OnActivateWindow(UINT uMsg, WPARAM wParam, LPARAM lParam) {
    Show();
    ::SetForegroundWindow(m_hWnd);
    return 0;
}

void CMainDlg::RedrawTreeQueue() {
    if (m_redrawTreeQueued) {
        return;
    }

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));

    auto selectedItem = treeView.GetSelectedItem();
    m_redrawTreeQueuedEnsureSelectionVisible =
        selectedItem && IsTreeItemInView(selectedItem);

    treeView.SetRedraw(FALSE);
    SetTimer(TIMER_ID_REDRAW_TREE, kRedrawTreeDelay);
    m_redrawTreeQueued = true;
}

bool CMainDlg::SetSelectedElementInformation() {
    KillTimer(TIMER_ID_REFRESH_SELECTED_ELEMENT_INFORMATION);

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto selectedItem = treeView.GetSelectedItem();
    if (!selectedItem) {
        return false;
    }

    auto handle = static_cast<InstanceHandle>(selectedItem.GetData());
    bool hasParent = !!selectedItem.GetParent();

    wf::IInspectable obj;
    try {
        winrt::check_hresult(m_xamlDiagnostics->GetIInspectableFromHandle(
            handle, reinterpret_cast<::IInspectable**>(winrt::put_abi(obj))));
        auto className = winrt::get_class_name(obj);
        SetDlgItemText(IDC_CLASS_EDIT, className.c_str());
    } catch (...) {
        obj = nullptr;

        HRESULT hr = winrt::to_hresult();
        auto errorMsg = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        SetDlgItemText(IDC_CLASS_EDIT, errorMsg.c_str());
    }

    if (auto frameworkElement = obj.try_as<wux::FrameworkElement>()) {
        try {
            SetDlgItemText(IDC_NAME_EDIT, frameworkElement.Name().c_str());
        } catch (...) {
            HRESULT hr = winrt::to_hresult();
            auto errorMsg =
                std::format(L"Error {:08X}", static_cast<DWORD>(hr));
            SetDlgItemText(IDC_NAME_EDIT, errorMsg.c_str());
        }
    } else {
        SetDlgItemText(IDC_NAME_EDIT, L"");
    }

    if (obj) {
        try {
            auto rect = hasParent ? GetRelativeElementRect(obj)
                                  : GetRootElementRect(obj);
            if (rect) {
                auto rectStr = std::format(
                    L"({},{}) - ({},{})  -  {}x{}", rect->left, rect->top,
                    rect->right, rect->bottom, rect->Width(), rect->Height());
                SetDlgItemText(IDC_RECT_EDIT, rectStr.c_str());
            } else {
                SetDlgItemText(IDC_RECT_EDIT, L"Unknown");
            }
        } catch (...) {
            HRESULT hr = winrt::to_hresult();
            auto errorMsg =
                std::format(L"Error {:08X}", static_cast<DWORD>(hr));
            SetDlgItemText(IDC_RECT_EDIT, errorMsg.c_str());
        }
    } else {
        SetDlgItemText(IDC_RECT_EDIT, L"");
    }

    if (hasParent) {
        PopulateAttributesList(handle);
    } else {
        auto attributesList = CListViewCtrl(GetDlgItem(IDC_ATTRIBUTE_LIST));
        attributesList.DeleteAllItems();

        auto propertiesComboBox = CComboBox(GetDlgItem(IDC_PROPERTY_NAME));
        propertiesComboBox.ResetContent();
    }

    DestroyFlashArea();

    if (m_highlightSelection) {
        CreateFlashArea(handle);
    }

    return true;
}

bool CMainDlg::RefreshSelectedElementInformation(UINT delay) {
    // Use a delay since some information is not available immediately after
    // changing properties, such as the new element position and size.
    if (delay > 0) {
        SetTimer(TIMER_ID_REFRESH_SELECTED_ELEMENT_INFORMATION, delay);
        return true;
    }

    auto attributesList = CListViewCtrl(GetDlgItem(IDC_ATTRIBUTE_LIST));
    int attributesListTopIndex = attributesList.GetTopIndex();

    if (!SetSelectedElementInformation()) {
        return false;
    }

    ListViewSetTopIndex(attributesList, attributesListTopIndex);
    return true;
}

void CMainDlg::ResetAttributesListColumns() {
    auto attributesList = CListViewCtrl(GetDlgItem(IDC_ATTRIBUTE_LIST));

    attributesList.SetRedraw(FALSE);

    attributesList.DeleteAllItems();

    // If columns exist, try to retain width.
    int width = attributesList.GetColumnWidth(0) * 2;
    if (!width) {
        CRect rect;
        attributesList.GetClientRect(rect);
        width = rect.Width() - ::GetSystemMetrics(SM_CXVSCROLL);
    }

    while (attributesList.DeleteColumn(0)) {
        // Keep deleting...
    }

    int c = 0;

    if (m_detailedProperties) {
        attributesList.InsertColumn(c++, L"Name", LVCFMT_LEFT, width / 2);
        attributesList.InsertColumn(c++, L"Value", LVCFMT_LEFT, width / 3);
        attributesList.InsertColumn(c++, L"Type", LVCFMT_LEFT, width / 3);
        attributesList.InsertColumn(c++, L"DeclaringType", LVCFMT_LEFT,
                                    width / 3);
        attributesList.InsertColumn(c++, L"ValueType", LVCFMT_LEFT, width / 3);
        attributesList.InsertColumn(c++, L"ItemType", LVCFMT_LEFT, width / 3);
        attributesList.InsertColumn(c++, L"Overridden", LVCFMT_LEFT, width / 3);
        attributesList.InsertColumn(c++, L"MetadataBits", LVCFMT_LEFT,
                                    width / 3);
        attributesList.InsertColumn(c++, L"Style TargetType", LVCFMT_LEFT,
                                    width / 3);
        attributesList.InsertColumn(c++, L"Style Name", LVCFMT_LEFT, width / 3);
        attributesList.InsertColumn(c++, L"Style Source", LVCFMT_LEFT,
                                    width / 3);
    } else {
        attributesList.InsertColumn(c++, L"Name", LVCFMT_LEFT, width / 2);
        attributesList.InsertColumn(c++, L"Value", LVCFMT_LEFT, width / 2);
    }

    attributesList.SetRedraw(TRUE);
}

void CMainDlg::PopulateAttributesList(InstanceHandle handle) {
    auto attributesList = CListViewCtrl(GetDlgItem(IDC_ATTRIBUTE_LIST));
    attributesList.SetRedraw(FALSE);

    attributesList.DeleteAllItems();

    auto propertiesComboBox = CComboBox(GetDlgItem(IDC_PROPERTY_NAME));
    propertiesComboBox.ResetContent();

    unsigned int sourceCount = 0;
    PropertyChainSource* pPropertySources = nullptr;
    unsigned int propertyCount = 0;
    PropertyChainValue* pPropertyValues = nullptr;
    HRESULT hr = m_visualTreeService->GetPropertyValuesChain(
        handle, &sourceCount, &pPropertySources, &propertyCount,
        &pPropertyValues);
    if (FAILED(hr)) {
        auto errorMsg = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        attributesList.AddItem(0, 0, errorMsg.c_str());

        attributesList.SetRedraw(TRUE);
        return;
    }

    const auto metadataBitsToString = [](hyper metadataBits) {
        std::wstring str;

#define APPEND_BIT_TO_STR(x)   \
    if (metadataBits & x) {    \
        str += TEXT(#x) L", "; \
        metadataBits &= ~x;    \
    }

        APPEND_BIT_TO_STR(IsValueHandle);
        APPEND_BIT_TO_STR(IsPropertyReadOnly);
        APPEND_BIT_TO_STR(IsValueCollection);
        APPEND_BIT_TO_STR(IsValueCollectionReadOnly);
        APPEND_BIT_TO_STR(IsValueBindingExpression);
        APPEND_BIT_TO_STR(IsValueNull);
        APPEND_BIT_TO_STR(IsValueHandleAndEvaluatedValue);

#undef APPEND_BIT_TO_STR

        if (metadataBits) {
            str += std::to_wstring(metadataBits) + L", ";
        }

        if (!str.empty()) {
            str.resize(str.length() - 2);
        }

        return str;
    };

    const auto sourceToString = [](BaseValueSource source) {
        std::wstring str;

#define CASE_TO_STR(x)  \
    case x:             \
        str = TEXT(#x); \
        break;

        switch (source) {
            CASE_TO_STR(BaseValueSourceUnknown);
            CASE_TO_STR(BaseValueSourceDefault);
            CASE_TO_STR(BaseValueSourceBuiltInStyle);
            CASE_TO_STR(BaseValueSourceStyle);
            CASE_TO_STR(BaseValueSourceLocal);
            CASE_TO_STR(Inherited);
            CASE_TO_STR(DefaultStyleTrigger);
            CASE_TO_STR(TemplateTrigger);
            CASE_TO_STR(StyleTrigger);
            CASE_TO_STR(ImplicitStyleReference);
            CASE_TO_STR(ParentTemplate);
            CASE_TO_STR(ParentTemplateTrigger);
            CASE_TO_STR(Animation);
            CASE_TO_STR(Coercion);
            CASE_TO_STR(BaseValueSourceVisualState);
            default:
                str = std::to_wstring(source);
                break;
        }

#undef CASE_TO_STR

        return str;
    };

    int row = 0;
    for (unsigned int i = 0; i < propertyCount; i++) {
        const auto& v = pPropertyValues[i];

        const auto& src = pPropertySources[v.PropertyChainIndex];

        if (!v.Overridden) {
            auto comboBoxItemText =
                std::format(L"{} ({})", v.PropertyName, v.Type);
            int index = propertiesComboBox.AddString(comboBoxItemText.data());
            if (index != CB_ERR && index != CB_ERRSPACE) {
                propertiesComboBox.SetItemData(index, v.Index);
            }
        }

        if (!m_detailedProperties && src.Source != BaseValueSourceLocal) {
            continue;
        }

        std::wstring value;
        bool valueShownAsIs = false;

        if (m_detailedProperties) {
            value = v.Value;
            valueShownAsIs = true;
        } else if (v.MetadataBits & IsValueNull) {
            value = L"(null)";
        } else if (v.MetadataBits & IsValueHandle) {
            InstanceHandle valueHandle =
                static_cast<InstanceHandle>(std::wcstoll(v.Value, nullptr, 10));

            std::wstring className;

            wf::IInspectable valueObj;
            HRESULT hr = m_xamlDiagnostics->GetIInspectableFromHandle(
                valueHandle,
                reinterpret_cast<::IInspectable**>(winrt::put_abi(valueObj)));
            if (SUCCEEDED(hr)) {
                className = winrt::get_class_name(valueObj);
            } else {
                className =
                    std::format(L"Error {:08X}", static_cast<DWORD>(hr));
            }

            value = std::format(
                L"({}; {})",
                (v.MetadataBits & IsValueCollection) ? L"collection" : L"data",
                className);
        } else {
            value = v.Value;
            valueShownAsIs = true;
        }

        int c = 0;

        attributesList.AddItem(row, c++, v.PropertyName);
        attributesList.AddItem(row, c++, value.c_str());

        if (m_detailedProperties) {
            attributesList.AddItem(row, c++, v.Type);
            attributesList.AddItem(row, c++, v.DeclaringType);
            attributesList.AddItem(row, c++, v.ValueType);
            attributesList.AddItem(row, c++, v.ItemType);
            attributesList.AddItem(row, c++, v.Overridden ? L"Yes" : L"No");
            attributesList.AddItem(
                row, c++, metadataBitsToString(v.MetadataBits).c_str());
            attributesList.AddItem(row, c++, src.TargetType);
            attributesList.AddItem(row, c++, src.Name);
            attributesList.AddItem(row, c++,
                                   sourceToString(src.Source).c_str());
        }

        // Passes on 64-bit, not on 32-bit. Can be fixed later if a 32-bit build
        // is needed.
        static_assert(sizeof(DWORD_PTR) >= sizeof(v.Index) + 1);
        static_assert(sizeof(v.Index) == 4);
        attributesList.SetItemData(
            row,
            (static_cast<DWORD_PTR>(valueShownAsIs ? 1 : 0) << 32) | v.Index);

        row++;
    }

    if (m_lastPropertySelection.IsEmpty() ||
        propertiesComboBox.SelectString(0, m_lastPropertySelection) == CB_ERR) {
        propertiesComboBox.SetCurSel(0);
    }

    propertiesComboBox.SetDroppedWidth(
        GetRequiredComboDroppedWidth(propertiesComboBox));

    // Not documented, but it makes sense that the arrays have to be
    // freed and this seems to be working.
    CoTaskMemFree(pPropertySources);
    CoTaskMemFree(pPropertyValues);

    attributesList.SetRedraw(TRUE);
}

void CMainDlg::AddItemToTree(HTREEITEM parentTreeItem,
                             HTREEITEM insertAfter,
                             InstanceHandle handle,
                             ElementItem* elementItem) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));

    ATLASSERT(!elementItem->treeItem);

    // Passes on 64-bit, not on 32-bit. Can be fixed later if a 32-bit build is
    // needed.
    static_assert(sizeof(LPARAM) >= sizeof(handle));

    TVINSERTSTRUCT insertStruct{
        .hParent = parentTreeItem,
        .hInsertAfter = insertAfter,
        .item =
            {
                .mask = TVIF_TEXT | TVIF_PARAM | TVIF_STATE,
                .state = TVIS_EXPANDED,
                .stateMask = TVIS_EXPANDED,
                .pszText = const_cast<PWSTR>(elementItem->itemTitle.c_str()),
                .lParam = static_cast<LPARAM>(handle),
            },
    };
    HTREEITEM insertedItem = treeView.InsertItem(&insertStruct);
    if (!insertedItem) {
        ATLASSERT(FALSE);
        return;
    }

    elementItem->treeItem = insertedItem;

    if (auto it = m_parentToChildren.find(handle);
        it != m_parentToChildren.end()) {
        for (const auto& childHandle : it->second) {
            auto childElementItem = m_elementItems.find(childHandle);
            if (childElementItem == m_elementItems.end()) {
                ATLASSERT(FALSE);
                continue;
            }

            AddItemToTree(insertedItem, TVI_LAST, childHandle,
                          &childElementItem->second);
        }
    }
}

bool CMainDlg::SelectElementFromCursor() {
    CPoint pt;
    ::GetCursorPos(&pt);

    InstanceHandle handle = ElementFromPoint(pt);
    if (!handle) {
        return false;
    }

    auto it = m_elementItems.find(handle);
    if (it == m_elementItems.end()) {
        return false;
    }

    auto treeItem = it->second.treeItem;
    if (!treeItem) {
        return false;
    }

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    treeView.SelectItem(treeItem);
    treeView.EnsureVisible(treeItem);
    return true;
}
