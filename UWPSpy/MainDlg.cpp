#include "stdafx.h"

#include "MainDlg.h"

#include "AboutDlg.h"
#include "flash_area.h"

namespace {

// Delay each tree redraw, redrawing at most once every kRedrawTreeDelay ms.
// Otherwise, multiple redraw operations can make the UI very slow.
constexpr UINT kRedrawTreeDelay = 200;

bool CopyTextToClipboard(HWND hWndParent, std::wstring_view text) {
    size_t bytes = (text.size() + 1) * sizeof(WCHAR);
    HGLOBAL hClipboardData = ::GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!hClipboardData) {
        return false;
    }

    void* pClipboardData = ::GlobalLock(hClipboardData);
    if (!pClipboardData) {
        ::GlobalFree(hClipboardData);
        return false;
    }

    memcpy(pClipboardData, text.data(), text.size() * sizeof(WCHAR));
    // Add null terminator.
    ((WCHAR*)pClipboardData)[text.size()] = L'\0';
    ::GlobalUnlock(hClipboardData);

    BOOL openedClipboard = ::OpenClipboard(hWndParent);
    if (!openedClipboard) {
        ::GlobalFree(hClipboardData);
        return false;
    }

    ::EmptyClipboard();

    // Let the OS synthesize CF_TEXT.
    HANDLE handle = ::SetClipboardData(CF_UNICODETEXT, hClipboardData);
    ::CloseClipboard();

    if (!handle) {
        ::GlobalFree(hClipboardData);
        return false;
    }

    return true;
}

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

void TreeViewSubtreeToStringRecursively(CTreeViewCtrlEx tree,
                                        CTreeItem treeItem,
                                        CString& str,
                                        int depth) {
    CString linePrefix(L' ', depth * 2);

    while (treeItem) {
        CString itemText;
        treeItem.GetText(itemText);
        str += linePrefix;
        str += itemText;
        str += L'\n';
        CTreeItem childItem = tree.GetChildItem(treeItem);
        if (childItem) {
            TreeViewSubtreeToStringRecursively(tree, childItem, str, depth + 1);
        }
        treeItem = tree.GetNextSiblingItem(treeItem);
    }
}

void TreeViewSubtreeToString(CTreeViewCtrlEx tree,
                             CTreeItem treeItem,
                             CString& str) {
    treeItem.GetText(str);
    str += L'\n';
    CTreeItem childItem = tree.GetChildItem(treeItem);
    if (childItem) {
        TreeViewSubtreeToStringRecursively(tree, childItem, str, 1);
    }
}

void TreeViewToString(CTreeViewCtrlEx tree, CString& str) {
    str.Empty();
    CTreeItem treeItem = tree.GetRootItem();
    while (treeItem) {
        CString itemText;
        treeItem.GetText(itemText);
        str += itemText;
        str += L'\n';
        CTreeItem childItem = tree.GetChildItem(treeItem);
        if (childItem) {
            TreeViewSubtreeToStringRecursively(tree, childItem, str, 1);
        }
        treeItem = tree.GetNextSiblingItem(treeItem);
    }
}

std::wstring BuildElementPath(
    const std::unordered_map<InstanceHandle, CMainDlg::ElementItem>&
        elementItems,
    InstanceHandle handle) {
    // Build path components from child to root
    std::vector<std::wstring> pathComponents;

    InstanceHandle currentHandle = handle;
    while (currentHandle) {
        auto it = elementItems.find(currentHandle);
        if (it == elementItems.end()) {
            break;
        }

        const auto& elementItem = it->second;

        // Stop before adding root element (root has parentHandle == 0)
        if (!elementItem.parentHandle) {
            break;
        }

        // Parse itemTitle to extract ClassName and Name
        // Format is "ClassName - Name" or just "ClassName"
        const std::wstring& itemTitle = elementItem.itemTitle;
        std::wstring component;

        size_t separatorPos = itemTitle.find(L" - ");
        if (separatorPos != std::wstring::npos) {
            // Has name, format as "ClassName#Name"
            std::wstring className = itemTitle.substr(0, separatorPos);
            // Skip " - "
            std::wstring name = itemTitle.substr(separatorPos + 3);
            component = className + L"#" + name;
        } else {
            // No name, just use ClassName
            component = itemTitle;
        }

        pathComponents.push_back(component);
        currentHandle = elementItem.parentHandle;
    }

    // Reverse to get root-to-child order
    std::reverse(pathComponents.begin(), pathComponents.end());

    // Join with " > " separator
    std::wstring path;
    for (size_t i = 0; i < pathComponents.size(); ++i) {
        if (i > 0) {
            path += L" > ";
        }
        path += pathComponents[i];
    }

    return path;
}

bool IsTreeItemInView(CTreeItem treeItem) {
    CRect rect;
    if (!treeItem.GetRect(rect, FALSE)) {
        return false;
    }

    CRect clientRect;
    treeItem.GetTreeView()->GetClientRect(clientRect);

    CRect intersectionRect;
    return intersectionRect.IntersectRect(rect, clientRect);
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

// Can be carefully used to pass a guid of a different version of an interface,
// and use the resulting object with functions that are the same in both
// interface versions.
template <typename To,
          typename From,
          std::enable_if_t<winrt::impl::is_com_interface_v<To>, int> = 0>
winrt::impl::com_ref<To> try_as_with_guid_unsafe(From* ptr,
                                                 winrt::guid guid) noexcept {
    if (!ptr) {
        return nullptr;
    }

    void* result{};
    ptr->QueryInterface(guid, &result);
    return winrt::impl::wrap_as_result<To>(result);
}

// 1.2.220727.1-experimental1 - 1.2.220909.2-experimental2
// {A81014CD-55C1-506E-AA79-D9AC96DB9B8E}
constexpr winrt::guid IID_IDesktopWindowXamlSource_WinUI_1{
    0xa81014cd,
    0x55c1,
    0x506e,
    {0xaa, 0x79, 0xd9, 0xac, 0x96, 0xdb, 0x9b, 0x8e}};

// 1.3.230202101-experimental1 - 1.4.230518007-experimental1
// {F2AA238F-1C21-581E-AADC-7D6EC5320F56}
constexpr winrt::guid IID_IDesktopWindowXamlSource_WinUI_2{
    0xf2aa238f,
    0x1c21,
    0x581e,
    {0xaa, 0xdc, 0x7d, 0x6e, 0xc5, 0x32, 0x0f, 0x56}};

std::optional<CRect> GetRelativeElementRect(wf::IInspectable element) {
    std::optional<wf::Point> offset;
    std::optional<wf::Numerics::float2> size;
    if (auto uiElement = element.try_as<wux::UIElement>()) {
        offset = uiElement.TransformToVisual(nullptr).TransformPoint(
            wf::Point(0, 0));
        size = uiElement.ActualSize();
    } else if (auto uiElement = element.try_as<mux::UIElement>()) {
        offset = uiElement.TransformToVisual(nullptr).TransformPoint(
            wf::Point(0, 0));
        size = uiElement.ActualSize();
    }

    if (offset && size) {
        CRect rect;
        rect.left = std::lroundf(offset->X);
        rect.top = std::lroundf(offset->Y);
        rect.right = rect.left + std::lroundf(size->x);
        rect.bottom = rect.top + std::lroundf(size->y);

        return rect;
    }

    return std::nullopt;
}

std::optional<CRect> GetRootElementRect(wf::IInspectable element,
                                        HWND* outWnd = nullptr) {
    std::optional<wf::Rect> bounds;
    winrt::Windows::UI::Core::CoreWindow coreWindow = nullptr;
    if (auto window = element.try_as<wux::Window>()) {
        bounds = window.Bounds();
        if (outWnd) {
            coreWindow = window.CoreWindow();
        }
    } else if (auto window = element.try_as<mux::Window>()) {
        bounds = window.Bounds();
        if (outWnd) {
            coreWindow = window.CoreWindow();
        }
    }

    if (bounds) {
        CRect rect;
        rect.left = std::lroundf(bounds->X);
        rect.top = std::lroundf(bounds->Y);
        rect.right = rect.left + std::lroundf(bounds->Width);
        rect.bottom = rect.top + std::lroundf(bounds->Height);

        if (outWnd && coreWindow) {
            if (auto coreInterop = coreWindow.try_as<ICoreWindowInterop>()) {
                coreInterop->get_WindowHandle(outWnd);
            }
        }

        return rect;
    }

    CWindow nativeWnd;
    if (auto nativeSource = element.try_as<IDesktopWindowXamlSourceNative>()) {
        nativeSource->get_WindowHandle(&nativeWnd.m_hWnd);
    } else if (auto desktopWindowXamlSource =
                   element.try_as<mux::Hosting::DesktopWindowXamlSource>()) {
        auto appWindowId = desktopWindowXamlSource.SiteBridge().WindowId();
        nativeWnd = winrt::Microsoft::UI::GetWindowFromWindowId(appWindowId);
    } else if (auto desktopWindowXamlSource = try_as_with_guid_unsafe<
                   mux::Hosting::DesktopWindowXamlSource>(
                   reinterpret_cast<::IInspectable*>(winrt::get_abi(element)),
                   IID_IDesktopWindowXamlSource_WinUI_2)) {
        auto appWindowId = desktopWindowXamlSource.SiteBridge().WindowId();
        nativeWnd = winrt::Microsoft::UI::GetWindowFromWindowId(appWindowId);
    } else if (auto nativeSource =
                   element.try_as<IDesktopWindowXamlSourceNative_WinUI>()) {
        nativeSource->get_WindowHandle(&nativeWnd.m_hWnd);
    }

    if (nativeWnd) {
        // A window which is no longer valid might be returned if the element is
        // being destroyed.
        if (!nativeWnd.IsWindow()) {
            return std::nullopt;
        }

        CRect rectWithDpi;
        nativeWnd.GetClientRect(rectWithDpi);
        nativeWnd.MapWindowPoints(nullptr, &rectWithDpi);

        UINT dpi = ::GetDpiForWindow(nativeWnd);
        CRect rect;
        rect.left = MulDiv(rectWithDpi.left, 96, dpi);
        rect.top = MulDiv(rectWithDpi.top, 96, dpi);
        // Round width and height, not right and bottom, for consistent results.
        rect.right =
            rect.left + MulDiv(rectWithDpi.right - rectWithDpi.left, 96, dpi);
        rect.bottom =
            rect.top + MulDiv(rectWithDpi.bottom - rectWithDpi.top, 96, dpi);

        if (outWnd) {
            *outWnd = nativeWnd;
        }

        return rect;
    }

    return std::nullopt;
}

template <typename GridLengthType>
std::wstring FormatGridLength(const GridLengthType& gridLength) {
    // GridLength is a struct with: double Value, GridUnitType GridUnitType
    // GridUnitType enum: Auto=0, Pixel=1, Star=2

    if (gridLength.GridUnitType == decltype(gridLength.GridUnitType)::Auto) {
        return L"Auto";
    } else if (gridLength.GridUnitType ==
               decltype(gridLength.GridUnitType)::Star) {
        if (gridLength.Value == 1.0) {
            return L"*";
        } else {
            return std::format(L"{}*", gridLength.Value);
        }
    } else {  // Pixel (Absolute)
        return std::format(L"{}px", gridLength.Value);
    }
}

template <typename CollectionType>
std::wstring FormatColumnDefinitionCollection(
    const CollectionType& collection) {
    uint32_t size = collection.Size();
    if (size == 0) {
        return L" {Size=0}";
    }

    std::wstring result = L" [";

    for (uint32_t i = 0; i < size; i++) {
        if (i > 0) {
            result += L", ";
        }

        auto definition = collection.GetAt(i);
        result += L"{Width=";
        result += FormatGridLength(definition.Width());

        // Only include MinWidth if != 0
        double minWidth = definition.MinWidth();
        if (minWidth != 0.0) {
            result += std::format(L", MinWidth={}", minWidth);
        }

        // Only include MaxWidth if != Infinity
        double maxWidth = definition.MaxWidth();
        if (!std::isinf(maxWidth)) {
            result += std::format(L", MaxWidth={}", maxWidth);
        }

        result += L"}";
    }

    result += L"]";
    return result;
}

template <typename CollectionType>
std::wstring FormatRowDefinitionCollection(const CollectionType& collection) {
    uint32_t size = collection.Size();
    if (size == 0) {
        return L" {Size=0}";
    }

    std::wstring result = L" [";

    for (uint32_t i = 0; i < size; i++) {
        if (i > 0) {
            result += L", ";
        }

        auto definition = collection.GetAt(i);
        result += L"{Height=";
        result += FormatGridLength(definition.Height());

        // Only include MinHeight if != 0
        double minHeight = definition.MinHeight();
        if (minHeight != 0.0) {
            result += std::format(L", MinHeight={}", minHeight);
        }

        // Only include MaxHeight if != Infinity
        double maxHeight = definition.MaxHeight();
        if (!std::isinf(maxHeight)) {
            result += std::format(L", MaxHeight={}", maxHeight);
        }

        result += L"}";
    }

    result += L"]";
    return result;
}

struct FormattedPropertyValue {
    std::wstring value;
    bool valueShownAsIs;
};

FormattedPropertyValue FormatPropertyValue(const PropertyChainValue& v,
                                           IXamlDiagnostics* xamlDiagnostics) {
    if (v.MetadataBits & IsValueNull) {
        return {L"(null)", false};
    }

    if (v.MetadataBits & IsValueHandle) {
        InstanceHandle valueHandle =
            static_cast<InstanceHandle>(std::wcstoll(v.Value, nullptr, 10));

        std::wstring className;

        wf::IInspectable valueObj;
        HRESULT hr = xamlDiagnostics->GetIInspectableFromHandle(
            valueHandle,
            reinterpret_cast<::IInspectable**>(winrt::put_abi(valueObj)));
        if (SUCCEEDED(hr)) {
            className = winrt::get_class_name(valueObj);
        } else {
            className = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        }

        std::wstring extraInfo;
        if (auto vector =
                valueObj.try_as<wf::Collections::IVector<wf::IInspectable>>()) {
            extraInfo = std::format(L" {{Size={}}}", vector.Size());
        } else if (auto val = valueObj.try_as<wux::Media::SolidColorBrush>()) {
            auto color = val.Color();
            extraInfo =
                color.A == 0xFF
                    ? std::format(L" {{Color=#{:02X}{:02X}{:02X}}}", color.R,
                                  color.G, color.B)
                    : std::format(L" {{Color=#{:02X}{:02X}{:02X}{:02X}}}",
                                  color.A, color.R, color.G, color.B);
        } else if (auto val = valueObj.try_as<mux::Media::SolidColorBrush>()) {
            auto color = val.Color();
            extraInfo =
                color.A == 0xFF
                    ? std::format(L" {{Color=#{:02X}{:02X}{:02X}}}", color.R,
                                  color.G, color.B)
                    : std::format(L" {{Color=#{:02X}{:02X}{:02X}{:02X}}}",
                                  color.A, color.R, color.G, color.B);
        } else if (auto val = valueObj.try_as<wux::Controls::Grid>()) {
            extraInfo = std::format(L" {{Columns={}, Rows={}}}",
                                    val.ColumnDefinitions().Size(),
                                    val.RowDefinitions().Size());
        } else if (auto val = valueObj.try_as<mux::Controls::Grid>()) {
            extraInfo = std::format(L" {{Columns={}, Rows={}}}",
                                    val.ColumnDefinitions().Size(),
                                    val.RowDefinitions().Size());
        } else if (auto val = valueObj.try_as<
                              wux::Controls::ColumnDefinitionCollection>()) {
            extraInfo = FormatColumnDefinitionCollection(val);
        } else if (auto val = valueObj.try_as<
                              mux::Controls::ColumnDefinitionCollection>()) {
            extraInfo = FormatColumnDefinitionCollection(val);
        } else if (auto val =
                       valueObj
                           .try_as<wux::Controls::RowDefinitionCollection>()) {
            extraInfo = FormatRowDefinitionCollection(val);
        } else if (auto val =
                       valueObj
                           .try_as<mux::Controls::RowDefinitionCollection>()) {
            extraInfo = FormatRowDefinitionCollection(val);
        }

        return {std::format(L"({}; {}{})",
                            (v.MetadataBits & IsValueCollection) ? L"collection"
                                                                 : L"data",
                            className, extraInfo),
                false};
    }

    return {v.Value, true};
}

// https://stackoverflow.com/a/5665377
std::wstring EscapeXmlAttribute(std::wstring_view data) {
    std::wstring buffer;
    buffer.reserve(data.size());
    for (size_t pos = 0; pos != data.size(); ++pos) {
        switch (data[pos]) {
            case '&':
                buffer.append(L"&amp;");
                break;
            case '\"':
                buffer.append(L"&quot;");
                break;
            // case '\'':
            //     buffer.append(L"&apos;");
            //     break;
            case '<':
                buffer.append(L"&lt;");
                break;
            case '>':
                buffer.append(L"&gt;");
                break;
            default:
                buffer.append(&data[pos], 1);
                break;
        }
    }

    return buffer;
}

std::wstring GetResourceDictionaryXamlFromXamlSetters(
    const std::wstring_view type,
    const std::wstring_view xamlStyleSetters) {
    std::wstring xaml =
        LR"(<ResourceDictionary
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:muxc="using:Microsoft.UI.Xaml.Controls")";

    if (auto pos = type.rfind('.'); pos != type.npos) {
        auto typeNamespace = std::wstring_view(type).substr(0, pos);
        auto typeName = std::wstring_view(type).substr(pos + 1);

        xaml += L"\n    xmlns:uwpspy=\"using:";
        xaml += EscapeXmlAttribute(typeNamespace);
        xaml +=
            L"\">\n"
            L"    <Style TargetType=\"uwpspy:";
        xaml += EscapeXmlAttribute(typeName);
        xaml += L"\">\n";
    } else {
        xaml +=
            L">\n"
            L"    <Style TargetType=\"";
        xaml += EscapeXmlAttribute(type);
        xaml += L"\">\n";
    }

    xaml += xamlStyleSetters;

    xaml +=
        L"    </Style>\n"
        L"</ResourceDictionary>";

    return xaml;
}

wux::Style GetStyleFromXamlSettersWux(
    const std::wstring_view type,
    const std::wstring_view xamlStyleSetters) {
    std::wstring xaml =
        GetResourceDictionaryXamlFromXamlSetters(type, xamlStyleSetters);

    auto resourceDictionary =
        wux::Markup::XamlReader::Load(xaml).as<wux::ResourceDictionary>();

    auto [styleKey, styleInspectable] = resourceDictionary.First().Current();
    return styleInspectable.as<wux::Style>();
}

mux::Style GetStyleFromXamlSettersMux(
    const std::wstring_view type,
    const std::wstring_view xamlStyleSetters) {
    std::wstring xaml =
        GetResourceDictionaryXamlFromXamlSetters(type, xamlStyleSetters);

    auto resourceDictionary =
        mux::Markup::XamlReader::Load(xaml).as<mux::ResourceDictionary>();

    auto [styleKey, styleInspectable] = resourceDictionary.First().Current();
    return styleInspectable.as<mux::Style>();
}

wf::IInspectable StyleValueFromXaml(wf::IInspectable element,
                                    const std::wstring_view name,
                                    const std::wstring_view value) {
    std::wstring xaml;

    xaml += L"        <Setter Property=\"";
    xaml += EscapeXmlAttribute(name);
    xaml += L"\"";
    xaml +=
        L">\n"
        L"            <Setter.Value>\n";
    xaml += value;
    xaml +=
        L"\n"
        L"            </Setter.Value>\n"
        L"        </Setter>\n";

    auto className = winrt::get_class_name(element);

    if (auto uiElement = element.try_as<wux::UIElement>()) {
        auto style = GetStyleFromXamlSettersWux(className, xaml);
        return style.Setters().GetAt(0).as<wux::Setter>().Value();
    } else if (auto uiElement = element.try_as<mux::UIElement>()) {
        auto style = GetStyleFromXamlSettersMux(className, xaml);
        return style.Setters().GetAt(0).as<mux::Setter>().Value();
    }

    winrt::throw_hresult(E_NOTIMPL);
}

// Helper functions for 32-bit/64-bit compatibility for storing 64-bit values in
// LPARAM/DWORD_PTR fields.
//
// On 64-bit builds, LPARAM and DWORD_PTR are 64 bits, so we can store 64-bit
// values directly.
//
// On 32-bit builds, LPARAM and DWORD_PTR are only 32 bits, which is too small
// to hold a 64-bit InstanceHandle or a 64-bit composite value. As a workaround,
// we allocate the 64-bit value on the heap and store the pointer instead.
//
// Deliberate memory leak: The 32-bit *ToLParam/*ToDWordPtr functions allocate
// memory that is never freed. This is a quick and dirty fix to enable 32-bit
// builds without extensive refactoring of the data structures. For a proper
// fix, we would need to track and free these allocations when UI elements are
// removed.

#ifdef _WIN64
inline InstanceHandle HandleFromLParam(LPARAM lparam) {
    return static_cast<InstanceHandle>(lparam);
}

inline LPARAM HandleToLParam(InstanceHandle handle) {
    return static_cast<LPARAM>(handle);
}

inline bool ValueShownAsIsFromDWordPtr(DWORD_PTR ptr) {
    return static_cast<bool>((ptr >> 32) & 1);
}

inline unsigned int PropertyIndexFromDWordPtr(DWORD_PTR ptr) {
    return static_cast<unsigned int>(ptr & 0xFFFFFFFF);
}

inline DWORD_PTR ListItemDataToDWordPtr(bool valueShownAsIs,
                                        unsigned int propertyIndex) {
    return (static_cast<uint64_t>(valueShownAsIs ? 1 : 0) << 32) |
           propertyIndex;
}
#else
inline InstanceHandle HandleFromLParam(LPARAM lparam) {
    return *reinterpret_cast<InstanceHandle*>(lparam);
}

inline LPARAM HandleToLParam(InstanceHandle handle) {
    // Memory leak: This allocation is never freed (see comment above).
    auto* pHandle = new InstanceHandle;
    *pHandle = handle;
    return reinterpret_cast<LPARAM>(pHandle);
}

inline bool ValueShownAsIsFromDWordPtr(DWORD_PTR ptr) {
    uint64_t value = *reinterpret_cast<uint64_t*>(ptr);
    return static_cast<bool>((value >> 32) & 1);
}

inline unsigned int PropertyIndexFromDWordPtr(DWORD_PTR ptr) {
    uint64_t value = *reinterpret_cast<uint64_t*>(ptr);
    return static_cast<unsigned int>(value & 0xFFFFFFFF);
}

inline DWORD_PTR ListItemDataToDWordPtr(bool valueShownAsIs,
                                        unsigned int propertyIndex) {
    // Memory leak: This allocation is never freed (see comment above).
    auto* pData = new uint64_t;
    *pData =
        (static_cast<uint64_t>(valueShownAsIs ? 1 : 0) << 32) | propertyIndex;
    return reinterpret_cast<DWORD_PTR>(pData);
}
#endif

}  // namespace

void CMainDlg::DumpElementRecursive(std::wstring& output,
                                    InstanceHandle handle,
                                    bool isFirst) {
    // Skip separator for first element
    if (!isFirst) {
        output += L"\n";
    }

    // Build and write path
    std::wstring path = BuildElementPath(m_elementItems, handle);
    output += L"Path: ";
    output += path.empty() ? L"(root)" : path;
    output += L"\n";

    // Get IInspectable object
    wf::IInspectable obj;
    HRESULT hr = m_xamlDiagnostics->GetIInspectableFromHandle(
        handle, reinterpret_cast<::IInspectable**>(winrt::put_abi(obj)));

    if (SUCCEEDED(hr)) {
        // Write class name
        auto className = winrt::get_class_name(obj);
        output += L"Class: ";
        output += className;
        output += L"\n";

        // Write element name
        std::wstring elementName;
        if (auto frameworkElement = obj.try_as<wux::FrameworkElement>()) {
            elementName = frameworkElement.Name();
        } else if (auto frameworkElement =
                       obj.try_as<mux::FrameworkElement>()) {
            elementName = frameworkElement.Name();
        }
        output += L"Name: ";
        output += elementName.empty() ? L"(none)" : elementName;
        output += L"\n";

        // Write rectangle
        auto it = m_elementItems.find(handle);
        bool isRoot =
            (it != m_elementItems.end() && it->second.parentHandle == 0);

        std::optional<CRect> rect;
        if (isRoot) {
            rect = GetRootElementRect(obj);
        } else {
            rect = GetRelativeElementRect(obj);
        }

        if (rect) {
            output += std::format(L"Rectangle: ({},{}) - ({},{})  -  {}x{}\n",
                                  rect->left, rect->top, rect->right,
                                  rect->bottom, rect->Width(), rect->Height());
        } else {
            output += L"Rectangle: (unavailable)\n";
        }
    }

    // Write properties
    unsigned int sourceCount = 0;
    PropertyChainSource* pPropertySources = nullptr;
    unsigned int propertyCount = 0;
    PropertyChainValue* pPropertyValues = nullptr;

    hr = m_visualTreeService->GetPropertyValuesChain(
        handle, &sourceCount, &pPropertySources, &propertyCount,
        &pPropertyValues);

    if (SUCCEEDED(hr)) {
        // First, output local properties
        bool hasLocalProperties = false;
        for (unsigned int i = 0; i < propertyCount; i++) {
            const auto& prop = pPropertyValues[i];
            const auto& src = pPropertySources[prop.PropertyChainIndex];

            if (src.Source == BaseValueSourceLocal) {
                if (!hasLocalProperties) {
                    output += L"Local properties:\n";
                    hasLocalProperties = true;
                }

                auto [value, valueShownAsIs] =
                    FormatPropertyValue(prop, m_xamlDiagnostics.get());
                output += L"- ";
                output += prop.PropertyName;
                output += L": ";
                output += value;
                output += L"\n";
            }
        }

        // Then, output other properties
        bool hasOtherProperties = false;
        for (unsigned int i = 0; i < propertyCount; i++) {
            const auto& prop = pPropertyValues[i];
            const auto& src = pPropertySources[prop.PropertyChainIndex];

            if (src.Source != BaseValueSourceLocal) {
                if (!hasOtherProperties) {
                    output += L"Other Properties:\n";
                    hasOtherProperties = true;
                }

                auto [value, valueShownAsIs] =
                    FormatPropertyValue(prop, m_xamlDiagnostics.get());
                output += L"- ";
                output += prop.PropertyName;
                output += L": ";
                output += value;
                output += L"\n";
            }
        }

        if (!hasLocalProperties && !hasOtherProperties) {
            output += L"Properties:\n";
            output += L"(no properties available)\n";
        }

        // Free memory
        CoTaskMemFree(pPropertySources);
        CoTaskMemFree(pPropertyValues);
    } else {
        output += L"Properties:\n";
        output += L"(no properties available)\n";
    }

    // Recursively dump children
    auto childrenIt = m_parentToChildren.find(handle);
    if (childrenIt != m_parentToChildren.end()) {
        for (InstanceHandle childHandle : childrenIt->second) {
            DumpElementRecursive(output, childHandle, false);
        }
    }
}

CMainDlg::CMainDlg(winrt::com_ptr<IXamlDiagnostics> diagnostics,
                   OnEventCallback_t eventCallback)
    : m_elementTree(this, 1),
      m_visualTreeService(diagnostics.as<IVisualTreeService3>()),
      m_xamlDiagnostics(std::move(diagnostics)),
      m_eventCallback(std::move(eventCallback)) {}

void CMainDlg::Hide() {
    ShowWindow(SW_HIDE);

    if (m_flashAreaWindow) {
        m_flashAreaWindow.ShowWindow(SW_HIDE);
    }

    if (m_eventCallback) {
        m_eventCallback(m_hWnd, EventId::Hidden);
    }
}

void CMainDlg::Show() {
    ShowWindow(SW_SHOW);

    if (m_flashAreaWindow) {
        m_flashAreaWindow.ShowWindow(SW_SHOW);
    }

    if (m_eventCallback) {
        m_eventCallback(m_hWnd, EventId::Shown);
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
    m_icon = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR,
                              ::GetSystemMetrics(SM_CXICON),
                              ::GetSystemMetrics(SM_CYICON));
    SetIcon(m_icon, TRUE);
    m_smallIcon = AtlLoadIconImage(IDR_MAINFRAME, LR_DEFAULTCOLOR,
                                   ::GetSystemMetrics(SM_CXSMICON),
                                   ::GetSystemMetrics(SM_CYSMICON));
    SetIcon(m_smallIcon, FALSE);

    // Init resizing.
    DlgResize_Init();

    // Init UI elements.
    SetWindowText(std::format(L"UWPSpy - PID: {} TID: {}",
                              GetCurrentProcessId(), GetCurrentThreadId())
                      .c_str());

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    m_elementTree.SubclassWindow(treeView);
    treeView.SetExtendedStyle(TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);
    ::SetWindowTheme(treeView, L"Explorer", nullptr);

    auto detailsTabs = CTabCtrl(GetDlgItem(IDC_DETAILS_TABS));
    detailsTabs.InsertItem(0, L"Attributes");
    detailsTabs.InsertItem(1, L"Detailed attributes");
    detailsTabs.InsertItem(2, L"Visual states");
    CRect detailsTabsRect;
    detailsTabs.GetWindowRect(&detailsTabsRect);
    ::MapWindowPoints(nullptr, m_hWnd,
                      reinterpret_cast<POINT*>(&detailsTabsRect),
                      sizeof(RECT) / sizeof(POINT));

    CListViewCtrl attributesList(GetDlgItem(IDC_ATTRIBUTE_LIST));
    m_attributesList.SubclassWindow(attributesList);
    attributesList.SetExtendedListViewStyle(
        LVS_EX_FULLROWSELECT | LVS_EX_LABELTIP | LVS_EX_DOUBLEBUFFER);
    ::SetWindowTheme(attributesList, L"Explorer", nullptr);
    ResetAttributesListColumns();
    RECT attributesListRect = {};
    attributesList.GetWindowRect(&attributesListRect);
    ::MapWindowPoints(nullptr, m_hWnd,
                      reinterpret_cast<POINT*>(&attributesListRect),
                      sizeof(RECT) / sizeof(POINT));
    attributesListRect.top = detailsTabsRect.bottom;
    // Height will be adjusted automatically.
    attributesListRect.bottom = attributesListRect.top;
    attributesList.SetWindowPos(nullptr, &attributesListRect,
                                SWP_NOZORDER | SWP_NOACTIVATE);

    auto visualStatesTree = CTreeViewCtrlEx(GetDlgItem(IDC_VISUAL_STATE_TREE));
    visualStatesTree.SetExtendedStyle(TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);
    ::SetWindowTheme(visualStatesTree, L"Explorer", nullptr);
    visualStatesTree.SetWindowPos(nullptr, &attributesListRect,
                                  SWP_NOZORDER | SWP_NOACTIVATE);

    CButton(GetDlgItem(IDC_HIGHLIGHT_SELECTION))
        .SetCheck(m_highlightSelection ? BST_CHECKED : BST_UNCHECKED);

    CButton(GetDlgItem(IDC_STICKY))
        .SetCheck(m_sticky ? BST_CHECKED : BST_UNCHECKED);

    return TRUE;
}

void CMainDlg::OnDestroy() {}

BOOL CMainDlg::OnNcActivate(BOOL bActive) {
    if (!bActive && m_sticky && IsWindowEnabled()) {
        // Prevent the window from being deactivated.
        return FALSE;
    }

    SetMsgHandled(FALSE);
    return TRUE;
}

void CMainDlg::OnFinalMessage(HWND hWnd) {
    if (m_eventCallback) {
        m_eventCallback(hWnd, EventId::FinalMessage);
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

        case TIMER_ID_SET_SELECTED_ELEMENT_INFORMATION:
            KillTimer(nIDEvent);
            SetSelectedElementInformation();
            break;

        case TIMER_ID_REFRESH_SELECTED_ELEMENT_INFORMATION:
            KillTimer(nIDEvent);
            RefreshSelectedElementInformation(0);
            break;
    }
}

void CMainDlg::OnContextMenu(CWindow wnd, CPoint point) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    if (wnd == treeView) {
        OnElementTreeContextMenu(treeView, point);
        return;
    }

    auto attributesList = CListViewCtrl(GetDlgItem(IDC_ATTRIBUTE_LIST));
    if (wnd == attributesList) {
        OnAttributesListContextMenu(attributesList, point);
        return;
    }

    auto visualStatesTree = CTreeViewCtrlEx(GetDlgItem(IDC_VISUAL_STATE_TREE));
    if (wnd == visualStatesTree) {
        OnVisualStateContextMenu(visualStatesTree, point);
        return;
    }
}

InstanceHandle CMainDlg::ElementFromPoint(CPoint pt) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));

    for (auto item = treeView.GetRootItem(); item;
         item = item.GetNextSibling()) {
        auto handle = HandleFromLParam(item.GetData());

        wf::IInspectable rootElement;
        HRESULT hr = m_xamlDiagnostics->GetIInspectableFromHandle(
            handle,
            reinterpret_cast<::IInspectable**>(winrt::put_abi(rootElement)));
        if (FAILED(hr) || !rootElement) {
            continue;
        }

        CWindow rootWnd;
        CRect rootElementRect;
        if (auto rect = GetRootElementRect(rootElement, &rootWnd.m_hWnd)) {
            rootElementRect = *rect;
        }

        wux::UIElement wsubtree = nullptr;
        mux::UIElement msubtree = nullptr;

        if (auto window = rootElement.try_as<wux::Window>()) {
            wsubtree = window.Content();
        } else if (auto desktopWindowXamlSource =
                       rootElement
                           .try_as<wux::Hosting::DesktopWindowXamlSource>()) {
            wsubtree = desktopWindowXamlSource.Content();
        } else if (auto window = rootElement.try_as<mux::Window>()) {
            msubtree = window.Content();
        } else if (auto desktopWindowXamlSource =
                       rootElement
                           .try_as<mux::Hosting::DesktopWindowXamlSource>()) {
            msubtree = desktopWindowXamlSource.Content();
        } else if (auto desktopWindowXamlSource = try_as_with_guid_unsafe<
                       mux::Hosting::DesktopWindowXamlSource>(
                       reinterpret_cast<::IInspectable*>(
                           winrt::get_abi(rootElement)),
                       IID_IDesktopWindowXamlSource_WinUI_2)) {
            msubtree = desktopWindowXamlSource.Content();
        } else if (auto desktopWindowXamlSource = try_as_with_guid_unsafe<
                       mux::Hosting::DesktopWindowXamlSource>(
                       reinterpret_cast<::IInspectable*>(
                           winrt::get_abi(rootElement)),
                       IID_IDesktopWindowXamlSource_WinUI_1)) {
            msubtree = desktopWindowXamlSource.Content();
        }

        if (!wsubtree && !msubtree) {
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

        if (!rootElementRect.PtInRect(ptWithoutDpi)) {
            continue;
        }

        CPoint ptWithoutDpiRelative = ptWithoutDpi;
        ptWithoutDpiRelative.Offset(-rootElementRect.TopLeft());

        InstanceHandle foundHandle = 0;
        if (wsubtree) {
            foundHandle =
                ElementFromPointInSubtree(wsubtree, ptWithoutDpiRelative);
        } else if (msubtree) {
            foundHandle =
                ElementFromPointInSubtree(msubtree, ptWithoutDpiRelative);
        }

        if (foundHandle) {
            return foundHandle;
        }
    }

    return 0;
}

InstanceHandle CMainDlg::ElementFromPointInSubtree(wux::UIElement subtree,
                                                   CPoint pt) {
    wf::Rect rect{wf::Point{static_cast<float>(pt.x), static_cast<float>(pt.y)},
                  wf::Size{1, 1}};

    auto elements = wux::Media::VisualTreeHelper::FindElementsInHostCoordinates(
        rect, subtree);

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

InstanceHandle CMainDlg::ElementFromPointInSubtree(mux::UIElement subtree,
                                                   CPoint pt) {
    wf::Rect rect{wf::Point{static_cast<float>(pt.x), static_cast<float>(pt.y)},
                  wf::Size{1, 1}};

    auto elements = mux::Media::VisualTreeHelper::FindElementsInHostCoordinates(
        rect, subtree);

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
    CRect rootElementRect;
    if (auto rect = GetRootElementRect(rootElement, &rootWnd.m_hWnd)) {
        rootElementRect = *rect;
    }

    CRect rect;
    if (element) {
        auto elementRect = GetRelativeElementRect(element);
        if (!elementRect) {
            return false;
        }

        rect = *elementRect;
        rect.OffsetRect(rootElementRect.TopLeft());
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

    m_flashAreaWindow = FlashArea(rootWnd, _Module.GetModuleInstance(), rect,
                                  IsWindowVisible());
    if (!m_flashAreaWindow) {
        return false;
    }

    // Keep the current window on top.
    if (GetForegroundWindow() == m_hWnd) {
        SetWindowPos(nullptr, 0, 0, 0, 0,
                     SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
    }

    return true;
}

void CMainDlg::DestroyFlashArea() {
    if (m_flashAreaWindow) {
        m_flashAreaWindow.DestroyWindow();
    }
}

LRESULT CMainDlg::OnElementTreeSelChanged(LPNMHDR pnmh) {
    NMTREEVIEW* pnmtv = (NMTREEVIEW*)pnmh;

    switch (pnmtv->action) {
        case TVC_BYKEYBOARD:
        case TVC_BYMOUSE:
            SetSelectedElementInformation();
            break;

        default:
            // Action is TVC_UNKNOWN when the selection is changed after the
            // previous selection was deleted. Use a delay since the app might
            // be busy deleting many items at once.
            SetTimer(TIMER_ID_SET_SELECTED_ELEMENT_INFORMATION,
                     kRedrawTreeDelay);
            break;
    }

    return 0;
}

void CMainDlg::OnSplitToggle(UINT uNotifyCode, int nID, CWindow wndCtl) {
    m_splitModeAttributesExpanded = !m_splitModeAttributesExpanded;
    wndCtl.SetWindowText(m_splitModeAttributesExpanded ? L">" : L"<");

    // Below is a very ugly hack to switch to another DlgResize layout.

    CRect rectDlg;
    GetClientRect(&rectDlg);

    SetWindowPos(nullptr, 0, 0, m_sizeDialog.cx, m_sizeDialog.cy,
                 SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

    DlgResize_Init();

    ResizeClient(rectDlg.right, rectDlg.bottom);
}

LRESULT CMainDlg::OnDetailsTabsSelChange(LPNMHDR pnmh) {
    auto detailsTabs = CTabCtrl(GetDlgItem(IDC_DETAILS_TABS));
    int index = detailsTabs.GetCurSel();

    bool needsReconfigure = false;
    if (index == 0 || index == 1) {
        // Determine if we're in detailed mode based on tab selection.
        bool newDetailedProperties = (index == 1);

        // Check if we need to reconfigure columns.
        needsReconfigure = (m_detailedProperties != newDetailedProperties);

        m_detailedProperties = newDetailedProperties;
    }

    // Show/hide appropriate controls
    m_attributesList.ShowWindow((index == 0 || index == 1) ? SW_SHOW : SW_HIDE);

    auto visualStatesTree = CTreeViewCtrlEx(GetDlgItem(IDC_VISUAL_STATE_TREE));
    visualStatesTree.ShowWindow(index == 2 ? SW_SHOW : SW_HIDE);

    // If switching between attribute tabs, reconfigure and refresh.
    if (needsReconfigure) {
        ResetAttributesListColumns();

        auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
        auto selectedItem = treeView.GetSelectedItem();
        if (selectedItem && selectedItem.GetParent()) {
            auto handle = HandleFromLParam(selectedItem.GetData());
            PopulateAttributesList(handle);
        }
    }

    return 0;
}

LRESULT CMainDlg::OnAttributesListDblClk(LPNMHDR pnmh) {
    auto itemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pnmh);
    if (itemActivate->iItem == -1 ||
        (itemActivate->iSubItem != 0 && itemActivate->iSubItem != 1)) {
        return 0;
    }

    auto attributesList = CListViewCtrl(pnmh->hwndFrom);

    if (itemActivate->iSubItem == 0) {
        unsigned int clickedPropertyIndex = PropertyIndexFromDWordPtr(
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
        bool valueShownAsIs = ValueShownAsIsFromDWordPtr(
            attributesList.GetItemData(itemActivate->iItem));
        if (valueShownAsIs) {
            CString itemValue;
            attributesList.GetItemText(itemActivate->iItem,
                                       itemActivate->iSubItem, itemValue);

            auto propertyValueEdit = CEdit(GetDlgItem(IDC_PROPERTY_VALUE));
            propertyValueEdit.SetWindowText(itemValue);

            auto propertyValueXamlCheckbox =
                CButton(GetDlgItem(IDC_PROPERTY_IS_XAML));
            if (propertyValueXamlCheckbox.GetCheck() != BST_UNCHECKED) {
                propertyValueXamlCheckbox.SetCheck(BST_UNCHECKED);
                propertyValueEdit.ShowWindow(SW_SHOW);
                auto propertyValueEditXaml =
                    CEdit(GetDlgItem(IDC_PROPERTY_VALUE_XAML));
                propertyValueEditXaml.ShowWindow(SW_HIDE);
            }
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

void CMainDlg::OnPropertyIsXaml(UINT uNotifyCode, int nID, CWindow wndCtl) {
    bool checked = CButton(wndCtl).GetCheck() != BST_UNCHECKED;

    auto propertyValueEdit = CEdit(GetDlgItem(IDC_PROPERTY_VALUE));
    auto propertyValueEditXaml = CEdit(GetDlgItem(IDC_PROPERTY_VALUE_XAML));

    propertyValueEdit.ShowWindow(checked ? SW_HIDE : SW_SHOW);
    propertyValueEditXaml.ShowWindow(checked ? SW_SHOW : SW_HIDE);
}

void CMainDlg::OnPropertyRemove(UINT uNotifyCode, int nID, CWindow wndCtl) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto selectedItem = treeView.GetSelectedItem();
    if (!selectedItem) {
        return;
    }

    auto handle = HandleFromLParam(selectedItem.GetData());

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

    auto handle = HandleFromLParam(selectedItem.GetData());

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

    InstanceHandle newValueHandle;
    HRESULT hr = S_OK;
    std::wstring errorExtraMsg;

    auto propertyValueXamlCheckbox = CButton(GetDlgItem(IDC_PROPERTY_IS_XAML));
    if (propertyValueXamlCheckbox.GetCheck() != BST_UNCHECKED) {
        CString propertyXaml;
        auto propertyValueEditXaml = CEdit(GetDlgItem(IDC_PROPERTY_VALUE_XAML));
        propertyValueEditXaml.GetWindowText(propertyXaml);

        try {
            wf::IInspectable obj;
            winrt::check_hresult(m_xamlDiagnostics->GetIInspectableFromHandle(
                handle,
                reinterpret_cast<::IInspectable**>(winrt::put_abi(obj))));

            auto value = StyleValueFromXaml(
                obj,
                {propertyName.GetString(), (size_t)propertyName.GetLength()},
                {propertyXaml.GetString(), (size_t)propertyXaml.GetLength()});

            winrt::check_hresult(m_xamlDiagnostics->RegisterInstance(
                static_cast<IInspectable*>(winrt::get_abi(value)),
                &newValueHandle));
        } catch (winrt::hresult_error const& ex) {
            hr = ex.code();
            errorExtraMsg = ex.message();
        } catch (...) {
            hr = winrt::to_hresult();
        }
    } else {
        CString propertyValue;
        auto propertyValueEdit = CEdit(GetDlgItem(IDC_PROPERTY_VALUE));
        propertyValueEdit.GetWindowText(propertyValue);

        hr = m_visualTreeService->CreateInstance(
            _bstr_t(propertyType), _bstr_t(propertyValue), &newValueHandle);
    }

    if (SUCCEEDED(hr)) {
        hr = m_visualTreeService->SetProperty(handle, newValueHandle,
                                              propertyIndex);
    }

    if (FAILED(hr)) {
        auto errorMsg = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        if (!errorExtraMsg.empty()) {
            errorMsg += L'\n';
            errorMsg += errorExtraMsg;
        }
        MessageBox(errorMsg.c_str(), L"Error");
        return;
    }

    RefreshSelectedElementInformation();
}

void CMainDlg::OnCollapseAll(UINT uNotifyCode, int nID, CWindow wndCtl) {
    auto button = CButton(wndCtl);

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto root = treeView.GetRootItem();
    if (root) {
        if (!m_redrawTreeQueued) {
            treeView.SetRedraw(FALSE);
        }

        if (!m_listCollapsed) {
            treeView.SelectItem(root);
            TreeViewExpandRecursively(treeView, root, TVE_COLLAPSE);
            treeView.EnsureVisible(root);
            button.SetWindowText(L"Expand all");
        } else {
            TreeViewExpandRecursively(treeView, root, TVE_EXPAND);
            auto selected = treeView.GetSelectedItem();
            treeView.EnsureVisible(selected ? selected : root);
            button.SetWindowText(L"Collapse all");
        }

        m_listCollapsed = !m_listCollapsed;

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
            auto handle = HandleFromLParam(selectedItem.GetData());
            CreateFlashArea(handle);
        }
    }
}

void CMainDlg::OnSticky(UINT uNotifyCode, int nID, CWindow wndCtl) {
    m_sticky = CButton(wndCtl).GetCheck() != BST_UNCHECKED;
}

void CMainDlg::OnAppAbout(UINT uNotifyCode, int nID, CWindow wndCtl) {
    CAboutDlg dlgAbout;
    dlgAbout.DoModal(m_hWnd);
}

void CMainDlg::OnCancel(UINT uNotifyCode, int nID, CWindow wndCtl) {
    Hide();
}

LRESULT CMainDlg::OnActivateWindow(UINT uMsg, WPARAM wParam, LPARAM lParam) {
    Show();
    ::SetForegroundWindow(m_hWnd);
    return 0;
}

LRESULT CMainDlg::OnDestroyWindow(UINT uMsg, WPARAM wParam, LPARAM lParam) {
    DestroyWindow();
    return 0;
}

void CMainDlg::ElementTreeOnChar(TCHAR chChar, UINT nRepCnt, UINT nFlags) {
    if (chChar == 4) {
        // Ctrl+D.
        SelectElementFromCursor();
    } else {
        SetMsgHandled(FALSE);
    }
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
    KillTimer(TIMER_ID_SET_SELECTED_ELEMENT_INFORMATION);
    KillTimer(TIMER_ID_REFRESH_SELECTED_ELEMENT_INFORMATION);

    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));
    auto selectedItem = treeView.GetSelectedItem();
    if (!selectedItem) {
        SetDlgItemText(IDC_CLASS_EDIT, L"");
        SetDlgItemText(IDC_NAME_EDIT, L"");
        SetDlgItemText(IDC_RECT_EDIT, L"");

        m_attributesList.DeleteAllItems();
        m_attributesList.SetSortColumn(-1);

        auto visualStatesTree =
            CTreeViewCtrlEx(GetDlgItem(IDC_VISUAL_STATE_TREE));
        visualStatesTree.DeleteAllItems();

        auto propertiesComboBox = CComboBox(GetDlgItem(IDC_PROPERTY_NAME));
        propertiesComboBox.ResetContent();

        DestroyFlashArea();

        return false;
    }

    auto handle = HandleFromLParam(selectedItem.GetData());
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

    std::wstring frameworkElementName;
    try {
        if (auto frameworkElement = obj.try_as<wux::FrameworkElement>()) {
            frameworkElementName = frameworkElement.Name();
        } else if (auto frameworkElement =
                       obj.try_as<mux::FrameworkElement>()) {
            frameworkElementName = frameworkElement.Name();
        }
    } catch (...) {
        HRESULT hr = winrt::to_hresult();
        frameworkElementName =
            std::format(L"Error {:08X}", static_cast<DWORD>(hr));
    }

    SetDlgItemText(IDC_NAME_EDIT, frameworkElementName.c_str());

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
        PopulateVisualStatesTree(handle);
    } else {
        m_attributesList.DeleteAllItems();
        m_attributesList.SetSortColumn(-1);

        auto visualStatesTree =
            CTreeViewCtrlEx(GetDlgItem(IDC_VISUAL_STATE_TREE));
        visualStatesTree.DeleteAllItems();

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

    CListViewCtrl attributesList(m_attributesList);
    int attributesListTopIndex = attributesList.GetTopIndex();

    if (!SetSelectedElementInformation()) {
        return false;
    }

    ListViewSetTopIndex(attributesList, attributesListTopIndex);
    return true;
}

void CMainDlg::ResetAttributesListColumns() {
    CListViewCtrl attributesList(m_attributesList);

    attributesList.SetRedraw(FALSE);

    attributesList.DeleteAllItems();

    int firstColumnWidth = attributesList.GetColumnWidth(0);

    if (attributesList.DeleteColumn(0)) {
        while (attributesList.DeleteColumn(0)) {
            // Keep deleting...
        }
    } else {
        firstColumnWidth = 0;
    }

    // If columns existed, try to retain width.
    int width = firstColumnWidth * 2;
    if (!width) {
        CRect rect;
        attributesList.GetClientRect(rect);
        width = rect.Width() - ::GetSystemMetrics(SM_CXVSCROLL);
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
    CListViewCtrl attributesList(m_attributesList);

    attributesList.SetRedraw(FALSE);

    attributesList.DeleteAllItems();
    m_attributesList.SetSortColumn(-1);

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

        auto [value, valueShownAsIs] =
            FormatPropertyValue(v, m_xamlDiagnostics.get());

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

        attributesList.SetItemData(
            row, ListItemDataToDWordPtr(valueShownAsIs, v.Index));

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

void CMainDlg::PopulateVisualStatesTree(InstanceHandle handle) {
    auto visualStatesTree = CTreeViewCtrlEx(GetDlgItem(IDC_VISUAL_STATE_TREE));
    visualStatesTree.SetRedraw(FALSE);

    visualStatesTree.DeleteAllItems();

    try {
        wf::IInspectable element;
        winrt::check_hresult(m_xamlDiagnostics->GetIInspectableFromHandle(
            handle,
            reinterpret_cast<::IInspectable**>(winrt::put_abi(element))));
        if (!element) {
            throw std::runtime_error("Element can't be retrieved");
        }

        auto wuiFrameworkElement = element.try_as<wux::FrameworkElement>();
        auto muiFrameworkElement =
            wuiFrameworkElement ? mux::FrameworkElement{nullptr}
                                : element.try_as<mux::FrameworkElement>();

        // A workaround for a taskbar search box inspection crash.
        if (wuiFrameworkElement) {
            auto& element = wuiFrameworkElement;

            // The TaskListButtonPanel child element of the search box (with
            // "Icon and label" configuration) returns a list of size 1, but
            // accessing the first item leads to a null dereference crash. Skip
            // this element.
            if (winrt::get_class_name(element) ==
                L"Taskbar.TaskListButtonPanel") {
                auto parent = wux::Media::VisualTreeHelper::GetParent(element)
                                  .try_as<wux::FrameworkElement>();
                if (parent && winrt::get_class_name(parent) ==
                                  L"Taskbar.SearchBoxLaunchListButton") {
                    throw std::runtime_error("Unsupported element");
                }
            }

            // Same as above for an updated element layout (around Jun 2025).
            if (winrt::get_class_name(element) ==
                L"SearchUx.SearchUI.SearchButtonRootGrid") {
                auto parent = wux::Media::VisualTreeHelper::GetParent(element)
                                  .try_as<wux::FrameworkElement>();
                if (parent && winrt::get_class_name(parent) ==
                                  L"SearchUx.SearchUI.SearchPillButton") {
                    throw std::runtime_error("Unsupported element");
                }
            }
        }

        auto populateList = [&visualStatesTree](auto visualStateGroups) {
            for (const auto& group : visualStateGroups) {
                auto groupName = group.Name();
                if (groupName.empty()) {
                    groupName = L"(unnamed)";
                }

                auto currentState = group.CurrentState();

                auto groupItem = visualStatesTree.InsertItem(
                    groupName.c_str(), TVI_ROOT, TVI_LAST);

                for (auto state : group.States()) {
                    std::wstring name(state.Name());
                    if (name.empty()) {
                        name = L"(unnamed)";
                    }

                    if (state == currentState) {
                        name += L" (current)";
                    }

                    visualStatesTree.InsertItem(name.c_str(), groupItem,
                                                TVI_LAST);
                }

                groupItem.Expand();
            }
        };

        if (wuiFrameworkElement) {
            auto visualStateGroups =
                wux::VisualStateManager::GetVisualStateGroups(
                    wuiFrameworkElement);
            populateList(visualStateGroups);
        } else if (muiFrameworkElement) {
            auto visualStateGroups =
                mux::VisualStateManager::GetVisualStateGroups(
                    muiFrameworkElement);
            populateList(visualStateGroups);
        }
    } catch (...) {
        HRESULT hr = winrt::to_hresult();
        auto errorMsg = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        visualStatesTree.InsertItem(errorMsg.c_str(), TVI_ROOT, TVI_LAST);

        visualStatesTree.SetRedraw(TRUE);
        return;
    }

    visualStatesTree.SetRedraw(TRUE);
}

void CMainDlg::AddItemToTree(HTREEITEM parentTreeItem,
                             HTREEITEM insertAfter,
                             InstanceHandle handle,
                             ElementItem* elementItem) {
    auto treeView = CTreeViewCtrlEx(GetDlgItem(IDC_ELEMENT_TREE));

    ATLASSERT(!elementItem->treeItem);

    TVINSERTSTRUCT insertStruct{
        .hParent = parentTreeItem,
        .hInsertAfter = insertAfter,
        .item =
            {
                .mask = TVIF_TEXT | TVIF_PARAM | TVIF_STATE,
                .state = TVIS_EXPANDED,
                .stateMask = TVIS_EXPANDED,
                .pszText = const_cast<PWSTR>(elementItem->itemTitle.c_str()),
                .lParam = HandleToLParam(handle),
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

void CMainDlg::OnElementTreeContextMenu(CTreeViewCtrlEx treeView,
                                        CPoint point) {
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
        MENU_ID_COPY_ITEM,
        MENU_ID_COPY_PATH,
        MENU_ID_COPY_SUBTREE,
        MENU_ID_COPY_SUBTREE_WITH_PROPERTIES,
    };

    auto handle = HandleFromLParam(targetItem.GetData());

    bool isRoot = !targetItem.GetParent();

    try {
        wf::IInspectable element;
        winrt::check_hresult(m_xamlDiagnostics->GetIInspectableFromHandle(
            handle,
            reinterpret_cast<::IInspectable**>(winrt::put_abi(element))));

        auto wuiElement = element.try_as<wux::UIElement>();
        auto muiElement = wuiElement ? mux::UIElement{nullptr}
                                     : element.try_as<mux::UIElement>();

        bool visible = false;
        bool visibleCanBeToggled = true;
        if (wuiElement) {
            visible = wuiElement.Visibility() == wux::Visibility::Visible;
        } else if (muiElement) {
            visible = muiElement.Visibility() == mux::Visibility::Visible;
        } else {
            visibleCanBeToggled = false;
        }

        menu.AppendMenu(MF_STRING | (visible ? MF_CHECKED : 0) |
                            (visibleCanBeToggled ? 0 : MF_GRAYED),
                        MENU_ID_VISIBLE, L"Visible");
        menu.AppendMenu(MF_SEPARATOR);
        menu.AppendMenu(MF_STRING, MENU_ID_COPY_ITEM, L"Copy item");
        menu.AppendMenu(MF_STRING | (isRoot ? MF_GRAYED : 0), MENU_ID_COPY_PATH,
                        L"Copy path");
        menu.AppendMenu(MF_STRING, MENU_ID_COPY_SUBTREE, L"Copy subtree");
        menu.AppendMenu(MF_STRING, MENU_ID_COPY_SUBTREE_WITH_PROPERTIES,
                        L"Copy subtree with properties");

        int nCmd = menu.TrackPopupMenu(TPM_RIGHTBUTTON | TPM_RETURNCMD,
                                       menuPoint.x, menuPoint.y, m_hWnd);
        switch (nCmd) {
            case MENU_ID_VISIBLE:
                if (wuiElement) {
                    wuiElement.Visibility(visible ? wux::Visibility::Collapsed
                                                  : wux::Visibility::Visible);
                } else if (muiElement) {
                    muiElement.Visibility(visible ? mux::Visibility::Collapsed
                                                  : mux::Visibility::Visible);
                }

                RefreshSelectedElementInformation();
                break;

            case MENU_ID_COPY_ITEM: {
                CString itemText;
                targetItem.GetText(itemText);
                if (!CopyTextToClipboard(
                        m_hWnd,
                        {itemText.GetString(), (size_t)itemText.GetLength()})) {
                    MessageBox(L"Failed to copy to clipboard", L"Error");
                }
                break;
            }

            case MENU_ID_COPY_PATH: {
                std::wstring path = BuildElementPath(m_elementItems, handle);
                if (!CopyTextToClipboard(m_hWnd, path)) {
                    MessageBox(L"Failed to copy to clipboard", L"Error");
                }
                break;
            }

            case MENU_ID_COPY_SUBTREE: {
                CString str;
                TreeViewSubtreeToString(treeView, targetItem, str);
                if (!CopyTextToClipboard(
                        m_hWnd, {str.GetString(), (size_t)str.GetLength()})) {
                    MessageBox(L"Failed to copy to clipboard", L"Error");
                }
                break;
            }

            case MENU_ID_COPY_SUBTREE_WITH_PROPERTIES: {
                std::wstring dump;
                DumpElementRecursive(dump, handle, true);

                if (!CopyTextToClipboard(m_hWnd, dump)) {
                    MessageBox(L"Failed to copy to clipboard", L"Error");
                    return;
                }

                MessageBox(L"Content copied successfully", L"Success",
                           MB_OK | MB_ICONINFORMATION);
                break;
            }
        }
    } catch (...) {
        HRESULT hr = winrt::to_hresult();
        auto errorMsg = std::format(L"Error {:08X}", static_cast<DWORD>(hr));
        MessageBox(errorMsg.c_str(), L"Error");
    }
}

void CMainDlg::OnAttributesListContextMenu(CListViewCtrl listView,
                                           CPoint point) {
    int targetItem = -1;

    if (point.x == -1 && point.y == -1) {
        // Keyboard context menu.
        targetItem = listView.GetSelectedIndex();
    } else {
        CPoint mappedPoint = point;
        listView.ScreenToClient(&mappedPoint);

        targetItem = listView.HitTest(mappedPoint, nullptr);
    }

    if (targetItem == -1) {
        return;
    }

    CPoint menuPoint = point;
    if (menuPoint.x == -1 && menuPoint.y == -1) {
        // Keyboard context menu.
        CRect rect;
        if (listView.GetItemRect(targetItem, rect, LVIR_BOUNDS)) {
            menuPoint = rect.CenterPoint();
            listView.ClientToScreen(&menuPoint);
        } else {
            ::GetCursorPos(&menuPoint);
        }
    }

    CMenu menu;
    menu.CreatePopupMenu();

    enum {
        MENU_ID_COPY_ITEM = 1,
        MENU_ID_COPY_ALL_ITEMS,
    };

    menu.AppendMenu(MF_STRING, MENU_ID_COPY_ITEM, L"Copy item");
    menu.AppendMenu(MF_STRING, MENU_ID_COPY_ALL_ITEMS, L"Copy all items");

    int nCmd = menu.TrackPopupMenu(TPM_RIGHTBUTTON | TPM_RETURNCMD, menuPoint.x,
                                   menuPoint.y, m_hWnd);
    switch (nCmd) {
        case MENU_ID_COPY_ITEM: {
            int columns = listView.GetHeader().GetItemCount();
            CString columnText;
            CString itemText;
            for (int i = 0; i < columns; i++) {
                if (i > 0) {
                    itemText += L'\t';
                }
                listView.GetItemText(targetItem, i, columnText);
                itemText += columnText;
            }

            if (!CopyTextToClipboard(m_hWnd, {itemText.GetString(),
                                              (size_t)itemText.GetLength()})) {
                MessageBox(L"Failed to copy item text to clipboard", L"Error");
            }
            break;
        }

        case MENU_ID_COPY_ALL_ITEMS: {
            int columns = listView.GetHeader().GetItemCount();
            CString str;
            CString columnText;
            CString itemText;
            for (int i = 0; i < listView.GetItemCount(); i++) {
                itemText.Empty();
                for (int j = 0; j < columns; j++) {
                    if (j > 0) {
                        itemText += L'\t';
                    }
                    listView.GetItemText(i, j, columnText);
                    itemText += columnText;
                }
                str += itemText + L'\n';
            }

            if (!CopyTextToClipboard(
                    m_hWnd, {str.GetString(), (size_t)str.GetLength()})) {
                MessageBox(L"Failed to copy all items text to clipboard",
                           L"Error");
            }
            break;
        }
    }
}

void CMainDlg::OnVisualStateContextMenu(CTreeViewCtrlEx treeView,
                                        CPoint point) {
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
        MENU_ID_COPY_ITEM = 1,
        MENU_ID_COPY_ALL,
    };

    menu.AppendMenu(MF_STRING, MENU_ID_COPY_ITEM, L"Copy item");
    menu.AppendMenu(MF_STRING, MENU_ID_COPY_ALL, L"Copy all items");

    int nCmd = menu.TrackPopupMenu(TPM_RIGHTBUTTON | TPM_RETURNCMD, menuPoint.x,
                                   menuPoint.y, m_hWnd);
    switch (nCmd) {
        case MENU_ID_COPY_ITEM: {
            CString itemText;
            targetItem.GetText(itemText);
            if (!CopyTextToClipboard(m_hWnd, {itemText.GetString(),
                                              (size_t)itemText.GetLength()})) {
                MessageBox(L"Failed to copy item text to clipboard", L"Error");
            }
            break;
        }

        case MENU_ID_COPY_ALL: {
            CString str;
            TreeViewToString(treeView, str);
            if (!CopyTextToClipboard(
                    m_hWnd, {str.GetString(), (size_t)str.GetLength()})) {
                MessageBox(L"Failed to copy subtree text to clipboard",
                           L"Error");
            }
            break;
        }
    }
}
