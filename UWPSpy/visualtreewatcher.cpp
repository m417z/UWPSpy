#include "stdafx.h"

#include "visualtreewatcher.hpp"

#define EXTRA_DEBUG 1

namespace {

// https://stackoverflow.com/a/10444161
ULONG_PTR EnableVisualStyles() {
    TCHAR dir[MAX_PATH];
    ULONG_PTR ulpActivationCookie = FALSE;
    ACTCTX actCtx = {sizeof(actCtx),
                     ACTCTX_FLAG_RESOURCE_NAME_VALID |
                         ACTCTX_FLAG_SET_PROCESS_DEFAULT |
                         ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID,
                     TEXT("shell32.dll"),
                     0,
                     0,
                     dir,
                     (LPCTSTR)124};
    UINT cch = GetSystemDirectory(dir, sizeof(dir) / sizeof(*dir));
    if (cch >= sizeof(dir) / sizeof(*dir)) {
        return FALSE; /*shouldn't happen*/
    }
    dir[cch] = TEXT('\0');
    ActivateActCtx(CreateActCtx(&actCtx), &ulpActivationCookie);
    return ulpActivationCookie;
}

}  // namespace

VisualTreeWatcher::VisualTreeWatcher(winrt::com_ptr<IUnknown> site)
    : m_xamlDiagnostics(site.as<IXamlDiagnostics>()) {
    const auto treeService = m_xamlDiagnostics.as<IVisualTreeService3>();

    HWND hChildWnd = nullptr;

    UINT32 length = 0;
    if (GetCurrentPackageFullName(&length, nullptr) !=
        APPMODEL_ERROR_NO_PACKAGE) {
        // Packaged UWP processes seem to lack visual styles, which causes
        // artifacts such as black squares and UI flickering. This call fixes
        // it.
        EnableVisualStyles();

        // Creating a dialog without a parent causes it to be occluded. Using
        // the core window as a parent window fixes it.
        if (const auto coreInterop =
                winrt::Windows::UI::Core::CoreWindow::GetForCurrentThread()
                    .try_as<ICoreWindowInterop>()) {
            coreInterop->get_WindowHandle(&hChildWnd);
        }
    }

    m_dlgMain.emplace(treeService, m_xamlDiagnostics);

    if (m_dlgMain->Create(hChildWnd)) {
        m_dlgMain->ShowWindow(SW_SHOWDEFAULT);
        SetForegroundWindow(m_dlgMain->m_hWnd);
    } else {
        ATLTRACE(_T("Main dialog creation failed!\n"));
    }

    winrt::check_hresult(treeService->AdviseVisualTreeChange(this));
}

void VisualTreeWatcher::Activate() {
    ATLASSERT(m_dlgMain && *m_dlgMain);
    m_dlgMain->Show();
    SetForegroundWindow(m_dlgMain->m_hWnd);
}

VisualTreeWatcher::~VisualTreeWatcher() {
    ATLASSERT(m_dlgMain && *m_dlgMain);
    m_dlgMain->DestroyWindow();
}

HRESULT VisualTreeWatcher::OnVisualTreeChange(
    ParentChildRelation parentChildRelation,
    VisualElement element,
    VisualMutationType mutationType) try {
#if EXTRA_DEBUG
    WCHAR szBuffer[1025];
    OutputDebugString(L"========================================\n");

    wsprintf(szBuffer, L"Parent: %p\n", (void*)parentChildRelation.Parent);
    OutputDebugString(szBuffer);
    wsprintf(szBuffer, L"Child: %p\n", (void*)parentChildRelation.Child);
    OutputDebugString(szBuffer);
    wsprintf(szBuffer, L"Child index: %u\n", parentChildRelation.ChildIndex);
    OutputDebugString(szBuffer);
    wsprintf(szBuffer, L"Handle: %p\n", (void*)element.Handle);
    OutputDebugString(szBuffer);
    wsprintf(szBuffer, L"Type: %s\n", element.Type);
    OutputDebugString(szBuffer);
    wsprintf(szBuffer, L"Name: %s\n", element.Name);
    OutputDebugString(szBuffer);
    wsprintf(szBuffer, L"NumChildren: %u\n", element.NumChildren);
    OutputDebugString(szBuffer);
    OutputDebugString(L"~~~~~~~~~~\n");

    switch (mutationType) {
        case Add:
            OutputDebugString(L"Mutation type: Add\n");
            break;

        case Remove:
            OutputDebugString(L"Mutation type: Remove\n");
            break;

        default:
            wsprintf(szBuffer, L"Mutation type: %d\n",
                     static_cast<int>(mutationType));
            OutputDebugString(szBuffer);
            break;
    }

    const auto inspectable = FromHandle<wf::IInspectable>(element.Handle);
    const auto frameworkElement = inspectable.try_as<wux::FrameworkElement>();
    if (frameworkElement) {
        wsprintf(szBuffer, L"FrameworkElement address: %p\n",
                 winrt::get_abi(frameworkElement));
        OutputDebugString(szBuffer);
        wsprintf(szBuffer, L"FrameworkElement class: %s\n",
                 winrt::get_class_name(frameworkElement).c_str());
        OutputDebugString(szBuffer);
        wsprintf(szBuffer, L"FrameworkElement name: %s\n",
                 frameworkElement.Name().c_str());
        OutputDebugString(szBuffer);
    }

    if (parentChildRelation.Parent) {
        if (frameworkElement) {
            wsprintf(szBuffer, L"Real parent address: %p\n",
                     winrt::get_abi(
                         frameworkElement.Parent().as<wf::IInspectable>()));
            OutputDebugString(szBuffer);
        }

        const auto inspectable =
            FromHandle<wf::IInspectable>(parentChildRelation.Parent);
        const auto frameworkElement =
            inspectable.try_as<wux::FrameworkElement>();
        if (frameworkElement) {
            wsprintf(szBuffer, L"Parent FrameworkElement address: %p\n",
                     winrt::get_abi(frameworkElement));
            OutputDebugString(szBuffer);
            wsprintf(szBuffer, L"Parent FrameworkElement class: %s\n",
                     winrt::get_class_name(frameworkElement).c_str());
            OutputDebugString(szBuffer);
            wsprintf(szBuffer, L"Parent FrameworkElement name: %s\n",
                     frameworkElement.Name().c_str());
            OutputDebugString(szBuffer);
        }
    }
#endif  // EXTRA_DEBUG

    switch (mutationType) {
        case Add:
            if (m_dlgMain && *m_dlgMain) {
                m_dlgMain->ElementAdded(parentChildRelation, element);
            }
            break;

        case Remove:
            if (m_dlgMain && *m_dlgMain) {
                m_dlgMain->ElementRemoved(element.Handle);
            }
            break;

        default:
            ATLASSERT(FALSE);
            break;
    }

    return S_OK;
} catch (...) {
    ATLASSERT(FALSE);
    return winrt::to_hresult();
}

HRESULT VisualTreeWatcher::OnElementStateChanged(InstanceHandle,
                                                 VisualElementState,
                                                 LPCWSTR) noexcept {
    return S_OK;
}
