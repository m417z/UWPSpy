#include "stdafx.h"

#include "visualtreewatcher.hpp"

#define EXTRA_DEBUG 0

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
    // const auto treeService = m_xamlDiagnostics.as<IVisualTreeService3>();
    // winrt::check_hresult(treeService->AdviseVisualTreeChange(this));

    // Calling AdviseVisualTreeChange from the current thread causes the app to
    // hang on Windows 10 in Advising::RunOnUIThread. Creating a new thread and
    // calling it from there fixes it.
    HANDLE thread = CreateThread(
        nullptr, 0,
        [](LPVOID lpParam) -> DWORD {
            auto watcher = reinterpret_cast<VisualTreeWatcher*>(lpParam);
            const auto treeService =
                watcher->m_xamlDiagnostics.as<IVisualTreeService3>();
            winrt::check_hresult(treeService->AdviseVisualTreeChange(watcher));
            return 0;
        },
        this, 0, nullptr);
    if (thread) {
        CloseHandle(thread);
    }
}

void VisualTreeWatcher::Activate() {
    std::shared_lock lock(m_dlgMainMutex);

    for (auto& [threadId, dlgMain] : m_dlgMainForEachThread) {
        dlgMain.PostMessage(CMainDlg::UWM_ACTIVATE_WINDOW);
    }
}

VisualTreeWatcher::~VisualTreeWatcher() {
    std::shared_lock lock(m_dlgMainMutex);

    for (auto& [threadId, dlgMain] : m_dlgMainForEachThread) {
        dlgMain.SendMessage(CMainDlg::UWM_DESTROY_WINDOW);
    }
}

HRESULT VisualTreeWatcher::OnVisualTreeChange(
    ParentChildRelation parentChildRelation,
    VisualElement element,
    VisualMutationType mutationType) try {
#if EXTRA_DEBUG
    try {
        WCHAR szBuffer[1025];
        OutputDebugString(L"========================================\n");

        wsprintf(szBuffer, L"Parent: %p\n", (void*)parentChildRelation.Parent);
        OutputDebugString(szBuffer);
        wsprintf(szBuffer, L"Child: %p\n", (void*)parentChildRelation.Child);
        OutputDebugString(szBuffer);
        wsprintf(szBuffer, L"Child index: %u\n",
                 parentChildRelation.ChildIndex);
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

        wux::FrameworkElement frameworkElement = nullptr;

        const auto inspectable = FromHandle<wf::IInspectable>(element.Handle);
        frameworkElement = inspectable.try_as<wux::FrameworkElement>();

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
    } catch (...) {
        HRESULT hr = winrt::to_hresult();
        WCHAR szBuffer[1025];
        wsprintf(szBuffer, L"Error %X\n", hr);
        OutputDebugString(szBuffer);
    }
#endif  // EXTRA_DEBUG

    auto* dlgMain = DlgMainForCurrentThread();

    switch (mutationType) {
        case Add:
            if (dlgMain) {
                dlgMain->ElementAdded(parentChildRelation, element);
            }
            break;

        case Remove:
            if (dlgMain) {
                dlgMain->ElementRemoved(element.Handle);
            }
            break;

        default:
            ATLASSERT(FALSE);
            break;
    }

    return S_OK;
} catch (...) {
    ATLASSERT(FALSE);
    // Returning an error prevents (some?) further messages, always return
    // success.
    // return winrt::to_hresult();
    return S_OK;
}

HRESULT VisualTreeWatcher::OnElementStateChanged(InstanceHandle,
                                                 VisualElementState,
                                                 LPCWSTR) noexcept {
    return S_OK;
}

CMainDlg* VisualTreeWatcher::DlgMainForCurrentThread() {
    DWORD dwCurrentThreadId = GetCurrentThreadId();

    {
        std::shared_lock lock(m_dlgMainMutex);

        auto it = m_dlgMainForEachThread.find(dwCurrentThreadId);
        if (it != m_dlgMainForEachThread.end()) {
            return &it->second;
        }
    }

    std::unique_lock lock(m_dlgMainMutex);

    auto [it, inserted] = m_dlgMainForEachThread.try_emplace(
        dwCurrentThreadId, m_xamlDiagnostics, this);

    CMainDlg& dlgMain = it->second;

    if (!inserted) {
        ATLASSERT(FALSE);
        return &dlgMain;
    }

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

    if (!dlgMain.Create(hChildWnd)) {
        ATLTRACE(L"Main dialog creation failed!\n");
        m_dlgMainForEachThread.erase(it);
        return nullptr;
    }

    dlgMain.SetOnFinalMessageCallback([](void* param, HWND hWnd) {
        reinterpret_cast<VisualTreeWatcher*>(param)->OnDlgMainFinalMessage(
            hWnd);
    });

    dlgMain.ShowWindow(SW_SHOWDEFAULT);
    SetForegroundWindow(dlgMain.m_hWnd);
    return &dlgMain;
}

void VisualTreeWatcher::OnDlgMainFinalMessage(HWND hWnd) {
    DWORD dwCurrentThreadId = GetCurrentThreadId();

    std::unique_lock lock(m_dlgMainMutex);

    auto it = m_dlgMainForEachThread.find(dwCurrentThreadId);
    if (it == m_dlgMainForEachThread.end()) {
        ATLASSERT(FALSE);
        return;
    }

    m_dlgMainForEachThread.erase(it);
}
