#include "stdafx.h"

#include "visualtreewatcher.hpp"

#define EXTRA_DEBUG 0

namespace {

// https://devblogs.microsoft.com/oldnewthing/20240319-00/?p=109552
template <typename T>
std::enable_if_t<std::is_base_of_v<::IUnknown, T>, winrt::com_ptr<T>>
make_com_ptr_from_copy(T* p) {
    winrt::com_ptr<T> result;
    result.copy_from(p);
    return result;
}

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
    : m_selfPtr(make_com_ptr_from_copy(this)),
      m_xamlDiagnostics(site.as<IXamlDiagnostics>()) {
    // AdviseVisualTreeChange();

    // Calling AdviseVisualTreeChange from the current thread causes the app to
    // hang on Windows 10 in Advising::RunOnUIThread. Creating a new thread and
    // calling it from there fixes it.
    HANDLE thread = CreateThread(
        nullptr, 0,
        [](LPVOID lpParam) -> DWORD {
            auto watcher = reinterpret_cast<VisualTreeWatcher*>(lpParam);
            watcher->AdviseVisualTreeChange();
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
    while (true) {
        HWND hWnd = nullptr;

        {
            std::shared_lock lock(m_dlgMainMutex);

            if (!m_dlgMainForEachThread.empty()) {
                hWnd = m_dlgMainForEachThread.begin()->second;
            }
        }

        if (!hWnd) {
            break;
        }

        if (!SendMessage(hWnd, CMainDlg::UWM_DESTROY_WINDOW, 0, 0)) {
            ATLASSERT(FALSE);
        }
    }
}

void VisualTreeWatcher::AdviseVisualTreeChange() {
    const auto treeService = m_xamlDiagnostics.as<IVisualTreeService3>();
    winrt::check_hresult(treeService->AdviseVisualTreeChange(this));
}

void VisualTreeWatcher::UnadviseVisualTreeChange() {
    const auto treeService = m_xamlDiagnostics.as<IVisualTreeService3>();
    winrt::check_hresult(treeService->UnadviseVisualTreeChange(this));
    m_selfPtr = nullptr;
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

    CMainDlg* dlgMain;

    {
        std::unique_lock lock(m_dlgMainMutex);

        auto [it, inserted] = m_dlgMainForEachThread.try_emplace(
            dwCurrentThreadId, m_xamlDiagnostics,
            std::bind(&VisualTreeWatcher::OnDlgMainEvent, this,
                      std::placeholders::_1, std::placeholders::_2));

        dlgMain = &it->second;

        if (!inserted) {
            ATLASSERT(FALSE);
            return dlgMain;
        }
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

    if (!dlgMain->Create(hChildWnd)) {
        ATLTRACE(L"Main dialog creation failed!\n");
        return nullptr;
    }

    dlgMain->ShowWindow(SW_SHOWDEFAULT);
    SetForegroundWindow(dlgMain->m_hWnd);

    return dlgMain;
}

void VisualTreeWatcher::OnDlgMainEvent(HWND hWnd, CMainDlg::EventId eventId) {
    if (eventId == CMainDlg::EventId::Hidden) {
        OnDlgMainHidden(hWnd);
    } else if (eventId == CMainDlg::EventId::FinalMessage) {
        OnDlgMainFinalMessage(hWnd);
    }
}

void VisualTreeWatcher::OnDlgMainHidden(HWND hWnd) {
    UnloadIfNoMoreVisibleDlgs();
}

void VisualTreeWatcher::OnDlgMainFinalMessage(HWND hWnd) {
    DWORD dwCurrentThreadId = GetCurrentThreadId();

    {
        std::unique_lock lock(m_dlgMainMutex);

        auto it = m_dlgMainForEachThread.find(dwCurrentThreadId);
        if (it == m_dlgMainForEachThread.end()) {
            ATLASSERT(FALSE);
            return;
        }

        m_dlgMainForEachThread.erase(it);
    }

    UnloadIfNoMoreVisibleDlgs();
}

void VisualTreeWatcher::UnloadIfNoMoreVisibleDlgs() {
    if (m_unloading) {
        return;
    }

    {
        std::shared_lock lock(m_dlgMainMutex);

        for (auto& [threadId, dlgMain] : m_dlgMainForEachThread) {
            if (dlgMain.IsWindowVisible()) {
                return;
            }
        }
    }

    if (m_unloading.exchange(true)) {
        return;
    }

    // Unload in a new thread to avoid deadlocks and to be able to unload self.
    HANDLE thread = CreateThread(
        nullptr, 0,
        [](LPVOID lpParam) -> DWORD {
            auto watcher = reinterpret_cast<VisualTreeWatcher*>(lpParam);
            watcher->UnadviseVisualTreeChange();
            Sleep(100);
            FreeLibraryAndExitThread(_Module.GetModuleInstance(), 0);
        },
        this, 0, nullptr);
    if (thread) {
        CloseHandle(thread);
    }
}
