#pragma once

#include "MainDlg.h"
#include "winrt.hpp"

class VisualTreeWatcher : public winrt::implements<VisualTreeWatcher,
                                                   IVisualTreeServiceCallback2,
                                                   winrt::non_agile> {
   public:
    VisualTreeWatcher(winrt::com_ptr<IUnknown> site);

    void Activate();

    VisualTreeWatcher(const VisualTreeWatcher&) = delete;
    VisualTreeWatcher& operator=(const VisualTreeWatcher&) = delete;

    VisualTreeWatcher(VisualTreeWatcher&&) = delete;
    VisualTreeWatcher& operator=(VisualTreeWatcher&&) = delete;

    ~VisualTreeWatcher();

    void AdviseVisualTreeChange();
    void UnadviseVisualTreeChange();

   private:
    HRESULT STDMETHODCALLTYPE
    OnVisualTreeChange(ParentChildRelation relation,
                       VisualElement element,
                       VisualMutationType mutationType) override;
    HRESULT STDMETHODCALLTYPE
    OnElementStateChanged(InstanceHandle element,
                          VisualElementState elementState,
                          LPCWSTR context) noexcept override;

    template <typename T>
    T FromHandle(InstanceHandle handle) {
        wf::IInspectable obj;
        winrt::check_hresult(m_xamlDiagnostics->GetIInspectableFromHandle(
            handle, reinterpret_cast<::IInspectable**>(winrt::put_abi(obj))));

        return obj.as<T>();
    }

    CMainDlg* DlgMainForCurrentThread();
    void OnDlgMainEvent(HWND hWnd, CMainDlg::EventId eventId);
    void OnDlgMainHidden(HWND hWnd);
    void OnDlgMainFinalMessage(HWND hWnd);
    void UnloadIfNoMoreVisibleDlgs();

    winrt::com_ptr<VisualTreeWatcher> m_selfPtr;
    winrt::com_ptr<IXamlDiagnostics> m_xamlDiagnostics;

    std::shared_mutex m_dlgMainMutex;
    std::unordered_map<DWORD, CMainDlg> m_dlgMainForEachThread;

    std::atomic<bool> m_unloading = false;
};
