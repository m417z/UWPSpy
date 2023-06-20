#pragma once

#include "MainDlg.h"
#include "winrt.hpp"

struct VisualTreeWatcher : winrt::implements<VisualTreeWatcher,
                                             IVisualTreeServiceCallback2,
                                             winrt::non_agile> {
    VisualTreeWatcher(winrt::com_ptr<IUnknown> site);

    void Activate();

    VisualTreeWatcher(const VisualTreeWatcher&) = delete;
    VisualTreeWatcher& operator=(const VisualTreeWatcher&) = delete;

    VisualTreeWatcher(VisualTreeWatcher&&) = delete;
    VisualTreeWatcher& operator=(VisualTreeWatcher&&) = delete;

    ~VisualTreeWatcher();

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

    winrt::com_ptr<IXamlDiagnostics> m_xamlDiagnostics;
    std::optional<CMainDlg> m_dlgMain;
};
