#include "stdafx.h"

#include "tap.hpp"

winrt::com_ptr<VisualTreeWatcher> ExplorerTAP::s_VisualTreeWatcher;

HRESULT ExplorerTAP::SetSite(IUnknown* pUnkSite) try {
    // Only ever 1 VTW at once.
    if (s_VisualTreeWatcher.get()) {
        s_VisualTreeWatcher->Activate();
        throw winrt::hresult_illegal_method_call();
    }

    site.copy_from(pUnkSite);

    if (site) {
        s_VisualTreeWatcher = winrt::make_self<VisualTreeWatcher>(site);
    }

    return S_OK;
} catch (...) {
    return winrt::to_hresult();
}

HRESULT ExplorerTAP::GetSite(REFIID riid, void** ppvSite) noexcept {
    return site.as(riid, ppvSite);
}
