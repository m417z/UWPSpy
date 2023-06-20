#pragma once

#include "visualtreewatcher.hpp"
#include "winrt.hpp"

// {4CE91FA6-B898-48AC-BE31-A5485A7435E7}
static constexpr CLSID CLSID_ExplorerTAP = {
    0x4ce91fa6,
    0xb898,
    0x48ac,
    {0xbe, 0x31, 0xa5, 0x48, 0x5a, 0x74, 0x35, 0xe7}};

struct ExplorerTAP
    : winrt::implements<ExplorerTAP, IObjectWithSite, winrt::non_agile> {
    HRESULT STDMETHODCALLTYPE SetSite(IUnknown* pUnkSite) override;
    HRESULT STDMETHODCALLTYPE GetSite(REFIID riid,
                                      void** ppvSite) noexcept override;

   private:
    static winrt::com_ptr<VisualTreeWatcher> s_VisualTreeWatcher;

    winrt::com_ptr<IUnknown> site;
};
