#include "stdafx.h"

#include "simplefactory.hpp"
#include "tap.hpp"
#include "winrt.hpp"

_Use_decl_annotations_ STDAPI DllGetClassObject(REFCLSID rclsid,
                                                REFIID riid,
                                                LPVOID* ppv) try {
    if (rclsid == CLSID_UWPSpyTAP) {
        *ppv = nullptr;
        return winrt::make<SimpleFactory<ExplorerTAP>>().as(riid, ppv);
    } else {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
} catch (...) {
    return winrt::to_hresult();
}

_Use_decl_annotations_ STDAPI DllCanUnloadNow() {
    if (winrt::get_module_lock()) {
        return S_FALSE;
    } else {
        return S_OK;
    }
}
