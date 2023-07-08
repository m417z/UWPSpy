#include "stdafx.h"

#include "MainDlg.h"
#include "process_spy.h"

CAppModule _Module;

int WINAPI wWinMain(_In_ HINSTANCE hInstance,
                    _In_opt_ HINSTANCE hPrevInstance,
                    _In_ LPWSTR lpCmdLine,
                    _In_ int nCmdShow) {
    HRESULT hRes = ::CoInitialize(NULL);
    ATLASSERT(SUCCEEDED(hRes));

    AtlInitCommonControls(
        ICC_BAR_CLASSES);  // add flags to support other controls

    hRes = _Module.Init(NULL, hInstance);
    ATLASSERT(SUCCEEDED(hRes));

    int nRet = 0;

    DWORD pid = 0;
    if (__argc >= 2) {
        pid = wcstoul(__wargv[1], nullptr, 0);
    }

    if (pid) {
        ProcessSpy(nullptr, pid);
    } else {
        CMainDlg dlgMain;
        nRet = (int)dlgMain.DoModal();
    }

    _Module.Term();
    ::CoUninitialize();

    return nRet;
}
