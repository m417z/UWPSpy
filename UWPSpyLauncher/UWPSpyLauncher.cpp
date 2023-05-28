#include "stdafx.h"

#include "resource.h"

using start_proc_t = int (WINAPI*)();

int WINAPI wWinMain(
	_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR lpCmdLine,
	_In_ int nCmdShow)
{
	HMODULE lib = LoadLibrary(L"UWPSpy.dll");
	if (!lib) {
		return 1;
	}

	start_proc_t start = (start_proc_t)GetProcAddress(lib, "start");
	if (!start) {
		return 1;
	}

	return start();
}
