#pragma once

enum ProcessSpyFramework : DWORD {
    kFrameworkUWP = 1,
    kFrameworkWinUI,
};

bool ProcessSpy(HWND hWnd, DWORD pid, ProcessSpyFramework framework);
