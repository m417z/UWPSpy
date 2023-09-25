#pragma once

#include "targetver.h"

// STL

#include <algorithm>
#include <format>
#include <functional>
#include <optional>
#include <shared_mutex>
#include <string_view>
#include <unordered_map>
#include <vector>

// WTL

#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#define NOMINMAX
#define _WTL_NO_CSTRING
#define _WTL_NO_WTYPES
#define _WTL_NO_UNION_CLASSES
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS

#include <atlbase.h>
#include <atlstr.h>
#include <atltypes.h>

#include <atlapp.h>

extern CAppModule _Module;

#include <atlwin.h>

#include <atlcrack.h>
#include <atlctrls.h>
#include <atldlgs.h>
#include <atlframe.h>

// Windows

#include <appmodel.h>
#include <combaseapi.h>
#include <comdef.h>
#include <corewindow.h>
#include <guiddef.h>
#include <libloaderapi.h>
#include <ocidl.h>
#include <tlhelp32.h>
#include <unknwn.h>
#include <wincodec.h>
#include <windows.ui.xaml.hosting.desktopwindowxamlsource.h>
#include <xamlom.h>

// C++/WinRT

#undef GetCurrentTime

#include <winrt/base.h>

#include <winrt/windows.foundation.collections.h>
#include <winrt/windows.foundation.h>
#include <winrt/windows.ui.core.h>
#include <winrt/windows.ui.xaml.h>
#include <winrt/windows.ui.xaml.hosting.h>
#include <winrt/windows.ui.xaml.media.h>

// WinUI

#include "winui/microsoft.ui.xaml.hosting.desktopwindowxamlsource_winui.h"

// For IVisualTreeServiceCallback3 if it ever becomes useful.
// #include "winui/xamlom.winui.h"

#include "winrt/microsoft.ui.xaml.h"
#include "winrt/microsoft.ui.xaml.hosting.h"
#include "winrt/microsoft.ui.xaml.media.h"
