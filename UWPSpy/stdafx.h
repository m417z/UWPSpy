#pragma once

// STL

#include <algorithm>
#include <format>
#include <functional>
#include <optional>
#include <string_view>
#include <unordered_map>
#include <vector>

// WTL

#define NTDDI_VERSION NTDDI_WIN10_RS5
#include <sdkddkver.h>

#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
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

#include <CoreWindow.h>
#include <Unknwn.h>
#include <appmodel.h>
#include <combaseapi.h>
#include <comdef.h>
#include <guiddef.h>
#include <libloaderapi.h>
#include <ocidl.h>
#include <wincodec.h>
#include <windows.ui.xaml.hosting.desktopwindowxamlsource.h>
#include <xamlOM.h>

// C++/WinRT

#undef GetCurrentTime

#include <winrt/base.h>

#include <winrt/windows.foundation.collections.h>
#include <winrt/windows.foundation.h>
#include <winrt/windows.ui.core.h>
#include <winrt/windows.ui.xaml.h>
#include <winrt/windows.ui.xaml.hosting.h>
#include <winrt/windows.ui.xaml.media.h>
