#pragma once

#include "targetver.h"

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
#include <atlutil.h>

#include <atlapp.h>

extern CAppModule _Module;

#include <atlwin.h>

#include <atlcrack.h>
#include <atlctrls.h>
#include <atlctrlx.h>
#include <atldlgs.h>
#include <atlframe.h>

// Windows

#include <aclapi.h>
#include <sddl.h>
#include <tlhelp32.h>
