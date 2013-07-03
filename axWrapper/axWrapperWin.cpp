/**********************************************************\

\**********************************************************/

#include "axWrapperAPI.h"
#include "axWrapper.h"
#include "axWrapperWin.h"

#ifdef SubclassWindow
#undef SubclassWindow // this is a macro defined in Windowsx.h - it conflicts with ATL CWindow::SubclassWindow function!
#endif

// Define the event function info structures for any events we want
// to catch. These structures are referenced in the event sink map
// used by IDispEventSimpleImpl (see axWrapper.h BEGIN_SINK_MAP).
_ATL_FUNC_INFO efiClick = {CC_STDCALL, VT_EMPTY, 0, {NULL}};
_ATL_FUNC_INFO efiKeyPress = {CC_STDCALL, VT_EMPTY, 1, {VT_I2 | VT_BYREF}};


/**********************************************************\

  implementation for the ActiveX control container class

\**********************************************************/

// ActiveX control event handlers

void __stdcall axWrapperAxWin::onKeyPress(short* pKey)
{
	if(pKey)
	{
		try {
			m_pPlugin->getJSAPI()->fire_keypress(*pKey);
		}
		catch(...)
		{
		}
	}
}

void __stdcall axWrapperAxWin::onClick()
{
	try {
		m_pPlugin->getJSAPI()->fire_click();
	}
	catch(...)
	{
	}
}



