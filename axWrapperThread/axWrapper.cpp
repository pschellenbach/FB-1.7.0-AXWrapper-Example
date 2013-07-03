/**********************************************************\

  Auto-generated axWrapper.cpp

  This file contains the auto-generated main plugin object
  implementation for the axWrapper project

\**********************************************************/

#include "NpapiTypes.h"
#include "variant_list.h"
#include "axWrapperAPI.h"
#include "axWrapperWin.h"
#include "axWrapper.h"

///////////////////////////////////////////////////////////////////////////////
/// @fn axWrapper::StaticInitialize()
///
/// @brief  Called from PluginFactory::globalPluginInitialize()
///
/// @see FB::FactoryBase::globalPluginInitialize
///////////////////////////////////////////////////////////////////////////////
void axWrapper::StaticInitialize()
{
    // Place one-time initialization stuff here; note that there isn't an absolute guarantee that
    // this will only execute once per process, just a guarantee that it won't execute again until
    // after StaticDeinitialize is called
}

///////////////////////////////////////////////////////////////////////////////
/// @fn axWrapper::StaticInitialize()
///
/// @brief  Called from PluginFactory::globalPluginDeinitialize()
///
/// @see FB::FactoryBase::globalPluginDeinitialize
///////////////////////////////////////////////////////////////////////////////
void axWrapper::StaticDeinitialize()
{
    // Place one-time deinitialization stuff here
}

///////////////////////////////////////////////////////////////////////////////
/// @brief  axWrapper constructor.  Note that your API is not available
///         at this point, nor the window.  For best results wait to use
///         the JSAPI object until the onPluginReady method is called
///////////////////////////////////////////////////////////////////////////////
axWrapper::axWrapper()
{
	m_WindowAttached = false;
	m_PlugInReady = false;
}

///////////////////////////////////////////////////////////////////////////////
/// @brief  axWrapper destructor.
///////////////////////////////////////////////////////////////////////////////
axWrapper::~axWrapper()
{
}

void axWrapper::onPluginReady()
{
    // When this is called, the BrowserHost is attached, the JSAPI object is
    // created, and we are ready to interact with the page and such.  The
    // PluginWindow may or may not have already fire the AttachedEvent at
    // this point.
	// Ensure that we have received WindowAttached event before creating the CAxWindow
	if(m_PlugInReady == false)
	{
		m_PlugInReady = true;
		if(m_WindowAttached)
			CreateAxWindow();
	}
}

///////////////////////////////////////////////////////////////////////////////
/// @brief  Creates an instance of the JSAPI object that provides your main
///         Javascript interface.
///
/// Note that m_host is your BrowserHost and shared_ptr returns a
/// FB::PluginCorePtr, which can be used to provide a
/// boost::weak_ptr<axWrapper> for your JSAPI class.
///
/// Be very careful where you hold a shared_ptr to your plugin class from,
/// as it could prevent your plugin class from getting destroyed properly.
///////////////////////////////////////////////////////////////////////////////
FB::JSAPIPtr axWrapper::createJSAPI()
{
    // m_host is the BrowserHost
    //return boost::make_shared<axWrapperAPI>(FB::ptr_cast<axWrapper>(shared_from_this()), m_host);
    m_axwrapperapi = boost::make_shared<axWrapperAPI>(FB::ptr_cast<axWrapper>(shared_from_this()), m_host);
    return m_axwrapperapi;
}

bool axWrapper::setReady()
{
	 // The call to getRootJSAPI() was removed from PluginCore::setReady() between version
	 // 1.5 and 1.6. Without this call, the plugin JSAPI object may not get created unless
	 // the html <object> tag includes an "onload" param. This plugin relies on the pointer
	 // to the JSAPI object being valid in order to fire events.

    // Ensure that the JSAPI object has been created, in case the browser hasn't requested it yet.
    getRootJSAPI(); 
	 return PluginCore::setReady();
}

bool axWrapper::onWindowsMessage(FB::WindowsEvent *evt, FB::PluginWindowWin *piw)
{
	bool rc = false;
	if(!m_axwin)
		return false;
	switch(evt->uMsg)
	{
		case WM_SIZE:
			// resize the ActiveX container window
			m_axwin->MoveWindow(0, 0, LOWORD(evt->lParam), HIWORD(evt->lParam));
			evt->lRes = 0;
			rc = true;
			break;
		case WM_MOUSEACTIVATE:
			// activate on mouse click
			evt->lRes = MA_ACTIVATE;
			rc = true;
			break;
		case WM_SETFOCUS:
			// forward focus to the ActiveX control container window
			m_axwin->SetFocus();
			evt->lRes = 0;
			rc = true;
			break;
	}
	return rc;
}

bool axWrapper::onWindowAttached(FB::AttachedEvent *evt, FB::PluginWindow *piw)
{
   // The window is attached; act appropriately

	// Strangely, we can get this function called multiple times by
	// some browsers. If we already have created the CAxWindow, just
	// check if the plugin window is still the same.

	// Ensure that we have received WindowAttached event before creating the CAxWindow
	if(m_WindowAttached == false)
	{
		m_WindowAttached = true;
		if(m_PlugInReady)
			CreateAxWindow();
	}
	return false;
}

bool axWrapper::onWindowDetached(FB::DetachedEvent *evt, FB::PluginWindow *)
{
    // The window is about to be detached; act appropriately
    DestroyAxWindow();
    return false;
}

bool axWrapper::CreateAxWindow()
{
	try {

		/* Now that we have the plugin window, create the ActiveX container
		   window as a child of the plugin, then create the ActiveX control
			as a child of the container.
		*/
		
		FB::PluginWindowWin* pwnd = static_cast<FB::PluginWindowWin*>(GetWindow());
		if(pwnd != NULL)
		{
			HWND hWnd = pwnd->getHWND();
			if(hWnd)
			{				

				// Get the initialization parameters
				std::string caption;
				int theme = 0;

				try {
					caption = m_params["caption"].convert_cast<std::string>();
				} catch(...) {} // ignore missing param
				try {
					theme = m_params["theme"].convert_cast<long>();
				} catch(...) {} // ignore missing param

				m_axwin = axWrapperAxWinPtr(new axWrapperAxWin);
				return m_axwin->Startup(FB::ptr_cast<axWrapper>(shared_from_this()), hWnd, caption, theme);
			}
		}
	} catch(...) {
		//TODO: should we throw a FB exception here?
	}
	return false;
}

void axWrapper::DestroyAxWindow()
{
	if(m_axwin)
		m_axwin->Shutdown();
	m_axwin.reset();
}

std::string axWrapper::get_Caption()
{
	if(m_axwin && m_axwin->getAxCtl()) // cross-thread call must be marshalled
	{
		try {
			CComBSTR bstr;
			HRESULT hr = m_axwin->getAxCtl()->get_Caption(&bstr);
			if(SUCCEEDED(hr))
				return FB::wstring_to_utf8(std::wstring(bstr.m_str, bstr.Length()));
		}
		catch(...) {
		}
	}
	return std::string(); // punt
}

void axWrapper::set_Caption(const std::string& caption)
{
	if(m_axwin && m_axwin->getAxCtl()) // cross-thread call must be marshalled
	{
		try {
			HRESULT hr = m_axwin->getAxCtl()->put_Caption(CComBSTR(FB::utf8_to_wstring(caption).c_str()));
		}
		catch(...) {
		}
	}
}

int axWrapper::get_Theme()
{
	if(m_axwin && m_axwin->getAxCtl()) // cross-thread call must be marshalled
	{
		try {
			ThemeConstants theme;
			HRESULT hr = m_axwin->getAxCtl()->get_Theme(&theme);
			if(SUCCEEDED(hr))
				return (int)theme;
		}
		catch(...) {
		}
	}
	return 0; // punt
}

void axWrapper::set_Theme(int theme)
{
	if(m_axwin && m_axwin->getAxCtl()) // cross-thread call must be marshalled
	{
		try {
			HRESULT hr = m_axwin->getAxCtl()->put_Theme(static_cast<ThemeConstants>(theme));
		}
		catch(...) {
		}
	}
}

void axWrapper::FireClick(int duration)
{
	if(m_axwin && m_axwin->getAxCtl()) // cross-thread call must be marshalled
	{
		try {
			HRESULT hr = m_axwin->getAxCtl()->FireClick(duration);
		}
		catch(...) {
		}
	}
}
