/**********************************************************\

\**********************************************************/

#include "axWrapperAPI.h"
#include "axWrapper.h"
#include "axWrapperWin.h"
#include "utf8_tools.h"

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

axWrapperAxWin::axWrapperAxWin() : m_ThisHwnd(NULL)
{
	m_bShutdownRequested = false;
	m_bThreadStarted = false;
}

axWrapperAxWin::~axWrapperAxWin()
{
}

bool axWrapperAxWin::Startup(axWrapperPtr pPlugin, HWND hWnd, const std::string& caption, int theme)
{
	try {
		m_pPlugin = pPlugin; // assign shared_ptr to weak_ptr
		m_thread = boost::thread(&axWrapperAxWin::ThreadProc, this, hWnd, caption, theme);
		return true;
	}
	catch(...) {
	}
	return false;
}

void axWrapperAxWin::Shutdown()
{
	bool did_shutdown = false;
	{
		boost::mutex::scoped_lock lock(m_mutex);
		if(!m_bShutdownRequested && m_bThreadStarted)
		{
			m_bShutdownRequested = true; // we have posted WM_CLOSE to the atCliAxWin window
			m_pPlugin.reset();
			if(::IsWindow(m_ThisHwnd))
				::PostMessage(m_ThisHwnd, WM_CLOSE, 0, 0);
			did_shutdown = true;
		}
	}
	if(did_shutdown)
	{
		// give thread a chance to wrap up
		#ifdef _DEBUG
			BOOL bRemoteDebugger = FALSE;
			if(IsDebuggerPresent() || (CheckRemoteDebuggerPresent(GetCurrentProcess(), &bRemoteDebugger) && bRemoteDebugger))
				m_thread.join(); // don't timeout if debugging cause breaking may cause timer to expire
			else
				m_thread.timed_join(boost::posix_time::time_duration(boost::posix_time::milliseconds(500))); // give thread a chance to wrap up
		#else
			m_thread.timed_join(boost::posix_time::time_duration(boost::posix_time::milliseconds(500))); // give thread a chance to wrap up
		#endif
	}
}

// return marshalled interface for ActiveX control so main thread can call control functions
AXCTLDEFINTF* axWrapperAxWin::getAxCtl()
{
	if(!m_marshal)
	{
		boost::mutex::scoped_lock lock(m_mutex);
		if(m_spstream)
		{
			CoGetInterfaceAndReleaseStream(m_spstream, __uuidof(AXCTLDEFINTF), (void**)&m_marshal);
			m_spstream.p = NULL; // prevent double-release - CoGetInterfaceAndReleaseStream always releases the stream interface
		}
	}
	return m_marshal;
}


//**********************************************************
// The following functions run in the worker thread
//**********************************************************


void axWrapperAxWin::ThreadProc(HWND hWnd, const std::string& caption, int theme)
{
	m_bDisconnected = false;
	m_nShutdownCycle = 0;
	axWrapperAxWinPtr self(shared_from_this()); // ensure I don't destruct until the thread terminates
	HRESULT hr = CoInitialize(NULL); // initialze COM for this thread
	try {
		RECT rc;
		::GetClientRect(hWnd, &rc);
		m_ThisHwnd = Create(hWnd, &rc, 0, WS_VISIBLE|WS_CHILD);
		if(m_ThisHwnd)
		{
			{
				CComPtr<IUnknown> spControl;
				hr = CreateControlEx(AXCTLPROGID, NULL, NULL, &spControl, GUID_NULL, NULL);
				m_spaxctl = spControl; // automatic QueryInterface for CComQIPtr assignment
			}
			if(m_spaxctl != NULL)
			{
				hr = DispEventAdvise((IUnknown*)m_spaxctl);
				if(SUCCEEDED(hr))
				{
					if(InitializeAxControl(caption, theme))
					{
						{
							boost::mutex::scoped_lock lock(m_mutex);
							m_bThreadStarted = true;
							// create an IStream to marshal the control's interface for use by the main thread
							CoMarshalInterThreadInterfaceInStream(__uuidof(AXCTLDEFINTF), m_spaxctl, &m_spstream);
						}
						RunMessagePump();
					}
				}
			}
		}
	} catch(...) { }

	// release resources
	try {
		if(m_spaxctl)
		{
			if(!m_bDisconnected)
			{
				short n = 0;
				DispEventUnadvise((IUnknown*)m_spaxctl);
				// shut down the ActiveX control
				HRESULT hr = /* m_spaxctl->Disconnect(&n) */ 0; // initiate shutdown of ActiveX control
			}
			m_spaxctl.Release(); // this should destroy the AX control
			if(::IsWindow(m_ThisHwnd))
			{
				::SetParent(m_ThisHwnd, NULL); // parent is owned by different thread, cannot destroy child without detaching
				::DestroyWindow(m_ThisHwnd); // this will destroy the AX container window (OnFinalMessage will clear m_ThisHwnd when handling WM_NCDESTROY)
			}			
		}
	} catch(...) { }
	CoUninitialize();
}

bool axWrapperAxWin::InitializeAxControl(const std::string& caption, int theme)
{
	HRESULT hr = E_FAIL;
	try {
		hr = m_spaxctl->put_Caption(CComBSTR(FB::utf8_to_wstring(caption).c_str()));
		if(SUCCEEDED(hr))
			hr = m_spaxctl->put_Theme(static_cast<ThemeConstants>(theme));
	}
	catch(...) {
	}
	if(SUCCEEDED(hr))
		return true;
	return false;
}

void axWrapperAxWin::RunMessagePump()
{
	MSG msg;
	while(GetMessage(&msg, 0, 0, 0) > 0)
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
}

void axWrapperAxWin::OnFinalMessage(HWND hWnd)
{
	m_ThisHwnd = NULL;
}

LRESULT axWrapperAxWin::onClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	bHandled = TRUE;
	m_bDisconnected = true;
	short n = 0;
	if(m_spaxctl)
	{
		DispEventUnadvise((IUnknown*)m_spaxctl);
		// If the ActiveX control may take some time to shut down, initiate the
		// shutdown here. If the control cannot be shut down immediately (n <> 0)
		// then set a timer (25ms) and try again, up to a specified number of attempts (10).
		HRESULT hr = /* m_spaxctl->Disconnect(&n) */ 0; // initiate shutdown of ActiveX control
		if(SUCCEEDED(hr))
		{
			if(n != 0)
			{
				// shutdown not complete - try again later (up to 10 times)
				if(++m_nShutdownCycle < 10)
				{
					if(::SetTimer(m_ThisHwnd, WM_CLOSE, 25, NULL))
					{
						// come back in 25ms & recheck
					} else
						n = 0; // give up if no timer
				} else
					n = 0; // give up after 10 attempts
			}
		} else
			n = 0; // give up if shutdown function fails
	}
	if(n == 0)
	{
		if(::IsWindow(m_ThisHwnd))
		{
			::SetParent(m_ThisHwnd, NULL); // parent is owned by different thread, cannot destroy child without detaching
			::DestroyWindow(m_ThisHwnd); // this will destroy the AX container window (OnFinalMessage will clear m_ThisHwnd when handling WM_NCDESTROY)
			m_ThisHwnd = NULL;
		}
		m_spaxctl.Release(); // destroy the AX control now
		// note: message pump (hence this thread) continues to run until OnFinalMessage (WM_NCDESTROY)
	}
	return 0;
}

LRESULT axWrapperAxWin::onDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	bHandled = FALSE;
	PostQuitMessage(0); // terminate the thread
	return 0;
}

// If the ActiveX control requires some time to shut down, the WM_CLOSE handler will set
// a timer (25ms) up to 10 times to allow the control time to shut down. Here we re-call
// the WM_CLOSE handler when the timer expires.
LRESULT axWrapperAxWin::onTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	if(wParam == WM_CLOSE)
	{
		KillTimer(WM_CLOSE);
		return onClose(uMsg, wParam, lParam, bHandled);
	}
	return 0;
}

LRESULT axWrapperAxWin::onMouseActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	bHandled = TRUE;
	return MA_ACTIVATE;
}

// ActiveX control event handlers

void __stdcall axWrapperAxWin::onKeyPress(short* pKey)
{
	// JSAPI FireEvent is thread safe
	if(axWrapperPtr plugin = m_pPlugin.lock())
	{
		if(pKey)
		{
			try {
				plugin->getJSAPI()->fire_keypress(*pKey);
			}
			catch(...)
			{
			}
		}
	}
}

void __stdcall axWrapperAxWin::onClick()
{
	// JSAPI FireEvent is thread safe
	if(axWrapperPtr plugin = m_pPlugin.lock())
	{
		try {
			plugin->getJSAPI()->fire_click();
		}
		catch(...)
		{
		}
	}
}

