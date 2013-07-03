/**********************************************************\

\**********************************************************/
#ifndef H_AXWRAPPERWIN
#define H_AXWRAPPERWIN

#include "PluginWindow.h"
#include "Win\PluginWindowWin.h"
#include "PluginEvents/MouseEvents.h"
#include "PluginEvents/AttachedEvent.h"
#include "PluginEvents/WindowsEvent.h"
#include "PluginCore.h"
#include <atlwin.h>

// Import the ActiveX control's typelib so we can easily call methods, etc.
// on the ActiveX control.
#import "PROGID:FBExampleCtl.xpcmdbutton" no_namespace, raw_interfaces_only

// Define the ProgID for the ActiveX control.
#define AXCTLPROGID L"FBExampleCtl.xpcmdbutton"

// Define the ActiveX control's default & event interfaces. You might
// want to use a type library view like OleView.exe to identify interface
// names. Or you can find them in the registry.
#define AXCTLDEFINTF _xpcmdbutton
#define AXCTLEVTINTF __xpcmdbutton

// If the ActiveX control you are wrapping requires a license key,
// define that key here. Check Microsoft knowledge base article
// Q151771 for information on obtaining the license key for an
// ActiveX control.
//#define AXCTLLICKEY L"xxxxxxxxxxxx"

// Event function info structures are defined in axWrapper.cpp
extern _ATL_FUNC_INFO efiClick;
extern _ATL_FUNC_INFO efiKeyPress;


class axWrapper;
typedef boost::shared_ptr<axWrapper> axWrapperPtr;
typedef boost::weak_ptr<axWrapper> axWrapperWeakPtr;


/**********************************************************\

 Define the ActiveX control container class using ATL's CAxWindow
 template with event support from ATL's IDispEventSimpleImpl
 tempalte.

 Note: if the ActiveX control that you are wrapping requires a
 license key, use CAxWindow2 instead of CAxWindow.

 \**********************************************************/
class axWrapperAxWin : public CWindowImpl<axWrapperAxWin, CAxWindow>,
		public IDispEventSimpleImpl<1, axWrapperAxWin, &__uuidof(AXCTLEVTINTF)>,
		public boost::enable_shared_from_this<axWrapperAxWin>
{
public:

	friend axWrapper;

	DECLARE_WND_SUPERCLASS(NULL, CAxWindow::GetWndClassName())

	axWrapperAxWin();
	virtual ~axWrapperAxWin();

	bool Startup(axWrapperPtr pPlugin, HWND hWnd, const std::string& caption, int theme);
	void Shutdown();

	// provide access to ActiveX control in worker thread via marshalled interface pointer
	AXCTLDEFINTF* getAxCtl();

	BEGIN_MSG_MAP(axWrapperAxWin)
		MESSAGE_HANDLER(WM_CLOSE, onClose)
		MESSAGE_HANDLER(WM_DESTROY, onDestroy)
		MESSAGE_HANDLER(WM_TIMER, onTimer)
		MESSAGE_HANDLER(WM_MOUSEACTIVATE, onMouseActivate)
	END_MSG_MAP()

	LRESULT onClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT onDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT onTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT onMouseActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	// Declare the events from the ActiveX control that we want to catch
	// in our FB plugin. See ATL documentation for IDispEventSimpleImpl for
	// details.
	BEGIN_SINK_MAP(axWrapperAxWin)
		SINK_ENTRY_INFO(1, __uuidof(AXCTLEVTINTF), DISPID_CLICK, onClick, &efiClick)
		SINK_ENTRY_INFO(1, __uuidof(AXCTLEVTINTF), DISPID_KEYPRESS, onKeyPress, &efiKeyPress)
	END_SINK_MAP()

	// Declare the ActiveX control event handler functions
	void __stdcall onClick();
	void __stdcall onKeyPress(short* pKeyCode);

	// override CWindowImplBase::OnFinalMessage in case window is destroyed before shutdown process completes
	virtual void OnFinalMessage(HWND /*hWnd*/);

private:

	void ThreadProc(HWND hWnd, const std::string& caption, int theme);
	bool InitializeAxControl(const std::string& caption, int theme);
	void RunMessagePump();
	
	volatile HWND m_ThisHwnd; // copy of CWindow's hWnd, but with volatile qualifier because its accessed from the main thread and the AX thread
	volatile bool m_bShutdownRequested;
	volatile bool m_bThreadStarted;
	bool m_bDisconnected;
	int m_nShutdownCycle;
	axWrapperWeakPtr m_pPlugin;			// back pointer to containing plugin object
	CComQIPtr<AXCTLDEFINTF> m_spaxctl;	// ActiveX control instance
	CComQIPtr<IStream> m_spstream;		// IStream pointer for interthread marshalling
	CComPtr<AXCTLDEFINTF> m_marshal;		// marshalled interface pointer so main thread can call functions in ActiveX control running in worker thread
	boost::thread m_thread;
	boost::mutex m_mutex;

};

#endif
