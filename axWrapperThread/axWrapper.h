/**********************************************************\

  Auto-generated axWrapper.cpp

  This file contains the auto-generated main plugin object
  implementation for the axWrapper project

\**********************************************************/
#ifndef H_axWrapperPLUGIN
#define H_axWrapperPLUGIN

#include "PluginWindow.h"
#include "Win\PluginWindowWin.h"
#include "PluginEvents/MouseEvents.h"
#include "PluginEvents/AttachedEvent.h"
#include "PluginEvents/WindowsEvent.h"

#include "PluginCore.h"

/**********************************************************\
 This Firebreath example is a wrapper for the accompanying
 sample xpcmdbutton VB6 ActiveX user-control (FBExample.vbp
 project).
\**********************************************************/

// Forward declare
class axWrapperAPI;
class axWrapperAxWin;
typedef boost::shared_ptr<axWrapperAxWin> axWrapperAxWinPtr;

/**********************************************************\
  This is the main Firebreath plugin class
\**********************************************************/
class axWrapper : public FB::PluginCore
{
public:
    static void StaticInitialize();
    static void StaticDeinitialize();

public:
    axWrapper();
    virtual ~axWrapper();

public:
    void onPluginReady();
    virtual FB::JSAPIPtr createJSAPI();
    virtual bool IsWindowless() { return false; }
	 virtual bool setReady(); // override this so we can ensure that createJSAPI() gets called!

    BEGIN_PLUGIN_EVENT_MAP()
        EVENTTYPE_CASE(FB::WindowsEvent, onWindowsMessage, FB::PluginWindowWin)
        EVENTTYPE_CASE(FB::AttachedEvent, onWindowAttached, FB::PluginWindow)
        EVENTTYPE_CASE(FB::DetachedEvent, onWindowDetached, FB::PluginWindow)
    END_PLUGIN_EVENT_MAP()

    /** BEGIN EVENTDEF -- DON'T CHANGE THIS LINE **/
	 virtual bool onWindowsMessage(FB::WindowsEvent *evt, FB::PluginWindowWin *);
    virtual bool onWindowAttached(FB::AttachedEvent *evt, FB::PluginWindow *);
    virtual bool onWindowDetached(FB::DetachedEvent *evt, FB::PluginWindow *);
    /** END EVENTDEF -- DON'T CHANGE THIS LINE **/
private:

	FB::BrowserHostPtr getHost() {return m_host;}

	// ActiveX control initialization
	bool CreateAxWindow();
	void DestroyAxWindow();

public:

	// ActiveX control property accessors
	std::string get_Caption();
	void set_Caption(const std::string& caption);

	int get_Theme();
	void set_Theme(int theme);

	// ActiveX control methods
	void FireClick(int duration);

	// provide access to JSAPI functions
   typedef boost::shared_ptr<axWrapperAPI> axWrapperAPIPtr;
	axWrapperAPIPtr getJSAPI() {return m_axwrapperapi;}

private:

	// Flags to ensure we wait until both onPlugInReady & onWindowAttached
	// have been called before creating the CAxWindow.
	bool m_WindowAttached;
	bool m_PlugInReady;

	axWrapperAPIPtr m_axwrapperapi;	// easy access to JSAPI functions
	axWrapperAxWinPtr m_axwin;			// ActiveX control container instance (CAxWindow)

};
typedef boost::shared_ptr<axWrapper> axWrapperPtr;
typedef boost::weak_ptr<axWrapper> axWrapperWeakPtr;


#endif

