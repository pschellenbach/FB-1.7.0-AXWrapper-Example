Example Firebreath plugin wrapper for ActiveX control
-----------------------------------------------------

Created by Peter Schellenbach (pjs@asent.com) January 2011,
and updated July 2013 for Firebreath v. 1.7.0. This project
is for educational purposes. This is free software. Use at
your own risk.

This example includes a VB6 ActiveX control project used
to create FBExampleCtl.ocx. The UserControl in this project
acts like a command button. It has public properties,
methods and events which are manipulated by the Firebreath
plugin. The VB6 project and ocx are in the fbexamplectrl
directory.

Two versions of the plugin project are included. The first,
in the axWrapper directory, runs the ActiveX control in the
same thread as the plugin (browser UI thread?). The second
project, in the axWrapperThread directory, creates a worker
thread to run the ActiveX control. In the threaded example,
properties and methods of the ActiveX control are accessed
by the main thread using standard COM marshalling.

The plugin uses ATL CAxWindow as an ActiveX control container,
and IDispEventSimpleImpl as an event sink for the wrapped
ActiveX control.

The technique used in this example should be fairly
straightforward to apply to other ActiveX controls.

This example plugin was tested with IE10, FireFox 15.0,
Chrome 27.0.1453.116 and Safari 5.1.7. It works as expected
in all tested browsers, except that focus is not handled
properly in most browsers (IE works best).

Sample html files are included in the test directory to
illustrate basic JavaScript functions to get/set plugin
properties, call methods and handle events.

For convenience, compiled versions of the plugin DLLs and
OCX are provided in the test\bin directory.
