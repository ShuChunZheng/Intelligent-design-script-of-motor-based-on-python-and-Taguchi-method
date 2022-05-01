#!/usr/bin/env python

# ----------------------------------------------------------------------------
# Name:         run.py
# Purpose:      Simple framework for running individual demos
#
# Author:       Robin Dunn
#
# Created:      6-March-2000
# Copyright:    (c) 2000-2020 by Total Control Software
# Licence:      wxWindows license
# ----------------------------------------------------------------------------

"""
This program will load and run one of the individual demos in this
directory within its own frame window.  Just specify the module name
on the command line.
"""
import wx
import wx.lib.inspection
import wx.lib.mixins.inspection

assertMode = wx.APP_ASSERT_DIALOG


# assertMode = wx.APP_ASSERT_EXCEPTION


class Log:
    def WriteText(self, text):
        if text[-1:] == '\n':
            text = text[:-1]
        wx.LogMessage(text)

    write = WriteText


class RunDemoApp(wx.App, wx.lib.mixins.inspection.InspectionMixin):
    def __init__(self, name, module):
        self.name = name
        self.demoModule = module
        wx.App.__init__(self, redirect=False)

    def OnInit(self):
        wx.Log.SetActiveTarget(wx.LogStderr())

        self.SetAssertMode(assertMode)
        self.InitInspection()  # for the InspectionMixin base class

        frame = wx.Frame(None, -1, "电机优化程序控制界面", size=(800, 600),
                         style=wx.DEFAULT_FRAME_STYLE, name="run a sample")
        frame.CreateStatusBar()
        menuBar = wx.MenuBar()
        menu = wx.Menu()
        item = menu.Append(-1, "&Widget Inspector\tF6", "Show the wxPython Widget Inspection Tool")
        self.Bind(wx.EVT_MENU, self.OnWidgetInspector, item)
        item = menu.Append(wx.ID_EXIT, "E&xit\tCtrl-Q", "Exit demo")
        self.Bind(wx.EVT_MENU, self.OnExitApp, item)
        menuBar.Append(menu, "&File")

        ns = {'wx': wx, 'app': self, 'module': self.demoModule, 'frame': frame}

        frame.SetMenuBar(menuBar)
        frame.Show(True)
        frame.Bind(wx.EVT_CLOSE, self.OnCloseFrame)

        win = self.demoModule.runTest(frame, frame, Log())

        # a window will be returned if the demo does not create
        # its own top-level window
        if win:
            # so set the frame to a good size for showing stuff
            frame.SetSize((420, 600))
            win.SetFocus()
            self.window = win
            ns['win'] = win
            frect = frame.GetRect()

        else:
            # It was probably a dialog or something that is already
            # gone, so we're done.
            frame.Destroy()
            return True

        self.SetTopWindow(frame)
        self.frame = frame
        # wx.Log.SetActiveTarget(wx.LogStderr())
        # wx.Log.SetTraceMask(wx.TraceMessages)

        return True

    def OnExitApp(self, evt):
        self.frame.Close(True)

    def OnCloseFrame(self, evt):
        if hasattr(self, "window") and hasattr(self.window, "ShutdownDemo"):
            self.window.ShutdownDemo()
        evt.Skip()

    def OnWidgetInspector(self, evt):
        wx.lib.inspection.InspectionTool().Show()


def main():
    name = 'DelayedResult'
    module = __import__(name)
    app = RunDemoApp(name, module)
    locale = wx.Locale(wx.LANGUAGE_DEFAULT)
    app.MainLoop()


if __name__ == "__main__":
    main()
