#!/usr/bin/env python

"""
This demonstrates a simple use of delayedresult: get/compute
something that takes a long time, without hanging the GUI while this
is taking place.

The top button runs a small GUI that uses wx.lib.delayedresult.startWorker
to wrap a long-running function into a separate thread. Just click
Get, and move the slider, and click Get and Abort a few times, and
observe that GUI responds. The key functions to look for in the code
are startWorker() and __handleResult().

The second button runs the same GUI, but without delayedresult.  Click
Get: now the get/compute is taking place in main thread, so the GUI
does not respond to user actions until worker function returns, it's
not even possible to Abort.
"""
import os
import time

import wx
import wx.grid as gridlib
import wx.lib.delayedresult as delayedresult
from wx import Point, Size

from minitabexecute import pythoncomundo, killprocessall, dirgenerate, openexcel, inputfilehan, namelistsingle, \
    namelist, fulldatanam, creattable, numberinput, phasechecktabstr, numberget, quickdatanam, finishcheckstr, \
    rownumber, rownumbersingle, killprocess, minitabop, copyfile, deltefile, getdata, namelistthreer, \
    rownumbersthreerounbo, turnscaccheckstr, logconfi, loggenera, turnscachigeffcheckstr, killexcellprocess
from motorcadinputdata import openmcad

global minitbookaddr
global plotorkbookaddr
global minitbookaddr_temp
global threepmodeladdr
global MotorCAD_Fileshot
global singminitbookaddr
global singminitbookaddr_temp
global singlepmodeladdr
global logautorunaddr
global setupdir
global minitabaddr
global logautorunaddrtem
global minitabaddr_tem
global plotorkbookaddr_tem
global modeladdr
global Base_File_path
global minitbookroundboaddr
global minitbookroundboaddr_temp
global threepmodelroundboaddr
hadImportError = False

try:
    import numpy
    import wx.lib.plot
except ImportError:
    hadImportError = True


def paradifine():
    global minitbookaddr
    global plotorkbookaddr
    global minitbookaddr_temp
    global threepmodeladdr
    global MotorCAD_Fileshot
    global singminitbookaddr
    global singminitbookaddr_temp
    global singlepmodeladdr
    global logautorunaddr
    global setupdir
    global minitabaddr
    global logautorunaddrtem
    global minitabaddr_tem
    global plotorkbookaddr_tem
    global modeladdr
    global Base_File_path
    global minitbookroundboaddr
    global minitbookroundboaddr_temp
    global threepmodelroundboaddr
    diclist = dirgenerate()
    minitbookaddr = diclist['minitbookaddr']
    plotorkbookaddr = diclist['plotorkbookaddr']
    minitbookaddr_temp = diclist['minitbookaddr_temp']
    threepmodeladdr = diclist['threepmodeladdr']
    MotorCAD_Fileshot = diclist['MotorCAD_Fileshot']
    singminitbookaddr = diclist['singminitbookaddr']
    singminitbookaddr_temp = diclist['singminitbookaddr_temp']
    singlepmodeladdr = diclist['singlepmodeladdr']
    logautorunaddr = diclist['logautorunaddr']
    setupdir = diclist['setupdir']
    minitabaddr = diclist['minitabaddr']
    logautorunaddrtem = diclist['logautorunaddrtem']
    minitabaddr_tem = diclist['minitabaddr_tem']
    plotorkbookaddr_tem = diclist['plotorkbookaddr_tem']
    minitbookroundboaddr = diclist['minitbookroundboaddr']
    minitbookroundboaddr_temp = diclist['minitbookroundboaddr_temp']
    threepmodelroundboaddr = diclist['threepmodelroundboaddr']

    modeladdr = diclist['modeladdr']
    Base_File_path = MotorCAD_Fileshot + '.mot'


# ---------------------------------------------------------------------------

class HugeTable(gridlib.GridTableBase):

    def __init__(self, log):
        gridlib.GridTableBase.__init__(self)
        self.log = log

        self.odd = gridlib.GridCellAttr()
        self.odd.SetBackgroundColour("sky blue")
        self.even = gridlib.GridCellAttr()
        self.even.SetBackgroundColour("sea green")

    def GetAttr(self, row, col, kind):
        attr = [self.even, self.odd][row % 2]
        attr.IncRef()
        return attr

    # This is all it takes to make a custom data table to plug into a
    # wxGrid.  There are many more methods that can be overridden, but
    # the ones shown below are the required ones.  This table simply
    # provides strings containing the row and column values.

    def GetNumberRows(self):
        return 10000

    def GetNumberCols(self):
        return 10000

    def IsEmptyCell(self, row, col):
        return False

    def GetValue(self, row, col):
        return str((row, col))

    def SetValue(self, row, col, value):
        self.log.write('SetValue(%d, %d, "%s") ignored.\n' % (row, col, value))


# ---------------------------------------------------------------------------


class HugeTableGrid(gridlib.Grid):
    def __init__(self, parent, log):
        gridlib.Grid.__init__(self, parent, -1)

        table = HugeTable(log)

        # The second parameter means that the grid is to take
        # ownership of the table and will destroy it when done.
        # Otherwise you would need to keep a reference to it, but that
        # would allow other grids to use the same table.
        self.SetTable(table, True)

        self.Bind(gridlib.EVT_GRID_CELL_RIGHT_CLICK, self.OnRightDown)

    def OnRightDown(self, event):
        i = 0


# ---------------------------------------------------------------------------


class TestPanel(wx.Panel):

    def __init__(self, parent, log):
        self.log = log
        wx.Panel.__init__(self, parent, -1)
        self.result = 0
        self.jobID = 0
        self.signal = 1
        self.runingsignal = 1
        self.phasesignal = 0
        self.abortEvent = delayedresult.AbortEvent()
        self.timecount1 = 0
        self.num = 0
        self.j = 0
        self.minitbookaddr_ = minitbookaddr
        self.rownumber_ = rownumber
        self.threepmodeladdr_ = threepmodeladdr
        self.namelist_ = namelist
        self.minitbookaddr_temp_ = minitbookaddr_temp
        vsizer = wx.BoxSizer(wx.VERTICAL)
        b = wx.Button(self, pos=Point(20, 30), size=Size(100, 60), label="三相圆底槽录入")
        vsizer.Add(b, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonthreeinproundbo, b)

        bs = wx.Button(self, pos=Point(20, 130), size=Size(100, 60), label="三相圆底槽运行")
        vsizer.Add(bs, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonrunthreeroundbo, bs)

        b = wx.Button(self, pos=Point(153, 30), size=Size(100, 60), label="三相录入")
        vsizer.Add(b, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonthreeinp, b)

        bs = wx.Button(self, pos=Point(153, 130), size=Size(100, 60), label="三相运行")
        vsizer.Add(bs, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonrunthree, bs)

        b = wx.Button(self, pos=Point(285, 30), size=Size(100, 60), label="单相录入")
        vsizer.Add(b, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonsingleinp, b)

        bs = wx.Button(self, pos=Point(285, 130), size=Size(100, 60), label="单相运行")
        vsizer.Add(bs, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonrunsingle, bs)

        bsm = wx.Button(self, pos=Point(20, 240), size=Size(100, 60), label="当前模板")
        vsizer.Add(bsm, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonshowmo, bsm)

        bss = wx.Button(self, pos=Point(20, 330), size=Size(100, 60), label="显示结果")
        vsizer.Add(bss, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonshowresul, bss)

        bsm = wx.Button(self, pos=Point(153, 240), size=Size(100, 60), label="新建模板")
        vsizer.Add(bsm, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonopenfilefold, bsm)

        bss = wx.Button(self, pos=Point(153, 330), size=Size(100, 60), label="重置")
        vsizer.Add(bss, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonreset, bss)

        bsmm = wx.Button(self, pos=Point(285, 240), size=Size(100, 60), label="趋势分析")
        vsizer.Add(bsmm, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonshowminitab, bsmm)

        bssm = wx.Button(self, pos=Point(285, 330), size=Size(100, 60), label="显示日志")
        vsizer.Add(bssm, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonshowlog, bssm)

        # b2 = wx.Button(self, -1, "显示曲线")
        # vsizer.Add(b2, 0, wx.ALL, 5)
        # self.Bind(wx.EVT_BUTTON, self.OnButton2, b2)

        b3 = wx.Button(self, pos=Point(130, 430), size=Size(150, 80), label="退出")
        vsizer.Add(b3, 0, wx.ALL, 20)
        self.Bind(wx.EVT_BUTTON, self.OnButtonstop, b3)

        # grid = HugeTableGrid(self, log)
        # vsizer.Add(grid, 1, wx.EXPAND)

    def handleGet(self):
        """Compute result in separate thread, doesn't affect GUI response."""
        self.abortEvent.clear()
        self.jobID += 1
        delayedresult.startWorker(self._resultConsumer, self._resultProducer,
                                  wargs=(), jobID=self.jobID)

    def _resultProducer(self):
        """
        Pretend to be a complex worker function or something that takes
        long time to run due to network access etc. GUI will freeze if this
        method is not called in separate thread.
        """
        result = self.resultprofun()
        return result

    def resultprofun(self):
        turnscaccheck = numberget(quickdatanam, turnscaccheckstr)
        turnscachigeffcheck = numberget(quickdatanam, turnscachigeffcheckstr)

        if turnscachigeffcheck > 0:
            while 1:
                time.sleep(10)
                finishcheck = numberget(quickdatanam, finishcheckstr)
                turnscachigeffcheck = numberget(quickdatanam, turnscachigeffcheckstr)
                if finishcheck == 1 and turnscachigeffcheck == 8:
                    return 1

        elif turnscaccheck > 0 and turnscachigeffcheck == 0:
            while 1:
                time.sleep(10)
                finishcheck = numberget(quickdatanam, finishcheckstr)
                turnscaccheck = numberget(quickdatanam, turnscaccheckstr)
                if finishcheck == 1 and turnscaccheck == 8:
                    return 1

        else:
            while 1:
                time.sleep(10)
                finishcheck = numberget(quickdatanam, finishcheckstr)
                if finishcheck == 1:
                    return 1


    def _resultConsumer(self, delayedResult):
        jobID = delayedResult.getJobID()
        try:
            self.result = delayedResult.get()
            if self.result == 1:
                self.whilefini('计算完成!', '电机优化程序通知')
        except Exception as exc:
            return

    def whilefini(self, content, title):
        self.signal = 0
        self.runingsignal = 0
        dig = wx.MessageDialog(self, content, title, wx.OK | wx.ICON_INFORMATION)
        dig.ShowModal()
        dig.Destroy()

    def OnButtonstop(self, evt):
        self.signal = 0
        self.runingsignal = 1
        self.phasesignal = 0
        pythoncomundo()
        killprocessall()
        killexcellprocess()

    def OnButton2(self, evt):
        from wx.lib.plot.examples.demo import PlotDemoMainFrame
        win = PlotDemoMainFrame(self, -1, "wx.lib.plot Demo")
        win.Show()

    def OnButtonrunthree(self, evt):
        if self.signal == 0 and self.phasesignal == 3:
            self.signal = 1
            self.runingsignal = 0
            self.startcalc(self.phasesignal)
            # self.listgenerate(self.phasesignal)
            self.handleGet()

    def OnButtonrunthreeroundbo(self, evt):
        if self.signal == 0 and self.phasesignal == 4:
            self.signal = 1
            self.runingsignal = 0
            self.startcalc(self.phasesignal)
            # self.listgenerate(self.phasesignal)
            self.handleGet()

    def OnButtonrunsingle(self, evt):
        if self.signal == 0 and self.phasesignal == 1:
            self.signal = 1
            self.runingsignal = 0
            self.startcalc(self.phasesignal)
            # self.listgenerate(self.phasesignal)
            self.handleGet()

    def OnButtonthreeinp(self, evt):
        openexcel(minitbookaddr)
        if self.runingsignal == 1:
            self.phasesignal = 3
            self.signal = 0

    def OnButtonthreeinproundbo(self, evt):
        openexcel(minitbookroundboaddr)
        if self.runingsignal == 1:
            self.phasesignal = 4
            self.signal = 0

    def OnButtonsingleinp(self, evt):
        openexcel(singminitbookaddr)
        if self.runingsignal == 1:
            self.phasesignal = 1
            self.signal = 0

    def OnButtonshowresul(self, evt):
        deltefile(plotorkbookaddr_tem)
        copyfile(plotorkbookaddr, plotorkbookaddr_tem)
        openexcel(plotorkbookaddr_tem)

    def OnButtonshowmo(self, evt):
        modelname = getdata(self.minitbookaddr_, 17, 10)
        if modelname is not None:
            openmcad(self.threepmodeladdr_ + r'\%s.mot' % modelname)
        else:
            self.whilefini('请指定模板!', '错误提示')

    def OnButtonshowminitab(self, evt):
        deltefile(minitabaddr_tem)
        copyfile(minitabaddr, minitabaddr_tem)
        minitab, extra = minitabop(minitabaddr_tem)
        minitab.UserInterface.Visible = True
        minitab.UserInterface.UserControl = True

    def OnButtonshowlog(self, evt):
        deltefile(logautorunaddrtem)
        copyfile(logautorunaddr, logautorunaddrtem)
        os.startfile(logautorunaddrtem)

    def OnButtonopenfilefold(self, evt):
        os.startfile(modeladdr)

    def OnButtonreset(self, evt):
        self.signal = 0
        self.runingsignal = 1
        self.phasesignal = 0
        pythoncomundo()
        killprocess()
        killexcellprocess()

    def startcalc(self, phaseinpu):
        self.threesingde(phaseinpu)
        numberinput(fulldatanam, phasechecktabstr, phaseinpu)
        numberinput(quickdatanam, finishcheckstr, 0)
        numberinput(quickdatanam, turnscachigeffcheckstr, 0)
        numberinput(quickdatanam, turnscaccheckstr, 0)
        inputfilehan(self.minitbookaddr_, self.minitbookaddr_temp_, self.threepmodeladdr_, MotorCAD_Fileshot,
                     self.namelist_)
        turnscacchecksign = getdata(self.minitbookaddr_, 21, 10)
        loggenera('turnscacchecksign：' + str(turnscacchecksign))
        if turnscacchecksign == 3:
            numberinput(quickdatanam, turnscachigeffcheckstr, 4)
            numberinput(quickdatanam, turnscaccheckstr, 2)
        elif turnscacchecksign == 1:
            numberinput(quickdatanam, turnscaccheckstr, 2)
        elif turnscacchecksign == 2:
            numberinput(quickdatanam, turnscachigeffcheckstr, 4)

        os.system('start turnscalculate.bat')
        os.system('start autorun.bat')
        os.system('start alertwindowscomferm.bat')




    def threesingde(self, phasecheckinpu):
        if phasecheckinpu == 3:
            self.minitbookaddr_ = minitbookaddr
            self.rownumber_ = rownumber
            self.threepmodeladdr_ = threepmodeladdr
            self.namelist_ = namelist
            self.minitbookaddr_temp_ = minitbookaddr_temp
        elif phasecheckinpu == 1:
            self.minitbookaddr_ = singminitbookaddr
            self.rownumber_ = rownumbersingle
            self.threepmodeladdr_ = singlepmodeladdr
            self.namelist_ = namelistsingle
            self.minitbookaddr_temp_ = singminitbookaddr_temp
        elif phasecheckinpu == 4:
            self.minitbookaddr_ = minitbookroundboaddr
            self.rownumber_ = rownumbersthreerounbo
            self.threepmodeladdr_ = threepmodelroundboaddr
            self.namelist_ = namelistthreer
            self.minitbookaddr_temp_ = minitbookroundboaddr_temp

# ---------------------------------------------------------------------------


def runTest(frame, nb, log):
    paradifine()
    creattable(fulldatanam, phasechecktabstr, 0)
    logconfi(logautorunaddr)
    if not hadImportError:
        win = TestPanel(nb, log)
    else:
        from wx.lib.msgpanel import MessagePanel
        win = MessagePanel(nb, """\
This demo requires the numpy module, which could not be imported.
It probably is not installed (it's not part of the standard Python
distribution). See https://pypi.python.org/pypi/numpy for information
about the numpy package.""", 'Sorry', wx.ICON_WARNING)

    return win


# ---------------------------------------------------------------------------


overview = __doc__
