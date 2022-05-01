import datetime
import os
import time

import win32com.client

from minitabexecute import dirgenerate, testresultexcelhandle, inputdatahandle, numberinput, turnscaccheckstr, \
    quickdatanam, numberget, fulldatanam, phasechecktabstr, finishcheckstr, deltefile, loggenera, logconfi, \
    inputdatahandlehigeff, copyfile, turnscachigeffcheckstr, killprocess, getdata
from motorcadinputdata import openmcad, close_instances

global testresultaddr
global testresultafteraddr
global minitbookaddr
global minitbookroundboaddr
global singminitbookaddr
global MotorCAD_File5
global logautorunaddr
global higheffresultfoldaddr
global plotorkbookaddr
global threepmodeladdr
global singlepmodeladdr
global threepmodelroundboaddr
rowmaxp = 6
rowmaxq = 5


def paradifine():
    global testresultaddr
    global testresultafteraddr
    global minitbookaddr
    global minitbookroundboaddr
    global singminitbookaddr
    global MotorCAD_File5
    global logautorunaddr
    global higheffresultfoldaddr
    global plotorkbookaddr
    global threepmodeladdr
    global singlepmodeladdr
    global threepmodelroundboaddr

    diclist = dirgenerate()
    testresultaddr = diclist['testresultaddr']
    testresultafteraddr = diclist['testresultafteraddr']
    minitbookaddr = diclist['minitbookaddr']
    minitbookroundboaddr = diclist['minitbookroundboaddr']
    singminitbookaddr = diclist['singminitbookaddr']
    MotorCAD_File5 = diclist['MotorCAD_File5']
    logautorunaddr = diclist['logautorunaddr']
    higheffresultfoldaddr = diclist['higheffresultfoldaddr']
    plotorkbookaddr = diclist['plotorkbookaddr']
    threepmodeladdr = diclist['threepmodeladdr']
    singlepmodeladdr = diclist['singlepmodeladdr']
    threepmodelroundboaddr = diclist['threepmodelroundboaddr']


def main():
    paradifine()
    deltefile(testresultafteraddr)
    deltefile(MotorCAD_File5)
    logconfi(logautorunaddr)
    numberinput(quickdatanam, finishcheckstr, 0)
    turnscaccheck = numberget(quickdatanam, turnscaccheckstr)
    turnscachigeffcheck = numberget(quickdatanam, turnscachigeffcheckstr)
    phasecheck = numberget(fulldatanam, phasechecktabstr)

    if phasecheck == 3:
        minitbookaddr_ = minitbookaddr
        threepmodeladdr_ = threepmodeladdr
    elif phasecheck == 4:
        minitbookaddr_ = minitbookroundboaddr
        threepmodeladdr_ = singlepmodeladdr
    else:
        minitbookaddr_ = singminitbookaddr
        threepmodeladdr_ = threepmodelroundboaddr
    number = 100
    datalist = []
    if turnscaccheck > 0 or turnscachigeffcheck > 0:
        dataarray = testresultexcelhandle(testresultaddr, testresultafteraddr, number)
        turn = getdata(minitbookaddr_, 24, 10)
        PhaseAdvance01 = getdata(minitbookaddr_, 11, 7)

        numbertures = 0
        for y in range(number):
            if dataarray[y].sum() < 1:
                numbertures = y
                break
        for e in range(9):
            datalist.append(dataarray[0][e][rowmaxp])
        numberinput(quickdatanam, finishcheckstr, 0)
        higheffindex = datalist.index(max(datalist))
        higheffspeed = round(dataarray[0][higheffindex][0], 0)
        higheffcurrent = dataarray[0][higheffindex][1]
        highefftorque = dataarray[0][higheffindex][2]
        datalist.clear()
        inputdatahandle(minitbookaddr_, higheffspeed, higheffcurrent, highefftorque, phasecheck)
        numberinput(quickdatanam, turnscaccheckstr, 1)
        while 1:
            time.sleep(10)
            finishcheck = numberget(quickdatanam, finishcheckstr)
            if finishcheck == 1:
                break

        mcad = None

        if turnscaccheck == 2 and turnscachigeffcheck == 4:

            maxturnscalcother(dataarray, phasecheck, numbertures, turn, mcad, rowmaxp, minitbookaddr_)
            maxturnscalcother(dataarray, phasecheck, numbertures, turn, mcad, rowmaxq, minitbookaddr_)

        elif turnscaccheck == 2:
            maxturnscalcother(dataarray, phasecheck, numbertures, turn, mcad, rowmaxp, minitbookaddr_)

        elif turnscachigeffcheck == 4:
            maxturnscalcother(dataarray, phasecheck, numbertures, turn, mcad, rowmaxq, minitbookaddr_)

        numberinput(quickdatanam, turnscaccheckstr, 8)
        numberinput(quickdatanam, finishcheckstr, 1)
        numberinput(quickdatanam, turnscachigeffcheckstr, 8)
        loggenera('结束时间：' + str(datetime.datetime.now()))
        killprocess()


def higheffcalc(dataarray, minitbookaddr_, phasecheck, number, turnsdatalist, converloslist, threepmodeladdr_):
    for i in range(number):
        datalist = []
        for e in range(9):
            datalist.append(dataarray[i][e][3])
        numberinput(quickdatanam, finishcheckstr, 0)
        higheffindex = datalist.index(max(datalist))
        higheffspeed = round(dataarray[i][higheffindex][0], 0)
        higheffcurrent = dataarray[i][higheffindex][1]
        highefftorque = dataarray[i][higheffindex][2]
        datalist.clear()
        setconverloss(converloslist, i, minitbookaddr_, threepmodeladdr_)
        inputdatahandlehigeff(minitbookaddr_, turnsdatalist[i][2], higheffspeed, higheffcurrent, highefftorque,
                              phasecheck)
        numberinput(quickdatanam, turnscachigeffcheckstr, 7)
        os.system('start autorun.bat')
        while 1:
            time.sleep(10)
            finishcheck = numberget(quickdatanam, finishcheckstr)
            if finishcheck == 1:
                break
        higheffresultfoldaddrfull = higheffresultfoldaddr + r'\%d-%d.xlsx' % (i + 1, higheffspeed)
        copyfile(plotorkbookaddr, higheffresultfoldaddrfull)
    numberinput(quickdatanam, turnscachigeffcheckstr, 0)


def setconverloss(converloslist, i, minitbookaddr_, threepmodeladdr_):
    modelname = getdata(minitbookaddr_, 17, 10)
    filepath = threepmodeladdr_ + r'\%s.mot' % modelname
    mcad = openmcad(filepath)
    mcad.SetVariable('ConverterLosses', converloslist[i][1])
    close_instances(filepath, mcad)


def converloscalc(dataarray, phasecheck, number, turn, mcad, rowsl, minitbookaddr_):
    if mcad is None:
        mcad = win32com.client.Dispatch('motorcad.appautomation')
    turnsdatalist = []
    higheffspeed = 0
    for i in range(number):
        loggenera('i：' + str(i))
        datalist = []
        for e in range(9):
            datalist.append(dataarray[i][e][rowsl])
        numberinput(quickdatanam, finishcheckstr, 0)
        higheffindex = datalist.index(max(datalist))
        higheffspeedpre = higheffspeed
        higheffspeed = round(dataarray[i][higheffindex][0], 0)
        higheffcurrent = dataarray[i][higheffindex][1]
        highefftorque = dataarray[i][higheffindex][2]
        higheffmotoreff = round(dataarray[i][higheffindex][4], 1)
        datalist.clear()
        if i == 0:
            mcad = openmcad(MotorCAD_File5)
        mcad.SetVariable('Shaft_Speed_[RPM]', higheffspeed)
        mcad.SetVariable('MagTurnsConductor', turn)
        higheffcurrent01 = round(higheffcurrent * 1.5, 3)
        mcad.SetVariable('PeakCurrent', higheffcurrent01)
        loggenera('higheffcurrent：' + str(higheffcurrent))
        loggenera('higheffcurrent01：' + str(higheffcurrent01))
        loggenera('higheffspeed：' + str(higheffspeed))
        PhaseAdvance01 = getdata(minitbookaddr_, 11, 7)
        mcad.SetVariable('PhaseAdvance', PhaseAdvance01)
        resultlist = get_parameteremf(mcad, phasecheck)
        DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0 = resultlist[0], resultlist[1], resultlist[2]
        ConverterLosses, RMSTerminalVoltage_PMDC, SystemEfficiency, MeanDCSupplyCurrent, ShaftTorque, PeakCurrent = \
            resultlist[3], resultlist[4], resultlist[5], resultlist[6], resultlist[7], resultlist[8]
        while ConverterLosses <= 0:
            ConverterLosses = ConverterLosses + 1
        mcad.SetVariable('ConverterLosses', ConverterLosses)

        higheffcurrentuplimit = higheffcurrent * 1.01
        higheffcurrentdownlimit = higheffcurrent * 0.99
        higheffmotoreffuplimit = higheffmotoreff * 1.01
        higheffmotoreffdownlimit = higheffmotoreff * 0.99
        loggenera(
            'higheffcurrent, higheffmotoreff, highefftorque' + str([higheffcurrent, higheffmotoreff, highefftorque]))
        loggenera('higheffmotoreffuplimit, higheffmotoreffdownlimit' + str(
            [higheffmotoreffuplimit, higheffmotoreffdownlimit]))

        loggenera(
            '[DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0, ConverterLosses, RMSTerminalVoltage_PMDC, SystemEfficiency, MeanDCSupplyCurrent, ShaftTorque, PeakCurrent]')
        higheffcurrent01, SystemEfficiency, RMSTerminalVoltage_PMDC, ShaftTorque, MeanDCSupplyCurrent = currentcalibrat(
            MeanDCSupplyCurrent,
            higheffcurrent01,
            higheffcurrentdownlimit,
            higheffcurrentuplimit,
            mcad, phasecheck)
        converlosspre = 0
        if i > 0:
            converlosspre = turnsdatalist[i - 1][1]

        while 1:
            if SystemEfficiency > higheffmotoreffuplimit:
                inputpower = DCBusVoltage * MeanDCSupplyCurrent
                miniconvlos = inputpower * 0.01

                ConverterLosses = round(
                    ConverterLosses + inputpower * (SystemEfficiency - higheffmotoreff) / 100, 3)
                if ConverterLosses < miniconvlos:
                    ConverterLosses = miniconvlos
                    PhaseAdvance01 = round(mcad.GetVariable('PhaseAdvance')[1], 2)
                    PhaseAdvance01 = round(
                        PhaseAdvance01 - 3, 2)
                    mcad.SetVariable('PhaseAdvance', PhaseAdvance01)

                    loggenera('SystemEfficiency > higheffmotoreffdownlimit PhaseAdvance01：' + str(PhaseAdvance01))
                if i > 0:
                    if ConverterLosses > converlosspre*3:
                        ConverterLosses = round(converlosspre * higheffspeed / higheffspeedpre, 3)
                        PhaseAdvance01 = round(mcad.GetVariable('PhaseAdvance')[1], 2)
                        PhaseAdvance01 = round(
                            PhaseAdvance01 + 3, 2)
                        mcad.SetVariable('PhaseAdvance', PhaseAdvance01)
                        loggenera('ConverterLosses >= converlosspre PhaseAdvance01：' + str(PhaseAdvance01))
                loggenera('SystemEfficiency < higheffmotoreffdownlimit ConverterLosses：' + str(ConverterLosses))
                mcad.SetVariable('ConverterLosses', ConverterLosses)
                resultlist = get_parameteremf(mcad, phasecheck)
                loggenera('SystemEfficiency > higheffmotoreffuplimit resultlist：' + str(resultlist))

                MeanDCSupplyCurrent = resultlist[6]
                SystemEfficiency = resultlist[5]
                RMSTerminalVoltage_PMDC = resultlist[4]
                ShaftTorque = resultlist[7]

                higheffcurrent01, SystemEfficiency1, RMSTerminalVoltage_PMDC1, ShaftTorque1, MeanDCSupplyCurrent = currentcalibrat(
                    MeanDCSupplyCurrent, higheffcurrent01, higheffcurrentdownlimit,
                    higheffcurrentuplimit, mcad, phasecheck)
                if SystemEfficiency1 != 0:
                    SystemEfficiency = SystemEfficiency1
                    RMSTerminalVoltage_PMDC = RMSTerminalVoltage_PMDC1
                    ShaftTorque = ShaftTorque1

            elif SystemEfficiency < higheffmotoreffdownlimit:
                inputpower = DCBusVoltage * MeanDCSupplyCurrent
                ConverterLosses = round(
                    ConverterLosses - inputpower * (higheffmotoreff - SystemEfficiency) / 100, 3)
                miniconvlos = inputpower * 0.01
                if ConverterLosses < miniconvlos:
                    ConverterLosses = miniconvlos
                    PhaseAdvance01 = round(mcad.GetVariable('PhaseAdvance')[1], 2)
                    PhaseAdvance01 = round(
                        PhaseAdvance01 - 3, 2)
                    mcad.SetVariable('PhaseAdvance', PhaseAdvance01)
                    loggenera('SystemEfficiency < higheffmotoreffdownlimit PhaseAdvance01：' + str(PhaseAdvance01))
                if i > 0:
                    if ConverterLosses > converlosspre*3:
                        ConverterLosses = round(converlosspre * higheffspeed / higheffspeedpre, 3)
                        PhaseAdvance01 = round(mcad.GetVariable('PhaseAdvance')[1], 2)
                        PhaseAdvance01 = round(
                            PhaseAdvance01 + 3, 2)
                        mcad.SetVariable('PhaseAdvance', PhaseAdvance01)
                        loggenera('ConverterLosses >= converlosspre PhaseAdvance01：' + str(PhaseAdvance01))
                loggenera('SystemEfficiency < higheffmotoreffdownlimit ConverterLosses：' + str(ConverterLosses))

                mcad.SetVariable('ConverterLosses', ConverterLosses)
                resultlist = get_parameteremf(mcad, phasecheck)
                loggenera('SystemEfficiency < higheffmotoreffdownlimit resultlist：' + str(resultlist))

                MeanDCSupplyCurrent = resultlist[6]
                SystemEfficiency = resultlist[5]
                RMSTerminalVoltage_PMDC = resultlist[4]
                ShaftTorque = resultlist[7]

                higheffcurrent01, SystemEfficiency1, RMSTerminalVoltage_PMDC1, ShaftTorque1, MeanDCSupplyCurrent = currentcalibrat(
                    MeanDCSupplyCurrent, higheffcurrent01, higheffcurrentdownlimit,
                    higheffcurrentuplimit, mcad, phasecheck)
                if SystemEfficiency1 != 0:
                    SystemEfficiency = SystemEfficiency1
                    RMSTerminalVoltage_PMDC = RMSTerminalVoltage_PMDC1
                    ShaftTorque = ShaftTorque1

            else:
                break
        PhaseAdvance01 = round(mcad.GetVariable('PhaseAdvance')[1], 2)
        turnsdatalist.append(
            [higheffspeed, round(ConverterLosses, 3), PhaseAdvance01, '原', higheffcurrent, highefftorque,
             higheffmotoreff, '计算',
             round(MeanDCSupplyCurrent, 2),
             round(ShaftTorque, 2), round(SystemEfficiency, 1)])

    turnsdataliststr = '\n'
    for i in turnsdatalist:
        turnsdataliststr = turnsdataliststr + str(i) + '\n'
    loggenera('控制损耗计算结果：' + turnsdataliststr)
    numberinput(quickdatanam, turnscachigeffcheckstr, 0)
    return turnsdatalist


def maxturnscalcother(dataarray, phasecheck, number, turn, mcad, rowsl, minitbookaddr_):
    if mcad is None:
        mcad = win32com.client.Dispatch('motorcad.appautomation')
    turnsdatalist = []
    turncuren = turn
    higheffspeed = 0
    converloslist = []

    for i in range(number):
        loggenera('i：' + str(i))
        datalist = []
        for e in range(9):
            datalist.append(dataarray[i][e][rowsl])
        numberinput(quickdatanam, finishcheckstr, 0)
        higheffindex = datalist.index(max(datalist))

        if i == 0:
            higheffspeedpre = round(dataarray[i][higheffindex][0], 0)
        else:
            higheffspeedpre = higheffspeed
        higheffspeed = round(dataarray[i][higheffindex][0], 0)
        higheffcurrent = dataarray[i][higheffindex][1]
        highefftorque = dataarray[i][higheffindex][2]
        higheffmotoreff = round(dataarray[i][higheffindex][4], 1)
        turnpre = turncuren
        turncuren = round(turnpre * higheffspeedpre / higheffspeed, 0)
        datalist.clear()
        if i == 0:
            mcad = openmcad(MotorCAD_File5)
            converloslist = converloscalc(dataarray, phasecheck, number, turn, mcad, rowsl, minitbookaddr_)

        mcad.SetVariable('ConverterLosses', converloslist[i][1])
        mcad.SetVariable('PhaseAdvance', converloslist[i][2])
        mcad.SetVariable('Shaft_Speed_[RPM]', higheffspeed)
        mcad.SetVariable('PeakCurrent', higheffcurrent)
        mcad.SetVariable('Wdg_Definition', 1)
        mcad.SetVariable('MagTurnsConductor', turncuren)

        resultlist = get_parameteremf(mcad, phasecheck)
        DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0 = resultlist[0], resultlist[1], resultlist[2]
        ConverterLosses, RMSTerminalVoltage_PMDC, SystemEfficiency, MeanDCSupplyCurrent, ShaftTorque, PeakCurrent = \
            resultlist[3], resultlist[4], resultlist[5], resultlist[6], resultlist[7], resultlist[8]
        higheffcurrent01 = round(higheffcurrent * 1.5, 3)
        torqueuplimit = highefftorque * 1.03
        torquedownlimit = highefftorque * 0.97
        DCBusVoltageuplimit = DCBusVoltage * 0.97
        DCBusVoltagedownlimit = DCBusVoltage * 0.93
        loggenera(
            'higheffcurrent, higheffmotoreff, highefftorque' + str([higheffcurrent, higheffmotoreff, highefftorque]))
        loggenera('DCBusVoltageuplimit, DCBusVoltagedownlimit' + str(
            [DCBusVoltageuplimit, DCBusVoltagedownlimit]))

        loggenera(
            '[DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0, ConverterLosses, RMSTerminalVoltage_PMDC, SystemEfficiency, MeanDCSupplyCurrent, ShaftTorque, PeakCurrent]')
        higheffcurrent01, SystemEfficiency1, RMSTerminalVoltage_PMDC1, ShaftTorque, PeakLineLineVoltage0 = currenttorquecalibrat(
            ShaftTorque, higheffcurrent01, torquedownlimit, torqueuplimit, mcad, phasecheck)

        while 1:
            if PeakLineLineVoltage0 > DCBusVoltageuplimit:
                turncuren = turncuren - 1
                loggenera('turncuren：' + str(turncuren))
                mcad.SetVariable('MagTurnsConductor', turncuren)
                resultlist = get_parameteremf(mcad, phasecheck)
                loggenera('PeakLineLineVoltage0 > DCBusVoltageuplimit resultlist：' + str(resultlist))

                MeanDCSupplyCurrent = resultlist[6]
                PeakLineLineVoltage0 = resultlist[2]
                ShaftTorque = resultlist[7]
                loggenera('ShaftTorque：' + str(ShaftTorque))

                higheffcurrent01, SystemEfficiency1, RMSTerminalVoltage_PMDC1, ShaftTorque1, PeakLineLineVoltage1 = currenttorquecalibrat(
                    ShaftTorque, higheffcurrent01, torquedownlimit, torqueuplimit, mcad, phasecheck)
                loggenera('ShaftTorque1：' + str(ShaftTorque1))

                if PeakLineLineVoltage1 != 0:
                    PeakLineLineVoltage0 = PeakLineLineVoltage1
                    ShaftTorque = ShaftTorque1
                    loggenera('ShaftTorque：' + str(ShaftTorque))
                    loggenera('PeakLineLineVoltage0：' + str(PeakLineLineVoltage0))
            elif PeakLineLineVoltage0 < DCBusVoltagedownlimit:

                turncuren = turncuren + 1
                loggenera('turncuren：' + str(turncuren))
                mcad.SetVariable('MagTurnsConductor', turncuren)
                resultlist = get_parameteremf(mcad, phasecheck)
                loggenera('PeakLineLineVoltage0 < DCBusVoltageuplimit resultlist：' + str(resultlist))
                MeanDCSupplyCurrent = resultlist[6]
                PeakLineLineVoltage0 = resultlist[2]
                ShaftTorque = resultlist[7]
                loggenera('ShaftTorque：' + str(ShaftTorque))
                higheffcurrent01, SystemEfficiency1, RMSTerminalVoltage_PMDC1, ShaftTorque1, PeakLineLineVoltage1 = currenttorquecalibrat(
                    ShaftTorque, higheffcurrent01, torquedownlimit, torqueuplimit, mcad, phasecheck)
                loggenera('ShaftTorque1：' + str(ShaftTorque1))

                if PeakLineLineVoltage1 != 0:
                    PeakLineLineVoltage0 = PeakLineLineVoltage1
                    ShaftTorque = ShaftTorque1
                    loggenera('ShaftTorque：' + str(ShaftTorque))
                    loggenera('PeakLineLineVoltage0：' + str(PeakLineLineVoltage0))
            else:
                break

        PhaseAdvance01 = round(mcad.GetVariable('PhaseAdvance')[1], 2)

        turnsdatalist.append(
            [higheffspeed, turncuren, round(ConverterLosses, 3), PhaseAdvance01, '原', higheffcurrent, highefftorque,
             higheffmotoreff, '计算',
             PeakLineLineVoltage0, ShaftTorque])

    turnsdataliststr = '\n'
    for i in turnsdatalist:
        turnsdataliststr = turnsdataliststr + str(i) + '\n'
    if rowsl == rowmaxp:
        loggenera(
            '全开匝数计算结果：\n' + '[higheffspeed, turncuren, ConverterLosses, 原, higheffcurrent, highefftorque, higheffmotoreff, 计算, PeakLineLineVoltage0, ShaftTorque] \n' + turnsdataliststr)
    else:
        loggenera(
            '背压匝数计算结果：\n' + '[higheffspeed, turncuren, ConverterLosses, 原, higheffcurrent, highefftorque, higheffmotoreff, 计算, PeakLineLineVoltage0, ShaftTorque] \n' + turnsdataliststr)
    turnsdatalist.clear()
    close_instances(MotorCAD_File5, mcad)


def currentcalibrat(MeanDCSupplyCurrent, higheffcurrent01, higheffcurrentdownlimit, higheffcurrentuplimit, mcad,
                    phasecheck):
    SystemEfficiency = 0
    RMSTerminalVoltage_PMDC = 0
    ShaftTorque = 0

    while 1:
        if MeanDCSupplyCurrent > higheffcurrentuplimit:
            higheffcurrent01pre = higheffcurrent01
            higheffcurrent01 = round(higheffcurrent01 * higheffcurrentuplimit / MeanDCSupplyCurrent, 2)
            if higheffcurrent01 == higheffcurrent01pre:
                higheffcurrent01 = higheffcurrent01 - 0.01
            mcad.SetVariable('PeakCurrent', higheffcurrent01)
            resultlist = get_parameteremf(mcad, phasecheck)
            loggenera('MeanDCSupplyCurrent > higheffcurrentuplimit resultlist：' + str(resultlist))

            MeanDCSupplyCurrent = resultlist[6]
            SystemEfficiency = resultlist[5]
            RMSTerminalVoltage_PMDC = resultlist[4]
            ShaftTorque = resultlist[7]

        elif MeanDCSupplyCurrent < higheffcurrentdownlimit:
            higheffcurrent01pre = higheffcurrent01
            higheffcurrent01 = round(higheffcurrent01 * higheffcurrentdownlimit / MeanDCSupplyCurrent, 2)
            if higheffcurrent01 == higheffcurrent01pre:
                higheffcurrent01 = higheffcurrent01 + 0.01
            mcad.SetVariable('PeakCurrent', higheffcurrent01)
            resultlist = get_parameteremf(mcad, phasecheck)
            loggenera('MeanDCSupplyCurrent < higheffcurrentdownlimit resultlist：' + str(resultlist))

            MeanDCSupplyCurrent = resultlist[6]
            SystemEfficiency = resultlist[5]
            RMSTerminalVoltage_PMDC = resultlist[4]
            ShaftTorque = resultlist[7]
        else:
            break

    return higheffcurrent01, SystemEfficiency, RMSTerminalVoltage_PMDC, ShaftTorque, round(MeanDCSupplyCurrent, 2)


def currenttorquecalibrat(targetorque, higheffcurrent01, torquedownlimit, torqueuplimit, mcad, phasecheck):
    SystemEfficiency = 0
    RMSTerminalVoltage_PMDC = 0
    shaft_torque = targetorque
    PeakLineLineVoltage0 = 0
    loggenera('targetorque torquedownlimit resultlistt：' + str([shaft_torque, torquedownlimit, torqueuplimit]))

    while 1:
        if shaft_torque > torqueuplimit:
            higheffcurrent01pre = higheffcurrent01
            higheffcurrent01 = round(higheffcurrent01 * torqueuplimit / shaft_torque, 2)
            if higheffcurrent01 == higheffcurrent01pre:
                higheffcurrent01 = higheffcurrent01 - 0.01
            mcad.SetVariable('PeakCurrent', higheffcurrent01)
            resultlistt = get_parameteremf(mcad, phasecheck)
            loggenera('targetorque > torqueuplimit resultlistt：' + str(resultlistt))

            MeanDCSupplyCurrent = resultlistt[6]
            SystemEfficiency = resultlistt[5]
            RMSTerminalVoltage_PMDC = resultlistt[4]
            shaft_torque = resultlistt[7]
            PeakLineLineVoltage0 = resultlistt[2]
            loggenera('shaft_torque：' + str(shaft_torque))

        elif shaft_torque < torquedownlimit:
            higheffcurrent01pre = higheffcurrent01
            higheffcurrent01 = round(higheffcurrent01 * torquedownlimit / shaft_torque, 2)
            if higheffcurrent01 == higheffcurrent01pre:
                higheffcurrent01 = higheffcurrent01 + 0.01
            mcad.SetVariable('PeakCurrent', higheffcurrent01)
            resultlistt = get_parameteremf(mcad, phasecheck)
            loggenera('targetorque < torquedownlimit resultlistt：' + str(resultlistt))

            SystemEfficiency = resultlistt[5]
            RMSTerminalVoltage_PMDC = resultlistt[4]
            shaft_torque = resultlistt[7]
            PeakLineLineVoltage0 = resultlistt[2]
            loggenera('shaft_torque：' + str(shaft_torque))
        else:
            break

    return higheffcurrent01, SystemEfficiency, RMSTerminalVoltage_PMDC, shaft_torque, PeakLineLineVoltage0


def maxturnscalc(dataarray, minitbookaddr_, phasecheck, number, mcad, threepmodeladdr_, turn):
    if mcad is None:
        mcad = win32com.client.Dispatch('motorcad.appautomation')
    turnsdatalist = []
    converloslist = []
    for i in range(number):
        datalist = []
        for e in range(9):
            datalist.append(dataarray[i][e][3])

        maxsspeed = dataarray[i][8][0] + 300
        maxscurrent = dataarray[i][8][1] * 1.7
        maxstorque = dataarray[i][8][2]
        datalist.clear()
        inputdatahandle(minitbookaddr_, maxsspeed, maxscurrent, maxstorque, phasecheck)
        if i == 0:
            numberinput(quickdatanam, turnscaccheckstr, 1)
        while 1:
            if i == 0:
                time.sleep(10)
                finishcheck = numberget(quickdatanam, finishcheckstr)
                if finishcheck == 1:
                    break
            else:
                break
        if i == 0:
            mcad = openmcad(MotorCAD_File5)
            converloslist = converloscalc(dataarray, phasecheck, number, turn, mcad, 0, minitbookaddr_)
        setconverloss(converloslist, i, minitbookaddr_, threepmodeladdr_)

        mcad.SetVariable('Shaft_Speed_[RPM]', maxsspeed)
        mcad.SetVariable('PeakCurrent', maxscurrent)
        resultlist = get_parameteremf(mcad, phasecheck)
        DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0 = resultlist[0], resultlist[1], resultlist[2]
        ArmatureTurnsPerCoil02 = round(DCBusVoltage * ArmatureTurnsPerCoil01 / PeakLineLineVoltage0, 0)
        mcad.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil02)
        resultlist = get_parameteremf(mcad, phasecheck)
        DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0 = resultlist[0], resultlist[1], resultlist[2]
        ArmatureTurnsPerCoil01, PeakLineLineVoltage01 = voltagecalebra(ArmatureTurnsPerCoil01, DCBusVoltage,
                                                                       PeakLineLineVoltage0, mcad, phasecheck)

        turnsdatalist.append([maxsspeed, PeakLineLineVoltage01, ArmatureTurnsPerCoil01])
    turnsdataliststr = '\n'
    for i in turnsdatalist:
        turnsdataliststr = turnsdataliststr + str(i) + '\n'
    loggenera('最大开环匝数计算结果：' + turnsdataliststr)
    close_instances(MotorCAD_File5, mcad)

    return turnsdatalist, converloslist


def voltagecalebra(ArmatureTurnsPerCoil01, DCBusVoltage, PeakLineLineVoltage0, mcad, phasecheck):
    PeakLineLineVoltage01 = PeakLineLineVoltage0
    PeakLineLineVoltage0 = round(PeakLineLineVoltage0, 2)
    if PeakLineLineVoltage0 > DCBusVoltage:
        while PeakLineLineVoltage0 >= DCBusVoltage:
            ArmatureTurnsPerCoil02 = round(ArmatureTurnsPerCoil01 - 1, 0)
            mcad.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil02)
            resultlist = get_parameteremf(mcad, phasecheck)
            DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0 = resultlist[0], resultlist[1], resultlist[2]
            PeakLineLineVoltage0 = round(PeakLineLineVoltage0, 2)
            PeakLineLineVoltage01 = PeakLineLineVoltage0

    elif PeakLineLineVoltage0 < DCBusVoltage:
        while PeakLineLineVoltage0 < DCBusVoltage:
            ArmatureTurnsPerCoil02 = round(ArmatureTurnsPerCoil01 + 1, 0)
            mcad.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil02)
            PeakLineLineVoltage01 = PeakLineLineVoltage0
            resultlist = get_parameteremf(mcad, phasecheck)
            DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0 = resultlist[0], resultlist[1], resultlist[2]
            PeakLineLineVoltage0 = round(PeakLineLineVoltage0, 2)
        ArmatureTurnsPerCoil01 = ArmatureTurnsPerCoil01 - 1

    else:
        ArmatureTurnsPerCoil02 = round(ArmatureTurnsPerCoil01 - 1, 0)
        mcad.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil02)
        resultlist = get_parameteremf(mcad, phasecheck)
        DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0 = resultlist[0], resultlist[1], resultlist[2]
        PeakLineLineVoltage0 = round(PeakLineLineVoltage0, 2)
        PeakLineLineVoltage01 = PeakLineLineVoltage0
    return ArmatureTurnsPerCoil01, PeakLineLineVoltage01


def get_parameteremf(mcad, phasecheck):
    if mcad is None:
        mcad = win32com.client.Dispatch('motorcad.appautomation')
    success = mcad.DoMagneticCalculation()
    messagestr = mcad.GetMessages(1)[1]
    PeakLineLineVoltage0 = 0.1
    if success == 0:
        if phasecheck == 3 or phasecheck == 4:
            PeakLineLineVoltage0 = round(mcad.GetVariable('RmsLineLineVoltage')[1], 1)
        else:
            PeakLineLineVoltage0 = round(mcad.GetVariable('RmsPhaseVoltage')[1], 1)

    elif success == -1 or 'fail' in messagestr or 'not' in messagestr or 'Unable' in messagestr or 'fatal' in messagestr or 'error' in messagestr:
        pass
    DCBusVoltage = round(mcad.GetVariable('DCBusVoltage')[1], 1)
    ArmatureTurnsPerCoil01 = round(mcad.GetVariable('MagTurnsConductor')[1], 0)
    ConverterLosses = round(mcad.GetVariable('ConverterLosses')[1], 3)
    RMSTerminalVoltage_PMDC = round(mcad.GetVariable('RMSTerminalVoltage_PMDC')[1], 1)
    SystemEfficiency = round(mcad.GetVariable('SystemEfficiency')[1], 1)
    MeanDCSupplyCurrent = round(mcad.GetVariable('MeanDCSupplyCurrent')[1], 2)
    ShaftTorque = round(mcad.GetVariable('ShaftTorque')[1] * 1000, 1)

    PeakCurrent = round(mcad.GetVariable('PeakCurrent')[1], 2)

    resultlistret = [DCBusVoltage, ArmatureTurnsPerCoil01, PeakLineLineVoltage0, ConverterLosses,
                     RMSTerminalVoltage_PMDC, SystemEfficiency, MeanDCSupplyCurrent, ShaftTorque, PeakCurrent]

    return resultlistret


if __name__ == '__main__':
    main()
