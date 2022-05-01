# -*- coding:utf-8 -*-
import datetime
import os
import sqlite3
import time

from collectdatatoexcel import main as maincollect
from minitabexecute import main as mainexecut, dirgenerate, setzero, appendlist, getdata, setdata, plotcurve, \
    runfatorn, getsysteminf, killprocess, creatplotexcel, pythoncomundo, quickedit, mainsqlite, plothodata, returnrown, \
    quickdatanam, numberget, fulldatanam, creattable, numberinput, numbercheckstr, quickcalcheckstr, namelist, \
    namelistsingle, phasechecktabstr, rownumber, rownumbersingle, timeratiocheckstr, threetimeratiocheckstr, \
    finishcheckstr, loggenera, mutation, extract_data_from_excellist, writelist, logdebug, \
    motorinputreducesigetstr, motorinpufincheckstr, logconfi, namelistthreer, rownumbersthreerounbo, turnscaccheckstr, \
    turnscachigeffcheckstr
from minitabgenerate import main as maingener
from minitabperformance import main as mainperform
from mintabextract import main as mainextra
from separatedatatoexcel import main as mainsepara

global plotorkbookaddr
global workbookaddr
global minitbookaddr
global minitbookaddr_temp
global skruratio
global MotorCAD_Fileshot
global singminitbookaddr
global timeratio
global logautorunaddr
global sqlitetablename
global minitbookroundboaddr
sleeptime = 20
separatest = '--------------------------------------------------------------------'
separatestq = '===================================================================='
separatesth = '****************************************************************************'
finishcutline = '________________________________________________完 成_____________________________________________________'


def paradifine():
    global plotorkbookaddr
    global workbookaddr
    global minitbookaddr
    global minitbookaddr_temp
    global MotorCAD_Fileshot
    global singminitbookaddr
    global logautorunaddr
    global minitbookroundboaddr

    diclist = dirgenerate()
    plotorkbookaddr = diclist['plotorkbookaddr']
    workbookaddr = diclist['workbookaddr']
    minitbookaddr = diclist['minitbookaddr']
    minitbookaddr_temp = diclist['minitbookaddr_temp']
    MotorCAD_Fileshot = diclist['MotorCAD_Fileshot']
    singminitbookaddr = diclist['singminitbookaddr']
    logautorunaddr = diclist['logautorunaddr']
    minitbookroundboaddr = diclist['minitbookroundboaddr']


def calculatemotorcad2(phasecheck, minitbookaddr_):
    setzero(minitbookaddr_)
    maincollect(phasecheck)
    mainextra(phasecheck=phasecheck)
    mainexecut(phasecheck=phasecheck)
    killprocess()


def calculatemotorcad21(phasecheck, runs1, minitbookaddr_):
    mainperform(phasecheck)
    mainsepara(phasecheck)
    setzero(minitbookaddr_)
    numberinput(fulldatanam, quickcalcheckstr, 1)
    os.system('start motorcad1.bat')
    time.sleep(15)
    os.system('start motorcad2.bat')
    time.sleep(15)
    os.system('start motorcad3.bat')
    time.sleep(15)
    os.system('start motorcad4.bat')
    if runs1 % 4 == 0:
        boolremain = 0
    else:
        boolremain = 1
    times1 = (runs1 // 4 + boolremain) * 5
    time.sleep(times1)


def calculatemotorcad4(minitbookaddr_, phasecheck):
    setzero(minitbookaddr_)
    maincollect(phasecheck)
    mainextra(signal=1, phasecheck=phasecheck)
    mainexecut(signal=1, phasecheck=phasecheck)
    killprocess()


def calculatemotorcad1(rownumber_, phasecheck, minitbookaddr_, runs1, reducesi=1):
    numberinput(fulldatanam, quickcalcheckstr, 1)
    motorcadinputhandle(reducesi, minitbookaddr_, rownumber_, phasecheck)
    maingener(phasecheck)
    mainsepara(phasecheck)
    setzero(minitbookaddr_)
    os.system('start motorcad1.bat')
    time.sleep(15)
    os.system('start motorcad2.bat')
    time.sleep(15)
    os.system('start motorcad3.bat')
    time.sleep(15)
    os.system('start motorcad4.bat')
    if runs1 % 4 == 0:
        boolremain = 0
    else:
        boolremain = 1
    times1 = (runs1 // 4 + boolremain) * 10
    time.sleep(times1)


def motorcadinputhandle(reducesi, minitbookaddr_, rownumber_, phasecheck):
    numberinput(fulldatanam, motorinputreducesigetstr, reducesi)
    numberinput(fulldatanam, motorinpufincheckstr, 2)
    os.system('start motorcadinputdata.bat')
    timecount1 = time.time()
    while 1:
        time.sleep(15)
        motorinpufin = numberget(fulldatanam, motorinpufincheckstr)
        timecount2 = time.time()
        logdebug(motorinpufin)
        logdebug(timecount2 - timecount1)
        if motorinpufin == 1:
            break
        elif timecount2 - timecount1 > 100:
            killprocess()
            timecount1 = timecount2
            mutationphase(minitbookaddr_, rownumber_, phasecheck)
            os.system('start motorcadinputdata.bat')
    global sqlitetablename
    if sqlitetablename == '':
        sqlitetablename = getdata(minitbookaddr_, 18, 10)
    logdebug(sqlitetablename)


def calculatemotorcad111(rownumber_, phasecheck, minitbookaddr_, runs1, reducesi=1):
    numberinput(fulldatanam, quickcalcheckstr, 1)
    motorcadinputhandle(reducesi, minitbookaddr_, rownumber_, phasecheck)
    maingener(phasecheck)
    mainsepara(phasecheck)
    setzero(minitbookaddr_)
    os.system('start motorcad5.bat')
    if runs1 % 4 == 0:
        boolremain = 0
    else:
        boolremain = 1
    times1 = (runs1 // 4 + boolremain) * 10
    time.sleep(times1)


def calculatemotorcad3(rownumber_, phasecheck, minitbookaddr_, runs1, reducesi=1):
    numberinput(fulldatanam, quickcalcheckstr, 2)
    motorcadinputhandle(reducesi, minitbookaddr_, rownumber_, phasecheck)
    maingener(phasecheck)
    mainsepara(phasecheck)
    setzero(minitbookaddr_)
    os.system('start motorcad1.bat')
    time.sleep(15)
    os.system('start motorcad2.bat')
    time.sleep(15)
    os.system('start motorcad3.bat')
    time.sleep(15)
    os.system('start motorcad4.bat')
    if runs1 % 4 == 0:
        boolremain = 0
    else:
        boolremain = 1
    times1 = (runs1 // 4 + boolremain) * 5
    time.sleep(times1)


def quickcalculate(phasecheck, rownumber_, runs1, factornumber, minitbookaddr_):
    quicktime = 0
    quicksign11 = 0
    timestart = time.time()
    timetemple = timestart
    overtimec = 0
    reducesi = 1
    while quicksign11 == 0:
        timecoump = time.time()
        timecoumpf = 0
        signal = 1
        listsum = 0
        calculatemotorcad3(rownumber_, phasecheck, minitbookaddr_, runs1, reducesi)
        if factornumber < 5:
            overtimec, quicksign11, reducesi, timetemple = quickcalmod(factornumber, listsum, minitbookaddr_, overtimec,
                                                                       phasecheck, quicksign11, quicktime, rownumber_,
                                                                       runs1, signal, timecoump, timecoumpf,
                                                                       timetemple, 3)

        else:
            overtimec, quicksign11, reducesi, timetemple = quickcalmod(factornumber, listsum, minitbookaddr_, overtimec,
                                                                       phasecheck, quicksign11, quicktime, rownumber_,
                                                                       runs1, signal, timecoump, timecoumpf,
                                                                       timetemple, 4)

        quicktime = quicktime + 1
        if quicktime > 15:
            break

    setdata(minitbookaddr_, 17, 9, 0)
    setdata(minitbookaddr_, 19, 9, 0)
    timetemple21 = time.time()
    loggenera('快算总耗时：' + str(int(timetemple21 - timestart)) + '  ' + '时间：' + str(datetime.datetime.now()))
    loggenera(separatestq)


def quickcalmod(factornumber, listsum, minitbookaddr_, overtimec, phasecheck, quicksign11, quicktime, rownumber_, runs1,
                signal, timecoump, timecoumpf, timetemple, motorcanum):
    while 1:
        if overtimec > 0:
            overtimec = 0
            reducesi = 0.5
            killprocess()
            break
        else:
            reducesi = 1
        time.sleep(sleeptime)
        listsin1 = appendlist(minitbookaddr_)
        signal, timecoumpf, listsum, timecoump, overtimec = restart(minitbookaddr_, rownumber_, phasecheck,
                                                                    factornumber, listsum, motorcanum, timecoump,
                                                                    timecoumpf, signal, listsin1, overtimec)

        if sum(listsin1) == motorcanum:
            overtimec = 0
            calculatemotorcad4(minitbookaddr_, phasecheck)
            quicksign11 = getdata(minitbookaddr_, 17, 9)
            if quicksign11 is None:
                quicksign11 = 0
            mainsqlite(rownumber_, minitbookaddr_, workbookaddr, runs1, sqlitetablename, 1, quickdatanam)
            timecount31 = time.time()
            loggenera('快算计次：' + str(quicktime + 1) + '  ' + '耗时：' + str(int(timecount31 - timetemple)))
            loggenera(separatest)
            timetemple = timecount31
            break
    return overtimec, quicksign11, reducesi, timetemple


def plotcureco(phasecheck, i, timetemple2, listsin2, minitbookaddr_, rownumber_):
    setzero(minitbookaddr_)
    maincollect(phasecheck)
    plotcurve(workbookaddr, plotorkbookaddr, str(i + 1), listsin2, rownumber_)
    setdata(minitbookaddr_, 14, 9, i)
    timecount3 = time.time()
    loggenera('细算计次：' + str(i + 1) + '  ' + '耗时：' + str(int(timecount3 - timetemple2)))
    loggenera(separatesth)
    timetemple2 = timecount3
    return timetemple2


def restartmc(inttime, signallist=None):
    if signallist is None:
        signallist = [1 for _ in range(4)]
    indexn = 1
    killprocess()
    for i in range(inttime):
        if signallist[i] == 0:
            commstr = 'start motorcad%d.bat' % (i + indexn)
            os.system(commstr)
            time.sleep(15)


def restart(minitbookaddr_, rownumber_, phasecheck, factornumber, listsum1, inttime, timecoump, timecoumpf,
            signal, listsin=None, overtimec=0):
    if listsin is None:
        listsin = [1 for _ in range(inttime)]
    timecoumpf1 = timecoumpf
    signal1 = signal
    listsumtem = sum(listsin)
    listsum = listsum1
    timecoump1 = timecoump
    overtimech = overtimec
    if listsumtem > 0 and signal1 == 1:
        timecoumpf1 = int(time.time() - timecoump1)
        signal1 = 0
        listsum = listsumtem
    elif listsumtem == 0:
        timecoumpf2 = int(time.time() - timecoump1)
        if factornumber < 5:
            if timecoumpf2 > 240 * skruratio * timeratio:
                mutationphase(minitbookaddr_, rownumber_, phasecheck)
                overtimech = overtimech + 1

        elif factornumber < 14:
            if timecoumpf2 > 560 * skruratio * timeratio:
                mutationphase(minitbookaddr_, rownumber_, phasecheck)
                overtimech = overtimech + 1

        elif factornumber < 26:
            if timecoumpf2 > 1120 * skruratio * timeratio:
                mutationphase(minitbookaddr_, rownumber_, phasecheck)
                overtimech = overtimech + 1

        else:
            if timecoumpf2 > 640 * skruratio * timeratio:
                mutationphase(minitbookaddr_, rownumber_, phasecheck)
                overtimech = overtimech + 1

    elif inttime > listsumtem >= listsum and signal1 == 0:
        timecoumpfe = int(time.time() - timecoump1)
        listsum = listsumtem
        if timecoumpfe > timecoumpf1 + 200 * skruratio * timeratio:
            timecoump1 = time.time()
            mutationphase(minitbookaddr_, rownumber_, phasecheck)
            overtimech = overtimech + 1

    logdebug((timecoumpf1, listsum, overtimech, signal1, listsumtem))
    return signal1, timecoumpf1, listsum, timecoump1, overtimech


def mutationphase(minitbookaddr_, rownumber_, phasecheck):
    if phasecheck == 1:
        listtemple = []
        extract_data_from_excellist(minitbookaddr_, rownumber_, 1, 2, 2, 'Sheet', listtemple)
        mutation(minitbookaddr_, 0, rownumber_, listtemple)
        exc1, exc2, get_num_lista = runfatorn(minitbookaddr_, rownumber_)
        writelist(minitbookaddr_, rownumber_, listtemple, get_num_lista)
        logdebug(get_num_lista)
        logdebug(listtemple)


def restartpl(runs, listsum1, inttime, timecoump, timecoumpf, signal, listsin=None, overtimec=0):
    if listsin is None:
        listsin = [1 for _ in range(inttime)]
    timecoumpf1 = timecoumpf
    signal1 = signal
    listsumtem = sum(listsin)
    listsum = listsum1
    timecoump1 = timecoump
    overtimech = overtimec
    if listsumtem > 0 and signal1 == 1:
        timecoumpf1 = int(time.time() - timecoump1)
        signal1 = 0
        listsum = listsumtem
    elif listsumtem == 0:
        timecoumpf2 = int(time.time() - timecoump1)
        # loggenera(timecoumpf2)
        if runs % 4 == 0:
            boolremain = 0
        else:
            boolremain = 1
        timelimi = ((runs // 4) + boolremain) * 80 * timeratio * skruratio + 100

        if timecoumpf2 > timelimi:
            timecoump1 = time.time()
            restartmc(inttime, listsin)

    elif inttime > listsumtem >= listsum and signal1 == 0:
        timecoumpfe = int(time.time() - timecoump1)
        listsum = listsumtem
        delaytime = timecoumpf1 * 0.5
        if delaytime < 200:
            delaytime = 200
        if timecoumpfe > timecoumpf1 + delaytime:
            timecoump1 = time.time()
            restartmc(inttime, listsin)
    logdebug((timecoumpf1, listsum, overtimech, signal1, listsumtem))
    return signal1, timecoumpf1, listsum, timecoump1, overtimech


def atstartset():
    quickedit(0)
    paradifine()
    logconfi(logautorunaddr)
    global sqlitetablename
    sqlitetablename = ''
    loggenera('开始时间：' + str(datetime.datetime.now()))
    #getsysteminf()
    killprocess()
    creattable(fulldatanam, numbercheckstr, 0)
    numbercn = numberget(fulldatanam, numbercheckstr)
    numbercn = numbercn + 1
    numberinput(fulldatanam, numbercheckstr, numbercn)
    creattable(fulldatanam, quickcalcheckstr, 0)

    return numbercn


def maincal():
    numbercn = atstartset()
    phasecheck = numberget(fulldatanam, phasechecktabstr)
    global timeratio
    if phasecheck == 3:
        minitbookaddr_ = minitbookaddr
        rownumber_ = rownumber
    elif phasecheck == 4:
        minitbookaddr_ = minitbookroundboaddr
        rownumber_ = rownumbersthreerounbo
    else:
        minitbookaddr_ = singminitbookaddr
        rownumber_ = rownumbersingle

    if phasecheck == 3 or phasecheck == 4:
        timeratioinpu = getdata(minitbookaddr_, 10, 9)
        timeratio = 1
        if timeratioinpu is not None:
            timeratio = timeratioinpu
            numberinput(fulldatanam, threetimeratiocheckstr, timeratioinpu)
        else:
            try:
                timeratio = numberget(fulldatanam, threetimeratiocheckstr)
            except sqlite3.OperationalError:
                loggenera('no table')

        if timeratio <= 0 or timeratio is None or timeratio > 10:
            timeratio = 1
        if phasecheck == 4:
            timeratio = timeratio * 2
    else:
        timeratioinpu = getdata(minitbookaddr_, 5, 9)
        timeratio = 5
        if timeratioinpu is not None:
            timeratio = timeratioinpu
            numberinput(fulldatanam, timeratiocheckstr, timeratioinpu)
        else:
            try:
                timeratio = numberget(fulldatanam, timeratiocheckstr)
            except sqlite3.OperationalError:
                loggenera('no table')
        if timeratio <= 0 or timeratio is None or timeratio > 10:
            timeratio = 5

    timestem = getdata(minitbookaddr_, 11, 9)
    if timestem is None:
        times = 3
    else:
        times = int(timestem)

        if times <= 0:
            times = 1
        elif times > 500:
            times = 500
    logdebug('times: ' + str(times))
    listsin2 = getdata(minitbookaddr_, 2, 9)
    startrowh = listsin2 + 5
    startrow = startrowh + returnrown + 3
    toque = getdata(minitbookaddr_, 3, 9)
    if toque is None:
        toque = 0
    cogging = getdata(minitbookaddr_, 4, 9)
    if cogging is None:
        cogging = toque * 0.3

    if phasecheck == 3:
        creatplotexcel(plotorkbookaddr, listsin2, times, 2, startrowh, startrow, returnrown, toque, cogging, namelist)
    elif phasecheck == 4:
        creatplotexcel(plotorkbookaddr, listsin2, times, 2, startrowh, startrow, returnrown, toque, cogging,
                       namelistthreer)
    else:
        creatplotexcel(plotorkbookaddr, listsin2, times, 2, startrowh, startrow, returnrown, toque, cogging,
                       namelistsingle)
    global skruratio
    skruchecks = getdata(minitbookaddr_, 20, 2)
    skrucheckr = getdata(minitbookaddr_, 21, 2)
    if skruchecks is None:
        skruchecks = 0
    if skrucheckr is None:
        skrucheckr = 0
    if skruchecks > 0 or skrucheckr > 0:
        skruratio = 2
    else:
        skruratio = 1

    timecount1 = time.time()
    runs, factornumber, exc = runfatorn(minitbookaddr_, rownumber_)
    logdebug((runs, factornumber))
    if numbercn < 10000:
        if (phasecheck == 3 or phasecheck == 4) and factornumber >= 2:
            quickcalculate(phasecheck, rownumber_, runs, factornumber, minitbookaddr_)
        timetemple2 = time.time()
        itimecount = 0
        overtimec = 0
        reducesi = 1
        while itimecount < times:
            timecoump = time.time()
            timecoumpf = 0
            signal = 1
            listsum = 0
            if factornumber < 1:
                calculatemotorcad111(rownumber_, phasecheck, minitbookaddr_, runs, reducesi)
                while 1:
                    listsin3 = appendlist(minitbookaddr_)
                    if listsin3[0] == 1:
                        break
                break
            elif factornumber < 2:
                calculatemotorcad1(rownumber_, phasecheck, minitbookaddr_, runs, reducesi)
                timecoump1 = time.time()
                timecoumpf1 = 0
                signal1 = 1
                listsum1 = 0
                while 1:
                    time.sleep(sleeptime)
                    listsin3 = appendlist(minitbookaddr_)
                    signal1, timecoumpf1, listsum1, timecoump1, overtimecp = restartpl(runs,
                                                                                       listsum1, 3,
                                                                                       timecoump1,
                                                                                       timecoumpf1,
                                                                                       signal1,
                                                                                       listsin3)
                    if listsin3[0] == 1 and listsin3[1] == 1 and listsin3[2] == 1:
                        plotcureco(phasecheck, 0, timetemple2, 3, minitbookaddr_, rownumber_)
                        break
                break

            elif factornumber < 5:
                overtimec, reducesi, timetemple2 = maincalfunction(factornumber, itimecount, listsin2, listsum,
                                                                   minitbookaddr_,
                                                                   overtimec, phasecheck, reducesi, rownumber_, runs,
                                                                   signal, timecoump, timecoumpf, times, timetemple2, 3)

            else:
                overtimec, reducesi, timetemple2 = maincalfunction(factornumber, itimecount, listsin2, listsum,
                                                                   minitbookaddr_,
                                                                   overtimec, phasecheck, reducesi, rownumber_, runs,
                                                                   signal, timecoump, timecoumpf, times, timetemple2, 4)

            if itimecount == times - 1:
                setzero(minitbookaddr_)

            itimecount = itimecount + 1

    loggenera('总耗时：' + str(int((time.time() - timecount1) / 60)) + ' 分钟' + '(' + str(
        round((time.time() - timecount1) / 3600, 1)) + ' 小时' + ')')
    if factornumber > 1:
        plothodata(plotorkbookaddr, sqlitetablename, returnrown, rownumber_ + 12, startrowh)
        if phasecheck == 3 or phasecheck == 4:
            plothodata(plotorkbookaddr, sqlitetablename, returnrown, rownumber_ + 12, startrow, 'Sheet1', quickdatanam)

    loggenera('数据库表名：' + sqlitetablename)
    loggenera('结束时间：' + str(datetime.datetime.now()))
    loggenera(finishcutline)
    pythoncomundo()
    turnscachigeffcheck = numberget(quickdatanam, turnscachigeffcheckstr)
    if factornumber >= 1 and turnscachigeffcheck == 0:
        killprocess()
    numberinput(quickdatanam, finishcheckstr, 1)


def maincalfunction(factornumber, itimecoun, listsin2, listsum, minitbookaddr_, overtimec, phasecheck, reducesi,
                    rownumber_,
                    runs, signal, timecoump, timecoumpf, times, timetemple2, motorct):
    calculatemotorcad1(rownumber_, phasecheck, minitbookaddr_, runs, reducesi)
    while 1:
        if overtimec > 0:
            overtimec = 0
            reducesi = 0.5
            killprocess()
            break
        else:
            reducesi = 1
        time.sleep(sleeptime)
        listsin = appendlist(minitbookaddr_)
        signal, timecoumpf, listsum, timecoump, overtimec = restart(minitbookaddr_, rownumber_, phasecheck,
                                                                    factornumber, listsum, motorct, timecoump,
                                                                    timecoumpf, signal, listsin, overtimec)

        if sum(listsin) == motorct:
            overtimec = 0
            calculatemotorcad2(phasecheck, minitbookaddr_)
            mainsqlite(rownumber_, minitbookaddr_, workbookaddr, runs, sqlitetablename)
            quicksign1 = getdata(minitbookaddr_, 16, 9)
            if quicksign1 is None:
                quicksign1 = 0

            if quicksign1 == 1 and (phasecheck == 3 or phasecheck == 4):
                if itimecoun == times - 1:
                    break
                setdata(minitbookaddr_, 16, 9, 0)
                quickcalculate(phasecheck, rownumber_, runs, factornumber, minitbookaddr_)
            elif quicksign1 == 1 and phasecheck == 1:
                if itimecoun == times - 1:
                    break
                setdata(minitbookaddr_, 16, 9, 0)
                break
            else:
                timecoump1 = time.time()
                timecoumpf1 = 0
                signal1 = 1
                listsum1 = 0
                calculatemotorcad21(phasecheck, listsin2, minitbookaddr_)
                timetemple2 = resultplothand(itimecoun, listsin2, listsum1, minitbookaddr_, phasecheck, rownumber_,
                                             runs,
                                             signal1, timecoump1, timecoumpf1, timetemple2)

            break
    return overtimec, reducesi, timetemple2


def resultplothand(i, listsin2, listsum1, minitbookaddr_, phasecheck, rownumber_, runs, signal1, timecoump1,
                   timecoumpf1, timetemple2):
    while 1:
        time.sleep(sleeptime)
        listsin3 = appendlist(minitbookaddr_)
        if listsin2 == 3 or listsin2 == 5 or listsin2 == 6 or listsin2 == 9:
            signal1, timecoumpf1, listsum1, timecoump1, overtimecp = restartpl(listsin2,
                                                                               listsum1, 3,
                                                                               timecoump1,
                                                                               timecoumpf1,
                                                                               signal1,
                                                                               listsin3)

            if listsin3[0] == 1 and listsin3[1] == 1 and listsin3[2] == 1:
                mainsqlite(rownumber_, minitbookaddr_, workbookaddr, runs, sqlitetablename)
                timetemple2 = plotcureco(phasecheck, i, timetemple2, listsin2, minitbookaddr_,
                                         rownumber_)
                break
        else:
            signal1, timecoumpf1, listsum1, timecoump1, overtimecp = restartpl(listsin2,
                                                                               listsum1, 4,
                                                                               timecoump1,
                                                                               timecoumpf1,
                                                                               signal1,
                                                                               listsin3)

            if sum(listsin3) == 4:
                mainsqlite(rownumber_, minitbookaddr_, workbookaddr, runs, sqlitetablename)
                timetemple2 = plotcureco(phasecheck, i, timetemple2, listsin2, minitbookaddr_,
                                         rownumber_)
                break
    return timetemple2


def testcod():
    while 1:
        testsignal = numberget(quickdatanam, finishcheckstr)
        time.sleep(3)
        if testsignal == 2:
            maincal()


def main():
    turnscaccheck = numberget(quickdatanam, turnscaccheckstr)
    turnscachigeffcheck = numberget(quickdatanam, turnscachigeffcheckstr)

    if turnscaccheck == 2 or turnscachigeffcheck == 4:
        while 1:
            time.sleep(10)
            turnscaccheck1 = numberget(quickdatanam, turnscaccheckstr)
            if turnscaccheck1 == 1:
                maincal()
                break
    else:
        maincal()

    # testcod()


if __name__ == '__main__':
    main()
