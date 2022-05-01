import datetime
import os
import random
import time

import numpy

from autorun import main as automain
from minitabexecute import dirgenerate, write_data_to_excelhol, numberinput, fulldatanam, phasechecktabstr, \
    inputfilehan, namelist, numberget, quickdatanam, finishcheckstr, rownumber, rownumbersingle, killprocess, \
    filename, setdata, namelistsingle, extract_data_from_excelarray, copyfile, loggenera, logconfi, namelistthreer, \
    rownumbersthreerounbo

global minitbookaddr
global minitbookaddr_temp
global threepmodeladdr
global MotorCAD_Fileshot
global singminitbookaddr
global logautorunaddr
global setupdir
global singlepmodeladdr
global singminitbookaddr_temp
global minitbookroundboaddr
global minitbookroundboaddr_temp
global threepmodelroundboaddr

def paradifine():
    global minitbookaddr
    global minitbookaddr_temp
    global singminitbookaddr
    global logautorunaddr
    global setupdir
    global threepmodeladdr
    global singlepmodeladdr
    global MotorCAD_Fileshot
    global singminitbookaddr_temp
    global minitbookroundboaddr
    global minitbookroundboaddr_temp
    global threepmodelroundboaddr

    diclist = dirgenerate()
    minitbookaddr = diclist['minitbookaddr']
    minitbookaddr_temp = diclist['minitbookaddr_temp']
    MotorCAD_Fileshot = diclist['MotorCAD_Fileshot']
    singminitbookaddr = diclist['singminitbookaddr']
    logautorunaddr = diclist['logautorunaddr']
    setupdir = diclist['setupdir']
    threepmodeladdr = diclist['threepmodeladdr']
    singlepmodeladdr = diclist['singlepmodeladdr']
    singminitbookaddr_temp = diclist['singminitbookaddr_temp']
    minitbookroundboaddr = diclist['minitbookroundboaddr']
    minitbookroundboaddr_temp = diclist['minitbookroundboaddr_temp']
    threepmodelroundboaddr = diclist['threepmodelroundboaddr']


def listgenerate(phasecheckinpu):
    if phasecheckinpu == 3:
        minitbookaddr_ = minitbookaddr
        rownumber_ = rownumber
        threepmodeladdr_ = threepmodeladdr
        namelist_ = namelist
        minitbookaddr_temp_ = minitbookaddr_temp
    elif phasecheckinpu == 1:
        minitbookaddr_ = singminitbookaddr
        rownumber_ = rownumbersingle
        threepmodeladdr_ = singlepmodeladdr
        namelist_ = namelistsingle
        minitbookaddr_temp_ = singminitbookaddr_temp
    else:
        minitbookaddr_ = minitbookroundboaddr
        rownumber_ = rownumbersthreerounbo
        threepmodeladdr_ = threepmodelroundboaddr
        namelist_ = namelistthreer
        minitbookaddr_temp_ = minitbookroundboaddr_temp

    namestrinlist = filename(threepmodeladdr_)

    listindex = [i for i in range(rownumber_)]
    limitarray = numpy.zeros((rownumber_, 2))
    extract_data_from_excelarray(minitbookaddr_, rownumber_, 2, 2, 7, 'Sheet', limitarray)
    loggenera(datetime.datetime.now())

    subtestuni(limitarray, listindex, minitbookaddr_, minitbookaddr_temp_, namelist_, namestrinlist, phasecheckinpu,
               rownumber_, threepmodeladdr_, 1, 3)
    namestrinlisttem = [namestrinlist[random.randint(0, len(namestrinlist))]]
    subtestuni(limitarray, listindex, minitbookaddr_, minitbookaddr_temp_, namelist_, namestrinlisttem, phasecheckinpu,
               rownumber_, threepmodeladdr_, 3, rownumber_)


def subtestuni(limitarray, listindex, minitbookaddr_, minitbookaddr_temp_, namelist_, namestrinlist, phasecheckinpu,
               rownumber_, threepmodeladdr_, low, up):
    for e in namestrinlist:
        setdata(minitbookaddr_, 17, 10, e)
        for num in range(low, up):
            os.system('start alertwindowscomferm.bat')
            arrayfact = numpy.zeros((rownumber_, 1))
            listindexnum = random.sample(listindex, num)
            for i in range(num):
                arrayfact[listindexnum[i]] = 1
            if arrayfact[13] > 0 and arrayfact[23] > 0:
                if random.randint(0, 1):
                    arrayfact[13] = 0
                else:
                    arrayfact[23] = 0
            sheet = 'Sheet'

            if phasecheckinpu == 1 or phasecheckinpu == 4:
                limitarraytem = numpy.zeros((rownumber_, 2))

                for i in range(rownumber_):
                    if arrayfact[i] == 0:
                        limitarraytem[i][0] = limitarray[i][0]
                        limitarraytem[i][1] = limitarray[i][0]
                    else:
                        limitarraytem[i][0] = limitarray[i][0]
                        limitarraytem[i][1] = limitarray[i][1]
                write_data_to_excelhol(minitbookaddr_, rownumber_, 2, 2, 7, sheet, limitarraytem)

            write_data_to_excelhol(minitbookaddr_, rownumber_, 1, 2, 2, sheet, arrayfact)
            numberinput(fulldatanam, phasechecktabstr, phasecheckinpu)
            inputfilehan(minitbookaddr_, minitbookaddr_temp_, threepmodeladdr_, MotorCAD_Fileshot, namelist_)
            numberinput(quickdatanam, finishcheckstr, 2)
            automain()

            while 1:
                time.sleep(3)
                finishcheck = numberget(quickdatanam, finishcheckstr)
                if finishcheck == 1:
                    filenamestr = r'\%d_logautorun%d_%s.txt' % (phasecheckinpu, num, e)
                    loggenera((phasecheckinpu, num, e))
                    copyfile(logautorunaddr, setupdir + filenamestr)
                    loggenera(datetime.datetime.now())
                    break
            killprocess()


def main():
    paradifine()
    logconfi(logautorunaddr)
    listgenerate(3)
    listgenerate(1)


if __name__ == '__main__':
    main()
