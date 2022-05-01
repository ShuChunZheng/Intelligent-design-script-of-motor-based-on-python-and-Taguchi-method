# -*- coding:utf-8 -*-

from openpyxl import Workbook

from minitabexecute import extract_data_from_excellist, dirgenerate, getdata, rownumber, rownumbersingle, deltefile, \
    rownumbersthreerounbo
from motorcadinputdata import inputminijugsing, inputminijug, inputminijugthreeroundbo

space = ' '
equalestr = '='
ccolumn = 'C'
middle = '-'
semicolon = ';'
periodstr = '.'
lowerisbetter = 'Lower'
higherisbetter = 'Higher'
gsncd = 'Gsn'
gmeanscd = 'Gmeans'
plotmap = 'GSN;  Gmeans;  TSN;  Tmeans.'

global minitbookaddr
global mworkbookaddr
global workbookaddr
global singminitbookaddr
global minitbookroundboaddr


def paradifine():
    global minitbookaddr
    global mworkbookaddr
    global workbookaddr
    global singminitbookaddr
    global minitbookroundboaddr

    diclist = dirgenerate()
    minitbookaddr = diclist['minitbookaddr']
    mworkbookaddr = diclist['mworkbookaddr']
    workbookaddr = diclist['workbookaddr']
    singminitbookaddr = diclist['singminitbookaddr']
    minitbookroundboaddr = diclist['minitbookroundboaddr']


def saveasexcel(addr):
    wsavecda = 'WSave ' + '"' + addr + '";' + 'FType;    Excel 97;  Missing;    Numeric \'*\' \'*\';    Text "" "";  ' \
                                              'Replace. '
    return wsavecda


def main(phasecheck=3):
    paradifine()

    if phasecheck == 3:
        rownumber_ = rownumber
        minitbookaddr_ = minitbookaddr
    elif phasecheck == 4:
        rownumber_ = rownumbersthreerounbo
        minitbookaddr_ = minitbookroundboaddr
    else:
        rownumber_ = rownumbersingle
        minitbookaddr_ = singminitbookaddr

    deltefile(workbookaddr)
    wbmotorcad = Workbook()
    get_num_perform_lista = []
    extract_data_from_excellist(minitbookaddr_, rownumber_, 1, 2, 2, 'Sheet', get_num_perform_lista)
    if phasecheck == 3:
        get_num_lista = inputminijug(rownumber_, get_num_perform_lista)
    elif phasecheck == 1:
        get_num_lista = inputminijugsing(rownumber_, get_num_perform_lista)
    else:
        get_num_lista = inputminijugthreeroundbo(rownumber_, get_num_perform_lista)

    currentlist = []
    extract_data_from_excellist(minitbookaddr_, 1, 2, 13, 7, 'Sheet', currentlist)
    span = currentlist[1] - currentlist[0]
    calculatenum = getdata(minitbookaddr_, 2, 9)
    if calculatenum is None or calculatenum < 2:
        calculatenum = 4

    calculatspan = round(span / (calculatenum - 1), 3)

    for j in range(calculatenum):
        get_num_lista[11] = round(currentlist[0] + j * calculatspan, 1)
        for i in range(rownumber_):
            row1 = j + 6
            column1 = i + 2
            wbmotorcad['Sheet'].cell(row=row1, column=column1).value = get_num_lista[i]
    wbmotorcad.save(workbookaddr)


if __name__ == '__main__':
    main()
