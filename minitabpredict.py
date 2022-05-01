# -*- coding:utf-8 -*-

import numpy as np

from minitabexecute import extract_data_from_excellist, minitabop, minitabquit, dirgenerate, pythoncomundo

space = ' '
equalestr = '='
ccolumn = 'C'
middle = '-'
semicolon = ';'
lowerisbetter = 'Lower'
higherisbetter = 'Higher'
gsncd = 'Gsn'
gmeanscd = 'Gmeans'
periodstr = '.'
plotmap = 'GSN;  Gmeans;  TSN;  Tmeans.'

global minitabaddr
global minitbookaddr
global mworkbookaddr
global workbookaddr


def paradifine():
    global minitabaddr
    global minitbookaddr
    global mworkbookaddr
    global workbookaddr
    diclist = dirgenerate()
    minitbookaddr = diclist['minitbookaddr']
    mworkbookaddr = diclist['mworkbookaddr']
    workbookaddr = diclist['workbookaddr']
    minitabaddr = diclist['minitabaddr']


def levelscda(listlevels=None):
    if listlevels is None:
        listlevels = []
    levelscd = 'Levels'
    listlevelsstr = ' '.join(listlevels)
    levelscda1 = levelscd + space + listlevelsstr
    return levelscda1


def predictcda(firstcc, finalcc, column1, column2):
    predictcd = 'OAPREDICT'
    clist = ['C' + str(i1) for i1 in range(firstcc, finalcc + firstcc)]
    cliststr = ' '.join(clist)
    predictcda1 = predictcd + space + ccolumn + str(column1) + space + ccolumn + str(
        column2) + space + equalestr + space + cliststr
    return predictcda1


def main():
    paradifine()

    minitab = minitabop(minitabaddr)
    mproject = minitab.ActiveProject

    rownumber = 24
    columnnumber1a = 4
    columnnumber1 = 3

    get_num_lista = []
    get_num_list = np.zeros((rownumber, columnnumber1))
    extract_data_from_excellist(minitbookaddr, rownumber, columnnumber1a, 2, 3, 'Sheet', get_num_lista)
    columnnamelist = []
    for i in range(0, rownumber):
        j = i * 4
        get_num_list[i][0] = get_num_lista[j + 1]
        get_num_list[i][1] = get_num_lista[j + 2]
        get_num_list[i][2] = get_num_lista[j + 3]
        columnnamelist.append(get_num_lista[j])

    elementnumber = np.array(np.array(np.array(get_num_list.sum(1)).nonzero())[0])
    factornumber = elementnumber.shape[0]

    get_num_predict_lista = []

    get_num_predict_list = []

    for i in range(factornumber):
        get_num_predict_list.append(get_num_predict_lista[elementnumber[i]])
    elementnumberp = np.array(np.array(np.array(get_num_predict_list).nonzero())[0])
    factornumberp = elementnumberp.shape[0]
    get_num_predict_listap = []
    for i in range(factornumberp):
        get_num_predict_listap.append(get_num_predict_list[elementnumberp[i]])
    get_num_predict_liststr = [str(i) for i in get_num_predict_listap]

    currentlist = []
    extract_data_from_excellist(minitbookaddr, 5, 1, 3, 9, 'Sheet', currentlist)

    if factornumber < 14:
        if currentlist[4] is not None and currentlist[3] is not None:
            mproject.ExecuteCommand(predictcda(1, factornumberp, 67, 68) + semicolon + space + space +
                                    'Labels "信噪比" "均值";' + space + space +
                                    levelscda(get_num_predict_liststr) + semicolon + space + space +
                                    'PFits C47 C48;' + '  Units 1 1.' + 'Erase C4000 C3999')
            mproject.ExecuteCommand(predictcda(1, factornumberp, 69, 70) + semicolon + space + space +
                                    'Labels "信噪比" "均值";' + space + space +
                                    levelscda(get_num_predict_liststr) + semicolon + space + space +
                                    'PFits C49 C50;' + '  Units 1 1.' + 'Erase C4000 C3999')
            mproject.ExecuteCommand('Erase C4000 C3999')
        else:
            mproject.ExecuteCommand(predictcda(1, factornumberp, 67, 68) + semicolon + space + space +
                                    'Labels "信噪比" "均值";' + space + space +
                                    levelscda(get_num_predict_liststr) + semicolon + space + space +
                                    'PFits C47 C48;' + '  Units 1 1.' + 'Erase C4000 C3999')
            mproject.ExecuteCommand('Erase C4000 C3999')

    elif factornumber < 26:
        if currentlist[4] is not None and currentlist[3] is not None:
            mproject.ExecuteCommand(predictcda(2, factornumberp, 67, 68) + semicolon + space + space +
                                    'Labels "信噪比" "均值";' + space + space +
                                    levelscda(get_num_predict_liststr) + semicolon + space + space +
                                    'PFits C47 C48;' + '  Units 1 1.')
            mproject.ExecuteCommand(predictcda(2, factornumberp, 69, 70) + semicolon + space + space +
                                    'Labels "信噪比" "均值";' + space + space +
                                    levelscda(get_num_predict_liststr) + semicolon + space + space +
                                    'PFits C49 C50;' + '  Units 1 1.')
            mproject.ExecuteCommand('Erase C4000 C3999' + 'Abort.')
        else:
            mproject.ExecuteCommand(predictcda(2, factornumberp, 67, 68) + semicolon + space + space +
                                    'Labels "信噪比" "均值";' + space + space +
                                    levelscda(get_num_predict_liststr) + semicolon + space + space +
                                    'PFits C47 C48;' + '  Units 1 1.')
            mproject.ExecuteCommand('Erase C4000 C3999' + 'Abort.')

    mproject.SaveAs(minitabaddr)
    minitabquit(minitab)
    pythoncomundo()


if __name__ == '__main__':
    main()
