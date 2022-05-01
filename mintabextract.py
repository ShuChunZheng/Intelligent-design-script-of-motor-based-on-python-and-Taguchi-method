# -*- coding:utf-8 -*-

from minitabexecute import minitabop, minitabquit, dirgenerate, extract_data_from_excellist, \
    funcflatpara, funfinal, fuctionflat, fuctionflatc, rownumber, rownumbersingle, \
    runfatorn, listcut, loggenera, rownumbersthreerounbo

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

global minitabaddr
global workbookaddr
global minitbookaddr
global singminitbookaddr
global logautorunaddr
global minitbookroundboaddr


def paradifine():
    global minitabaddr
    global workbookaddr
    global minitbookaddr
    global singminitbookaddr
    global logautorunaddr
    global minitbookroundboaddr
    diclist = dirgenerate()
    minitbookaddr = diclist['minitbookaddr']
    workbookaddr = diclist['workbookaddr']
    minitabaddr = diclist['minitabaddr']
    singminitbookaddr = diclist['singminitbookaddr']
    logautorunaddr = diclist['logautorunaddr']
    minitbookroundboaddr = diclist['minitbookroundboaddr']


def namecd(factornumber, rownumber_, factnamelist=None):
    if factnamelist is None:
        factnamelist = []
    namecd1 = 'Name'
    combinelist = []
    factnamelistp = ['"' + i + '"' for i in factnamelist]
    clist = ['C' + str(i) for i in range(rownumber_ + 4, factornumber + rownumber_ + 4)]
    for i in range(factornumber):
        combinelist.append(clist[i])
        combinelist.append(factnamelistp[i])
    listcut(combinelist)
    combineliststr = ' '.join(combinelist)
    namecda = namecd1 + space + combineliststr + periodstr
    return namecda


def inputliststr(inputlist=None):
    if inputlist is None:
        inputlist = []
    inputliststr1 = [str(i) for i in inputlist]
    listcut(inputliststr1)
    inputliststrsp = ' '.join(inputliststr1)
    return inputliststrsp


def main(signal=0, phasecheck=3):
    paradifine()
    if phasecheck == 3:
        minitbookaddr_ = minitbookaddr
        rownumber_ = rownumber
    elif phasecheck == 4:
        minitbookaddr_ = minitbookroundboaddr
        rownumber_ = rownumbersthreerounbo
    else:
        minitbookaddr_ = singminitbookaddr
        rownumber_ = rownumbersingle

    minitab, mproject = minitabop(minitabaddr)
    namelist = ['效率', '齿槽转矩波动', '转矩', '转矩波动', '铜耗', '定子铁耗', '转子铁耗', '线端电压峰值', '母线平均电流', '转子轭磁密']

    get_num_lista = []

    columnnumber2 = 10
    rowoffset1 = 6
    coloffset1 = rownumber_ + 3
    efficiencylist = []
    cogginglist = []
    torquelist = []
    runs1, factornumber1, get_num_lista21 = runfatorn(minitbookaddr_, rownumber_)

    extract_data_from_excellist(workbookaddr, runs1, columnnumber2, rowoffset1, coloffset1, 'Sheet', get_num_lista)

    for i in range(0, runs1 * 10, 10):
        efficiencylist.append(get_num_lista[i])

    for i in range(1, runs1 * 10, 10):
        cogginglist.append(get_num_lista[i])

    for i in range(2, runs1 * 10, 10):
        torquelist.append(get_num_lista[i])

    currentlist = []
    extract_data_from_excellist(minitbookaddr_, 2, 1, 3, 9, 'Sheet', currentlist)
    loggenera('计算转矩：\n' + str(torquelist))
    loggenera('计算齿槽转矩：\n' + str(cogginglist))
    loggenera('计算效率：\n' + str(efficiencylist))
    namelist[0] = '效率（额定转矩,齿槽转矩）'

    paralistt = funcflatpara(currentlist[0])
    paralistc = []
    if currentlist[1] == 0 and currentlist[0] != 0:
        paralistc = funcflatpara(currentlist[0] * 0.3)

    elif currentlist[0] != 0 and currentlist[1] != 0:
        paralistc = funcflatpara(currentlist[1])

    for i in range(0, runs1 * 10, 10):
        j = int(i / 10)
        if j == len(torquelist):
            break
        get_num_lista[i] = funfinal(signal, efficiencylist[j], fuctionflat(torquelist[j], paralistt), fuctionflatc(cogginglist[j], paralistc))

    mproject.ExecuteCommand(namecd(10, rownumber_, namelist))

    for j in range(10):
        columnlist = []
        for i in range(j, runs1 * 10, 10):
            columnlist.append(get_num_lista[i])

        mproject.ExecuteCommand('SET C' + str(j + rownumber_ + 4) + '\n' +
                                inputliststr(columnlist) + '\n' +
                                'END.')

    mproject.SaveAs(minitabaddr)
    minitabquit(minitab)


if __name__ == '__main__':
    main()
