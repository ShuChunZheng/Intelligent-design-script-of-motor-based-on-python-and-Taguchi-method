# import os
import numpy

# this script contains three functions to initialize, run and close the
# parallel Motor-CAD sessions
from minitabexecute import dirgenerate, rownumbermax, rownumber, write_data_to_excel_after, rownumbersingle, \
    extract_data_from_excel, extract_data_from_excel_after, rownumbersthreerounbo

global workbookaddr
global workbookaddr1
global workbookaddr2
global workbookaddr3
global workbookaddr4
global get_num_list
global return_num_list

columnnumber2 = 11


def paradifine():
    global workbookaddr
    global workbookaddr1
    global workbookaddr2
    global workbookaddr3
    global workbookaddr4
    diclist = dirgenerate()
    workbookaddr = diclist['workbookaddr']
    workbookaddr1 = diclist['workbookaddr1']
    workbookaddr2 = diclist['workbookaddr2']
    workbookaddr3 = diclist['workbookaddr3']
    workbookaddr4 = diclist['workbookaddr4']


def main(phasecheck=3):
    paradifine()
    if phasecheck == 3:
        rownumber_ = rownumber
    elif phasecheck == 1:
        rownumber_ = rownumbersingle
    else:
        rownumber_ = rownumbersthreerounbo
    rownumberholl = rownumber_ + 12

    global get_num_list
    global return_num_list
    get_num_list = numpy.zeros((rownumbermax, rownumberholl))
    return_num_list = numpy.zeros((rownumbermax, rownumberholl))

    joffset1 = 2
    joffset2 = rownumber_+3
    rownumberseperate = 14
    get_num_list0 = numpy.zeros((rownumbermax, rownumberholl))

    rownumberture1 = rownumberseperate
    rownumberture2 = rownumberseperate
    rownumberture3 = rownumberseperate
    rownumberture = extract_data_from_excel(workbookaddr, rownumbermax, rownumberholl, joffset1, get_num_list0)
    if rownumberture % 4 == 0:
        boolremain = 0
    else:
        boolremain = 1
    rownumberseperate = (rownumberture // 4) + boolremain
    for i in range(rownumbermax):
        if i < rownumberseperate:
            rownumberture1 = extract_data_from_excel(workbookaddr1, rownumbermax, rownumberholl, joffset2, get_num_list)
            extract_data_from_excel_after(workbookaddr1, rownumber_, i, i, rownumberseperate, columnnumber2, return_num_list)
        elif i < rownumberseperate * 2:
            rownumberture2 = extract_data_from_excel(workbookaddr2, rownumbermax, rownumberholl, joffset2, get_num_list)
            extract_data_from_excel_after(workbookaddr2, rownumber_, i - rownumberseperate, i, rownumberseperate, columnnumber2, return_num_list)
            if rownumberture1 < rownumberseperate:
                break
        elif i < rownumberseperate * 3:
            rownumberture3 = extract_data_from_excel(workbookaddr3, rownumbermax, rownumberholl, joffset2, get_num_list)
            extract_data_from_excel_after(workbookaddr3, rownumber_, i - rownumberseperate * 2, i, rownumberseperate, columnnumber2,
                                          return_num_list)
            if rownumberture2 < rownumberseperate:
                break
        elif i < rownumberseperate * 4:
            extract_data_from_excel(workbookaddr4, rownumbermax, rownumberholl, joffset2, get_num_list)
            extract_data_from_excel_after(workbookaddr4, rownumber_, i - rownumberseperate * 3, i, rownumberseperate, columnnumber2,
                                          return_num_list)
            if rownumberture3 < rownumberseperate:
                break

    for i in range(rownumbermax):
        if write_data_to_excel_after(rownumber_, workbookaddr, i, rownumbermax, columnnumber2, return_num_list) == 0:
            break


if __name__ == '__main__':
    main(1)
