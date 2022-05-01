# import os

import time
from math import atan, asin, cos, sin, sqrt, pi

import numpy
import win32com.client
from dxfwrite import DXFEngine
from numpy.linalg import det
from win32com.universal import com_error

from minitabexecute import dirgenerate, readdatainput, extract_data_from_excellist, getdata, \
    coefficientratiochange, extract_data_from_excelarray, mutation, rownumber, rownumbersingle, deltefile, loggenera, \
    logdebug, quickedit, numberget, fulldatanam, phasechecktabstr, quickcalcheckstr, motorinputreducesigetstr, \
    numberinput, motorinpufincheckstr, logconfidebug, setdata, rownumbersthreerounbo

global MotorCAD_Fileshot
global minitbookaddr
global inputparalist
global limitarray
global Base_File_path
global singminitbookaddr
global dexfileloca
global logmotorcadinputaddr
global minitbookroundboaddr

arcttran = 180 / pi
columnnumber1a = 1


def paradifine():
    global MotorCAD_Fileshot
    global minitbookaddr
    global Base_File_path
    global singminitbookaddr
    global dexfileloca
    global logmotorcadinputaddr
    global minitbookroundboaddr
    diclist = dirgenerate()
    MotorCAD_Fileshot = diclist['MotorCAD_Fileshot']
    minitbookaddr = diclist['minitbookaddr']
    singminitbookaddr = diclist['singminitbookaddr']
    dexfileloca = diclist['dexfileloca']
    logmotorcadinputaddr = diclist['logmotorcadinputaddr']
    minitbookroundboaddr = diclist['minitbookroundboaddr']

    Base_File_path = MotorCAD_Fileshot + '.mot'


###########################################################################
# 由圆上三点确定圆心和半径
###########################################################################
# INPUT
# p1   :  - 第一个点坐标, list或者array 1x3
# p2   :  - 第二个点坐标, list或者array 1x3
# p3   :  - 第三个点坐标, list或者array 1x3
# 若输入1x2的行向量, 末位自动补0, 变为1x3的行向量
###########################################################################
# OUTPUT
# pc   :  - 圆心坐标, array 1x3
# r    :  - 半径, 标量
###########################################################################
# 调用示例1 - 平面上三个点
# pc1, r1 = points2circle([1, 2], [-2, 1], [0, -3])
# 调用示例2 - 空间中三个点
# pc2, r2 = points2circle([1, 2, -1], [-2, 1, 2], [0, -3, -3])
###########################################################################
def points2circle(p1, p2, p3):
    p1 = numpy.array(p1)
    p2 = numpy.array(p2)
    p3 = numpy.array(p3)
    num1 = len(p1)
    num2 = len(p2)
    num3 = len(p3)

    # 输入检查
    if (num1 == num2) and (num2 == num3):
        if num1 == 2:
            p1 = numpy.append(p1, 0)
            p2 = numpy.append(p2, 0)
            p3 = numpy.append(p3, 0)
        elif num1 != 3:
            loggenera('\t仅支持二维或三维坐标输入')
            return None
    else:
        loggenera('\t输入坐标的维数不一致')
        return None

    # 共线检查
    temp01 = p1 - p2
    temp02 = p3 - p2
    temp03 = numpy.cross(temp01, temp02)
    temp = (temp03 @ temp03) / (temp01 @ temp01) / (temp02 @ temp02)
    if temp < 10 ** -6:
        loggenera('\t三点共线, 无法确定圆')
        return None

    temp1 = numpy.vstack((p1, p2, p3))
    temp2 = numpy.ones(3).reshape(3, 1)
    mat1 = numpy.hstack((temp1, temp2))  # size = 3x4

    m = +det(mat1[:, 1:])
    n = -det(numpy.delete(mat1, 1, axis=1))
    p = +det(numpy.delete(mat1, 2, axis=1))
    q = -det(temp1)

    temp3 = numpy.array([p1 @ p1, p2 @ p2, p3 @ p3]).reshape(3, 1)
    temp4 = numpy.hstack((temp3, mat1))
    temp5 = numpy.array([2 * q, -m, -n, -p, 0])
    mat2 = numpy.vstack((temp4, temp5))  # size = 4x5

    A = +det(mat2[:, 1:])
    B = -det(numpy.delete(mat2, 1, axis=1))
    C = +det(numpy.delete(mat2, 2, axis=1))
    D = -det(numpy.delete(mat2, 3, axis=1))
    E = +det(mat2[:, :-1])

    pc = -numpy.array([B, C, D]) / 2 / A
    r = numpy.sqrt(B * B + C * C + D * D - 4 * A * E) / 2 / abs(A)

    return pc, r


def tortcoord(x, y):
    r = sqrt(x * x + y * y)
    t = 0
    if x > 0 and y >= 0:
        t = atan(y / x)
    elif x < 0:
        t = atan(y / x) + pi
    elif x > 0 > y:
        t = atan(y / x) + pi * 2
    elif x == 0:
        if y > 0:
            t = pi / 2
        elif y < 0:
            t = pi * 3 / 2
        else:
            t = 0
    return r, t


def toxycoord(r, t):
    x = r * cos(t)
    y = r * sin(t)
    return x, y


def fillradtopbott(x1, y1, y, b):
    tem = b ** 2 - (y - y1) ** 2
    if tem < 0:
        tem = 0
    return x1 + tem ** 0.5


def lineabsolv(x1, y1, x2, y2):
    x = (y1 - y2) / (x1 - x2)
    y = (x1 * y2 - x2 * y1) / (x1 - x2)
    return x, y


def toothtipfilcecb(x1, y1, r, d, br):
    rrl = br - r
    trl = asin((d - y1 + r) / rrl)
    xrl, yrl = toxycoord(rrl, trl)
    x, y = xrl + x1, yrl + y1
    ro, to = tortcoord(x, y)

    xa, ya = x, y - r
    ra, ta = tortcoord(xa, ya)

    rrlb = br
    trlb = trl
    xrlb, yrlb = toxycoord(rrlb, trlb)
    xb, yb = xrlb + x1, yrlb + y1
    rb, tb = tortcoord(xb, yb)

    return ro, to, ra, ta, rb, tb


def toothtipfilceca(x1, y1, r1, t1, r, d, br, c):
    rrl = r + br
    drl = r + d + r1 * sin(pi - t1 + c)
    trl = c - asin(drl / rrl)
    xrl, yrl = toxycoord(rrl, trl)
    x, y = xrl + x1, yrl + y1
    ro, to = tortcoord(x, y)

    rrla = br
    trla = trl
    xrla, yrla = toxycoord(rrla, trla)
    xa, ya = xrla + x1, yrla + y1
    ra, ta = tortcoord(xa, ya)

    tem = rrl ** 2 - drl ** 2
    if tem <= 0:
        tem = 0
    tem1 = drl - r
    rrlb = (tem + tem1 ** 2) ** 0.5
    if tem != 0:
        trlb = c - atan(tem1 / tem ** 0.5)
    else:
        if tem1 > 0:
            t = pi / 2
        else:
            t = 0
        trlb = c - t

    xrlb, yrlb = toxycoord(rrlb, trlb)
    xb, yb = xrlb + x1, yrlb + y1
    rb, tb = tortcoord(xb, yb)

    return ro, to, ra, ta, rb, tb


def outercircpoincal(e, m, r, d, c):
    tem = r * cos(asin(d / r))
    xrll = -d
    yrll = tem
    rrll, trll = tortcoord(xrll, yrll)
    rl = rrll
    tl = trll - (pi / 2 - c)

    xrlh = d
    yrlh = tem - e
    rrlh, trlh = tortcoord(xrlh, yrlh)
    rh = rrlh
    th = trlh - (pi / 2 - c)

    rmb, tmb = r - m * e, 0
    rmu, tmu = r - m * e, c * 2

    rlr = rl
    tlr = tl - c * 2
    x1, y1 = toxycoord(rh, th)
    x2, y2 = toxycoord(rmb, tmb)
    x3, y3 = toxycoord(rlr, tlr)
    pout, rrout = points2circle([x1, y1], [x2, y2], [x3, y3])
    rout, tout = tortcoord(pout[0], pout[1])
    routm, toutm = rout, tout + 2 * c

    return rl, toarc(tl), rh, toarc(th), rmb, toarc(tmb), rmu, toarc(tmu), rout, toarc(tout), routm, toarc(toutm)


def outercircpoincalthree(r, d, c):
    tem = r * cos(asin(d / r))
    xrll = -d
    yrll = tem
    rrll, trll = tortcoord(xrll, yrll)
    rl = rrll
    tl = trll - (pi / 2 - c)

    xrlh = d
    yrlh = tem
    rrlh, trlh = tortcoord(xrlh, yrlh)
    rh = rrlh
    th = trlh - (pi / 2 - c)

    rmb, tmb = r, 0
    rmu, tmu = r, c * 2

    return rl, toarc(tl), rh, toarc(th), rmb, toarc(tmb), rmu, toarc(tmu), 0, 0, 0, 0


def toarc(x):
    return x * arcttran


def funcd(r0, x1, y1, r1, y2, x3, y3, r3, x4, y4, r4, c1, c2, c3, c4, c5, x):
    yout = 0
    if c1 <= x <= c2:
        tem = r0 ** 2 - x ** 2
        if tem < 0:
            tem = 0
        yout = tem ** 0.5
    elif c2 < x <= c3:
        tem1 = r1 ** 2 - (x - x1) ** 2
        if tem1 < 0:
            tem1 = 0
        yout = y1 - tem1 ** 0.5
    elif c3 < x <= c4:
        yout = y2
    elif c4 < x <= c5:
        tem2 = r3 ** 2 - (x - x3) ** 2
        if tem2 < 0:
            tem2 = 0
        yout = y3 - tem2 ** 0.5
    elif c5 < x:
        tem3 = r4 ** 2 - (x - x4) ** 2
        if tem3 < 0:
            tem3 = 0
        yout = y4 - tem3 ** 0.5

    return yout


def funcu(k1, b, x1, y1, r1, x2, y2, r2, c1, c2, c3, c4, x):
    yout = 0
    if c1 <= x <= c2:
        yout = k1 * x
    elif c2 < x <= c3:
        yout = x / k1 + b
    elif c4 > c3 and c3 < x <= c4:
        tem1 = r1 ** 2 - (x - x1) ** 2
        if tem1 < 0:
            tem1 = 0
        yout = y1 + tem1 ** 0.5
    elif c3 < c4 < x:
        tem2 = r2 ** 2 - (x - x2) ** 2
        if tem2 < 0:
            tem2 = 0
        yout = y2 + tem2 ** 0.5
    elif c4 <= c3 < x:
        tem3 = r1 ** 2 - (x - x1) ** 2
        if tem3 < 0:
            tem3 = 0
        yout = y1 + tem3 ** 0.5

    return yout


def slotareacal(x1, y1, x2, y2, x3, y3, x4, y4, x5, y5, x6, y6, x7, y7, xo1, yo1, xo2, yo2, xo3, yo3, r1, r2, r3, r4):
    lol = x1
    if y4 >= yo2:
        upl = xo2 + r2
    else:
        if xo1 + r1 >= x3 and yo1 <= y3:
            upl = xo1 + r1
        else:
            upl = x3
    sepa = int((x4 - x7) / 0.01)
    logdebug('divid: ' + str(sepa))
    dex = (upl - lol) / sepa
    sumarea = 0
    x = lol
    k1 = (y2 - y1) / (x2 - x1)
    b = y3 - x3 / k1
    for i in range(sepa):
        yupout = funcu(k1, b, xo1, yo1, r1, xo2, yo2, r2, x1, x2, x3, x4, x)
        ydowout = funcd(r4, xo3, yo3, r3, y6, xo2, yo2, r2, xo1, yo1, r1, x1, x7, x6, x5, x4, x)
        yupout2 = funcu(k1, b, xo1, yo1, r1, xo2, yo2, r2, x1, x2, x3, x4, x + dex)
        ydowout2 = funcd(r4, xo3, yo3, r3, y6, xo2, yo2, r2, xo1, yo1, r1, x1, x7, x6, x5, x4, x + dex)
        sumarea = sumarea + (yupout - ydowout + yupout2 - ydowout2) * dex / 2
        x = x + dex

    lenl = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
    sl1 = ((x4 - x3) ** 2 + (y4 - y3) ** 2) ** 0.5
    sl2 = ((x5 - x4) ** 2 + (y5 - y4) ** 2) ** 0.5
    sl3 = ((x7 - x6) ** 2 + (y7 - y6) ** 2) ** 0.5
    sl4 = ((x1 - x7) ** 2 + (y1 - y7) ** 2) ** 0.5
    lenlh = x5 - x6 + r1 * 2 * asin(sl1 / (2 * r1)) + r2 * 2 * asin(sl2 / (2 * r2)) + r3 * 2 * asin(
        sl3 / (2 * r3)) + r4 * 2 * asin(sl4 / (2 * r4))
    windhig = y3 - y5
    tem1 = x3 - x4
    if tem1 == 0:
        equatoothtoplina = 90
    else:
        equatoothtoplina = toarc(atan((y3 - y4) / tem1))

    return sumarea, lenl, lenlh, windhig, equatoothtoplina


def drawstator(outerdia, innerdiameter, polenumber, toothwhi, slotdepth, toothtipd, toothtipdm, toothtipdb, toothtipop,
               slotbora, slottora, slottipra, outermag, outermidg, airgap, magnetthi, magnetarc, backironthi,
               mcad=None):
    if mcad is None:
        mcad = win32com.client.Dispatch('motorcad.appautomation')
    pole = polenumber * 2
    slotnumber = pole
    Single_SlotPole = 0

    mcad.AvoidImmediateUpdate(True)
    mcad.DisplayScreen('FE')
    mcad.ClearAllData()
    mcad.SetVariable("DXFImportType", Single_SlotPole)
    mcad.SetVariable("DXFAutoCentre", False)
    mcad.SetVariable("RotorRotation", 45)
    mcad.InitiateGeometryFromScript()

    innerradius = innerdiameter / 2
    hoangle = 360 / slotnumber
    hoanglehalf = 180 / slotnumber
    hoanglehalfa = hoanglehalf / arcttran
    outerradiu = outerdia / 2
    slotradius = outerradiu - slotdepth
    halftoothwhi = toothwhi / 2

    mcad.AddArc_CentreStartEnd_RT(0, 0, innerradius, 0, innerradius, hoangle)

    rslotbora = slotradius + slotbora
    tslotbora = asin((slotbora + halftoothwhi) / rslotbora)
    temtlslotborab = rslotbora * cos(tslotbora)
    if temtlslotborab == 0:
        tlslotbora = pi / 2
    else:
        tlslotbora = atan(halftoothwhi / temtlslotborab)

    rlslotbora = halftoothwhi / sin(tlslotbora)
    tl = tlslotbora * arcttran
    ta = tslotbora * arcttran
    tlmirr = hoangle - tl
    tamirr = hoangle - ta

    mcad.AddArc_CentreStartEnd_RT(rslotbora, ta, slotradius, ta, rlslotbora, tl)
    mcad.AddArc_CentreStartEnd_RT(0, 0, slotradius, ta, slotradius, tamirr)
    mcad.AddArc_CentreStartEnd_RT(rslotbora, tamirr, rlslotbora, tlmirr, slotradius, tamirr)

    rtoothtipb, ttoothtipb = tortcoord(
        sqrt(outerradiu * outerradiu - halftoothwhi * halftoothwhi) - toothtipdb * toothtipd,
        halftoothwhi)
    ttoothtip = hoanglehalfa - atan(toothtipop / (2 * (outerradiu - toothtipd)))
    rtoothtip = sqrt(toothtipop * toothtipop / 4 + (outerradiu - toothtipd) * (outerradiu - toothtipd))
    ttoothtipdm = (ttoothtipb + ttoothtip) / 2
    rtoothtipdm = outerradiu - toothtipdm * toothtipd
    xt, yt = toxycoord(rtoothtipb, ttoothtipb)
    xm, ym = toxycoord(rtoothtipdm, ttoothtipdm)
    xb, yb = toxycoord(rtoothtip, ttoothtip)
    pctooths, rtooths = points2circle([xt, yt], [xm, ym], [xb, yb])
    rctooths, tctooths = tortcoord(pctooths[0], pctooths[1])

    rctooths, tctoothsa = rctooths, toarc(tctooths)
    rctoothsm, tctoothsam = rctooths, hoangle - toarc(tctooths)

    rslottorac, tslottorac, rslottoraa, tslottoraa, rslottorab, tslottorab = toothtipfilcecb(pctooths[0], pctooths[1],
                                                                                             slottora, halftoothwhi,
                                                                                             rtooths)
    rslottorac, tslottoraca, rslottoraa, tslottoraaa, rslottorab, tslottoraba = rslottorac, toarc(
        tslottorac), rslottoraa, toarc(tslottoraa), rslottorab, toarc(tslottorab)
    rslottoracm, tslottoracm, rslottoraam, tslottoraam, rslottorabm, tslottorabm = rslottorac, hoangle - tslottoraca, rslottoraa, hoangle - tslottoraaa, rslottorab, hoangle - tslottoraba

    rslottiprac, tslottiprac, rslottipraa, tslottipraa, rslottiprab, tslottiprab = toothtipfilceca(pctooths[0],
                                                                                                   pctooths[1],
                                                                                                   rctooths, tctooths,
                                                                                                   slottipra,
                                                                                                   toothtipop / 2,
                                                                                                   rtooths,
                                                                                                   hoanglehalfa)
    rslottiprac, tslottipraca, rslottipraa, tslottipraaa, rslottiprab, tslottipraba = rslottiprac, toarc(
        tslottiprac), rslottipraa, toarc(tslottipraa), rslottiprab, toarc(tslottiprab)
    rslottipracm, tslottipracm, rslottipraam, tslottipraam, rslottiprabm, tslottiprabm = rslottiprac, hoangle - tslottipraca, rslottipraa, hoangle - tslottipraaa, rslottiprab, hoangle - tslottipraba

    rlouterci, tlouterci, rhouterci, thouterci, rmbouterci, tmbouterci, rmuouterci, tmuouterci, rout, tout, routm, toutm = outercircpoincal(
        outermag, outermidg, outerradiu, toothtipop / 2, hoanglehalfa)

    mcad.AddLine_RT(rlslotbora, tl, rslottoraa, tslottoraaa)
    mcad.AddLine_RT(rlslotbora, tlmirr, rslottoraam, tslottoraam)

    mcad.AddArc_CentreStartEnd_RT(rslottorac, tslottoraca, rslottoraa, tslottoraaa, rslottorab, tslottoraba)
    mcad.AddArc_CentreStartEnd_RT(rslottoracm, tslottoracm, rslottorabm, tslottorabm, rslottoraam, tslottoraam)

    mcad.AddArc_CentreStartEnd_RT(rslottiprac, tslottipraca, rslottiprab, tslottipraba, rslottipraa, tslottipraaa)
    mcad.AddArc_CentreStartEnd_RT(rslottipracm, tslottipracm, rslottipraam, tslottipraam, rslottiprabm, tslottiprabm)

    mcad.AddArc_CentreStartEnd_RT(rctooths, tctoothsa, rslottorab, tslottoraba, rslottipraa, tslottipraaa)
    mcad.AddArc_CentreStartEnd_RT(rctoothsm, tctoothsam, rslottipraam, tslottipraam, rslottorabm, tslottorabm)

    mcad.AddLine_RT(rslottiprab, tslottipraba, rhouterci, thouterci)
    mcad.AddLine_RT(rslottiprabm, tslottiprabm, rlouterci, tlouterci)

    mcad.AddArc_CentreStartEnd_RT(rout, tout, rmbouterci, tmbouterci, rhouterci, thouterci)
    mcad.AddArc_CentreStartEnd_RT(routm, toutm, rlouterci, tlouterci, rmuouterci, tmuouterci)

    mcad.AddLine_RT(rlouterci, tlouterci, rhouterci, thouterci)
    tem = rslottipraa * cos(hoanglehalfa - tslottipraa)
    mcad.AddLine_RT(rslottipraa, tslottipraaa, rslottipraam, tslottipraam)
    mcad.AddLine_RT(slotradius, hoanglehalf, tem, hoanglehalf)
    mcad.AddLine_RT(0, 0, rmbouterci, 0)
    mcad.AddLine_RT(0, 0, rmbouterci, hoangle)

    magnetinnerr = outerradiu + airgap
    magnetouterr = magnetinnerr + magnetthi
    backouterr = magnetouterr + backironthi
    tbackiron = 360 / pole
    tmanet = magnetarc * 2 / pole
    tmagnetst = (tbackiron - tmanet) / 2
    tmagnetend = tmagnetst + tmanet

    mcad.AddArc_CentreStartEnd_RT(0, 0, magnetouterr, 0, magnetouterr, tbackiron)
    mcad.AddArc_CentreStartEnd_RT(0, 0, backouterr, 0, backouterr, tbackiron)
    mcad.AddArc_CentreStartEnd_RT(0, 0, magnetinnerr, 0, magnetinnerr, tbackiron)

    mcad.AddLine_RT(magnetinnerr, 0, backouterr, 0)
    mcad.AddLine_RT(magnetinnerr, tbackiron, backouterr, tbackiron)

    mcad.AddLine_RT(magnetinnerr, tmagnetst, magnetouterr, tmagnetst)
    mcad.AddLine_RT(magnetinnerr, tmagnetend, magnetouterr, tmagnetend)

    sumarea, lenl, lenlh, windhig, equatoothtoplina = slotareacal(x1=toxycoord(slotradius, hoanglehalfa)[0],
                                                                  y1=toxycoord(slotradius, hoanglehalfa)[1],
                                                                  x2=toxycoord(tem, hoanglehalfa)[0],
                                                                  y2=toxycoord(tem, hoanglehalfa)[1],
                                                                  x3=toxycoord(rslottipraa, tslottipraa)[0],
                                                                  y3=toxycoord(rslottipraa, tslottipraa)[1],
                                                                  x4=toxycoord(rslottorab, tslottorab)[0],
                                                                  y4=toxycoord(rslottorab, tslottorab)[1],
                                                                  x5=toxycoord(rslottoraa, tslottoraa)[0],
                                                                  y5=toxycoord(rslottoraa, tslottoraa)[1],
                                                                  x6=toxycoord(rlslotbora, tlslotbora)[0],
                                                                  y6=toxycoord(rlslotbora, tlslotbora)[1],
                                                                  x7=toxycoord(slotradius, tslotbora)[0],
                                                                  y7=toxycoord(slotradius, tslotbora)[1],
                                                                  xo1=pctooths[0], yo1=pctooths[1],
                                                                  xo2=toxycoord(rslottorac, tslottorac)[0],
                                                                  yo2=toxycoord(rslottorac, tslottorac)[1],
                                                                  xo3=toxycoord(rslotbora, tslotbora)[0],
                                                                  yo3=toxycoord(rslotbora, tslotbora)[1], r1=rtooths,
                                                                  r2=slottora, r3=slotbora, r4=slotradius)

    if equatoothtoplina <= 0:
        equatoothtipang = hoanglehalf - 90 - equatoothtoplina
    else:
        equatoothtipang = 90 - equatoothtoplina + hoanglehalf
    return equatoothtipang, sumarea, lenl, lenlh, windhig
    # mcad.AddPoint_Magnetic_RT((magnetinnerr+magnetouterr)/2, hoanglehalf, "1Magnet1", hoanglehalf, 1, MagPole_North)
    # mcad.AddPoint_RT((rslottiprac+slotradius)/2, (hoanglehalf+tslottipraca)/2, "ArmatureSlotL1")
    # mcad.AddPoint_RT((rslottiprac+slotradius)/2, (hoanglehalf+tslottipracm)/2, "ArmatureSlotR1")
    # mcad.AddPoint_RT(innerdiameter/4, hoanglehalf, "Axle")
    # mcad.AddPoint_RT((magnetouterr+backouterr)/2, hoanglehalf, "Rotor")
    # mcad.AddPoint_RT((magnetinnerr+magnetouterr)/2, (tmagnetend+tbackiron)/2, "RotorAir")
    # mcad.AddPoint_RT((magnetinnerr+magnetouterr)/2, tmagnetst/2, "RotorAir")
    # mcad.AddPoint_RT((innerdiameter+slotradius)/2, hoanglehalf, "Stator")
    # mcad.AddPoint_RT((tem+rlouterci)/2, hoanglehalf, "StatorAir")


def xyarcdraw(x4, y4, x3, y3, xo1, yo1, dxfdraw=None):
    if dxfdraw is None:
        dxfdraw = DXFEngine.drawing()
    x3r1, y3r1 = x3 - xo1, y3 - yo1
    x4r1, y4r1 = x4 - xo1, y4 - yo1
    r3r1, t3r1 = tortcoord(x3r1, y3r1)
    r4r1, t4r1 = tortcoord(x4r1, y4r1)
    t3r1 = toarc(t3r1)
    t4r1 = toarc(t4r1)
    arc1 = DXFEngine.arc(r3r1, (xo1, yo1), t4r1, t3r1)
    dxfdraw.add(arc1)
    dxfdraw.save()


def toarcang(x):
    return x / arcttran


def rtarcdraw(ro1, to1, r4, t4, r3, t3, dxfdraw=None):
    if dxfdraw is None:
        dxfdraw = DXFEngine.drawing()
    to1 = toarcang(to1)
    t4 = toarcang(t4)
    t3 = toarcang(t3)
    x3, y3 = toxycoord(r3, t3)
    x4, y4 = toxycoord(r4, t4)
    xo1, yo1 = toxycoord(ro1, to1)
    x3r1, y3r1 = x3 - xo1, y3 - yo1
    x4r1, y4r1 = x4 - xo1, y4 - yo1
    r3r1, t3r1 = tortcoord(x3r1, y3r1)
    r4r1, t4r1 = tortcoord(x4r1, y4r1)
    t3r1 = toarc(t3r1)
    t4r1 = toarc(t4r1)
    arc1 = DXFEngine.arc(r3r1, (xo1, yo1), t4r1, t3r1)
    dxfdraw.add(arc1)
    dxfdraw.save()


def rtlinedraw(r1, t1, r2, t2, dxfdraw=None):
    if dxfdraw is None:
        dxfdraw = DXFEngine.drawing()
    t1 = toarcang(t1)
    t2 = toarcang(t2)
    x1, y1 = toxycoord(r1, t1)
    x2, y2 = toxycoord(r2, t2)
    line = DXFEngine.line((x1, y1), (x2, y2))
    dxfdraw.add(line)
    dxfdraw.save()


def drawstatot(outerdia, innerdiameter, polenumber, toothwhi, slotdepth, toothtipd, toothtipdm, toothtipdb, toothtipop,
               slotbora, slottora, slottipra, outermag, outermidg, airgap, magnetthi, magnetarc, backironthi,
               dexfileloca_):
    dxfdraw = DXFEngine.drawing(dexfileloca_)
    pole = polenumber * 2
    slotnumber = pole
    innerradius = innerdiameter / 2
    hoangle = 360 / slotnumber
    hoanglehalf = 180 / slotnumber
    hoanglehalfa = hoanglehalf / arcttran
    outerradiu = outerdia / 2
    slotradius = outerradiu - slotdepth
    halftoothwhi = toothwhi / 2

    rslotbora = slotradius + slotbora
    tslotbora = asin((slotbora + halftoothwhi) / rslotbora)
    temtlslotborab = rslotbora * cos(tslotbora)
    if temtlslotborab == 0:
        tlslotbora = pi / 2
    else:
        tlslotbora = atan(halftoothwhi / temtlslotborab)

    rlslotbora = halftoothwhi / sin(tlslotbora)
    tl = tlslotbora * arcttran
    ta = tslotbora * arcttran
    tlmirr = hoangle - tl
    tamirr = hoangle - ta

    rtoothtipb, ttoothtipb = tortcoord(
        sqrt(outerradiu * outerradiu - halftoothwhi * halftoothwhi) - toothtipdb * toothtipd,
        halftoothwhi)
    ttoothtip = hoanglehalfa - atan(toothtipop / (2 * (outerradiu - toothtipd)))
    rtoothtip = sqrt(toothtipop * toothtipop / 4 + (outerradiu - toothtipd) * (outerradiu - toothtipd))
    ttoothtipdm = (ttoothtipb + ttoothtip) / 2
    rtoothtipdm = outerradiu - toothtipdm * toothtipd
    xt, yt = toxycoord(rtoothtipb, ttoothtipb)
    xm, ym = toxycoord(rtoothtipdm, ttoothtipdm)
    xb, yb = toxycoord(rtoothtip, ttoothtip)
    pctooths, rtooths = points2circle([xt, yt], [xm, ym], [xb, yb])
    rctooths, tctooths = tortcoord(pctooths[0], pctooths[1])

    rctooths, tctoothsa = rctooths, toarc(tctooths)
    rctoothsm, tctoothsam = rctooths, hoangle - toarc(tctooths)

    rslottorac, tslottorac, rslottoraa, tslottoraa, rslottorab, tslottorab = toothtipfilcecb(pctooths[0], pctooths[1],
                                                                                             slottora, halftoothwhi,
                                                                                             rtooths)
    rslottorac, tslottoraca, rslottoraa, tslottoraaa, rslottorab, tslottoraba = rslottorac, toarc(
        tslottorac), rslottoraa, toarc(tslottoraa), rslottorab, toarc(tslottorab)
    rslottoracm, tslottoracm, rslottoraam, tslottoraam, rslottorabm, tslottorabm = rslottorac, hoangle - tslottoraca, rslottoraa, hoangle - tslottoraaa, rslottorab, hoangle - tslottoraba

    rslottiprac, tslottiprac, rslottipraa, tslottipraa, rslottiprab, tslottiprab = toothtipfilceca(pctooths[0],
                                                                                                   pctooths[1],
                                                                                                   rctooths, tctooths,
                                                                                                   slottipra,
                                                                                                   toothtipop / 2,
                                                                                                   rtooths,
                                                                                                   hoanglehalfa)
    rslottiprac, tslottipraca, rslottipraa, tslottipraaa, rslottiprab, tslottipraba = rslottiprac, toarc(
        tslottiprac), rslottipraa, toarc(tslottipraa), rslottiprab, toarc(tslottiprab)
    rslottipracm, tslottipracm, rslottipraam, tslottipraam, rslottiprabm, tslottiprabm = rslottiprac, hoangle - tslottipraca, rslottipraa, hoangle - tslottipraaa, rslottiprab, hoangle - tslottipraba

    rlouterci, tlouterci, rhouterci, thouterci, rmbouterci, tmbouterci, rmuouterci, tmuouterci, rout, tout, routm, toutm = outercircpoincal(
        outermag, outermidg, outerradiu, toothtipop / 2, hoanglehalfa)

    tem = rslottipraa * cos(hoanglehalfa - tslottipraa)

    magnetinnerr = outerradiu + airgap
    magnetouterr = magnetinnerr + magnetthi
    backouterr = magnetouterr + backironthi
    tbackiron = 360 / pole
    tmanet = magnetarc * 2 / pole
    tmagnetst = (tbackiron - tmanet) / 2
    tmagnetend = tmagnetst + tmanet

    rtarcdraw(0, 0, innerradius, 0, innerradius, hoangle, dxfdraw)

    rtarcdraw(rslotbora, ta, slotradius, ta, rlslotbora, tl, dxfdraw)

    rtarcdraw(0, 0, slotradius, ta, slotradius, tamirr, dxfdraw)

    rtarcdraw(rslotbora, tamirr, rlslotbora, tlmirr, slotradius, tamirr, dxfdraw)

    rtlinedraw(rlslotbora, tl, rslottoraa, tslottoraaa, dxfdraw)

    rtlinedraw(rlslotbora, tlmirr, rslottoraam, tslottoraam, dxfdraw)

    rtarcdraw(rslottorac, tslottoraca, rslottoraa, tslottoraaa, rslottorab, tslottoraba, dxfdraw)

    rtarcdraw(rslottoracm, tslottoracm, rslottorabm, tslottorabm, rslottoraam, tslottoraam, dxfdraw)

    rtarcdraw(rslottiprac, tslottipraca, rslottiprab, tslottipraba, rslottipraa, tslottipraaa, dxfdraw)

    rtarcdraw(rslottipracm, tslottipracm, rslottipraam, tslottipraam, rslottiprabm, tslottiprabm, dxfdraw)

    rtarcdraw(rctooths, tctoothsa, rslottorab, tslottoraba, rslottipraa, tslottipraaa, dxfdraw)

    rtarcdraw(rctoothsm, tctoothsam, rslottipraam, tslottipraam, rslottorabm, tslottorabm, dxfdraw)

    rtlinedraw(rslottiprab, tslottipraba, rhouterci, thouterci, dxfdraw)

    rtlinedraw(rslottiprabm, tslottiprabm, rlouterci, tlouterci, dxfdraw)

    rtarcdraw(rout, tout, rmbouterci, tmbouterci, rhouterci, thouterci, dxfdraw)

    rtarcdraw(routm, toutm, rlouterci, tlouterci, rmuouterci, tmuouterci, dxfdraw)

    rtlinedraw(rlouterci, tlouterci, rhouterci, thouterci, dxfdraw)

    rtlinedraw(rslottipraa, tslottipraaa, rslottipraam, tslottipraam, dxfdraw)

    rtlinedraw(slotradius, hoanglehalf, tem, hoanglehalf, dxfdraw)

    rtlinedraw(0, 0, rmbouterci, 0, dxfdraw)

    rtlinedraw(0, 0, rmbouterci, hoangle, dxfdraw)

    rtarcdraw(0, 0, magnetouterr, 0, magnetouterr, tbackiron, dxfdraw)

    rtarcdraw(0, 0, backouterr, 0, backouterr, tbackiron, dxfdraw)

    rtarcdraw(0, 0, magnetinnerr, 0, magnetinnerr, tbackiron, dxfdraw)

    rtlinedraw(magnetinnerr, 0, backouterr, 0, dxfdraw)

    rtlinedraw(magnetinnerr, tbackiron, backouterr, tbackiron, dxfdraw)

    rtlinedraw(magnetinnerr, tmagnetst, magnetouterr, tmagnetst, dxfdraw)

    rtlinedraw(magnetinnerr, tmagnetend, magnetouterr, tmagnetend, dxfdraw)
    dxfdraw.save()

    sumarea, lenl, lenlh, windhig, equatoothtoplina = slotareacal(x1=toxycoord(slotradius, hoanglehalfa)[0],
                                                                  y1=toxycoord(slotradius, hoanglehalfa)[1],
                                                                  x2=toxycoord(tem, hoanglehalfa)[0],
                                                                  y2=toxycoord(tem, hoanglehalfa)[1],
                                                                  x3=toxycoord(rslottipraa, tslottipraa)[0],
                                                                  y3=toxycoord(rslottipraa, tslottipraa)[1],
                                                                  x4=toxycoord(rslottorab, tslottorab)[0],
                                                                  y4=toxycoord(rslottorab, tslottorab)[1],
                                                                  x5=toxycoord(rslottoraa, tslottoraa)[0],
                                                                  y5=toxycoord(rslottoraa, tslottoraa)[1],
                                                                  x6=toxycoord(rlslotbora, tlslotbora)[0],
                                                                  y6=toxycoord(rlslotbora, tlslotbora)[1],
                                                                  x7=toxycoord(slotradius, tslotbora)[0],
                                                                  y7=toxycoord(slotradius, tslotbora)[1],
                                                                  xo1=pctooths[0], yo1=pctooths[1],
                                                                  xo2=toxycoord(rslottorac, tslottorac)[0],
                                                                  yo2=toxycoord(rslottorac, tslottorac)[1],
                                                                  xo3=toxycoord(rslotbora, tslotbora)[0],
                                                                  yo3=toxycoord(rslotbora, tslotbora)[1], r1=rtooths,
                                                                  r2=slottora, r3=slotbora, r4=slotradius)

    if equatoothtoplina <= 0:
        equatoothtipang = hoanglehalf - 90 - equatoothtoplina
    else:
        equatoothtipang = 90 - equatoothtoplina + hoanglehalf

    return equatoothtipang, sumarea, lenl, lenlh, windhig


def drawstatotroundbo(slotnumber, outerdia, innerdiameter, polenumber, toothwhi, slotdepth, toothtipd, toothtipdm,
                      toothtipdb, toothtipop, slotbora, slottora, slottipra, airgap, magnetthi, magnetarc, backironthi,
                      dexfileloca_):
    dxfdraw = DXFEngine.drawing(dexfileloca_)
    pole = polenumber * 2
    innerradius = innerdiameter / 2
    hoangle = 360 / slotnumber
    hoanglehalf = 180 / slotnumber
    hoanglehalfa = hoanglehalf / arcttran
    outerradiu = outerdia / 2
    slotradius = outerradiu - slotdepth
    halftoothwhi = toothwhi / 2

    rslotbora = slotradius + slotbora
    tslotbora = asin((slotbora + halftoothwhi) / rslotbora)
    temtlslotborab = rslotbora * cos(tslotbora)
    if temtlslotborab == 0:
        tlslotbora = pi / 2
    else:
        tlslotbora = atan(halftoothwhi / temtlslotborab)

    rlslotbora = halftoothwhi / sin(tlslotbora)
    tl = tlslotbora * arcttran
    ta = tslotbora * arcttran
    tlmirr = hoangle - tl
    tamirr = hoangle - ta

    rtoothtipb, ttoothtipb = tortcoord(
        sqrt(outerradiu * outerradiu - halftoothwhi * halftoothwhi) - toothtipdb * toothtipd,
        halftoothwhi)
    ttoothtip = hoanglehalfa - atan(toothtipop / (2 * (outerradiu - toothtipd)))
    rtoothtip = sqrt(toothtipop * toothtipop / 4 + (outerradiu - toothtipd) * (outerradiu - toothtipd))
    ttoothtipdm = (ttoothtipb + ttoothtip) / 2
    rtoothtipdm = outerradiu - toothtipdm * toothtipd
    xt, yt = toxycoord(rtoothtipb, ttoothtipb)
    xm, ym = toxycoord(rtoothtipdm, ttoothtipdm)
    xb, yb = toxycoord(rtoothtip, ttoothtip)
    pctooths, rtooths = points2circle([xt, yt], [xm, ym], [xb, yb])
    rctooths, tctooths = tortcoord(pctooths[0], pctooths[1])

    rctooths, tctoothsa = rctooths, toarc(tctooths)
    rctoothsm, tctoothsam = rctooths, hoangle - toarc(tctooths)

    rslottorac, tslottorac, rslottoraa, tslottoraa, rslottorab, tslottorab = toothtipfilcecb(pctooths[0], pctooths[1],
                                                                                             slottora, halftoothwhi,
                                                                                             rtooths)
    rslottorac, tslottoraca, rslottoraa, tslottoraaa, rslottorab, tslottoraba = rslottorac, toarc(
        tslottorac), rslottoraa, toarc(tslottoraa), rslottorab, toarc(tslottorab)
    rslottoracm, tslottoracm, rslottoraam, tslottoraam, rslottorabm, tslottorabm = rslottorac, hoangle - tslottoraca, rslottoraa, hoangle - tslottoraaa, rslottorab, hoangle - tslottoraba

    rslottiprac, tslottiprac, rslottipraa, tslottipraa, rslottiprab, tslottiprab = toothtipfilceca(pctooths[0],
                                                                                                   pctooths[1],
                                                                                                   rctooths, tctooths,
                                                                                                   slottipra,
                                                                                                   toothtipop / 2,
                                                                                                   rtooths,
                                                                                                   hoanglehalfa)
    rslottiprac, tslottipraca, rslottipraa, tslottipraaa, rslottiprab, tslottipraba = rslottiprac, toarc(
        tslottiprac), rslottipraa, toarc(tslottipraa), rslottiprab, toarc(tslottiprab)
    rslottipracm, tslottipracm, rslottipraam, tslottipraam, rslottiprabm, tslottiprabm = rslottiprac, hoangle - tslottipraca, rslottipraa, hoangle - tslottipraaa, rslottiprab, hoangle - tslottipraba

    rlouterci, tlouterci, rhouterci, thouterci, rmbouterci, tmbouterci, rmuouterci, tmuouterci, rout, tout, routm, toutm = outercircpoincalthree(
        outerradiu, toothtipop / 2, hoanglehalfa)

    tem = rslottipraa * cos(hoanglehalfa - tslottipraa)

    magnetinnerr = outerradiu + airgap
    magnetouterr = magnetinnerr + magnetthi
    backouterr = magnetouterr + backironthi
    tbackiron = 360 / pole
    tmanet = magnetarc * 2 / pole
    tmagnetst = (tbackiron - tmanet) / 2
    tmagnetend = tmagnetst + tmanet

    rtarcdraw(0, 0, innerradius, 0, innerradius, hoangle, dxfdraw)

    rtarcdraw(rslotbora, ta, slotradius, ta, rlslotbora, tl, dxfdraw)

    rtarcdraw(0, 0, slotradius, ta, slotradius, tamirr, dxfdraw)

    rtarcdraw(rslotbora, tamirr, rlslotbora, tlmirr, slotradius, tamirr, dxfdraw)

    rtlinedraw(rlslotbora, tl, rslottoraa, tslottoraaa, dxfdraw)

    rtlinedraw(rlslotbora, tlmirr, rslottoraam, tslottoraam, dxfdraw)

    rtarcdraw(rslottorac, tslottoraca, rslottoraa, tslottoraaa, rslottorab, tslottoraba, dxfdraw)

    rtarcdraw(rslottoracm, tslottoracm, rslottorabm, tslottorabm, rslottoraam, tslottoraam, dxfdraw)

    rtarcdraw(rslottiprac, tslottipraca, rslottiprab, tslottipraba, rslottipraa, tslottipraaa, dxfdraw)

    rtarcdraw(rslottipracm, tslottipracm, rslottipraam, tslottipraam, rslottiprabm, tslottiprabm, dxfdraw)

    rtarcdraw(rctooths, tctoothsa, rslottorab, tslottoraba, rslottipraa, tslottipraaa, dxfdraw)

    rtarcdraw(rctoothsm, tctoothsam, rslottipraam, tslottipraam, rslottorabm, tslottorabm, dxfdraw)

    rtlinedraw(rslottiprab, tslottipraba, rhouterci, thouterci, dxfdraw)

    rtlinedraw(rslottiprabm, tslottiprabm, rlouterci, tlouterci, dxfdraw)

    rtarcdraw(rout, tout, rmbouterci, tmbouterci, rhouterci, thouterci, dxfdraw)

    rtarcdraw(routm, toutm, rlouterci, tlouterci, rmuouterci, tmuouterci, dxfdraw)

    rtlinedraw(rlouterci, tlouterci, rhouterci, thouterci, dxfdraw)

    rtlinedraw(rslottipraa, tslottipraaa, rslottipraam, tslottipraam, dxfdraw)

    rtlinedraw(slotradius, hoanglehalf, tem, hoanglehalf, dxfdraw)

    rtlinedraw(0, 0, rmbouterci, 0, dxfdraw)

    rtlinedraw(0, 0, rmbouterci, hoangle, dxfdraw)

    rtarcdraw(0, 0, magnetouterr, 0, magnetouterr, tbackiron, dxfdraw)

    rtarcdraw(0, 0, backouterr, 0, backouterr, tbackiron, dxfdraw)

    rtarcdraw(0, 0, magnetinnerr, 0, magnetinnerr, tbackiron, dxfdraw)

    rtlinedraw(magnetinnerr, 0, backouterr, 0, dxfdraw)

    rtlinedraw(magnetinnerr, tbackiron, backouterr, tbackiron, dxfdraw)

    rtlinedraw(magnetinnerr, tmagnetst, magnetouterr, tmagnetst, dxfdraw)

    rtlinedraw(magnetinnerr, tmagnetend, magnetouterr, tmagnetend, dxfdraw)
    dxfdraw.save()

    sumarea, lenl, lenlh, windhig, equatoothtoplina = slotareacal(x1=toxycoord(slotradius, hoanglehalfa)[0],
                                                                  y1=toxycoord(slotradius, hoanglehalfa)[1],
                                                                  x2=toxycoord(tem, hoanglehalfa)[0],
                                                                  y2=toxycoord(tem, hoanglehalfa)[1],
                                                                  x3=toxycoord(rslottipraa, tslottipraa)[0],
                                                                  y3=toxycoord(rslottipraa, tslottipraa)[1],
                                                                  x4=toxycoord(rslottorab, tslottorab)[0],
                                                                  y4=toxycoord(rslottorab, tslottorab)[1],
                                                                  x5=toxycoord(rslottoraa, tslottoraa)[0],
                                                                  y5=toxycoord(rslottoraa, tslottoraa)[1],
                                                                  x6=toxycoord(rlslotbora, tlslotbora)[0],
                                                                  y6=toxycoord(rlslotbora, tlslotbora)[1],
                                                                  x7=toxycoord(slotradius, tslotbora)[0],
                                                                  y7=toxycoord(slotradius, tslotbora)[1],
                                                                  xo1=pctooths[0], yo1=pctooths[1],
                                                                  xo2=toxycoord(rslottorac, tslottorac)[0],
                                                                  yo2=toxycoord(rslottorac, tslottorac)[1],
                                                                  xo3=toxycoord(rslotbora, tslotbora)[0],
                                                                  yo3=toxycoord(rslotbora, tslotbora)[1], r1=rtooths,
                                                                  r2=slottora, r3=slotbora, r4=slotradius)

    if equatoothtoplina <= 0:
        equatoothtipang = hoanglehalf - 90 - equatoothtoplina
    else:
        equatoothtipang = 90 - equatoothtoplina + hoanglehalf

    return equatoothtipang, sumarea, lenl, lenlh, windhig


def wirediametcal(s, f, n):
    logdebug('squaresumefi: ' + str(s))

    if s <= 0:
        s = 0
    if f <= 0:
        f = 0
    if n <= 0:
        n = 1
    diam = (s * f / n) ** 0.5
    if diam <= 0.11:
        d = diam - 0.015
    elif diam <= 0.21:
        d = diam - 0.02
    elif diam <= 0.28:
        d = diam - 0.025
    elif diam <= 0.45:
        d = diam - 0.04
    elif diam <= 0.75:
        d = diam - 0.05
    elif diam <= 1.02:
        d = diam - 0.06
    elif diam <= 1.64:
        d = diam - 0.08
    elif diam <= 2.04:
        d = diam - 0.09
    else:
        d = diam - 0.1
    return diam, d


def read_parameter_single(dexfileloca_, rownumber_, coefficient1=None, mcad1=None, num_list=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if num_list is None:
        num_list = [i * 0 for i in range(rownumber_)]
    if coefficient1 is None:
        coefficient1 = [0 * i1 for i1 in range(rownumber_)]

    ArmatureTurnsPerCoil0 = int(num_list[0])
    Shaft_Speed_RPM0 = float(num_list[1])
    Stator_Lam_Length0 = float(num_list[2])
    Back_Iron_Thickness0 = float(num_list[3])
    Tooth_Width0 = float(num_list[4])
    Airgap0 = float(num_list[5])
    RequestedGrossSlotFillFactor0 = float(num_list[6])
    Magnet_Br_Multiplier0 = float(num_list[7])
    Magnet_Arc_ED0 = float(num_list[8])
    PhaseAdvance0 = float(num_list[9])
    DCBusVoltage0 = float(num_list[10])
    PeakCurrent0 = float(num_list[11])
    Magnet_Thickness0 = float(num_list[12])
    Armature_Diameter0 = float(num_list[13])
    Slot_Opening0 = float(num_list[14])
    Tooth_Tip_Depth0 = float(num_list[15])
    Slot_Corner_Radius0 = float(num_list[16])
    Tooth_Tip_Radius0 = float(num_list[17])
    StatorSkew0 = float(num_list[18])
    RotorSkewAngle0 = float(num_list[19])
    StatorIronLossBuildFactor0 = float(num_list[20])
    RotorIronLossBuildFactor0 = float(num_list[21])
    Slot_Depth0 = float(num_list[22])
    fixrotoroutdiameter0 = float(num_list[23])
    Axle_Dia0 = float(num_list[24])
    Pole_Number0 = float(num_list[25])
    Tooth_Tip_thickmid0 = float(num_list[26])
    Tooth_Tip_thickbig0 = float(num_list[27])
    outdiametersinkbig0 = float(num_list[28])
    outdiametersinkmid0 = float(num_list[29])
    slottopradiu0 = float(num_list[30])

    if Tooth_Tip_Radius0 != 0:
        uplimit1 = 100
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[17][0], limitarray[17][1])
        Tooth_Tip_Radius0 = factorelementcorrect(Tooth_Tip_Radius0, coefficient1[17], uplimit1, lowlimit1, 1)
        coefficient1[17] = coefficientcorrect(Tooth_Tip_Radius0, coefficient1[17], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_Radius0 = factorelementcorrect(Tooth_Tip_Radius0, coefficient1[17], uplimit1, lowlimit1, 1)

        num_list[17] = Tooth_Tip_Radius0
        loggenera('槽尖圆角' + ' : ' + str(Tooth_Tip_Radius0))

    if Axle_Dia0 != 0:
        Axle_Dia01 = 0
        uplimit1 = 300
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[24][0], limitarray[24][1])
        Axle_Dia0 = factorelementcorrect(Axle_Dia0, coefficient1[24], uplimit1, lowlimit1, 1)
        coefficient1[24] = coefficientcorrect(Axle_Dia0, coefficient1[24], uplimit1, lowlimit1, 0.1)
        Axle_Dia0 = factorelementcorrect(Axle_Dia0, coefficient1[24], uplimit1, lowlimit1, 1)

        while Axle_Dia0 != Axle_Dia01:
            mcad1.SetVariable('Axle_Dia', Axle_Dia0)
            Axle_Dia01 = round(mcad1.GetVariable('Axle_Dia')[1], 1)
        num_list[24] = Axle_Dia0
        loggenera('定子内径' + ' : ' + str(Axle_Dia0))

    if Pole_Number0 != 0:
        Pole_Number01 = 0
        uplimit1 = 20
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[25][0], limitarray[25][1])
        Pole_Number0 = factorelementcorrect(Pole_Number0, coefficient1[25], uplimit1, lowlimit1, 0)
        coefficient1[25] = coefficientcorrect(Pole_Number0, coefficient1[25], uplimit1, lowlimit1, 1)
        Pole_Number0 = factorelementcorrect(Pole_Number0, coefficient1[25], uplimit1, lowlimit1, 0)

        while Pole_Number0 * 2 != Pole_Number01:
            mcad1.SetVariable('Pole_Number', Pole_Number0 * 2)
            mcad1.SetVariable('Slot_Number', Pole_Number0 * 2)
            Pole_Number01 = round(mcad1.GetVariable('Pole_Number')[1], 0)
        num_list[25] = Pole_Number0
        loggenera('极对数' + ' : ' + str(Pole_Number0))

    if Tooth_Tip_thickmid0 != 0:
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[26][0], limitarray[26][1])
        Tooth_Tip_thickmid0 = factorelementcorrect(Tooth_Tip_thickmid0, coefficient1[26], uplimit1, lowlimit1, 1)
        coefficient1[26] = coefficientcorrect(Tooth_Tip_thickmid0, coefficient1[26], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_thickmid0 = factorelementcorrect(Tooth_Tip_thickmid0, coefficient1[26], uplimit1, lowlimit1, 1)

        num_list[26] = Tooth_Tip_thickmid0
        loggenera('齿顶厚度中' + ' : ' + str(Tooth_Tip_thickmid0))

    if Tooth_Tip_thickbig0 != 0:
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[27][0], limitarray[27][1])
        Tooth_Tip_thickbig0 = factorelementcorrect(Tooth_Tip_thickbig0, coefficient1[27], uplimit1, lowlimit1, 1)
        coefficient1[27] = coefficientcorrect(Tooth_Tip_thickbig0, coefficient1[27], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_thickbig0 = factorelementcorrect(Tooth_Tip_thickbig0, coefficient1[27], uplimit1, lowlimit1, 1)

        num_list[27] = Tooth_Tip_thickbig0
        loggenera('齿顶厚度大' + ' : ' + str(Tooth_Tip_thickbig0))

    if outdiametersinkbig0 != 0:
        uplimit1 = 3
        lowlimit1 = 0.01
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[28][0], limitarray[28][1])
        outdiametersinkbig0 = factorelementcorrect(outdiametersinkbig0, coefficient1[28], uplimit1, lowlimit1, 2)
        coefficient1[28] = coefficientcorrect(outdiametersinkbig0, coefficient1[28], uplimit1, lowlimit1, 0.01)
        outdiametersinkbig0 = factorelementcorrect(outdiametersinkbig0, coefficient1[28], uplimit1, lowlimit1, 2)

        num_list[28] = outdiametersinkbig0
        loggenera('外径下沉大' + ' : ' + str(outdiametersinkbig0))

    if outdiametersinkmid0 != 0:
        uplimit1 = 3
        lowlimit1 = 0.01
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[29][0], limitarray[29][1])
        outdiametersinkmid0 = factorelementcorrect(outdiametersinkmid0, coefficient1[29], uplimit1, lowlimit1, 2)
        coefficient1[29] = coefficientcorrect(outdiametersinkmid0, coefficient1[29], uplimit1, lowlimit1, 0.01)
        outdiametersinkmid0 = factorelementcorrect(outdiametersinkmid0, coefficient1[29], uplimit1, lowlimit1, 2)

        num_list[29] = outdiametersinkmid0
        loggenera('外径下沉中' + ' : ' + str(outdiametersinkmid0))

    if slottopradiu0 != 0:
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[30][0], limitarray[30][1])
        slottopradiu0 = factorelementcorrect(slottopradiu0, coefficient1[30], uplimit1, lowlimit1, 1)
        coefficient1[30] = coefficientcorrect(slottopradiu0, coefficient1[30], uplimit1, lowlimit1, 0.1)
        slottopradiu0 = factorelementcorrect(slottopradiu0, coefficient1[30], uplimit1, lowlimit1, 1)

        num_list[30] = slottopradiu0
        loggenera('槽顶圆角' + ' : ' + str(slottopradiu0))

    if ArmatureTurnsPerCoil0 != 0:
        ArmatureTurnsPerCoil01 = 0
        uplimit1 = 300
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[0][0], limitarray[0][1])
        ArmatureTurnsPerCoil0 = factorelementcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 0)
        coefficient1[0] = coefficientcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 1)
        ArmatureTurnsPerCoil0 = factorelementcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 0)

        while ArmatureTurnsPerCoil0 != ArmatureTurnsPerCoil01:
            mcad1.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil0)
            ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)
        num_list[0] = ArmatureTurnsPerCoil0
        loggenera('匝数' + ' : ' + str(ArmatureTurnsPerCoil01))

    if Shaft_Speed_RPM0 != 0:
        Shaft_Speed_RPM01 = 0
        uplimit1 = 50000
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[1][0], limitarray[1][1])
        Shaft_Speed_RPM0 = factorelementcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 0)
        coefficient1[1] = coefficientcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 1)
        Shaft_Speed_RPM0 = factorelementcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 0)

        while Shaft_Speed_RPM0 != Shaft_Speed_RPM01:
            mcad1.SetVariable('Shaft_Speed_[RPM]', Shaft_Speed_RPM0)
            Shaft_Speed_RPM01 = round(mcad1.GetVariable('Shaft_Speed_[RPM]')[1], 0)
        num_list[1] = Shaft_Speed_RPM0
        loggenera('转速' + ' : ' + str(Shaft_Speed_RPM01))

    if Back_Iron_Thickness0 != 0:
        Back_Iron_Thickness01 = 0
        uplimit1 = 20
        lowlimit1 = 0.3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[3][0], limitarray[3][1])
        Back_Iron_Thickness0 = factorelementcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 1)
        coefficient1[3] = coefficientcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 0.1)
        Back_Iron_Thickness0 = factorelementcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 1)

        while Back_Iron_Thickness0 != Back_Iron_Thickness01:
            mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
            Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
        num_list[3] = Back_Iron_Thickness0
        loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness01))

    if Airgap0 != 0:
        Airgap01 = 0
        uplimit1 = 1
        lowlimit1 = 0.4
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[5][0], limitarray[5][1])
        Airgap0 = factorelementcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 2)
        coefficient1[5] = coefficientcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 0.01)
        Airgap0 = factorelementcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 2)

        while Airgap0 != Airgap01:
            mcad1.SetVariable('Airgap', Airgap0)
            Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
        num_list[5] = Airgap0
        loggenera('气隙' + ' : ' + str(Airgap01))

    if RequestedGrossSlotFillFactor0 != 0:
        uplimit1 = 0.9
        lowlimit1 = 0.05
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[6][0], limitarray[6][1])
        RequestedGrossSlotFillFactor0 = factorelementcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1,
                                                             lowlimit1, 2)
        coefficient1[6] = coefficientcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1, lowlimit1, 0.01)
        RequestedGrossSlotFillFactor0 = factorelementcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1,
                                                             lowlimit1, 2)

        num_list[6] = RequestedGrossSlotFillFactor0
    else:
        RequestedGrossSlotFillFactor0 = 0.4
        num_list[6] = RequestedGrossSlotFillFactor0
    loggenera('槽满率' + ' : ' + str(RequestedGrossSlotFillFactor0))

    if Magnet_Br_Multiplier0 != 0:
        Magnet_Br_Multiplier01 = 0
        uplimit1 = 20
        lowlimit1 = 0.3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[7][0], limitarray[7][1])
        Magnet_Br_Multiplier0 = factorelementcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 1)
        coefficient1[7] = coefficientcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 0.1)
        Magnet_Br_Multiplier0 = factorelementcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 1)
        while Magnet_Br_Multiplier0 != Magnet_Br_Multiplier01:
            mcad1.SetVariable('Magnet_Br_Multiplier', Magnet_Br_Multiplier0)
            Magnet_Br_Multiplier01 = round(mcad1.GetVariable('Magnet_Br_Multiplier')[1], 1)
        num_list[7] = Magnet_Br_Multiplier0
        loggenera('Br 系数' + ' : ' + str(Magnet_Br_Multiplier01))

    if Magnet_Arc_ED0 != 0:
        Magnet_Arc_ED01 = 0
        uplimit1 = 166
        lowlimit1 = 90
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[8][0], limitarray[8][1])
        Magnet_Arc_ED0 = factorelementcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 0)
        coefficient1[8] = coefficientcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 1)
        Magnet_Arc_ED0 = factorelementcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 0)
        while Magnet_Arc_ED0 != Magnet_Arc_ED01:
            mcad1.SetVariable('Magnet_Arc_[ED]', Magnet_Arc_ED0)
            Magnet_Arc_ED01 = round(mcad1.GetVariable('Magnet_Arc_[ED]')[1], 0)
        loggenera('极弧系数' + ' : ' + str(Magnet_Arc_ED01))
        num_list[8] = Magnet_Arc_ED0

    if PhaseAdvance0 != 0:
        PhaseAdvance01 = 0
        uplimit1 = 30
        lowlimit1 = -30
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[9][0], limitarray[9][1])
        PhaseAdvance0 = factorelementcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 1)
        coefficient1[9] = coefficientcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 0.1)
        PhaseAdvance0 = factorelementcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 1)

        while PhaseAdvance0 != PhaseAdvance01:
            mcad1.SetVariable('PhaseAdvance', PhaseAdvance0)
            PhaseAdvance01 = round(mcad1.GetVariable('PhaseAdvance')[1], 1)
        loggenera('超前角' + ' : ' + str(PhaseAdvance01))
        num_list[9] = PhaseAdvance0

    if DCBusVoltage0 != 0:
        DCBusVoltage01 = 0
        uplimit1 = 100
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[10][0], limitarray[10][1])
        DCBusVoltage0 = factorelementcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 1)
        coefficient1[10] = coefficientcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 0.1)
        DCBusVoltage0 = factorelementcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 1)

        while DCBusVoltage0 != DCBusVoltage01:
            mcad1.SetVariable('DCBusVoltage', DCBusVoltage0)
            DCBusVoltage01 = round(mcad1.GetVariable('DCBusVoltage')[1], 1)
        loggenera('电压' + ' : ' + str(DCBusVoltage01))
        num_list[10] = DCBusVoltage0

    if PeakCurrent0 != 0:
        PeakCurrent01 = 0
        uplimit1 = 100
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[11][0], limitarray[11][1])

        if uplimit1 < 0.2 or lowlimit1 < 0.1:
            uplimit1 = 0.2
            lowlimit1 = 0.1
        PeakCurrent0 = factorelementcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 1)
        coefficient1[11] = coefficientcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 0.1)
        PeakCurrent0 = factorelementcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 1)
        while PeakCurrent0 != PeakCurrent01:
            mcad1.SetVariable('PeakCurrent', PeakCurrent0)
            PeakCurrent01 = round(mcad1.GetVariable('PeakCurrent')[1], 1)
        loggenera('峰值电流' + ' : ' + str(PeakCurrent01))
        num_list[11] = PeakCurrent0

    if Magnet_Thickness0 != 0:
        Magnet_Thickness01 = 0
        uplimit1 = 10
        lowlimit1 = 0.5
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[12][0], limitarray[12][1])

        Magnet_Thickness0 = factorelementcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 1)
        coefficient1[12] = coefficientcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 0.1)
        Magnet_Thickness0 = factorelementcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 1)
        while Magnet_Thickness0 != Magnet_Thickness01:
            mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
            Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
        num_list[12] = Magnet_Thickness0
        loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness01))

    if StatorSkew0 != 0:
        mcad1.SetVariable('SkewType', 1)
        mcad1.SetVariable('FluxSkewFactorCalc', True)
        StatorSkew01 = 0
        uplimit1 = 30
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[18][0], limitarray[18][1])

        StatorSkew0 = factorelementcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 1)
        coefficient1[18] = coefficientcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 0.1)
        StatorSkew0 = factorelementcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 1)

        while StatorSkew0 != StatorSkew01:
            mcad1.SetVariable('StatorSkew', StatorSkew0)
            StatorSkew01 = round(mcad1.GetVariable('StatorSkew')[1], 1)
        num_list[18] = StatorSkew0

        loggenera('定子扭角' + ' : ' + str(StatorSkew01))

    if RotorSkewAngle0 != 0:
        mcad1.SetVariable('SkewType', 2)
        mcad1.SetVariable('RotorSkewSlices', 3)
        mcad1.SetVariable('AxialSegments', 3)

        RotorSkewAngle01 = 0
        uplimit1 = 30
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[19][0], limitarray[19][1])

        RotorSkewAngle0 = factorelementcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 1)
        coefficient1[19] = coefficientcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 0.1)
        RotorSkewAngle0 = factorelementcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 1)
        RotorSkewAngle02 = -RotorSkewAngle0

        while RotorSkewAngle0 != RotorSkewAngle01:
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 0, RotorSkewAngle02)
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 2, RotorSkewAngle0)
            RotorSkewAngle01 = round(mcad1.GetArrayVariable('RotorSkewAngle_Array', 2)[1], 1)

        num_list[19] = RotorSkewAngle0
        loggenera('转子扭角' + ' : ' + str(RotorSkewAngle01))

    if StatorIronLossBuildFactor0 != 0:
        StatorIronLossBuildFactor01 = 0
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[20][0], limitarray[20][1])

        StatorIronLossBuildFactor0 = factorelementcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1,
                                                          lowlimit1, 1)
        coefficient1[20] = coefficientcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1, lowlimit1, 0.1)
        StatorIronLossBuildFactor0 = factorelementcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1,
                                                          lowlimit1, 1)

        while StatorIronLossBuildFactor0 != StatorIronLossBuildFactor01:
            mcad1.SetVariable('StatorIronLossBuildFactor', StatorIronLossBuildFactor0)
            StatorIronLossBuildFactor01 = round(mcad1.GetVariable('StatorIronLossBuildFactor')[1], 1)
        num_list[20] = StatorIronLossBuildFactor0
        loggenera('定子铁耗系数' + ' : ' + str(StatorIronLossBuildFactor01))

    if RotorIronLossBuildFactor0 != 0:
        RotorIronLossBuildFactor01 = 0
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[21][0], limitarray[21][1])

        RotorIronLossBuildFactor0 = factorelementcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1,
                                                         lowlimit1, 1)
        coefficient1[21] = coefficientcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1, lowlimit1, 0.1)
        RotorIronLossBuildFactor0 = factorelementcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1,
                                                         lowlimit1, 1)
        while RotorIronLossBuildFactor0 != RotorIronLossBuildFactor01:
            mcad1.SetVariable('RotorIronLossBuildFactor', RotorIronLossBuildFactor0)
            RotorIronLossBuildFactor01 = round(mcad1.GetVariable('RotorIronLossBuildFactor')[1], 1)
        num_list[21] = RotorIronLossBuildFactor0
        loggenera('转子铁耗系数' + ' : ' + str(RotorIronLossBuildFactor01))

    if Armature_Diameter0 != 0:
        Armature_Diameter01 = 0
        uplimit1 = 120
        lowlimit1 = Axle_Dia0 + 5
        if lowlimit1 <= 0:
            lowlimit1 = 5
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[13][0], limitarray[13][1])
        Armature_Diameter0 = factorelementcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 1)
        coefficient1[13] = coefficientcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 0.1)
        Armature_Diameter0 = factorelementcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 1)

        while Armature_Diameter0 != Armature_Diameter01:
            mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
            Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
        num_list[13] = Armature_Diameter0
        loggenera('定子外径' + ' : ' + str(Armature_Diameter01))

    if fixrotoroutdiameter0 != 0:
        uplimit1 = 150
        lowlimit1 = 10
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[23][0], limitarray[23][1])
        fixrotoroutdiameter0 = factorelementcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 1)
        coefficient1[23] = coefficientcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 0.1)
        fixrotoroutdiameter0 = factorelementcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 1)

        if Armature_Diameter0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Armature_Diameter0 = round(fixrotoroutdiameter0 - Airgap01 * 2 - Back_Iron_Thickness01 * 2 - Magnet_Thickness01 * 2, 1)
            Armature_Diameter01 = 0
            while Armature_Diameter0 != Armature_Diameter01:
                mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
                Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
            loggenera('定子外径' + ' : ' + str(Armature_Diameter0))
        if Magnet_Thickness0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Magnet_Thickness0 = round((fixrotoroutdiameter0 - Airgap01 * 2 - Back_Iron_Thickness01 * 2 - Armature_Diameter01)/2, 1)
            Magnet_Thickness01 = 0
            while Magnet_Thickness0 != Magnet_Thickness01:
                mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
                Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
            loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness0))

        if Airgap0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Airgap0 = round((fixrotoroutdiameter0 - Magnet_Thickness01 * 2 - Back_Iron_Thickness01 * 2 - Armature_Diameter01)/2, 2)
            Airgap01 = 0
            while Airgap0 != Airgap01:
                mcad1.SetVariable('Airgap', Airgap0)
                Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
            loggenera('气隙' + ' : ' + str(Airgap0))

        if Back_Iron_Thickness0 == 0:
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Back_Iron_Thickness0 = round((fixrotoroutdiameter0 - Magnet_Thickness01 * 2 - Airgap01 * 2 - Armature_Diameter01)/2, 1)
            Back_Iron_Thickness01 = 0
            while Back_Iron_Thickness0 != Back_Iron_Thickness01:
                mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
                Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
            loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness0))

        num_list[23] = fixrotoroutdiameter0

    if Tooth_Tip_Depth0 != 0:
        Tooth_Tip_Depth01 = 0
        uplimit1 = 5
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[15][0], limitarray[15][1])
        Tooth_Tip_Depth0 = factorelementcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 1)
        coefficient1[15] = coefficientcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_Depth0 = factorelementcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 1)

        while Tooth_Tip_Depth0 != Tooth_Tip_Depth01:
            mcad1.SetVariable('Tooth_Tip_Depth', Tooth_Tip_Depth0)
            Tooth_Tip_Depth01 = round(mcad1.GetVariable('Tooth_Tip_Depth')[1], 1)
        num_list[15] = Tooth_Tip_Depth0

        loggenera('槽顶深' + ' : ' + str(Tooth_Tip_Depth01))

    if Tooth_Width0 != 0:
        Tooth_Width01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        uplimit1 = Armature_Diameter01 * 0.75 / Slot_Number01
        lowlimit1 = 2
        if uplimit1 < 3:
            uplimit1 = 3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[4][0], limitarray[4][1])

        Tooth_Width0 = factorelementcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 1)
        coefficient1[4] = coefficientcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 0.1)
        Tooth_Width0 = factorelementcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 1)

        while Tooth_Width0 != Tooth_Width01:
            mcad1.SetVariable('Tooth_Width', Tooth_Width0)
            Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)

        loggenera('齿宽' + ' : ' + str(Tooth_Width01))
        num_list[4] = Tooth_Width0

    if Slot_Depth0 != 0:
        Tooth_Tip_Depth01 = Tooth_Tip_Depth0 * (1 - coefficient1[15])
        Armature_Diameter01 = Armature_Diameter0 * (1 + coefficient1[13])
        Axle_Dia1 = Axle_Dia0 * (1 - coefficient1[24])
        Tooth_Tip_Depth011 = Tooth_Tip_Depth0 * (1 + coefficient1[15])
        Armature_Diameter011 = Armature_Diameter0 * (1 - coefficient1[13])
        Axle_Dia11 = Axle_Dia0 * (1 + coefficient1[24])
        Slot_Depth0tem = 0
        uplimit1 = ((Armature_Diameter011 - Axle_Dia11) / 2 - Tooth_Tip_Depth011) * 0.95
        lowlimit1 = ((Armature_Diameter01 - Axle_Dia1) / 2 - Tooth_Tip_Depth01) * 0.05
        if lowlimit1 <= 0:
            lowlimit1 = 3
        if uplimit1 <= 0 or uplimit1 <= lowlimit1:
            uplimit1 = lowlimit1 + 3
        while Slot_Depth0tem <= 1:
            if uplimit1 <= lowlimit1:
                uplimit1 = lowlimit1 + 3
            uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[22][0], limitarray[22][1])

            Slot_Depth0 = factorelementcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 1)
            coefficient1[22] = coefficientcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 0.1)
            Slot_Depth0 = factorelementcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 1)
            Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
            Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)
            Slot_Depth0tem = fakeslotdepth(Armature_Diameter0, Slot_Depth0, Tooth_Width01, Slot_Number01)
            lowlimit1 = lowlimit1 + 0.1
            if limitarray[22][0] != 0:
                limitarray[22][0] = limitarray[22][0] + 0.1

        Slot_Depth01 = 0
        while Slot_Depth0tem != Slot_Depth01:
            mcad1.SetVariable('Slot_Depth', Slot_Depth0tem)
            Slot_Depth01 = round(mcad1.GetVariable('Slot_Depth')[1], 1)

        loggenera('槽深' + ' : ' + str(Slot_Depth0))
        num_list[22] = Slot_Depth0

    if Slot_Corner_Radius0 != 0:
        Slot_Corner_Radius01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = Tooth_Width0 * (1 + coefficient1[4])
        Slot_Depth01 = Slot_Depth0 * (1 - coefficient1[22])
        Tooth_Tip_Depth01 = Tooth_Tip_Depth0 * (1 + coefficient1[15])
        uplimit1 = ((Armature_Diameter01 - (
                Slot_Depth01 + Tooth_Tip_Depth01) * 2) * 3.14 / Slot_Number01 - Tooth_Width01 - 1) * 0.8
        lowlimit1 = 0.2
        if uplimit1 < 0.3:
            uplimit1 = 0.3
        elif uplimit1 > 20:
            uplimit1 = 20
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[16][0], limitarray[16][1])

        Slot_Corner_Radius0 = factorelementcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 1)
        coefficient1[16] = coefficientcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 0.1)
        Slot_Corner_Radius0 = factorelementcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 1)

        while Slot_Corner_Radius0 != Slot_Corner_Radius01:
            mcad1.SetVariable('Slot_Corner_Radius', Slot_Corner_Radius0)
            Slot_Corner_Radius01 = round(mcad1.GetVariable('Slot_Corner_Radius')[1], 1)

        loggenera('槽底圆角' + ' : ' + str(Slot_Corner_Radius01))
        num_list[16] = Slot_Corner_Radius0

    if Slot_Opening0 != 0:
        Slot_Opening01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = Tooth_Width0 * (1 + coefficient1[4])
        Slot_Depth01 = Slot_Depth0 * (1 + coefficient1[22])

        uplimit1 = ((Armature_Diameter01 - Slot_Depth01 * 2) * 3.14 / Slot_Number01 - Tooth_Width01) * 0.7
        lowlimit1 = 2.5
        if uplimit1 < 2.6:
            uplimit1 = 2.6
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[14][0], limitarray[14][1])

        Slot_Opening0 = factorelementcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 1)
        coefficient1[14] = coefficientcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 0.1)
        Slot_Opening0 = factorelementcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 1)

        while Slot_Opening0 != Slot_Opening01:
            mcad1.SetVariable('Slot_Opening', Slot_Opening0)
            Slot_Opening01 = round(mcad1.GetVariable('Slot_Opening')[1], 1)

        loggenera('槽口宽' + ' : ' + str(Slot_Opening01))
        num_list[14] = Slot_Opening0

    if Stator_Lam_Length0 != 0:
        Stator_Lam_Length01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        uplimit1 = Armature_Diameter01 * 100
        lowlimit1 = 5
        if uplimit1 <= lowlimit1:
            uplimit1 = 500

        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[2][0], limitarray[2][1])

        Stator_Lam_Length0 = factorelementcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 1)
        coefficient1[2] = coefficientcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 0.1)
        Stator_Lam_Length0 = factorelementcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 1)

        while Stator_Lam_Length0 != Stator_Lam_Length01:
            mcad1.SetVariable('Stator_Lam_Length', Stator_Lam_Length0)
            Stator_Lam_Length01 = round(mcad1.GetVariable('Stator_Lam_Length')[1], 1)
        num_list[2] = Stator_Lam_Length0
        mcad1.SetVariable('Magnet_Length', Stator_Lam_Length0)
        mcad1.SetVariable('Rotor_Lam_Length', Stator_Lam_Length0 + 20)

        loggenera('铁长' + ' : ' + str(Stator_Lam_Length01))

    Armature_Diameter0 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
    Axle_Dia0 = round(mcad1.GetVariable('Axle_Dia')[1], 1)
    Pole_Number0 = round(mcad1.GetVariable('Pole_Number')[1]/2, 0)
    Tooth_Width0 = round(mcad1.GetVariable('Tooth_Width')[1], 1)
    Tooth_Tip_Depth0 = round(mcad1.GetVariable('Tooth_Tip_Depth')[1], 1)
    Slot_Opening0 = round(mcad1.GetVariable('Slot_Opening')[1], 1)
    Slot_Corner_Radius0 = round(mcad1.GetVariable('Slot_Corner_Radius')[1], 1)
    Airgap0 = round(mcad1.GetVariable('Airgap')[1], 2)
    Magnet_Thickness0 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
    Magnet_Arc_ED0 = round(mcad1.GetVariable('Magnet_Arc_[ED]')[1], 0)
    Back_Iron_Thickness0 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
    Tooth_Tip_Angle0, sumarea, lenl, lenlh, EWdg_Overhang = drawstatot(outerdia=Armature_Diameter0,
                                                                       innerdiameter=Axle_Dia0,
                                                                       polenumber=Pole_Number0,
                                                                       toothwhi=Tooth_Width0, slotdepth=Slot_Depth0,
                                                                       toothtipd=Tooth_Tip_Depth0,
                                                                       toothtipdm=Tooth_Tip_thickmid0,
                                                                       toothtipdb=Tooth_Tip_thickbig0,
                                                                       toothtipop=Slot_Opening0,
                                                                       slotbora=Slot_Corner_Radius0,
                                                                       slottora=slottopradiu0,
                                                                       slottipra=Tooth_Tip_Radius0,
                                                                       outermag=outdiametersinkbig0,
                                                                       outermidg=outdiametersinkmid0, airgap=Airgap0,
                                                                       magnetthi=Magnet_Thickness0,
                                                                       magnetarc=Magnet_Arc_ED0,
                                                                       backironthi=Back_Iron_Thickness0,
                                                                       dexfileloca_=dexfileloca_)
    EWdg_Overhang = round(EWdg_Overhang * RequestedGrossSlotFillFactor0, 1)
    mcad1.SetVariable('EWdg_Overhang_[F]', EWdg_Overhang)
    mcad1.SetVariable('EWdg_Overhang_[R]', EWdg_Overhang)
    loggenera('绕组端部长' + ' : ' + str(EWdg_Overhang))
    Liner_Thickness = mcad1.GetVariable('Liner_Thickness')[1]
    sumareaeff = sumarea - (lenl - Liner_Thickness) * Slot_Opening0 / 2 - lenlh * Liner_Thickness
    ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)

    wirediam, wirediamco = wirediametcal(sumareaeff, RequestedGrossSlotFillFactor0, ArmatureTurnsPerCoil01)
    Wire_Diameter0, Copper_Diameter0 = round(wirediam, 3), round(wirediamco, 3)

    mcad1.SetVariable('Wire_Diameter', Wire_Diameter0)
    mcad1.SetVariable('Copper_Diameter', Copper_Diameter0)
    loggenera('裸线径' + ' : ' + str(Copper_Diameter0))
    loggenera('线径' + ' : ' + str(Wire_Diameter0))

    Tooth_Tip_Angle0 = round(Tooth_Tip_Angle0, 0)
    if Tooth_Tip_Angle0 != 0:
        Tooth_Tip_Angle01 = 0

        while Tooth_Tip_Angle0 != Tooth_Tip_Angle01:
            mcad1.SetVariable('Tooth_Tip_Angle', Tooth_Tip_Angle0)
            Tooth_Tip_Angle01 = round(mcad1.GetVariable('Tooth_Tip_Angle')[1], 0)

        loggenera('槽顶角度' + ' : ' + str(Tooth_Tip_Angle01))


def read_parameter_threeroundbo(dexfileloca_, rownumber_, coefficient1=None, mcad1=None, num_list=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if num_list is None:
        num_list = [i * 0 for i in range(rownumber_)]
    if coefficient1 is None:
        coefficient1 = [0 * i1 for i1 in range(rownumber_)]

    ArmatureTurnsPerCoil0 = int(num_list[0])
    Shaft_Speed_RPM0 = float(num_list[1])
    Stator_Lam_Length0 = float(num_list[2])
    Back_Iron_Thickness0 = float(num_list[3])
    Tooth_Width0 = float(num_list[4])
    Airgap0 = float(num_list[5])
    RequestedGrossSlotFillFactor0 = float(num_list[6])
    Magnet_Br_Multiplier0 = float(num_list[7])
    Magnet_Arc_ED0 = float(num_list[8])
    PhaseAdvance0 = float(num_list[9])
    DCBusVoltage0 = float(num_list[10])
    PeakCurrent0 = float(num_list[11])
    Magnet_Thickness0 = float(num_list[12])
    Armature_Diameter0 = float(num_list[13])
    Slot_Opening0 = float(num_list[14])
    Tooth_Tip_Depth0 = float(num_list[15])
    Slot_Corner_Radius0 = float(num_list[16])
    Tooth_Tip_Radius0 = float(num_list[17])
    StatorSkew0 = float(num_list[18])
    RotorSkewAngle0 = float(num_list[19])
    StatorIronLossBuildFactor0 = float(num_list[20])
    RotorIronLossBuildFactor0 = float(num_list[21])
    Slot_Depth0 = float(num_list[22])
    fixrotoroutdiameter0 = float(num_list[23])
    Axle_Dia0 = float(num_list[24])
    Pole_Number0 = float(num_list[25])
    Tooth_Tip_thickmid0 = float(num_list[26])
    Tooth_Tip_thickbig0 = float(num_list[27])
    slottopradiu0 = float(num_list[28])
    slotnumbm30 = float(num_list[29])

    if Tooth_Tip_Radius0 != 0:
        uplimit1 = 100
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[17][0], limitarray[17][1])
        Tooth_Tip_Radius0 = factorelementcorrect(Tooth_Tip_Radius0, coefficient1[17], uplimit1, lowlimit1, 1)
        coefficient1[17] = coefficientcorrect(Tooth_Tip_Radius0, coefficient1[17], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_Radius0 = factorelementcorrect(Tooth_Tip_Radius0, coefficient1[17], uplimit1, lowlimit1, 1)

        num_list[17] = Tooth_Tip_Radius0
        loggenera('槽尖圆角' + ' : ' + str(Tooth_Tip_Radius0))

    if Axle_Dia0 != 0:
        Axle_Dia01 = 0
        uplimit1 = 300
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[24][0], limitarray[24][1])
        Axle_Dia0 = factorelementcorrect(Axle_Dia0, coefficient1[24], uplimit1, lowlimit1, 1)
        coefficient1[24] = coefficientcorrect(Axle_Dia0, coefficient1[24], uplimit1, lowlimit1, 0.1)
        Axle_Dia0 = factorelementcorrect(Axle_Dia0, coefficient1[24], uplimit1, lowlimit1, 1)

        while Axle_Dia0 != Axle_Dia01:
            mcad1.SetVariable('Axle_Dia', Axle_Dia0)
            Axle_Dia01 = round(mcad1.GetVariable('Axle_Dia')[1], 1)
        num_list[24] = Axle_Dia0
        loggenera('定子内径' + ' : ' + str(Axle_Dia0))

    if Pole_Number0 != 0:
        Pole_Number01 = 0
        uplimit1 = 20
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[25][0], limitarray[25][1])
        Pole_Number0 = factorelementcorrect(Pole_Number0, coefficient1[25], uplimit1, lowlimit1, 0)
        coefficient1[25] = coefficientcorrect(Pole_Number0, coefficient1[25], uplimit1, lowlimit1, 1)
        Pole_Number0 = factorelementcorrect(Pole_Number0, coefficient1[25], uplimit1, lowlimit1, 0)

        while Pole_Number0 * 2 != Pole_Number01:
            mcad1.SetVariable('Pole_Number', Pole_Number0 * 2)
            mcad1.SetVariable('Slot_Number', Pole_Number0 * 2)
            Pole_Number01 = round(mcad1.GetVariable('Pole_Number')[1], 0)
        num_list[25] = Pole_Number0
        loggenera('极对数' + ' : ' + str(Pole_Number0))

    if slotnumbm30 != 0:
        slotnumbm31 = 0
        uplimit1 = 20
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[29][0], limitarray[29][1])
        slotnumbm30 = factorelementcorrect(slotnumbm30, coefficient1[29], uplimit1, lowlimit1, 0)
        coefficient1[29] = coefficientcorrect(slotnumbm30, coefficient1[29], uplimit1, lowlimit1, 1)
        slotnumbm30 = factorelementcorrect(slotnumbm30, coefficient1[29], uplimit1, lowlimit1, 0)
        slotnumb = slotnumbm30 * 3
        while slotnumb != slotnumbm31:
            mcad1.SetVariable('Slot_Number', slotnumb)
            slotnumbm31 = round(mcad1.GetVariable('Slot_Number')[1], 0)
        num_list[29] = slotnumbm30
        loggenera('槽数/3' + ' : ' + str(slotnumbm30))

    if Tooth_Tip_thickmid0 != 0:
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[26][0], limitarray[26][1])
        Tooth_Tip_thickmid0 = factorelementcorrect(Tooth_Tip_thickmid0, coefficient1[26], uplimit1, lowlimit1, 1)
        coefficient1[26] = coefficientcorrect(Tooth_Tip_thickmid0, coefficient1[26], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_thickmid0 = factorelementcorrect(Tooth_Tip_thickmid0, coefficient1[26], uplimit1, lowlimit1, 1)

        num_list[26] = Tooth_Tip_thickmid0
        loggenera('齿顶厚度中' + ' : ' + str(Tooth_Tip_thickmid0))

    if Tooth_Tip_thickbig0 != 0:
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[27][0], limitarray[27][1])
        Tooth_Tip_thickbig0 = factorelementcorrect(Tooth_Tip_thickbig0, coefficient1[27], uplimit1, lowlimit1, 1)
        coefficient1[27] = coefficientcorrect(Tooth_Tip_thickbig0, coefficient1[27], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_thickbig0 = factorelementcorrect(Tooth_Tip_thickbig0, coefficient1[27], uplimit1, lowlimit1, 1)

        num_list[27] = Tooth_Tip_thickbig0
        loggenera('齿顶厚度大' + ' : ' + str(Tooth_Tip_thickbig0))

    if slottopradiu0 != 0:
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[28][0], limitarray[28][1])
        slottopradiu0 = factorelementcorrect(slottopradiu0, coefficient1[28], uplimit1, lowlimit1, 1)
        coefficient1[28] = coefficientcorrect(slottopradiu0, coefficient1[28], uplimit1, lowlimit1, 0.1)
        slottopradiu0 = factorelementcorrect(slottopradiu0, coefficient1[28], uplimit1, lowlimit1, 1)

        num_list[28] = slottopradiu0
        loggenera('槽顶圆角' + ' : ' + str(slottopradiu0))

    if ArmatureTurnsPerCoil0 != 0:
        ArmatureTurnsPerCoil01 = 0
        uplimit1 = 300
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[0][0], limitarray[0][1])
        ArmatureTurnsPerCoil0 = factorelementcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 0)
        coefficient1[0] = coefficientcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 1)
        ArmatureTurnsPerCoil0 = factorelementcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 0)

        while ArmatureTurnsPerCoil0 != ArmatureTurnsPerCoil01:
            mcad1.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil0)
            ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)
        num_list[0] = ArmatureTurnsPerCoil0
        loggenera('匝数' + ' : ' + str(ArmatureTurnsPerCoil01))

    if Shaft_Speed_RPM0 != 0:
        Shaft_Speed_RPM01 = 0
        uplimit1 = 50000
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[1][0], limitarray[1][1])
        Shaft_Speed_RPM0 = factorelementcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 0)
        coefficient1[1] = coefficientcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 1)
        Shaft_Speed_RPM0 = factorelementcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 0)

        while Shaft_Speed_RPM0 != Shaft_Speed_RPM01:
            mcad1.SetVariable('Shaft_Speed_[RPM]', Shaft_Speed_RPM0)
            Shaft_Speed_RPM01 = round(mcad1.GetVariable('Shaft_Speed_[RPM]')[1], 0)
        num_list[1] = Shaft_Speed_RPM0
        loggenera('转速' + ' : ' + str(Shaft_Speed_RPM01))

    if Back_Iron_Thickness0 != 0:
        Back_Iron_Thickness01 = 0
        uplimit1 = 20
        lowlimit1 = 0.3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[3][0], limitarray[3][1])
        Back_Iron_Thickness0 = factorelementcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 1)
        coefficient1[3] = coefficientcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 0.1)
        Back_Iron_Thickness0 = factorelementcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 1)

        while Back_Iron_Thickness0 != Back_Iron_Thickness01:
            mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
            Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
        num_list[3] = Back_Iron_Thickness0
        loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness01))

    if Airgap0 != 0:
        Airgap01 = 0
        uplimit1 = 1
        lowlimit1 = 0.4
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[5][0], limitarray[5][1])
        Airgap0 = factorelementcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 2)
        coefficient1[5] = coefficientcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 0.01)
        Airgap0 = factorelementcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 2)

        while Airgap0 != Airgap01:
            mcad1.SetVariable('Airgap', Airgap0)
            Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
        num_list[5] = Airgap0
        loggenera('气隙' + ' : ' + str(Airgap01))

    if RequestedGrossSlotFillFactor0 != 0:
        uplimit1 = 0.9
        lowlimit1 = 0.05
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[6][0], limitarray[6][1])
        RequestedGrossSlotFillFactor0 = factorelementcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1,
                                                             lowlimit1, 2)
        coefficient1[6] = coefficientcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1, lowlimit1, 0.01)
        RequestedGrossSlotFillFactor0 = factorelementcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1,
                                                             lowlimit1, 2)

        num_list[6] = RequestedGrossSlotFillFactor0
    else:
        RequestedGrossSlotFillFactor0 = 0.4
        num_list[6] = RequestedGrossSlotFillFactor0
    loggenera('槽满率' + ' : ' + str(RequestedGrossSlotFillFactor0))

    if Magnet_Br_Multiplier0 != 0:
        Magnet_Br_Multiplier01 = 0
        uplimit1 = 20
        lowlimit1 = 0.3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[7][0], limitarray[7][1])
        Magnet_Br_Multiplier0 = factorelementcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 1)
        coefficient1[7] = coefficientcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 0.1)
        Magnet_Br_Multiplier0 = factorelementcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 1)
        while Magnet_Br_Multiplier0 != Magnet_Br_Multiplier01:
            mcad1.SetVariable('Magnet_Br_Multiplier', Magnet_Br_Multiplier0)
            Magnet_Br_Multiplier01 = round(mcad1.GetVariable('Magnet_Br_Multiplier')[1], 1)
        num_list[7] = Magnet_Br_Multiplier0
        loggenera('Br 系数' + ' : ' + str(Magnet_Br_Multiplier01))

    if Magnet_Arc_ED0 != 0:
        Magnet_Arc_ED01 = 0
        uplimit1 = 166
        lowlimit1 = 90
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[8][0], limitarray[8][1])
        Magnet_Arc_ED0 = factorelementcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 0)
        coefficient1[8] = coefficientcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 1)
        Magnet_Arc_ED0 = factorelementcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 0)
        while Magnet_Arc_ED0 != Magnet_Arc_ED01:
            mcad1.SetVariable('Magnet_Arc_[ED]', Magnet_Arc_ED0)
            Magnet_Arc_ED01 = round(mcad1.GetVariable('Magnet_Arc_[ED]')[1], 0)
        loggenera('极弧系数' + ' : ' + str(Magnet_Arc_ED01))
        num_list[8] = Magnet_Arc_ED0

    if PhaseAdvance0 != 0:
        PhaseAdvance01 = 0
        uplimit1 = 30
        lowlimit1 = -30
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[9][0], limitarray[9][1])
        PhaseAdvance0 = factorelementcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 1)
        coefficient1[9] = coefficientcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 0.1)
        PhaseAdvance0 = factorelementcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 1)

        while PhaseAdvance0 != PhaseAdvance01:
            mcad1.SetVariable('PhaseAdvance', PhaseAdvance0)
            PhaseAdvance01 = round(mcad1.GetVariable('PhaseAdvance')[1], 1)
        loggenera('超前角' + ' : ' + str(PhaseAdvance01))
        num_list[9] = PhaseAdvance0

    if DCBusVoltage0 != 0:
        DCBusVoltage01 = 0
        uplimit1 = 100
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[10][0], limitarray[10][1])
        DCBusVoltage0 = factorelementcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 1)
        coefficient1[10] = coefficientcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 0.1)
        DCBusVoltage0 = factorelementcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 1)

        while DCBusVoltage0 != DCBusVoltage01:
            mcad1.SetVariable('DCBusVoltage', DCBusVoltage0)
            DCBusVoltage01 = round(mcad1.GetVariable('DCBusVoltage')[1], 1)
        loggenera('电压' + ' : ' + str(DCBusVoltage01))
        num_list[10] = DCBusVoltage0

    if PeakCurrent0 != 0:
        PeakCurrent01 = 0
        uplimit1 = 100
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[11][0], limitarray[11][1])

        if uplimit1 < 0.2 or lowlimit1 < 0.1:
            uplimit1 = 0.2
            lowlimit1 = 0.1
        PeakCurrent0 = factorelementcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 1)
        coefficient1[11] = coefficientcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 0.1)
        PeakCurrent0 = factorelementcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 1)
        while PeakCurrent0 != PeakCurrent01:
            mcad1.SetVariable('PeakCurrent', PeakCurrent0)
            PeakCurrent01 = round(mcad1.GetVariable('PeakCurrent')[1], 1)
        loggenera('峰值电流' + ' : ' + str(PeakCurrent01))
        num_list[11] = PeakCurrent0

    if Magnet_Thickness0 != 0:
        Magnet_Thickness01 = 0
        uplimit1 = 10
        lowlimit1 = 0.5
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[12][0], limitarray[12][1])

        Magnet_Thickness0 = factorelementcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 1)
        coefficient1[12] = coefficientcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 0.1)
        Magnet_Thickness0 = factorelementcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 1)
        while Magnet_Thickness0 != Magnet_Thickness01:
            mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
            Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
        num_list[12] = Magnet_Thickness0
        loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness01))

    if StatorSkew0 != 0:
        mcad1.SetVariable('SkewType', 1)
        mcad1.SetVariable('FluxSkewFactorCalc', True)
        StatorSkew01 = 0
        uplimit1 = 30
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[18][0], limitarray[18][1])

        StatorSkew0 = factorelementcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 1)
        coefficient1[18] = coefficientcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 0.1)
        StatorSkew0 = factorelementcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 1)

        while StatorSkew0 != StatorSkew01:
            mcad1.SetVariable('StatorSkew', StatorSkew0)
            StatorSkew01 = round(mcad1.GetVariable('StatorSkew')[1], 1)
        num_list[18] = StatorSkew0

        loggenera('定子扭角' + ' : ' + str(StatorSkew01))

    if RotorSkewAngle0 != 0:
        mcad1.SetVariable('SkewType', 2)
        mcad1.SetVariable('RotorSkewSlices', 3)
        mcad1.SetVariable('AxialSegments', 3)

        RotorSkewAngle01 = 0
        uplimit1 = 30
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[19][0], limitarray[19][1])

        RotorSkewAngle0 = factorelementcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 1)
        coefficient1[19] = coefficientcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 0.1)
        RotorSkewAngle0 = factorelementcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 1)
        RotorSkewAngle02 = -RotorSkewAngle0

        while RotorSkewAngle0 != RotorSkewAngle01:
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 0, RotorSkewAngle02)
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 2, RotorSkewAngle0)
            RotorSkewAngle01 = round(mcad1.GetArrayVariable('RotorSkewAngle_Array', 2)[1], 1)

        num_list[19] = RotorSkewAngle0
        loggenera('转子扭角' + ' : ' + str(RotorSkewAngle01))

    if StatorIronLossBuildFactor0 != 0:
        StatorIronLossBuildFactor01 = 0
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[20][0], limitarray[20][1])

        StatorIronLossBuildFactor0 = factorelementcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1,
                                                          lowlimit1, 1)
        coefficient1[20] = coefficientcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1, lowlimit1, 0.1)
        StatorIronLossBuildFactor0 = factorelementcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1,
                                                          lowlimit1, 1)

        while StatorIronLossBuildFactor0 != StatorIronLossBuildFactor01:
            mcad1.SetVariable('StatorIronLossBuildFactor', StatorIronLossBuildFactor0)
            StatorIronLossBuildFactor01 = round(mcad1.GetVariable('StatorIronLossBuildFactor')[1], 1)
        num_list[20] = StatorIronLossBuildFactor0
        loggenera('定子铁耗系数' + ' : ' + str(StatorIronLossBuildFactor01))

    if RotorIronLossBuildFactor0 != 0:
        RotorIronLossBuildFactor01 = 0
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[21][0], limitarray[21][1])

        RotorIronLossBuildFactor0 = factorelementcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1,
                                                         lowlimit1, 1)
        coefficient1[21] = coefficientcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1, lowlimit1, 0.1)
        RotorIronLossBuildFactor0 = factorelementcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1,
                                                         lowlimit1, 1)
        while RotorIronLossBuildFactor0 != RotorIronLossBuildFactor01:
            mcad1.SetVariable('RotorIronLossBuildFactor', RotorIronLossBuildFactor0)
            RotorIronLossBuildFactor01 = round(mcad1.GetVariable('RotorIronLossBuildFactor')[1], 1)
        num_list[21] = RotorIronLossBuildFactor0
        loggenera('转子铁耗系数' + ' : ' + str(RotorIronLossBuildFactor01))

    if Armature_Diameter0 != 0:
        Armature_Diameter01 = 0
        uplimit1 = 120
        lowlimit1 = Axle_Dia0 + 5
        if lowlimit1 <= 0:
            lowlimit1 = 5
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[13][0], limitarray[13][1])
        Armature_Diameter0 = factorelementcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 1)
        coefficient1[13] = coefficientcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 0.1)
        Armature_Diameter0 = factorelementcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 1)

        while Armature_Diameter0 != Armature_Diameter01:
            mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
            Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
        num_list[13] = Armature_Diameter0
        loggenera('定子外径' + ' : ' + str(Armature_Diameter01))

    if fixrotoroutdiameter0 != 0:
        uplimit1 = 150
        lowlimit1 = 10
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[23][0], limitarray[23][1])
        fixrotoroutdiameter0 = factorelementcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 1)
        coefficient1[23] = coefficientcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 0.1)
        fixrotoroutdiameter0 = factorelementcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 1)

        if Armature_Diameter0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Armature_Diameter0 = round(fixrotoroutdiameter0 - Airgap01 * 2 - Back_Iron_Thickness01 * 2 - Magnet_Thickness01 * 2, 1)
            Armature_Diameter01 = 0
            while Armature_Diameter0 != Armature_Diameter01:
                mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
                Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
            loggenera('定子外径' + ' : ' + str(Armature_Diameter0))
        if Magnet_Thickness0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Magnet_Thickness0 = round((fixrotoroutdiameter0 - Airgap01 * 2 - Back_Iron_Thickness01 * 2 - Armature_Diameter01)/2, 1)
            Magnet_Thickness01 = 0
            while Magnet_Thickness0 != Magnet_Thickness01:
                mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
                Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
            loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness0))

        if Airgap0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Airgap0 = round((fixrotoroutdiameter0 - Magnet_Thickness01 * 2 - Back_Iron_Thickness01 * 2 - Armature_Diameter01)/2, 2)
            Airgap01 = 0
            while Airgap0 != Airgap01:
                mcad1.SetVariable('Airgap', Airgap0)
                Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
            loggenera('气隙' + ' : ' + str(Airgap0))

        if Back_Iron_Thickness0 == 0:
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Back_Iron_Thickness0 = round((fixrotoroutdiameter0 - Magnet_Thickness01 * 2 - Airgap01 * 2 - Armature_Diameter01)/2, 1)
            Back_Iron_Thickness01 = 0
            while Back_Iron_Thickness0 != Back_Iron_Thickness01:
                mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
                Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
            loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness0))

        num_list[23] = fixrotoroutdiameter0

    if Tooth_Tip_Depth0 != 0:
        Tooth_Tip_Depth01 = 0
        uplimit1 = 5
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[15][0], limitarray[15][1])
        Tooth_Tip_Depth0 = factorelementcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 1)
        coefficient1[15] = coefficientcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_Depth0 = factorelementcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 1)

        while Tooth_Tip_Depth0 != Tooth_Tip_Depth01:
            mcad1.SetVariable('Tooth_Tip_Depth', Tooth_Tip_Depth0)
            Tooth_Tip_Depth01 = round(mcad1.GetVariable('Tooth_Tip_Depth')[1], 1)
        num_list[15] = Tooth_Tip_Depth0

        loggenera('槽顶深' + ' : ' + str(Tooth_Tip_Depth01))

    if Tooth_Width0 != 0:
        Tooth_Width01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        uplimit1 = Armature_Diameter01 * 0.75 / Slot_Number01
        lowlimit1 = 2
        if uplimit1 < 3:
            uplimit1 = 3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[4][0], limitarray[4][1])

        Tooth_Width0 = factorelementcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 1)
        coefficient1[4] = coefficientcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 0.1)
        Tooth_Width0 = factorelementcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 1)

        while Tooth_Width0 != Tooth_Width01:
            mcad1.SetVariable('Tooth_Width', Tooth_Width0)
            Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)

        loggenera('齿宽' + ' : ' + str(Tooth_Width01))
        num_list[4] = Tooth_Width0

    if Slot_Depth0 != 0:
        Tooth_Tip_Depth01 = Tooth_Tip_Depth0 * (1 - coefficient1[15])
        Armature_Diameter01 = Armature_Diameter0 * (1 + coefficient1[13])
        Axle_Dia1 = Axle_Dia0 * (1 - coefficient1[24])
        Tooth_Tip_Depth011 = Tooth_Tip_Depth0 * (1 + coefficient1[15])
        Armature_Diameter011 = Armature_Diameter0 * (1 - coefficient1[13])
        Axle_Dia11 = Axle_Dia0 * (1 + coefficient1[24])
        Slot_Depth0tem = 0
        uplimit1 = ((Armature_Diameter011 - Axle_Dia11) / 2 - Tooth_Tip_Depth011) * 0.95
        lowlimit1 = ((Armature_Diameter01 - Axle_Dia1) / 2 - Tooth_Tip_Depth01) * 0.05
        if lowlimit1 <= 0:
            lowlimit1 = 3
        if uplimit1 <= 0 or uplimit1 <= lowlimit1:
            uplimit1 = lowlimit1 + 3
        while Slot_Depth0tem <= 1:
            if uplimit1 <= lowlimit1:
                uplimit1 = lowlimit1 + 3
            uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[22][0], limitarray[22][1])

            Slot_Depth0 = factorelementcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 1)
            coefficient1[22] = coefficientcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 0.1)
            Slot_Depth0 = factorelementcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 1)
            Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
            Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)
            Slot_Depth0tem = fakeslotdepth(Armature_Diameter0, Slot_Depth0, Tooth_Width01, Slot_Number01)
            lowlimit1 = lowlimit1 + 0.1
            if limitarray[22][0] != 0:
                limitarray[22][0] = limitarray[22][0] + 0.1

        Slot_Depth01 = 0
        while Slot_Depth0tem != Slot_Depth01:
            mcad1.SetVariable('Slot_Depth', Slot_Depth0tem)
            Slot_Depth01 = round(mcad1.GetVariable('Slot_Depth')[1], 1)

        loggenera('槽深' + ' : ' + str(Slot_Depth0))
        num_list[22] = Slot_Depth0

    if Slot_Corner_Radius0 != 0:
        Slot_Corner_Radius01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = Tooth_Width0 * (1 + coefficient1[4])
        Slot_Depth01 = Slot_Depth0 * (1 - coefficient1[22])
        Tooth_Tip_Depth01 = Tooth_Tip_Depth0 * (1 + coefficient1[15])
        uplimit1 = ((Armature_Diameter01 - (
                Slot_Depth01 + Tooth_Tip_Depth01) * 2) * 3.14 / Slot_Number01 - Tooth_Width01 - 1) * 0.8
        lowlimit1 = 0.2
        if uplimit1 < 0.3:
            uplimit1 = 0.3
        elif uplimit1 > 20:
            uplimit1 = 20
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[16][0], limitarray[16][1])

        Slot_Corner_Radius0 = factorelementcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 1)
        coefficient1[16] = coefficientcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 0.1)
        Slot_Corner_Radius0 = factorelementcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 1)

        while Slot_Corner_Radius0 != Slot_Corner_Radius01:
            mcad1.SetVariable('Slot_Corner_Radius', Slot_Corner_Radius0)
            Slot_Corner_Radius01 = round(mcad1.GetVariable('Slot_Corner_Radius')[1], 1)

        loggenera('槽底圆角' + ' : ' + str(Slot_Corner_Radius01))
        num_list[16] = Slot_Corner_Radius0

    if Slot_Opening0 != 0:
        Slot_Opening01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = Tooth_Width0 * (1 + coefficient1[4])
        Slot_Depth01 = Slot_Depth0 * (1 + coefficient1[22])

        uplimit1 = ((Armature_Diameter01 - Slot_Depth01 * 2) * 3.14 / Slot_Number01 - Tooth_Width01) * 0.7
        lowlimit1 = 2.5
        if uplimit1 < 2.6:
            uplimit1 = 2.6
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[14][0], limitarray[14][1])

        Slot_Opening0 = factorelementcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 1)
        coefficient1[14] = coefficientcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 0.1)
        Slot_Opening0 = factorelementcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 1)

        while Slot_Opening0 != Slot_Opening01:
            mcad1.SetVariable('Slot_Opening', Slot_Opening0)
            Slot_Opening01 = round(mcad1.GetVariable('Slot_Opening')[1], 1)

        loggenera('槽口宽' + ' : ' + str(Slot_Opening01))
        num_list[14] = Slot_Opening0

    if Stator_Lam_Length0 != 0:
        Stator_Lam_Length01 = 0
        Armature_Diameter01 = Armature_Diameter0 * (1 - coefficient1[13])
        uplimit1 = Armature_Diameter01 * 100
        lowlimit1 = 5
        if uplimit1 <= lowlimit1:
            uplimit1 = 500

        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[2][0], limitarray[2][1])

        Stator_Lam_Length0 = factorelementcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 1)
        coefficient1[2] = coefficientcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 0.1)
        Stator_Lam_Length0 = factorelementcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 1)

        while Stator_Lam_Length0 != Stator_Lam_Length01:
            mcad1.SetVariable('Stator_Lam_Length', Stator_Lam_Length0)
            Stator_Lam_Length01 = round(mcad1.GetVariable('Stator_Lam_Length')[1], 1)
        num_list[2] = Stator_Lam_Length0
        mcad1.SetVariable('Magnet_Length', Stator_Lam_Length0)
        mcad1.SetVariable('Rotor_Lam_Length', Stator_Lam_Length0 + 20)

        loggenera('铁长' + ' : ' + str(Stator_Lam_Length01))
    Armature_Diameter0 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
    Axle_Dia0 = round(mcad1.GetVariable('Axle_Dia')[1], 1)
    Pole_Number0 = round(mcad1.GetVariable('Pole_Number')[1]/2, 0)
    Tooth_Width0 = round(mcad1.GetVariable('Tooth_Width')[1], 1)
    Tooth_Tip_Depth0 = round(mcad1.GetVariable('Tooth_Tip_Depth')[1], 1)
    Slot_Opening0 = round(mcad1.GetVariable('Slot_Opening')[1], 1)
    Slot_Corner_Radius0 = round(mcad1.GetVariable('Slot_Corner_Radius')[1], 1)
    Airgap0 = round(mcad1.GetVariable('Airgap')[1], 2)
    Magnet_Thickness0 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
    Magnet_Arc_ED0 = round(mcad1.GetVariable('Magnet_Arc_[ED]')[1], 0)
    Back_Iron_Thickness0 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
    Slot_Number0 = round(mcad1.GetVariable('Slot_Number')[1], 0)

    Tooth_Tip_Angle0, sumarea, lenl, lenlh, EWdg_Overhang = drawstatotroundbo(outerdia=Armature_Diameter0,
                                                                              innerdiameter=Axle_Dia0,
                                                                              polenumber=Pole_Number0,
                                                                              toothwhi=Tooth_Width0,
                                                                              slotdepth=Slot_Depth0,
                                                                              toothtipd=Tooth_Tip_Depth0,
                                                                              toothtipdm=Tooth_Tip_thickmid0,
                                                                              toothtipdb=Tooth_Tip_thickbig0,
                                                                              toothtipop=Slot_Opening0,
                                                                              slotbora=Slot_Corner_Radius0,
                                                                              slottora=slottopradiu0,
                                                                              slottipra=Tooth_Tip_Radius0,
                                                                              airgap=Airgap0,
                                                                              magnetthi=Magnet_Thickness0,
                                                                              magnetarc=Magnet_Arc_ED0,
                                                                              backironthi=Back_Iron_Thickness0,
                                                                              dexfileloca_=dexfileloca_,
                                                                              slotnumber=Slot_Number0)
    EWdg_Overhang = round(EWdg_Overhang * RequestedGrossSlotFillFactor0, 1)
    mcad1.SetVariable('EWdg_Overhang_[F]', EWdg_Overhang)
    mcad1.SetVariable('EWdg_Overhang_[R]', EWdg_Overhang)
    loggenera('绕组端部长' + ' : ' + str(EWdg_Overhang))
    Liner_Thickness = mcad1.GetVariable('Liner_Thickness')[1]
    sumareaeff = sumarea - (lenl - Liner_Thickness) * Slot_Opening0 / 2 - lenlh * Liner_Thickness
    ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)

    wirediam, wirediamco = wirediametcal(sumareaeff, RequestedGrossSlotFillFactor0, ArmatureTurnsPerCoil01)
    Wire_Diameter0, Copper_Diameter0 = round(wirediam, 3), round(wirediamco, 3)

    mcad1.SetVariable('Wire_Diameter', Wire_Diameter0)
    mcad1.SetVariable('Copper_Diameter', Copper_Diameter0)
    loggenera('裸线径' + ' : ' + str(Copper_Diameter0))
    loggenera('线径' + ' : ' + str(Wire_Diameter0))

    Tooth_Tip_Angle0 = round(Tooth_Tip_Angle0, 0)
    if Tooth_Tip_Angle0 != 0:
        Tooth_Tip_Angle01 = 0

        while Tooth_Tip_Angle0 != Tooth_Tip_Angle01:
            mcad1.SetVariable('Tooth_Tip_Angle', Tooth_Tip_Angle0)
            Tooth_Tip_Angle01 = round(mcad1.GetVariable('Tooth_Tip_Angle')[1], 0)

        loggenera('槽顶角度' + ' : ' + str(Tooth_Tip_Angle01))


def fakeslotdepth(Armature_Diameter0, Slot_Depth0, d, n):
    r = Armature_Diameter0 / 2 - Slot_Depth0
    rr = Armature_Diameter0 / 2
    Slot_Depth0tem = round(rr - r * cos(pi / n - asin(d / (2 * r))), 1)
    return Slot_Depth0tem


def lowlimit(input1, coefficient1, lowlimit1, roundnb):
    roundnbture = 0
    if roundnb == 0:
        roundnbture = 1
    elif roundnb == 1:
        roundnbture = 0.1
    elif roundnb == 2:
        roundnbture = 0.01
    elif roundnb == 3:
        roundnbture = 0.001

    if lowlimit1 < 0:
        lowlimit1 = 0
    if round(input1 * (1 - coefficient1), roundnb) <= lowlimit1:
        inputchange = round(lowlimit1 / (1 - coefficient1), roundnb)
        inputchange1 = lowlimit1 / (1 - coefficient1)
        if inputchange < inputchange1:
            inputchange = inputchange + roundnbture
    else:
        inputchange = input1
    logdebug('  lowlimit: ' + str(inputchange))
    return round(inputchange, roundnb)


def uplimit(input1, coefficient1, uplimit1, roundnb):
    roundnbture = 0
    if roundnb == 0:
        roundnbture = 1
    elif roundnb == 1:
        roundnbture = 0.1
    elif roundnb == 2:
        roundnbture = 0.01
    elif roundnb == 3:
        roundnbture = 0.001

    if round(input1 * (1 + coefficient1), roundnb) >= uplimit1:
        inputchange = round(uplimit1 / (1 + coefficient1), roundnb)
        inputchange1 = uplimit1 / (1 + coefficient1)
        if inputchange > inputchange1:
            inputchange = inputchange - roundnbture
    else:
        inputchange = input1
    logdebug('  uplimit: ' + str(inputchange))

    return round(inputchange, roundnb)


def coefficientcorrect(factorelement, coefficient, uplimit1, lowlimit1, steplimit):
    if coefficient != 0:
        coefficient1 = coefficient
        while factorelement * (1 + coefficient1) > uplimit1 or factorelement * (1 - coefficient1) < lowlimit1:
            coefficient1 = coefficient1 - 0.001
            if coefficient1 < 0.001:
                coefficient1 = 0.001
                break
        while factorelement * coefficient1 < steplimit:
            coefficient1 = coefficient1 + 0.001
            if coefficient1 > 0.999:
                coefficient1 = 0.999
                break
    else:
        coefficient1 = 0
    logdebug(coefficient1)
    return coefficient1


def factorelementcorrect(factorelement, coefficient1, uplimit1, lowlimit1, roundn):
    roundnbture = 0
    if roundn == 0:
        roundnbture = 1
    elif roundn == 1:
        roundnbture = 0.1
    elif roundn == 2:
        roundnbture = 0.01
    elif roundn == 3:
        roundnbture = 0.001
    if uplimit1 == lowlimit1:
        factorelement = uplimit1
    else:
        factorelementu = uplimit(factorelement, coefficient1, uplimit1, roundn)
        factorelementl = lowlimit(factorelement, coefficient1, lowlimit1, roundn)
        if factorelementu > uplimit1:
            factorelementu = uplimit1
        if factorelementl < lowlimit1:
            factorelementl = lowlimit1
        if factorelementl > uplimit1:
            factorelementl = uplimit1
        if factorelementu < lowlimit1:
            factorelementu = lowlimit1

        factorelement1 = (factorelementl + factorelementu) / 2
        factorelement = round((factorelementl + factorelementu) / 2, roundn)
        if factorelement < factorelement1:
            factorelement = factorelement + roundnbture
        # loggenera('  factorelement: ' + str(factorelement))
    while factorelement * (1 - coefficient1) < roundnbture:
        factorelement = factorelement + roundnbture
    logdebug(factorelement)
    return round(factorelement, roundn)


def limitdip(uplimit1, lowlimit1, lowlimitin, uplimitin):
    if lowlimitin == uplimitin:
        uplimitii = lowlimitin
        lowlimitii = uplimitii
    elif lowlimitin < uplimitin:
        if lowlimitin > lowlimit1:
            lowlimitii = lowlimitin
        else:
            lowlimitii = lowlimit1

        if uplimitin < uplimit1:
            uplimitii = uplimitin
        else:
            uplimitii = uplimit1
    else:
        uplimitii = uplimit1
        lowlimitii = lowlimit1
    logdebug((uplimitii, lowlimitii))
    return uplimitii, lowlimitii


def read_parameter(rownumber_, coefficient1=None, mcad1=None, num_list=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if num_list is None:
        num_list = [i * 0 for i in range(rownumber_)]
    if coefficient1 is None:
        coefficient1 = [0 * i1 for i1 in range(rownumber_)]
    ArmatureTurnsPerCoil0 = int(num_list[0])
    Shaft_Speed_RPM0 = float(num_list[1])
    Stator_Lam_Length0 = float(num_list[2])
    Back_Iron_Thickness0 = float(num_list[3])
    Tooth_Width0 = float(num_list[4])
    Airgap0 = float(num_list[5])
    RequestedGrossSlotFillFactor0 = float(num_list[6])
    Magnet_Br_Multiplier0 = float(num_list[7])
    Magnet_Arc_ED0 = float(num_list[8])
    PhaseAdvance0 = float(num_list[9])
    DCBusVoltage0 = float(num_list[10])
    PeakCurrent0 = float(num_list[11])
    Magnet_Thickness0 = float(num_list[12])
    Armature_Diameter0 = float(num_list[13])
    Slot_Opening0 = float(num_list[14])
    Tooth_Tip_Depth0 = float(num_list[15])
    Slot_Corner_Radius0 = float(num_list[16])
    Tooth_Tip_Angle0 = float(num_list[17])
    StatorSkew0 = float(num_list[18])
    RotorSkewAngle0 = float(num_list[19])
    StatorIronLossBuildFactor0 = float(num_list[20])
    RotorIronLossBuildFactor0 = float(num_list[21])
    Slot_Depth0 = float(num_list[22])
    fixrotoroutdiameter0 = float(num_list[23])

    if ArmatureTurnsPerCoil0 != 0:
        ArmatureTurnsPerCoil01 = 0
        uplimit1 = 300
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[0][0], limitarray[0][1])
        ArmatureTurnsPerCoil0 = factorelementcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 0)
        coefficient1[0] = coefficientcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 1)
        ArmatureTurnsPerCoil0 = factorelementcorrect(ArmatureTurnsPerCoil0, coefficient1[0], uplimit1, lowlimit1, 0)

        while ArmatureTurnsPerCoil0 != ArmatureTurnsPerCoil01:
            mcad1.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil0)
            ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)
        num_list[0] = ArmatureTurnsPerCoil0
        loggenera('匝数' + ' : ' + str(ArmatureTurnsPerCoil01))

    if Shaft_Speed_RPM0 != 0:
        Shaft_Speed_RPM01 = 0
        uplimit1 = 50000
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[1][0], limitarray[1][1])
        Shaft_Speed_RPM0 = factorelementcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 0)
        coefficient1[1] = coefficientcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 1)
        Shaft_Speed_RPM0 = factorelementcorrect(Shaft_Speed_RPM0, coefficient1[1], uplimit1, lowlimit1, 0)

        while Shaft_Speed_RPM0 != Shaft_Speed_RPM01:
            mcad1.SetVariable('Shaft_Speed_[RPM]', Shaft_Speed_RPM0)
            Shaft_Speed_RPM01 = round(mcad1.GetVariable('Shaft_Speed_[RPM]')[1], 0)
        num_list[1] = Shaft_Speed_RPM0
        loggenera('转速' + ' : ' + str(Shaft_Speed_RPM01))

    if Back_Iron_Thickness0 != 0:
        Back_Iron_Thickness01 = 0
        uplimit1 = 20
        lowlimit1 = 0.3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[3][0], limitarray[3][1])
        Back_Iron_Thickness0 = factorelementcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 1)
        coefficient1[3] = coefficientcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 0.1)
        Back_Iron_Thickness0 = factorelementcorrect(Back_Iron_Thickness0, coefficient1[3], uplimit1, lowlimit1, 1)

        while Back_Iron_Thickness0 != Back_Iron_Thickness01:
            mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
            Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
        num_list[3] = Back_Iron_Thickness0
        loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness01))

    if Airgap0 != 0:
        Airgap01 = 0
        uplimit1 = 1
        lowlimit1 = 0.4
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[5][0], limitarray[5][1])
        Airgap0 = factorelementcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 2)
        coefficient1[5] = coefficientcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 0.01)
        Airgap0 = factorelementcorrect(Airgap0, coefficient1[5], uplimit1, lowlimit1, 2)

        while Airgap0 != Airgap01:
            mcad1.SetVariable('Airgap', Airgap0)
            Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
        num_list[5] = Airgap0
        loggenera('气隙' + ' : ' + str(Airgap01))

    if RequestedGrossSlotFillFactor0 != 0:
        RequestedGrossSlotFillFactor01 = 0
        uplimit1 = 0.3
        lowlimit1 = 0.05
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[6][0], limitarray[6][1])
        RequestedGrossSlotFillFactor0 = factorelementcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1,
                                                             lowlimit1, 2)
        coefficient1[6] = coefficientcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1, lowlimit1, 0.01)
        RequestedGrossSlotFillFactor0 = factorelementcorrect(RequestedGrossSlotFillFactor0, coefficient1[6], uplimit1,
                                                             lowlimit1, 2)

        while RequestedGrossSlotFillFactor0 != RequestedGrossSlotFillFactor01:
            mcad1.SetVariable('RequestedGrossSlotFillFactor', RequestedGrossSlotFillFactor0)
            RequestedGrossSlotFillFactor01 = round(mcad1.GetVariable('RequestedGrossSlotFillFactor')[1], 2)
        num_list[6] = RequestedGrossSlotFillFactor0
        loggenera('槽满率' + ' : ' + str(RequestedGrossSlotFillFactor01))

    if Magnet_Br_Multiplier0 != 0:
        Magnet_Br_Multiplier01 = 0
        uplimit1 = 20
        lowlimit1 = 0.3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[7][0], limitarray[7][1])
        Magnet_Br_Multiplier0 = factorelementcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 1)
        coefficient1[7] = coefficientcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 0.1)
        Magnet_Br_Multiplier0 = factorelementcorrect(Magnet_Br_Multiplier0, coefficient1[7], uplimit1, lowlimit1, 1)
        while Magnet_Br_Multiplier0 != Magnet_Br_Multiplier01:
            mcad1.SetVariable('Magnet_Br_Multiplier', Magnet_Br_Multiplier0)
            Magnet_Br_Multiplier01 = round(mcad1.GetVariable('Magnet_Br_Multiplier')[1], 1)
        num_list[7] = Magnet_Br_Multiplier0
        loggenera('Br 系数' + ' : ' + str(Magnet_Br_Multiplier01))

    if Magnet_Arc_ED0 != 0:
        Magnet_Arc_ED01 = 0
        uplimit1 = 166
        lowlimit1 = 90
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[8][0], limitarray[8][1])
        Magnet_Arc_ED0 = factorelementcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 0)
        coefficient1[8] = coefficientcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 1)
        Magnet_Arc_ED0 = factorelementcorrect(Magnet_Arc_ED0, coefficient1[8], uplimit1, lowlimit1, 0)
        while Magnet_Arc_ED0 != Magnet_Arc_ED01:
            mcad1.SetVariable('Magnet_Arc_[ED]', Magnet_Arc_ED0)
            Magnet_Arc_ED01 = round(mcad1.GetVariable('Magnet_Arc_[ED]')[1], 0)
        loggenera('极弧系数' + ' : ' + str(Magnet_Arc_ED01))
        num_list[8] = Magnet_Arc_ED0

    if PhaseAdvance0 != 0:
        PhaseAdvance01 = 0
        uplimit1 = 30
        lowlimit1 = -30
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[9][0], limitarray[9][1])
        PhaseAdvance0 = factorelementcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 1)
        coefficient1[9] = coefficientcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 0.1)
        PhaseAdvance0 = factorelementcorrect(PhaseAdvance0, coefficient1[9], uplimit1, lowlimit1, 1)

        while PhaseAdvance0 != PhaseAdvance01:
            mcad1.SetVariable('PhaseAdvance', PhaseAdvance0)
            PhaseAdvance01 = round(mcad1.GetVariable('PhaseAdvance')[1], 1)
        loggenera('超前角' + ' : ' + str(PhaseAdvance01))
        num_list[9] = PhaseAdvance0

    if DCBusVoltage0 != 0:
        DCBusVoltage01 = 0
        uplimit1 = 100
        lowlimit1 = 1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[10][0], limitarray[10][1])
        DCBusVoltage0 = factorelementcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 1)
        coefficient1[10] = coefficientcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 0.1)
        DCBusVoltage0 = factorelementcorrect(DCBusVoltage0, coefficient1[10], uplimit1, lowlimit1, 1)

        while DCBusVoltage0 != DCBusVoltage01:
            mcad1.SetVariable('DCBusVoltage', DCBusVoltage0)
            DCBusVoltage01 = round(mcad1.GetVariable('DCBusVoltage')[1], 1)
        loggenera('电压' + ' : ' + str(DCBusVoltage01))
        num_list[10] = DCBusVoltage0

    if PeakCurrent0 != 0:
        PeakCurrent01 = 0
        uplimit1 = 100
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[11][0], limitarray[11][1])

        if uplimit1 < 0.2 or lowlimit1 < 0.1:
            uplimit1 = 0.2
            lowlimit1 = 0.1
        PeakCurrent0 = factorelementcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 1)
        coefficient1[11] = coefficientcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 0.1)
        PeakCurrent0 = factorelementcorrect(PeakCurrent0, coefficient1[11], uplimit1, lowlimit1, 1)
        while PeakCurrent0 != PeakCurrent01:
            mcad1.SetVariable('PeakCurrent', PeakCurrent0)
            PeakCurrent01 = round(mcad1.GetVariable('PeakCurrent')[1], 1)
        loggenera('峰值电流' + ' : ' + str(PeakCurrent01))
        num_list[11] = PeakCurrent0

    if Magnet_Thickness0 != 0:
        Magnet_Thickness01 = 0
        uplimit1 = 10
        lowlimit1 = 0.5
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[12][0], limitarray[12][1])

        Magnet_Thickness0 = factorelementcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 1)
        coefficient1[12] = coefficientcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 0.1)
        Magnet_Thickness0 = factorelementcorrect(Magnet_Thickness0, coefficient1[12], uplimit1, lowlimit1, 1)
        while Magnet_Thickness0 != Magnet_Thickness01:
            mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
            Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
        num_list[12] = Magnet_Thickness0
        loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness01))

    if Armature_Diameter0 != 0:
        Armature_Diameter01 = 0
        Axle_Dia1 = mcad1.GetVariable('Axle_Dia')[1]
        uplimit1 = 120
        lowlimit1 = Axle_Dia1 + 5
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[13][0], limitarray[13][1])
        Armature_Diameter0 = factorelementcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 1)
        coefficient1[13] = coefficientcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 0.1)
        Armature_Diameter0 = factorelementcorrect(Armature_Diameter0, coefficient1[13], uplimit1, lowlimit1, 1)

        while Armature_Diameter0 != Armature_Diameter01:
            mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
            Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
        num_list[13] = Armature_Diameter0
        loggenera('定子外径' + ' : ' + str(Armature_Diameter01))

    if fixrotoroutdiameter0 != 0:
        Axle_Dia1 = mcad1.GetVariable('Axle_Dia')[1]
        uplimit1 = 150
        lowlimit1 = Axle_Dia1 + 7
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[23][0], limitarray[23][1])
        fixrotoroutdiameter0 = factorelementcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 1)
        coefficient1[23] = coefficientcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 0.1)
        fixrotoroutdiameter0 = factorelementcorrect(fixrotoroutdiameter0, coefficient1[23], uplimit1, lowlimit1, 1)

        if Armature_Diameter0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Armature_Diameter0 = round(fixrotoroutdiameter0 - Airgap01 * 2 - Back_Iron_Thickness01 * 2 - Magnet_Thickness01 * 2, 1)
            Armature_Diameter01 = 0
            while Armature_Diameter0 != Armature_Diameter01:
                mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
                Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
            loggenera('定子外径' + ' : ' + str(Armature_Diameter0))
        if Magnet_Thickness0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Magnet_Thickness0 = round((fixrotoroutdiameter0 - Airgap01 * 2 - Back_Iron_Thickness01 * 2 - Armature_Diameter01)/2, 1)
            Magnet_Thickness01 = 0
            while Magnet_Thickness0 != Magnet_Thickness01:
                mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
                Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
            loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness0))

        if Airgap0 == 0:
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Airgap0 = round((fixrotoroutdiameter0 - Magnet_Thickness01 * 2 - Back_Iron_Thickness01 * 2 - Armature_Diameter01)/2, 2)
            Airgap01 = 0
            while Airgap0 != Airgap01:
                mcad1.SetVariable('Airgap', Airgap0)
                Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
            loggenera('气隙' + ' : ' + str(Airgap0))

        if Back_Iron_Thickness0 == 0:
            Airgap01 = mcad1.GetVariable('Airgap')[1]
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
            Back_Iron_Thickness0 = round((fixrotoroutdiameter0 - Magnet_Thickness01 * 2 - Airgap01 * 2 - Armature_Diameter01)/2, 1)
            Back_Iron_Thickness01 = 0
            while Back_Iron_Thickness0 != Back_Iron_Thickness01:
                mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
                Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
            loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness0))

        num_list[23] = fixrotoroutdiameter0

    if Tooth_Tip_Depth0 != 0:
        Tooth_Tip_Depth01 = 0
        uplimit1 = 5
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[15][0], limitarray[15][1])
        Tooth_Tip_Depth0 = factorelementcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 1)
        coefficient1[15] = coefficientcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 0.1)
        Tooth_Tip_Depth0 = factorelementcorrect(Tooth_Tip_Depth0, coefficient1[15], uplimit1, lowlimit1, 1)

        while Tooth_Tip_Depth0 != Tooth_Tip_Depth01:
            mcad1.SetVariable('Tooth_Tip_Depth', Tooth_Tip_Depth0)
            Tooth_Tip_Depth01 = round(mcad1.GetVariable('Tooth_Tip_Depth')[1], 1)
        num_list[15] = Tooth_Tip_Depth0

        loggenera('槽顶深' + ' : ' + str(Tooth_Tip_Depth01))

    if StatorSkew0 != 0:
        mcad1.SetVariable('SkewType', 1)
        mcad1.SetVariable('FluxSkewFactorCalc', True)
        StatorSkew01 = 0
        uplimit1 = 30
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[18][0], limitarray[18][1])

        StatorSkew0 = factorelementcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 1)
        coefficient1[18] = coefficientcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 0.1)
        StatorSkew0 = factorelementcorrect(StatorSkew0, coefficient1[18], uplimit1, lowlimit1, 1)

        while StatorSkew0 != StatorSkew01:
            mcad1.SetVariable('StatorSkew', StatorSkew0)
            StatorSkew01 = round(mcad1.GetVariable('StatorSkew')[1], 1)
        num_list[18] = StatorSkew0

        loggenera('定子扭角' + ' : ' + str(StatorSkew01))

    if RotorSkewAngle0 != 0:
        mcad1.SetVariable('SkewType', 2)
        mcad1.SetVariable('RotorSkewSlices', 3)
        mcad1.SetVariable('AxialSegments', 3)

        RotorSkewAngle01 = 0
        uplimit1 = 30
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[19][0], limitarray[19][1])

        RotorSkewAngle0 = factorelementcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 1)
        coefficient1[19] = coefficientcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 0.1)
        RotorSkewAngle0 = factorelementcorrect(RotorSkewAngle0, coefficient1[19], uplimit1, lowlimit1, 1)
        RotorSkewAngle02 = -RotorSkewAngle0

        while RotorSkewAngle0 != RotorSkewAngle01:
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 0, RotorSkewAngle02)
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 2, RotorSkewAngle0)
            RotorSkewAngle01 = round(mcad1.GetArrayVariable('RotorSkewAngle_Array', 2)[1], 1)

        num_list[19] = RotorSkewAngle0
        loggenera('转子扭角' + ' : ' + str(RotorSkewAngle01))

    if StatorIronLossBuildFactor0 != 0:
        StatorIronLossBuildFactor01 = 0
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[20][0], limitarray[20][1])

        StatorIronLossBuildFactor0 = factorelementcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1,
                                                          lowlimit1, 1)
        coefficient1[20] = coefficientcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1, lowlimit1, 0.1)
        StatorIronLossBuildFactor0 = factorelementcorrect(StatorIronLossBuildFactor0, coefficient1[20], uplimit1,
                                                          lowlimit1, 1)

        while StatorIronLossBuildFactor0 != StatorIronLossBuildFactor01:
            mcad1.SetVariable('StatorIronLossBuildFactor', StatorIronLossBuildFactor0)
            StatorIronLossBuildFactor01 = round(mcad1.GetVariable('StatorIronLossBuildFactor')[1], 1)
        num_list[20] = StatorIronLossBuildFactor0
        loggenera('定子铁耗系数' + ' : ' + str(StatorIronLossBuildFactor01))

    if RotorIronLossBuildFactor0 != 0:
        RotorIronLossBuildFactor01 = 0
        uplimit1 = 10
        lowlimit1 = 0.1
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[21][0], limitarray[21][1])

        RotorIronLossBuildFactor0 = factorelementcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1,
                                                         lowlimit1, 1)
        coefficient1[21] = coefficientcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1, lowlimit1, 0.1)
        RotorIronLossBuildFactor0 = factorelementcorrect(RotorIronLossBuildFactor0, coefficient1[21], uplimit1,
                                                         lowlimit1, 1)
        while RotorIronLossBuildFactor0 != RotorIronLossBuildFactor01:
            mcad1.SetVariable('RotorIronLossBuildFactor', RotorIronLossBuildFactor0)
            RotorIronLossBuildFactor01 = round(mcad1.GetVariable('RotorIronLossBuildFactor')[1], 1)
        num_list[21] = RotorIronLossBuildFactor0
        loggenera('转子铁耗系数' + ' : ' + str(RotorIronLossBuildFactor01))

    if Tooth_Width0 != 0:
        Tooth_Width01 = 0
        Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1] * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        uplimit1 = Armature_Diameter01 * 0.75 / Slot_Number01
        lowlimit1 = 2
        if uplimit1 < 3:
            uplimit1 = 3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[4][0], limitarray[4][1])

        Tooth_Width0 = factorelementcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 1)
        coefficient1[4] = coefficientcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 0.1)
        Tooth_Width0 = factorelementcorrect(Tooth_Width0, coefficient1[4], uplimit1, lowlimit1, 1)

        while Tooth_Width0 != Tooth_Width01:
            mcad1.SetVariable('Tooth_Width', Tooth_Width0)
            Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)

        loggenera('齿宽' + ' : ' + str(Tooth_Width01))
        num_list[4] = Tooth_Width0

    if Slot_Depth0 != 0:
        Tooth_Tip_Depth01 = mcad1.GetVariable('Tooth_Tip_Depth')[1] * (1 - coefficient1[15])
        Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1] * (1 + coefficient1[13])
        Axle_Dia1 = mcad1.GetVariable('Axle_Dia')[1]
        Tooth_Tip_Depth011 = mcad1.GetVariable('Tooth_Tip_Depth')[1] * (1 + coefficient1[15])
        Armature_Diameter011 = mcad1.GetVariable('Armature_Diameter')[1] * (1 - coefficient1[13])
        Axle_Dia11 = mcad1.GetVariable('Axle_Dia')[1]

        uplimit1 = ((Armature_Diameter011 - Axle_Dia11) / 2 - Tooth_Tip_Depth011) * 0.95
        lowlimit1 = ((Armature_Diameter01 - Axle_Dia1) / 2 - Tooth_Tip_Depth01) * 0.05
        if lowlimit1 <= 0:
            lowlimit1 = 3
        if uplimit1 <= 0 or uplimit1 <= lowlimit1:
            uplimit1 = lowlimit1 + 3
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[22][0], limitarray[22][1])

        Slot_Depth0 = factorelementcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 1)
        coefficient1[22] = coefficientcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 0.1)
        Slot_Depth0 = factorelementcorrect(Slot_Depth0, coefficient1[22], uplimit1, lowlimit1, 1)

        Slot_Depth01 = 0
        while Slot_Depth0 != Slot_Depth01:
            mcad1.SetVariable('Slot_Depth', Slot_Depth0)
            Slot_Depth01 = round(mcad1.GetVariable('Slot_Depth')[1], 1)

        loggenera('槽深' + ' : ' + str(Slot_Depth01))
        num_list[22] = Slot_Depth0

    if Slot_Corner_Radius0 != 0:
        Slot_Corner_Radius01 = 0
        Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1] * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = mcad1.GetVariable('Tooth_Width')[1] * (1 + coefficient1[4])
        Slot_Depth01 = mcad1.GetVariable('Slot_Depth')[1] * (1 - coefficient1[22])
        Tooth_Tip_Depth01 = mcad1.GetVariable('Tooth_Tip_Depth')[1] * (1 + coefficient1[15])
        uplimit1 = ((Armature_Diameter01 - (
                Slot_Depth01 + Tooth_Tip_Depth01) * 2) * 3.14 / Slot_Number01 - Tooth_Width01 - 1) * 0.8
        lowlimit1 = 0.2
        if uplimit1 < 0.3:
            uplimit1 = 0.3
        elif uplimit1 > 20:
            uplimit1 = 20
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[16][0], limitarray[16][1])

        Slot_Corner_Radius0 = factorelementcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 1)
        coefficient1[16] = coefficientcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 0.1)
        Slot_Corner_Radius0 = factorelementcorrect(Slot_Corner_Radius0, coefficient1[16], uplimit1, lowlimit1, 1)

        while Slot_Corner_Radius0 != Slot_Corner_Radius01:
            mcad1.SetVariable('Slot_Corner_Radius', Slot_Corner_Radius0)
            Slot_Corner_Radius01 = round(mcad1.GetVariable('Slot_Corner_Radius')[1], 1)

        loggenera('槽底圆角' + ' : ' + str(Slot_Corner_Radius01))
        num_list[16] = Slot_Corner_Radius0

    if Slot_Opening0 != 0:
        Slot_Opening01 = 0
        Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1] * (1 - coefficient1[13])
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = mcad1.GetVariable('Tooth_Width')[1] * (1 + coefficient1[4])
        Slot_Depth01 = mcad1.GetVariable('Slot_Depth')[1] * (1 + coefficient1[22])

        uplimit1 = ((Armature_Diameter01 - Slot_Depth01 * 2) * 3.14 / Slot_Number01 - Tooth_Width01) * 0.7
        lowlimit1 = 2.5
        if uplimit1 < 2.6:
            uplimit1 = 2.6
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[14][0], limitarray[14][1])

        Slot_Opening0 = factorelementcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 1)
        coefficient1[14] = coefficientcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 0.1)
        Slot_Opening0 = factorelementcorrect(Slot_Opening0, coefficient1[14], uplimit1, lowlimit1, 1)

        while Slot_Opening0 != Slot_Opening01:
            mcad1.SetVariable('Slot_Opening', Slot_Opening0)
            Slot_Opening01 = round(mcad1.GetVariable('Slot_Opening')[1], 1)

        loggenera('槽口宽' + ' : ' + str(Slot_Opening01))
        num_list[14] = Slot_Opening0

    if Tooth_Tip_Angle0 != 0:
        Tooth_Tip_Angle01 = 0
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]

        uplimit1 = 50
        lowlimit1 = 180 / Slot_Number01 - 25
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[17][0], limitarray[17][1])

        Tooth_Tip_Angle0 = factorelementcorrect(Tooth_Tip_Angle0, coefficient1[17], uplimit1, lowlimit1, 0)
        coefficient1[17] = coefficientcorrect(Tooth_Tip_Angle0, coefficient1[17], uplimit1, lowlimit1, 1)
        Tooth_Tip_Angle0 = factorelementcorrect(Tooth_Tip_Angle0, coefficient1[17], uplimit1, lowlimit1, 0)

        while Tooth_Tip_Angle0 != Tooth_Tip_Angle01:
            mcad1.SetVariable('Tooth_Tip_Angle', Tooth_Tip_Angle0)
            Tooth_Tip_Angle01 = round(mcad1.GetVariable('Tooth_Tip_Angle')[1], 0)

        loggenera('槽顶角度' + ' : ' + str(Tooth_Tip_Angle01))
        num_list[17] = Tooth_Tip_Angle0

    if Stator_Lam_Length0 != 0:
        Stator_Lam_Length01 = 0
        Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1] * (1 - coefficient1[13])
        uplimit1 = Armature_Diameter01 * 2
        lowlimit1 = 5
        if uplimit1 <= lowlimit1:
            uplimit1 = 500
        uplimit1, lowlimit1 = limitdip(uplimit1, lowlimit1, limitarray[2][0], limitarray[2][1])

        Stator_Lam_Length0 = factorelementcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 1)
        coefficient1[2] = coefficientcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 0.1)
        Stator_Lam_Length0 = factorelementcorrect(Stator_Lam_Length0, coefficient1[2], uplimit1, lowlimit1, 1)

        while Stator_Lam_Length0 != Stator_Lam_Length01:
            mcad1.SetVariable('Stator_Lam_Length', Stator_Lam_Length0)
            Stator_Lam_Length01 = round(mcad1.GetVariable('Stator_Lam_Length')[1], 1)
        num_list[2] = Stator_Lam_Length0
        mcad1.SetVariable('Magnet_Length', Stator_Lam_Length0)
        mcad1.SetVariable('Rotor_Lam_Length', Stator_Lam_Length0 + 20)
        loggenera('铁长' + ' : ' + str(Stator_Lam_Length01))

    RequestedGrossSlotFillFactor01 = round(mcad1.GetVariable('RequestedGrossSlotFillFactor')[1], 2)
    Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
    Tooth_Width01 = mcad1.GetVariable('Tooth_Width')[1]
    Slot_Opening01 = mcad1.GetVariable('Slot_Opening')[1]
    Slot_Number = mcad1.GetVariable('Slot_Number')[1]
    EWdg_Overhang = round(
        (Armature_Diameter01 * 3.14 / (
                    Slot_Number * 2) - Tooth_Width01 / 2 - Slot_Opening01 / 2) * RequestedGrossSlotFillFactor01 * 0.8 / 0.31,
        1)
    mcad1.SetVariable('EWdg_Overhang_[F]', EWdg_Overhang)
    mcad1.SetVariable('EWdg_Overhang_[R]', EWdg_Overhang)
    loggenera('绕组端部长' + ' : ' + str(EWdg_Overhang))


def openmcad(Base_File_path_):
    while 1:
        try:
            mcad1 = win32com.client.Dispatch('motorcad.appautomation')
            time.sleep(1)
            mcad1.LoadFromFile(Base_File_path_)
            return mcad1
        except (PermissionError, com_error):
            continue


def close_instances(Base_File_path_, mcad1=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    mcad1.SaveToFile(Base_File_path_)
    mcad1.Quit()  # quits each instance of Motor-CAD


def gettablename(phasecheck, rownumber_, torque, minitbookaddr_, mcad1=None, get_num_lista=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if get_num_lista is None:
        get_num_lista = [0 * i for i in range(rownumber_)]
    Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
    Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
    Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
    Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
    fixrotoroutdiameter = Armature_Diameter01 + (Magnet_Thickness01 + Back_Iron_Thickness01 + Airgap01) * 2

    IM_Pole_Number = round(mcad1.GetVariable('IM_Pole_Number')[1], 1)
    Slot_Number = round(mcad1.GetVariable('Slot_Number')[1], 1)
    DCBusVoltage = round(mcad1.GetVariable('DCBusVoltage')[1], 1)
    Shaft_Speed_RPM01 = round(mcad1.GetVariable('Shaft_Speed_[RPM]')[1], 0)
    Stator_Lam_Length01 = round(mcad1.GetVariable('Stator_Lam_Length')[1], 1)
    dbnamedist = getdata(minitbookaddr_, 17, 10)
    if dbnamedist is None:
        dbnamedist = ''
    else:
        dbnamedist = str(dbnamedist)
    if phasecheck == 3:
        if get_num_lista[2] == 0 and get_num_lista[23] == 0 and get_num_lista[1] == 0 and get_num_lista[10] == 0:
            tablename = "T%d%d%d%d%d%d%dE" % (
                IM_Pole_Number, Slot_Number, fixrotoroutdiameter, Stator_Lam_Length01, DCBusVoltage, Shaft_Speed_RPM01,
                torque) + dbnamedist
        elif get_num_lista[23] == 0 and get_num_lista[1] == 0 and get_num_lista[10] == 0:
            tablename = "T%d%d%d%d%d%dE" % (
                IM_Pole_Number, Slot_Number, fixrotoroutdiameter, DCBusVoltage, Shaft_Speed_RPM01, torque) + dbnamedist
        elif get_num_lista[1] == 0 and get_num_lista[10] == 0:
            tablename = "T%d%d%d%d%dE" % (
                IM_Pole_Number, Slot_Number, DCBusVoltage, Shaft_Speed_RPM01, torque) + dbnamedist
        elif get_num_lista[10] == 0:
            tablename = "T%d%d%d%dE" % (IM_Pole_Number, Slot_Number, DCBusVoltage, torque) + dbnamedist
        else:
            tablename = "T%d%d%dE" % (IM_Pole_Number, Slot_Number, torque) + dbnamedist
    else:
        if get_num_lista[2] == 0 and get_num_lista[23] == 0 and get_num_lista[1] == 0 and get_num_lista[10] == 0 and \
                get_num_lista[25] == 0:
            tablename = "T%d%d%d%d%d%dE" % (
                IM_Pole_Number, fixrotoroutdiameter, Stator_Lam_Length01, DCBusVoltage, Shaft_Speed_RPM01,
                torque) + dbnamedist
        elif get_num_lista[23] == 0 and get_num_lista[1] == 0 and get_num_lista[10] == 0 and get_num_lista[25] == 0:
            tablename = "T%d%d%d%d%dE" % (
                IM_Pole_Number, fixrotoroutdiameter, DCBusVoltage, Shaft_Speed_RPM01, torque) + dbnamedist
        elif get_num_lista[1] == 0 and get_num_lista[10] == 0 and get_num_lista[25] == 0:
            tablename = "T%d%d%d%dE" % (IM_Pole_Number, DCBusVoltage, Shaft_Speed_RPM01, torque) + dbnamedist
        elif get_num_lista[10] == 0 and get_num_lista[25] == 0:
            tablename = "T%d%d%dE" % (IM_Pole_Number, DCBusVoltage, torque) + dbnamedist
        elif get_num_lista[25] == 0:
            tablename = "T%d%dE" % (IM_Pole_Number, torque) + dbnamedist
        else:
            tablename = "T%dE" % torque + dbnamedist

    return tablename


def datainputhand(signal, reducesi, phasecheck, dexfileloca_):
    if phasecheck == 3:
        minitbookaddr_ = minitbookaddr
        rownumber_ = rownumber
    elif phasecheck == 1:
        minitbookaddr_ = singminitbookaddr
        rownumber_ = rownumbersingle
    else:
        minitbookaddr_ = minitbookroundboaddr
        rownumber_ = rownumbersthreerounbo
    get_num_lista = [0 * i for i in range(rownumber_)]
    mcad = openmcad(Base_File_path)

    while 1:
        try:
            inputdatahandlandcheck(dexfileloca_, minitbookaddr_, phasecheck, reducesi, rownumber_, signal,
                                   get_num_lista, mcad)
            break
        except ValueError:
            loggenera(ValueError)
            mutation(minitbookaddr_, signal, rownumber_, get_num_lista)
            continue
    close_instances(Base_File_path, mcad)


def inputdatahandlandcheck(dexfileloca_, minitbookaddr_, phasecheck, reducesi, rownumber_, signal, get_num_lista=None,
                           mcad=None):
    if get_num_lista is None:
        get_num_lista = [0 * i for i in range(rownumber_)]
    if mcad is None:
        mcad = win32com.client.Dispatch('motorcad.appautomation')
    get_num_listaorin = []
    extract_data_from_excellist(minitbookaddr_, rownumber_, columnnumber1a, 2, 2, 'Sheet', get_num_listaorin)
    logdebug(get_num_listaorin)
    global inputparalist
    inputparalist = []
    extract_data_from_excellist(minitbookaddr_, 10, 1, 2, 9, 'Sheet', inputparalist)
    global limitarray
    limitarray = numpy.zeros((rownumber_, 2))
    extract_data_from_excelarray(minitbookaddr_, rownumber_, 2, 2, 7, 'Sheet', limitarray)
    logdebug(limitarray)
    if reducesi != 1 and signal == 1:
        mutation(minitbookaddr_, signal, rownumber_, get_num_listaorin)
    logdebug(get_num_listaorin)
    if phasecheck == 3:
        get_num_lista = inputminijug(rownumber_, get_num_listaorin, get_num_lista)
    elif phasecheck == 1:
        get_num_lista = inputminijugsing(rownumber_, get_num_listaorin, get_num_lista)
    elif phasecheck == 4:
        get_num_lista = inputminijugthreeroundbo(rownumber_, get_num_listaorin, get_num_lista)
    coefficientlist = [0 * i1 for i1 in range(rownumber_)]
    quickstep = inputparalist[4]
    transtep = inputparalist[5]
    coefficient = coefficientratiochange(signal, get_num_lista[0], quickstep, transtep, minitbookaddr_) * reducesi
    for i1 in range(rownumber_):
        if get_num_lista[i1] > 0:
            coefficientlist[i1] = coefficient
            if limitarray[i1][0] == limitarray[i1][1] != 0:
                coefficientlist[i1] = 0
                get_num_lista[i1] = limitarray[i1][0]
        elif i1 == 9:
            coefficientlist[i1] = coefficient
            if limitarray[i1][0] == limitarray[i1][1] != 0:
                coefficientlist[i1] = 0
                get_num_lista[i1] = limitarray[i1][0]

    logdebug(get_num_lista)
    time.sleep(10)
    if phasecheck == 3:
        read_parameter(rownumber_, coefficientlist, mcad, get_num_lista)
    else:
        if phasecheck == 1:
            read_parameter_single(dexfileloca_, rownumber_, coefficientlist, mcad, get_num_lista)
        elif phasecheck == 4:
            read_parameter_threeroundbo(dexfileloca_, rownumber_, coefficientlist, mcad, get_num_lista)
        mcad.SetVariable('DXFFileName', dexfileloca_)
        mcad.LoadDXFFile(dexfileloca_)
        mcad.SetVariable('DXFImportType', 0)
        mcad.SetVariable('UseDXFImportForFEA_Magnetic', True)
        mcad.SaveToFile(Base_File_path)
    sqlitetablename = gettablename(phasecheck, rownumber_, inputparalist[1], minitbookaddr_, mcad, coefficientlist)
    setdata(minitbookaddr_, 18, 10, sqlitetablename)

    coefficientlist_r = [round(i, 3) for i in coefficientlist]
    num_list_generate = numpy.zeros((rownumber_, 3))
    logdebug(get_num_lista)
    if phasecheck == 3:
        listgenarat(rownumber_, num_list_generate, get_num_lista, coefficientlist_r)
    elif phasecheck == 1:
        listgenaratsingl(rownumber_, num_list_generate, get_num_lista, coefficientlist_r)
    elif phasecheck == 4:
        listgenaratthreeroundbo(rownumber_, num_list_generate, get_num_lista, coefficientlist_r)

    loggenera('输入参数列表：\n' + str(num_list_generate))
    readdatainput(minitbookaddr_, rownumber_, get_num_lista, num_list_generate)
    deltefile(dexfileloca_)
    get_num_listaorin.clear()
    coefficientlist.clear()
    coefficientlist_r.clear()


def listgenarat(rownumber_, num_list_generate=None, get_num_lista=None, coefficientlist_r=None):
    if num_list_generate is None:
        num_list_generate = numpy.zeros((rownumber_, 3))
    if get_num_lista is None:
        get_num_lista = [0 * i for i in range(rownumber_)]
    if coefficientlist_r is None:
        coefficientlist_r = [0 * i for i in range(rownumber_)]
    for i1 in range(rownumber_):
        if i1 == 0 or i1 == 1 or i1 == 17 or i1 == 8:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 0)
            num_list_generate[i1][1] = round(get_num_lista[i1], 0)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 0)
        elif i1 == 5 or i1 == 6:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 2)
            num_list_generate[i1][1] = round(get_num_lista[i1], 2)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 2)
        else:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 1)
            num_list_generate[i1][1] = round(get_num_lista[i1], 1)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 1)


def listgenaratsingl(rownumber_, num_list_generate=None, get_num_lista=None, coefficientlist_r=None):
    if num_list_generate is None:
        num_list_generate = numpy.zeros((rownumber_, 3))
    if get_num_lista is None:
        get_num_lista = [0 * i for i in range(rownumber_)]
    if coefficientlist_r is None:
        coefficientlist_r = [0 * i for i in range(rownumber_)]
    for i1 in range(rownumber_):
        if i1 == 0 or i1 == 1 or i1 == 8 or i1 == 25:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 0)
            num_list_generate[i1][1] = round(get_num_lista[i1], 0)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 0)
        elif i1 == 5 or i1 == 6 or i1 == 28 or i1 == 29 or i1 == 6:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 2)
            num_list_generate[i1][1] = round(get_num_lista[i1], 2)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 2)
        else:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 1)
            num_list_generate[i1][1] = round(get_num_lista[i1], 1)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 1)


def listgenaratthreeroundbo(rownumber_, num_list_generate=None, get_num_lista=None, coefficientlist_r=None):
    if num_list_generate is None:
        num_list_generate = numpy.zeros((rownumber_, 3))
    if get_num_lista is None:
        get_num_lista = [0 * i for i in range(rownumber_)]
    if coefficientlist_r is None:
        coefficientlist_r = [0 * i for i in range(rownumber_)]
    for i1 in range(rownumber_):
        if i1 == 0 or i1 == 1 or i1 == 8 or i1 == 25 or i1 == 29:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 0)
            num_list_generate[i1][1] = round(get_num_lista[i1], 0)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 0)
        elif i1 == 5 or i1 == 6 or i1 == 6:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 2)
            num_list_generate[i1][1] = round(get_num_lista[i1], 2)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 2)
        else:
            num_list_generate[i1][0] = round(get_num_lista[i1] - coefficientlist_r[i1]*abs(get_num_lista[i1]), 1)
            num_list_generate[i1][1] = round(get_num_lista[i1], 1)
            num_list_generate[i1][2] = round(get_num_lista[i1] + coefficientlist_r[i1]*abs(get_num_lista[i1]), 1)


def inputminijug(rownumber_, get_num_listaorin=None, get_num_lista=None):
    if get_num_listaorin is None:
        get_num_listaorin = [0 * i for i in range(rownumber_)]
    if get_num_lista is None:
        get_num_lista = [0 * i for i in range(rownumber_)]
    for i in range(rownumber_):
        if i == 0 or i == 1 or i == 17 or i == 8:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 0)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 1
            else:
                get_num_lista[i] = 0

        elif i == 5 or i == 6:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 2)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 0.01
            else:
                get_num_lista[i] = 0

        else:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 1)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 0.1
            else:
                get_num_lista[i] = 0
    return get_num_lista


def inputminijugsing(rownumber_, get_num_listaorin=None, get_num_lista=None):
    if get_num_listaorin is None:
        get_num_listaorin = [0 for _ in range(rownumber_)]
    if get_num_lista is None:
        get_num_lista = [0 for _ in range(rownumber_)]
    if get_num_listaorin[13] == 0 and get_num_listaorin[23] == 0:
        get_num_listaorin[13] = 1

    for i in range(rownumber_):
        if i == 26 or i == 27 or i == 30 or i == 17 or i == 28 or i == 29 or i == 22:
            if get_num_listaorin[i] == 0:
                get_num_listaorin[i] = 1

        if i == 0 or i == 1 or i == 8 or i == 25:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 0)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 1
            else:
                get_num_lista[i] = 0

        elif i == 5 or i == 6 or i == 28 or i == 29:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 2)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 0.01
            else:
                get_num_lista[i] = 0

        else:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 1)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 0.1
            else:
                get_num_lista[i] = 0
    return get_num_lista


def inputminijugthreeroundbo(rownumber_, get_num_listaorin=None, get_num_lista=None):
    if get_num_listaorin is None:
        get_num_listaorin = [0 for _ in range(rownumber_)]
    if get_num_lista is None:
        get_num_lista = [0 for _ in range(rownumber_)]
    if get_num_listaorin[13] == 0 and get_num_listaorin[23] == 0:
        get_num_listaorin[13] = 1

    for i in range(rownumber_):
        if i == 26 or i == 27 or i == 28 or i == 17 or i == 22:
            if get_num_listaorin[i] == 0:
                get_num_listaorin[i] = 1

        if i == 0 or i == 1 or i == 8 or i == 25 or i == 29:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 0)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 1
            else:
                get_num_lista[i] = 0

        elif i == 5 or i == 6:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 2)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 0.01
            else:
                get_num_lista[i] = 0

        else:
            if get_num_listaorin[i] > 0:
                get_num_lista[i] = round(get_num_listaorin[i], 1)
                if get_num_lista[i] == 0:
                    get_num_lista[i] = 0.1
            else:
                get_num_lista[i] = 0
    return get_num_lista


def main():
    quickedit(0)
    paradifine()
    logconfidebug(logmotorcadinputaddr)
    phasecheck = numberget(fulldatanam, phasechecktabstr)
    quickcalcheck = 0
    if phasecheck == 3 or phasecheck == 4:
        quickcalcheck = numberget(fulldatanam, quickcalcheckstr)
    if quickcalcheck == 0 or quickcalcheck == 1:
        signal = 0
    else:
        signal = 1
    reducesi = numberget(fulldatanam, motorinputreducesigetstr)
    if reducesi is None:
        reducesi = 0
    if 0 <= reducesi < 2:
        datainputhand(signal, reducesi, phasecheck, dexfileloca)
        numberinput(fulldatanam, motorinpufincheckstr, 1)


if __name__ == '__main__':
    main()
