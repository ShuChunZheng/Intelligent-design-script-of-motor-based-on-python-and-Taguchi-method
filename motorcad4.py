# import os
import time

import numpy
import win32com.client

from minitabexecute import dirgenerate, setdata, rownumbermax, rownumber, quickedit, fulldatanam, quickcalcheckstr, \
    numberget, phasechecktabstr, rownumbersingle, deltefile, loggenera, logconfi, extract_data_from_excel, \
    write_data_to_excelholexc, rownumbersthreerounbo
from motorcadinputdata import wirediametcal, drawstatot, openmcad, drawstatotroundbo, fakeslotdepth

columnnumber2 = 11
sequence = 4
global minitbookaddr
global MotorCAD_File
global pyseworkbookaddr
global get_num_list
global singminitbookaddr
global dexfilelocalist
global logmotorcadaddrlist
global minitbookroundboaddr


def paradifine():
    global minitbookaddr
    global singminitbookaddr
    global dexfilelocalist
    global MotorCAD_File
    global pyseworkbookaddr
    global logmotorcadaddrlist
    global minitbookroundboaddr
    diclist = dirgenerate()
    minitbookaddr = diclist['minitbookaddr']
    singminitbookaddr = diclist['singminitbookaddr']
    dexfileloca1 = diclist['dexfileloca1']
    dexfileloca2 = diclist['dexfileloca2']
    dexfileloca3 = diclist['dexfileloca3']
    dexfileloca4 = diclist['dexfileloca4']
    dexfilelocalist = [dexfileloca1, dexfileloca2, dexfileloca3, dexfileloca4]
    MotorCAD_File1 = diclist['MotorCAD_File1']
    pyseworkbookaddr1 = diclist['pyseworkbookaddr1']
    MotorCAD_File2 = diclist['MotorCAD_File2']
    pyseworkbookaddr2 = diclist['pyseworkbookaddr2']
    MotorCAD_File3 = diclist['MotorCAD_File3']
    pyseworkbookaddr3 = diclist['pyseworkbookaddr3']
    MotorCAD_File4 = diclist['MotorCAD_File4']
    pyseworkbookaddr4 = diclist['pyseworkbookaddr4']
    MotorCAD_Filelist = [MotorCAD_File1, MotorCAD_File2, MotorCAD_File3, MotorCAD_File4]
    pyseworkbookaddrlist = [pyseworkbookaddr1, pyseworkbookaddr2, pyseworkbookaddr3, pyseworkbookaddr4]
    MotorCAD_File = MotorCAD_Filelist[sequence - 1]
    pyseworkbookaddr = pyseworkbookaddrlist[sequence - 1]
    logmotorcad1addr = diclist['logmotorcad1addr']
    logmotorcad2addr = diclist['logmotorcad2addr']
    logmotorcad3addr = diclist['logmotorcad3addr']
    logmotorcad4addr = diclist['logmotorcad4addr']
    logmotorcadaddrlist = [logmotorcad1addr, logmotorcad2addr, logmotorcad3addr, logmotorcad4addr]
    minitbookroundboaddr = diclist['minitbookroundboaddr']


def read_parameter(rownumber0, columnnumber0, rownumberture, mcad1=None, num_list=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if num_list is None:
        num_list = numpy.zeros((rownumberture, columnnumber0))
    ArmatureTurnsPerCoil0 = int(num_list[rownumber0][0])
    Shaft_Speed_RPM0 = float(num_list[rownumber0][1])
    Stator_Lam_Length0 = float(num_list[rownumber0][2])
    Back_Iron_Thickness0 = float(num_list[rownumber0][3])
    Tooth_Width0 = float(num_list[rownumber0][4])
    Airgap0 = float(num_list[rownumber0][5])
    RequestedGrossSlotFillFactor0 = float(num_list[rownumber0][6])
    Magnet_Br_Multiplier0 = float(num_list[rownumber0][7])
    Magnet_Arc_ED0 = float(num_list[rownumber0][8])
    PhaseAdvance0 = float(num_list[rownumber0][9])
    DCBusVoltage0 = float(num_list[rownumber0][10])
    PeakCurrent0 = float(num_list[rownumber0][11])
    Magnet_Thickness0 = float(num_list[rownumber0][12])
    Armature_Diameter0 = float(num_list[rownumber0][13])
    Slot_Opening0 = float(num_list[rownumber0][14])
    Tooth_Tip_Depth0 = float(num_list[rownumber0][15])
    Slot_Corner_Radius0 = float(num_list[rownumber0][16])
    Tooth_Tip_Angle0 = float(num_list[rownumber0][17])
    StatorSkew0 = float(num_list[rownumber0][18])
    RotorSkewAngle0 = float(num_list[rownumber0][19])
    StatorIronLossBuildFactor0 = float(num_list[rownumber0][20])
    RotorIronLossBuildFactor0 = float(num_list[rownumber0][21])
    Slot_Depth0 = float(num_list[rownumber0][22])
    fixrotoroutdiameter0 = float(num_list[rownumber0][23])

    if ArmatureTurnsPerCoil0 != 0:
        ArmatureTurnsPerCoil01 = 0

        while ArmatureTurnsPerCoil0 != ArmatureTurnsPerCoil01:
            mcad1.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil0)
            ArmatureTurnsPerCoil01 = mcad1.GetVariable('MagTurnsConductor')[1]

        loggenera('匝数' + ' : ' + str(ArmatureTurnsPerCoil01))

    if Shaft_Speed_RPM0 != 0:
        Shaft_Speed_RPM01 = 0
        while Shaft_Speed_RPM0 != Shaft_Speed_RPM01:
            mcad1.SetVariable('Shaft_Speed_[RPM]', Shaft_Speed_RPM0)
            Shaft_Speed_RPM01 = mcad1.GetVariable('Shaft_Speed_[RPM]')[1]

        loggenera('转速' + ' : ' + str(Shaft_Speed_RPM01))

    if Stator_Lam_Length0 != 0:
        Stator_Lam_Length01 = 0
        while Stator_Lam_Length0 != Stator_Lam_Length01:
            mcad1.SetVariable('Stator_Lam_Length', Stator_Lam_Length0)
            Stator_Lam_Length01 = mcad1.GetVariable('Stator_Lam_Length')[1]
            mcad1.SetVariable('Magnet_Length', Stator_Lam_Length0)
            mcad1.SetVariable('Rotor_Lam_Length', Stator_Lam_Length0 + 20)
        loggenera('铁长' + ' : ' + str(Stator_Lam_Length01))

    if Back_Iron_Thickness0 != 0:
        Back_Iron_Thickness01 = 0
        while Back_Iron_Thickness0 != Back_Iron_Thickness01:
            mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
            Back_Iron_Thickness01 = mcad1.GetVariable('Back_Iron_Thickness')[1]

        loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness01))

    if Airgap0 != 0:
        Airgap01 = 0
        while Airgap0 != Airgap01:
            mcad1.SetVariable('Airgap', Airgap0)
            Airgap01 = mcad1.GetVariable('Airgap')[1]

        loggenera('气隙' + ' : ' + str(Airgap01))

    if RequestedGrossSlotFillFactor0 != 0:
        RequestedGrossSlotFillFactor01 = 0
        while RequestedGrossSlotFillFactor0 != RequestedGrossSlotFillFactor01:
            mcad1.SetVariable('RequestedGrossSlotFillFactor', RequestedGrossSlotFillFactor0)
            RequestedGrossSlotFillFactor01 = mcad1.GetVariable('RequestedGrossSlotFillFactor')[1]

        loggenera('槽满率' + ' : ' + str(RequestedGrossSlotFillFactor01))

    if Magnet_Br_Multiplier0 != 0:
        Magnet_Br_Multiplier01 = 0
        while Magnet_Br_Multiplier0 != Magnet_Br_Multiplier01:
            mcad1.SetVariable('Magnet_Br_Multiplier', Magnet_Br_Multiplier0)
            Magnet_Br_Multiplier01 = mcad1.GetVariable('Magnet_Br_Multiplier')[1]

        loggenera('Br 系数' + ' : ' + str(Magnet_Br_Multiplier0))

    if Magnet_Arc_ED0 != 0:
        Magnet_Arc_ED01 = 0
        while Magnet_Arc_ED0 != Magnet_Arc_ED01:
            mcad1.SetVariable('Magnet_Arc_[ED]', Magnet_Arc_ED0)
            Magnet_Arc_ED01 = mcad1.GetVariable('Magnet_Arc_[ED]')[1]
        loggenera('极弧系数' + ' : ' + str(Magnet_Arc_ED0))

    if PhaseAdvance0 != 0:
        PhaseAdvance01 = 0
        while PhaseAdvance0 != PhaseAdvance01:
            mcad1.SetVariable('PhaseAdvance', PhaseAdvance0)
            PhaseAdvance01 = mcad1.GetVariable('PhaseAdvance')[1]
        loggenera('超前角' + ' : ' + str(PhaseAdvance01))

    if DCBusVoltage0 != 0:
        DCBusVoltage01 = 0
        while DCBusVoltage0 != DCBusVoltage01:
            mcad1.SetVariable('DCBusVoltage', DCBusVoltage0)
            DCBusVoltage01 = mcad1.GetVariable('DCBusVoltage')[1]
        loggenera('电压' + ' : ' + str(DCBusVoltage01))

    if PeakCurrent0 != 0:
        PeakCurrent01 = 0
        while PeakCurrent0 != PeakCurrent01:
            mcad1.SetVariable('PeakCurrent', PeakCurrent0)
            PeakCurrent01 = mcad1.GetVariable('PeakCurrent')[1]
        loggenera('峰值电流' + ' : ' + str(PeakCurrent01))

    if Magnet_Thickness0 != 0:
        Magnet_Thickness01 = 0
        while Magnet_Thickness0 != Magnet_Thickness01:
            mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
            Magnet_Thickness01 = mcad1.GetVariable('Magnet_Thickness')[1]
        loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness01))

    if Armature_Diameter0 != 0:
        Armature_Diameter01 = 0

        while Armature_Diameter0 != Armature_Diameter01:
            mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
            Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
        loggenera('定子外径' + ' : ' + str(Armature_Diameter01))

    if fixrotoroutdiameter0 != 0:
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

    if Slot_Opening0 != 0:
        Slot_Opening01 = 0

        while Slot_Opening0 != Slot_Opening01:
            mcad1.SetVariable('Slot_Opening', Slot_Opening0)
            Slot_Opening01 = mcad1.GetVariable('Slot_Opening')[1]
        loggenera('槽口宽' + ' : ' + str(Slot_Opening01))

    if Tooth_Tip_Depth0 != 0:
        Tooth_Tip_Depth01 = 0

        while Tooth_Tip_Depth0 != Tooth_Tip_Depth01:
            mcad1.SetVariable('Tooth_Tip_Depth', Tooth_Tip_Depth0)
            Tooth_Tip_Depth01 = mcad1.GetVariable('Tooth_Tip_Depth')[1]
        loggenera('槽顶深' + ' : ' + str(Tooth_Tip_Depth01))

    if Slot_Corner_Radius0 != 0:
        Slot_Corner_Radius01 = 0

        while Slot_Corner_Radius0 != Slot_Corner_Radius01:
            mcad1.SetVariable('Slot_Corner_Radius', Slot_Corner_Radius0)
            Slot_Corner_Radius01 = mcad1.GetVariable('Slot_Corner_Radius')[1]
        loggenera('槽底圆角' + ' : ' + str(Slot_Corner_Radius01))

    if Tooth_Tip_Angle0 != 0:
        Tooth_Tip_Angle01 = 0

        while Tooth_Tip_Angle0 != Tooth_Tip_Angle01:
            mcad1.SetVariable('Tooth_Tip_Angle', Tooth_Tip_Angle0)
            Tooth_Tip_Angle01 = mcad1.GetVariable('Tooth_Tip_Angle')[1]
        loggenera('槽顶角度' + ' : ' + str(Tooth_Tip_Angle01))

    if StatorSkew0 != 0:
        StatorSkew01 = 0

        while StatorSkew0 != StatorSkew01:
            mcad1.SetVariable('StatorSkew', StatorSkew0)
            StatorSkew01 = mcad1.GetVariable('StatorSkew')[1]
        loggenera('定子扭角' + ' : ' + str(StatorSkew01))

    if RotorSkewAngle0 != 0:
        RotorSkewAngle01 = 0
        RotorSkewAngle02 = -RotorSkewAngle0

        while RotorSkewAngle0 != RotorSkewAngle01:
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 0, RotorSkewAngle02)
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 2, RotorSkewAngle0)
            RotorSkewAngle01 = mcad1.GetArrayVariable('RotorSkewAngle_Array', 2)[1]

        loggenera('转子扭角' + ' : ' + str(RotorSkewAngle01))

    if StatorIronLossBuildFactor0 != 0:
        StatorIronLossBuildFactor01 = 0
        while StatorIronLossBuildFactor0 != StatorIronLossBuildFactor01:
            mcad1.SetVariable('StatorIronLossBuildFactor', StatorIronLossBuildFactor0)
            StatorIronLossBuildFactor01 = mcad1.GetVariable('StatorIronLossBuildFactor')[1]

    if RotorIronLossBuildFactor0 != 0:
        RotorIronLossBuildFactor01 = 0
        while RotorIronLossBuildFactor0 != RotorIronLossBuildFactor01:
            mcad1.SetVariable('RotorIronLossBuildFactor', RotorIronLossBuildFactor0)
            RotorIronLossBuildFactor01 = mcad1.GetVariable('RotorIronLossBuildFactor')[1]
        loggenera('转子铁耗系数' + ' : ' + str(RotorIronLossBuildFactor01))

    if Tooth_Width0 != 0:
        Tooth_Width01 = 0

        while Tooth_Width0 != Tooth_Width01:
            mcad1.SetVariable('Tooth_Width', Tooth_Width0)
            Tooth_Width01 = mcad1.GetVariable('Tooth_Width')[1]
        loggenera('齿宽' + ' : ' + str(Tooth_Width01))

    if Slot_Depth0 != 0:
        Slot_Depth01 = 0
        while Slot_Depth0 != Slot_Depth01:
            mcad1.SetVariable('Slot_Depth', Slot_Depth0)
            Slot_Depth01 = mcad1.GetVariable('Slot_Depth')[1]
        loggenera('槽深' + ' : ' + str(Slot_Depth01))

    RequestedGrossSlotFillFactor01 = round(mcad1.GetVariable('RequestedGrossSlotFillFactor')[1], 2)
    Armature_Diameter01 = mcad1.GetVariable('Armature_Diameter')[1]
    Tooth_Width01 = mcad1.GetVariable('Tooth_Width')[1]
    Slot_Opening01 = mcad1.GetVariable('Slot_Opening')[1]
    Slot_Number = mcad1.GetVariable('Slot_Number')[1]
    EWdg_Overhang = round(
        (Armature_Diameter01 * 3.14 / (Slot_Number * 2) - Tooth_Width01 / 2 - Slot_Opening01 / 2) * RequestedGrossSlotFillFactor01 * 0.8/0.31, 1)
    mcad1.SetVariable('EWdg_Overhang_[F]', EWdg_Overhang)
    mcad1.SetVariable('EWdg_Overhang_[R]', EWdg_Overhang)
    loggenera('绕组端部长' + ' : ' + str(EWdg_Overhang))


def read_parameter_threeroundbo(rownumber0, rownumber_, rownumberture, dexfileloca_, mcad1=None, num_list=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if num_list is None:
        num_list = numpy.zeros((rownumberture, rownumber_))
    ArmatureTurnsPerCoil0 = int(num_list[rownumber0][0])
    Shaft_Speed_RPM0 = float(num_list[rownumber0][1])
    Stator_Lam_Length0 = float(num_list[rownumber0][2])
    Back_Iron_Thickness0 = float(num_list[rownumber0][3])
    Tooth_Width0 = float(num_list[rownumber0][4])
    Airgap0 = float(num_list[rownumber0][5])
    RequestedGrossSlotFillFactor0 = float(num_list[rownumber0][6])
    Magnet_Br_Multiplier0 = float(num_list[rownumber0][7])
    Magnet_Arc_ED0 = float(num_list[rownumber0][8])
    PhaseAdvance0 = float(num_list[rownumber0][9])
    DCBusVoltage0 = float(num_list[rownumber0][10])
    PeakCurrent0 = float(num_list[rownumber0][11])
    Magnet_Thickness0 = float(num_list[rownumber0][12])
    Armature_Diameter0 = float(num_list[rownumber0][13])
    Slot_Opening0 = float(num_list[rownumber0][14])
    Tooth_Tip_Depth0 = float(num_list[rownumber0][15])
    Slot_Corner_Radius0 = float(num_list[rownumber0][16])
    Tooth_Tip_Radius0 = float(num_list[rownumber0][17])
    StatorSkew0 = float(num_list[rownumber0][18])
    RotorSkewAngle0 = float(num_list[rownumber0][19])
    StatorIronLossBuildFactor0 = float(num_list[rownumber0][20])
    RotorIronLossBuildFactor0 = float(num_list[rownumber0][21])
    Slot_Depth0 = float(num_list[rownumber0][22])
    fixrotoroutdiameter0 = float(num_list[rownumber0][23])
    Axle_Dia0 = float(num_list[rownumber0][24])
    Pole_Number0 = float(num_list[rownumber0][25])
    Tooth_Tip_thickmid0 = float(num_list[rownumber0][26])
    Tooth_Tip_thickbig0 = float(num_list[rownumber0][27])
    slottopradiu0 = float(num_list[rownumber0][28])
    slotnumbm30 = float(num_list[rownumber0][29])

    if Axle_Dia0 != 0:
        Axle_Dia01 = 0
        while Axle_Dia0 != Axle_Dia01:
            mcad1.SetVariable('Axle_Dia', Axle_Dia0)
            Axle_Dia01 = round(mcad1.GetVariable('Axle_Dia')[1], 1)
        loggenera('定子内径' + ' : ' + str(Axle_Dia0))

    if Pole_Number0 != 0:
        Pole_Number01 = 0
        while Pole_Number0 * 2 != Pole_Number01:
            mcad1.SetVariable('Pole_Number', Pole_Number0 * 2)
            Pole_Number01 = round(mcad1.GetVariable('Pole_Number')[1], 0)
        loggenera('极对数' + ' : ' + str(Pole_Number0))

    if slotnumbm30 != 0:
        slotnumbm31 = 0
        slotnumb = slotnumbm30 * 3
        while slotnumb != slotnumbm31:
            mcad1.SetVariable('Slot_Number', slotnumb)
            slotnumbm31 = round(mcad1.GetVariable('Slot_Number')[1], 0)
        loggenera('槽数/3' + ' : ' + str(slotnumbm30))

    if ArmatureTurnsPerCoil0 != 0:
        ArmatureTurnsPerCoil01 = 0

        while ArmatureTurnsPerCoil0 != ArmatureTurnsPerCoil01:
            mcad1.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil0)
            ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)
        loggenera('匝数' + ' : ' + str(ArmatureTurnsPerCoil01))

    if Shaft_Speed_RPM0 != 0:
        Shaft_Speed_RPM01 = 0
        while Shaft_Speed_RPM0 != Shaft_Speed_RPM01:
            mcad1.SetVariable('Shaft_Speed_[RPM]', Shaft_Speed_RPM0)
            Shaft_Speed_RPM01 = round(mcad1.GetVariable('Shaft_Speed_[RPM]')[1], 0)
        loggenera('转速' + ' : ' + str(Shaft_Speed_RPM01))

    if Back_Iron_Thickness0 != 0:
        Back_Iron_Thickness01 = 0
        while Back_Iron_Thickness0 != Back_Iron_Thickness01:
            mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
            Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
        loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness01))

    if Airgap0 != 0:
        Airgap01 = 0
        while Airgap0 != Airgap01:
            mcad1.SetVariable('Airgap', Airgap0)
            Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
        loggenera('气隙' + ' : ' + str(Airgap01))

    if Magnet_Br_Multiplier0 != 0:
        Magnet_Br_Multiplier01 = 0
        while Magnet_Br_Multiplier0 != Magnet_Br_Multiplier01:
            mcad1.SetVariable('Magnet_Br_Multiplier', Magnet_Br_Multiplier0)
            Magnet_Br_Multiplier01 = round(mcad1.GetVariable('Magnet_Br_Multiplier')[1], 1)
        loggenera('Br 系数' + ' : ' + str(Magnet_Br_Multiplier01))

    if Magnet_Arc_ED0 != 0:
        Magnet_Arc_ED01 = 0
        while Magnet_Arc_ED0 != Magnet_Arc_ED01:
            mcad1.SetVariable('Magnet_Arc_[ED]', Magnet_Arc_ED0)
            Magnet_Arc_ED01 = round(mcad1.GetVariable('Magnet_Arc_[ED]')[1], 0)
        loggenera('极弧系数' + ' : ' + str(Magnet_Arc_ED01))

    if PhaseAdvance0 != 0:
        PhaseAdvance01 = 0
        while PhaseAdvance0 != PhaseAdvance01:
            mcad1.SetVariable('PhaseAdvance', PhaseAdvance0)
            PhaseAdvance01 = round(mcad1.GetVariable('PhaseAdvance')[1], 1)
        loggenera('超前角' + ' : ' + str(PhaseAdvance01))

    if DCBusVoltage0 != 0:
        DCBusVoltage01 = 0
        while DCBusVoltage0 != DCBusVoltage01:
            mcad1.SetVariable('DCBusVoltage', DCBusVoltage0)
            DCBusVoltage01 = round(mcad1.GetVariable('DCBusVoltage')[1], 1)
        loggenera('电压' + ' : ' + str(DCBusVoltage01))

    if PeakCurrent0 != 0:
        PeakCurrent01 = 0
        while PeakCurrent0 != PeakCurrent01:
            mcad1.SetVariable('PeakCurrent', PeakCurrent0)
            PeakCurrent01 = round(mcad1.GetVariable('PeakCurrent')[1], 1)
        loggenera('峰值电流' + ' : ' + str(PeakCurrent01))

    if Magnet_Thickness0 != 0:
        Magnet_Thickness01 = 0

        while Magnet_Thickness0 != Magnet_Thickness01:
            mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
            Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
        loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness01))

    if StatorSkew0 != 0:
        StatorSkew01 = 0
        while StatorSkew0 != StatorSkew01:
            mcad1.SetVariable('StatorSkew', StatorSkew0)
            StatorSkew01 = round(mcad1.GetVariable('StatorSkew')[1], 1)

        loggenera('定子扭角' + ' : ' + str(StatorSkew01))

    if RotorSkewAngle0 != 0:

        RotorSkewAngle01 = 0
        RotorSkewAngle02 = -RotorSkewAngle0

        while RotorSkewAngle0 != RotorSkewAngle01:
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 0, RotorSkewAngle02)
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 2, RotorSkewAngle0)
            RotorSkewAngle01 = round(mcad1.GetArrayVariable('RotorSkewAngle_Array', 2)[1], 1)

        loggenera('转子扭角' + ' : ' + str(RotorSkewAngle01))

    if StatorIronLossBuildFactor0 != 0:
        StatorIronLossBuildFactor01 = 0
        while StatorIronLossBuildFactor0 != StatorIronLossBuildFactor01:
            mcad1.SetVariable('StatorIronLossBuildFactor', StatorIronLossBuildFactor0)
            StatorIronLossBuildFactor01 = round(mcad1.GetVariable('StatorIronLossBuildFactor')[1], 1)
        loggenera('定子铁耗系数' + ' : ' + str(StatorIronLossBuildFactor01))

    if RotorIronLossBuildFactor0 != 0:
        RotorIronLossBuildFactor01 = 0

        while RotorIronLossBuildFactor0 != RotorIronLossBuildFactor01:
            mcad1.SetVariable('RotorIronLossBuildFactor', RotorIronLossBuildFactor0)
            RotorIronLossBuildFactor01 = round(mcad1.GetVariable('RotorIronLossBuildFactor')[1], 1)
        loggenera('转子铁耗系数' + ' : ' + str(RotorIronLossBuildFactor01))

    if Armature_Diameter0 != 0:
        Armature_Diameter01 = 0
        while Armature_Diameter0 != Armature_Diameter01:
            mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
            Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
        loggenera('定子外径' + ' : ' + str(Armature_Diameter01))

    if fixrotoroutdiameter0 != 0:
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

    if Tooth_Tip_Depth0 != 0:
        Tooth_Tip_Depth01 = 0
        while Tooth_Tip_Depth0 != Tooth_Tip_Depth01:
            mcad1.SetVariable('Tooth_Tip_Depth', Tooth_Tip_Depth0)
            Tooth_Tip_Depth01 = round(mcad1.GetVariable('Tooth_Tip_Depth')[1], 1)

        loggenera('槽顶深' + ' : ' + str(Tooth_Tip_Depth01))

    if Tooth_Width0 != 0:
        Tooth_Width01 = 0
        while Tooth_Width0 != Tooth_Width01:
            mcad1.SetVariable('Tooth_Width', Tooth_Width0)
            Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)

        loggenera('齿宽' + ' : ' + str(Tooth_Width01))

    if Slot_Depth0 != 0:
        Slot_Depth01 = 0
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)
        Slot_Depth0tem = fakeslotdepth(Armature_Diameter0, Slot_Depth0, Tooth_Width01, Slot_Number01)
        while Slot_Depth0tem != Slot_Depth01:
            mcad1.SetVariable('Slot_Depth', Slot_Depth0tem)
            Slot_Depth01 = round(mcad1.GetVariable('Slot_Depth')[1], 1)

        loggenera('槽深' + ' : ' + str(Slot_Depth0))

    if Slot_Corner_Radius0 != 0:
        Slot_Corner_Radius01 = 0
        while Slot_Corner_Radius0 != Slot_Corner_Radius01:
            mcad1.SetVariable('Slot_Corner_Radius', Slot_Corner_Radius0)
            Slot_Corner_Radius01 = round(mcad1.GetVariable('Slot_Corner_Radius')[1], 1)

        loggenera('槽底圆角' + ' : ' + str(Slot_Corner_Radius01))

    if Slot_Opening0 != 0:
        Slot_Opening01 = 0
        while Slot_Opening0 != Slot_Opening01:
            mcad1.SetVariable('Slot_Opening', Slot_Opening0)
            Slot_Opening01 = round(mcad1.GetVariable('Slot_Opening')[1], 1)

        loggenera('槽口宽' + ' : ' + str(Slot_Opening01))

    if Stator_Lam_Length0 != 0:
        Stator_Lam_Length01 = 0
        while Stator_Lam_Length0 != Stator_Lam_Length01:
            mcad1.SetVariable('Stator_Lam_Length', Stator_Lam_Length0)
            Stator_Lam_Length01 = round(mcad1.GetVariable('Stator_Lam_Length')[1], 1)
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

    EWdg_Overhang = round(EWdg_Overhang*RequestedGrossSlotFillFactor0, 1)
    mcad1.SetVariable('EWdg_Overhang_[F]', EWdg_Overhang)
    mcad1.SetVariable('EWdg_Overhang_[R]', EWdg_Overhang)
    loggenera('绕组端部长' + ' : ' + str(EWdg_Overhang))
    Liner_Thickness = mcad1.GetVariable('Liner_Thickness')[1]
    sumareaeff = sumarea - (lenl - Liner_Thickness) * Slot_Opening0 / 2 - lenlh * Liner_Thickness
    ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)

    wirediam, wirediamco = wirediametcal(sumareaeff, RequestedGrossSlotFillFactor0, ArmatureTurnsPerCoil01)
    Wire_Diameter0, Copper_Diameter0 = round(wirediam, 3), round(wirediamco, 3)
    Wire_Diameter01 = 0
    while Wire_Diameter0 != Wire_Diameter01:
        mcad1.SetVariable('Wire_Diameter', Wire_Diameter0)
        Wire_Diameter01 = round(mcad1.GetVariable('Wire_Diameter')[1], 3)
    loggenera('线径' + ' : ' + str(Wire_Diameter0))

    Copper_Diameter01 = 0
    while Copper_Diameter0 != Copper_Diameter01:
        mcad1.SetVariable('Copper_Diameter', Copper_Diameter0)
        Copper_Diameter01 = round(mcad1.GetVariable('Copper_Diameter')[1], 3)
    loggenera('裸线径' + ' : ' + str(Copper_Diameter0))

    Tooth_Tip_Angle0 = round(Tooth_Tip_Angle0, 0)
    if Tooth_Tip_Angle0 != 0:
        Tooth_Tip_Angle01 = 0

        while Tooth_Tip_Angle0 != Tooth_Tip_Angle01:
            mcad1.SetVariable('Tooth_Tip_Angle', Tooth_Tip_Angle0)
            Tooth_Tip_Angle01 = round(mcad1.GetVariable('Tooth_Tip_Angle')[1], 0)

        loggenera('槽顶角度' + ' : ' + str(Tooth_Tip_Angle01))


def read_parameter_single(rownumber0, rownumber_, rownumberture, dexfileloca_, mcad1=None, num_list=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if num_list is None:
        num_list = numpy.zeros((rownumberture, rownumber_))
    ArmatureTurnsPerCoil0 = int(num_list[rownumber0][0])
    Shaft_Speed_RPM0 = float(num_list[rownumber0][1])
    Stator_Lam_Length0 = float(num_list[rownumber0][2])
    Back_Iron_Thickness0 = float(num_list[rownumber0][3])
    Tooth_Width0 = float(num_list[rownumber0][4])
    Airgap0 = float(num_list[rownumber0][5])
    RequestedGrossSlotFillFactor0 = float(num_list[rownumber0][6])
    Magnet_Br_Multiplier0 = float(num_list[rownumber0][7])
    Magnet_Arc_ED0 = float(num_list[rownumber0][8])
    PhaseAdvance0 = float(num_list[rownumber0][9])
    DCBusVoltage0 = float(num_list[rownumber0][10])
    PeakCurrent0 = float(num_list[rownumber0][11])
    Magnet_Thickness0 = float(num_list[rownumber0][12])
    Armature_Diameter0 = float(num_list[rownumber0][13])
    Slot_Opening0 = float(num_list[rownumber0][14])
    Tooth_Tip_Depth0 = float(num_list[rownumber0][15])
    Slot_Corner_Radius0 = float(num_list[rownumber0][16])
    Tooth_Tip_Radius0 = float(num_list[rownumber0][17])
    StatorSkew0 = float(num_list[rownumber0][18])
    RotorSkewAngle0 = float(num_list[rownumber0][19])
    StatorIronLossBuildFactor0 = float(num_list[rownumber0][20])
    RotorIronLossBuildFactor0 = float(num_list[rownumber0][21])
    Slot_Depth0 = float(num_list[rownumber0][22])
    fixrotoroutdiameter0 = float(num_list[rownumber0][23])
    Axle_Dia0 = float(num_list[rownumber0][24])
    Pole_Number0 = float(num_list[rownumber0][25])
    Tooth_Tip_thickmid0 = float(num_list[rownumber0][26])
    Tooth_Tip_thickbig0 = float(num_list[rownumber0][27])
    outdiametersinkbig0 = float(num_list[rownumber0][28])
    outdiametersinkmid0 = float(num_list[rownumber0][29])
    slottopradiu0 = float(num_list[rownumber0][30])

    if Axle_Dia0 != 0:
        Axle_Dia01 = 0
        while Axle_Dia0 != Axle_Dia01:
            mcad1.SetVariable('Axle_Dia', Axle_Dia0)
            Axle_Dia01 = round(mcad1.GetVariable('Axle_Dia')[1], 1)
        loggenera('定子内径' + ' : ' + str(Axle_Dia0))

    if Pole_Number0 != 0:
        Pole_Number01 = 0
        while Pole_Number0 * 2 != Pole_Number01:
            mcad1.SetVariable('Pole_Number', Pole_Number0 * 2)
            mcad1.SetVariable('Slot_Number', Pole_Number0 * 2)
            Pole_Number01 = round(mcad1.GetVariable('Pole_Number')[1], 0)
        loggenera('极对数' + ' : ' + str(Pole_Number0))

    if ArmatureTurnsPerCoil0 != 0:
        ArmatureTurnsPerCoil01 = 0

        while ArmatureTurnsPerCoil0 != ArmatureTurnsPerCoil01:
            mcad1.SetVariable('MagTurnsConductor', ArmatureTurnsPerCoil0)
            ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)
        loggenera('匝数' + ' : ' + str(ArmatureTurnsPerCoil01))

    if Shaft_Speed_RPM0 != 0:
        Shaft_Speed_RPM01 = 0
        while Shaft_Speed_RPM0 != Shaft_Speed_RPM01:
            mcad1.SetVariable('Shaft_Speed_[RPM]', Shaft_Speed_RPM0)
            Shaft_Speed_RPM01 = round(mcad1.GetVariable('Shaft_Speed_[RPM]')[1], 0)
        loggenera('转速' + ' : ' + str(Shaft_Speed_RPM01))

    if Back_Iron_Thickness0 != 0:
        Back_Iron_Thickness01 = 0
        while Back_Iron_Thickness0 != Back_Iron_Thickness01:
            mcad1.SetVariable('Back_Iron_Thickness', Back_Iron_Thickness0)
            Back_Iron_Thickness01 = round(mcad1.GetVariable('Back_Iron_Thickness')[1], 1)
        loggenera('铁壳厚' + ' : ' + str(Back_Iron_Thickness01))

    if Airgap0 != 0:
        Airgap01 = 0
        while Airgap0 != Airgap01:
            mcad1.SetVariable('Airgap', Airgap0)
            Airgap01 = round(mcad1.GetVariable('Airgap')[1], 2)
        loggenera('气隙' + ' : ' + str(Airgap01))

    if Magnet_Br_Multiplier0 != 0:
        Magnet_Br_Multiplier01 = 0
        while Magnet_Br_Multiplier0 != Magnet_Br_Multiplier01:
            mcad1.SetVariable('Magnet_Br_Multiplier', Magnet_Br_Multiplier0)
            Magnet_Br_Multiplier01 = round(mcad1.GetVariable('Magnet_Br_Multiplier')[1], 1)
        loggenera('Br 系数' + ' : ' + str(Magnet_Br_Multiplier01))

    if Magnet_Arc_ED0 != 0:
        Magnet_Arc_ED01 = 0
        while Magnet_Arc_ED0 != Magnet_Arc_ED01:
            mcad1.SetVariable('Magnet_Arc_[ED]', Magnet_Arc_ED0)
            Magnet_Arc_ED01 = round(mcad1.GetVariable('Magnet_Arc_[ED]')[1], 0)
        loggenera('极弧系数' + ' : ' + str(Magnet_Arc_ED01))

    if PhaseAdvance0 != 0:
        PhaseAdvance01 = 0
        while PhaseAdvance0 != PhaseAdvance01:
            mcad1.SetVariable('PhaseAdvance', PhaseAdvance0)
            PhaseAdvance01 = round(mcad1.GetVariable('PhaseAdvance')[1], 1)
        loggenera('超前角' + ' : ' + str(PhaseAdvance01))

    if DCBusVoltage0 != 0:
        DCBusVoltage01 = 0
        while DCBusVoltage0 != DCBusVoltage01:
            mcad1.SetVariable('DCBusVoltage', DCBusVoltage0)
            DCBusVoltage01 = round(mcad1.GetVariable('DCBusVoltage')[1], 1)
        loggenera('电压' + ' : ' + str(DCBusVoltage01))

    if PeakCurrent0 != 0:
        PeakCurrent01 = 0
        while PeakCurrent0 != PeakCurrent01:
            mcad1.SetVariable('PeakCurrent', PeakCurrent0)
            PeakCurrent01 = round(mcad1.GetVariable('PeakCurrent')[1], 1)
        loggenera('峰值电流' + ' : ' + str(PeakCurrent01))

    if Magnet_Thickness0 != 0:
        Magnet_Thickness01 = 0

        while Magnet_Thickness0 != Magnet_Thickness01:
            mcad1.SetVariable('Magnet_Thickness', Magnet_Thickness0)
            Magnet_Thickness01 = round(mcad1.GetVariable('Magnet_Thickness')[1], 1)
        loggenera('磁铁厚度' + ' : ' + str(Magnet_Thickness01))

    if StatorSkew0 != 0:
        StatorSkew01 = 0
        while StatorSkew0 != StatorSkew01:
            mcad1.SetVariable('StatorSkew', StatorSkew0)
            StatorSkew01 = round(mcad1.GetVariable('StatorSkew')[1], 1)

        loggenera('定子扭角' + ' : ' + str(StatorSkew01))

    if RotorSkewAngle0 != 0:

        RotorSkewAngle01 = 0
        RotorSkewAngle02 = -RotorSkewAngle0

        while RotorSkewAngle0 != RotorSkewAngle01:
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 0, RotorSkewAngle02)
            mcad1.SetArrayVariable('RotorSkewAngle_Array', 2, RotorSkewAngle0)
            RotorSkewAngle01 = round(mcad1.GetArrayVariable('RotorSkewAngle_Array', 2)[1], 1)

        loggenera('转子扭角' + ' : ' + str(RotorSkewAngle01))

    if StatorIronLossBuildFactor0 != 0:
        StatorIronLossBuildFactor01 = 0
        while StatorIronLossBuildFactor0 != StatorIronLossBuildFactor01:
            mcad1.SetVariable('StatorIronLossBuildFactor', StatorIronLossBuildFactor0)
            StatorIronLossBuildFactor01 = round(mcad1.GetVariable('StatorIronLossBuildFactor')[1], 1)
        loggenera('定子铁耗系数' + ' : ' + str(StatorIronLossBuildFactor01))

    if RotorIronLossBuildFactor0 != 0:
        RotorIronLossBuildFactor01 = 0

        while RotorIronLossBuildFactor0 != RotorIronLossBuildFactor01:
            mcad1.SetVariable('RotorIronLossBuildFactor', RotorIronLossBuildFactor0)
            RotorIronLossBuildFactor01 = round(mcad1.GetVariable('RotorIronLossBuildFactor')[1], 1)
        loggenera('转子铁耗系数' + ' : ' + str(RotorIronLossBuildFactor01))

    if Armature_Diameter0 != 0:
        Armature_Diameter01 = 0
        while Armature_Diameter0 != Armature_Diameter01:
            mcad1.SetVariable('Armature_Diameter', Armature_Diameter0)
            Armature_Diameter01 = round(mcad1.GetVariable('Armature_Diameter')[1], 1)
        loggenera('定子外径' + ' : ' + str(Armature_Diameter01))

    if fixrotoroutdiameter0 != 0:
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

    if Tooth_Tip_Depth0 != 0:
        Tooth_Tip_Depth01 = 0
        while Tooth_Tip_Depth0 != Tooth_Tip_Depth01:
            mcad1.SetVariable('Tooth_Tip_Depth', Tooth_Tip_Depth0)
            Tooth_Tip_Depth01 = round(mcad1.GetVariable('Tooth_Tip_Depth')[1], 1)

        loggenera('槽顶深' + ' : ' + str(Tooth_Tip_Depth01))

    if Tooth_Width0 != 0:
        Tooth_Width01 = 0
        while Tooth_Width0 != Tooth_Width01:
            mcad1.SetVariable('Tooth_Width', Tooth_Width0)
            Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)

        loggenera('齿宽' + ' : ' + str(Tooth_Width01))

    if Slot_Depth0 != 0:
        Slot_Depth01 = 0
        Slot_Number01 = mcad1.GetVariable('Slot_Number')[1]
        Tooth_Width01 = round(mcad1.GetVariable('Tooth_Width')[1], 1)
        Slot_Depth0tem = fakeslotdepth(Armature_Diameter0, Slot_Depth0, Tooth_Width01, Slot_Number01)
        while Slot_Depth0tem != Slot_Depth01:
            mcad1.SetVariable('Slot_Depth', Slot_Depth0tem)
            Slot_Depth01 = round(mcad1.GetVariable('Slot_Depth')[1], 1)

        loggenera('槽深' + ' : ' + str(Slot_Depth0))

    if Slot_Corner_Radius0 != 0:
        Slot_Corner_Radius01 = 0
        while Slot_Corner_Radius0 != Slot_Corner_Radius01:
            mcad1.SetVariable('Slot_Corner_Radius', Slot_Corner_Radius0)
            Slot_Corner_Radius01 = round(mcad1.GetVariable('Slot_Corner_Radius')[1], 1)

        loggenera('槽底圆角' + ' : ' + str(Slot_Corner_Radius01))

    if Slot_Opening0 != 0:
        Slot_Opening01 = 0
        while Slot_Opening0 != Slot_Opening01:
            mcad1.SetVariable('Slot_Opening', Slot_Opening0)
            Slot_Opening01 = round(mcad1.GetVariable('Slot_Opening')[1], 1)

        loggenera('槽口宽' + ' : ' + str(Slot_Opening01))

    if Stator_Lam_Length0 != 0:
        Stator_Lam_Length01 = 0
        while Stator_Lam_Length0 != Stator_Lam_Length01:
            mcad1.SetVariable('Stator_Lam_Length', Stator_Lam_Length0)
            Stator_Lam_Length01 = round(mcad1.GetVariable('Stator_Lam_Length')[1], 1)
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

    EWdg_Overhang = round(EWdg_Overhang*RequestedGrossSlotFillFactor0, 1)
    mcad1.SetVariable('EWdg_Overhang_[F]', EWdg_Overhang)
    mcad1.SetVariable('EWdg_Overhang_[R]', EWdg_Overhang)
    loggenera('绕组端部长' + ' : ' + str(EWdg_Overhang))
    Liner_Thickness = mcad1.GetVariable('Liner_Thickness')[1]
    sumareaeff = sumarea - (lenl - Liner_Thickness) * Slot_Opening0 / 2 - lenlh * Liner_Thickness
    ArmatureTurnsPerCoil01 = round(mcad1.GetVariable('MagTurnsConductor')[1], 0)

    wirediam, wirediamco = wirediametcal(sumareaeff, RequestedGrossSlotFillFactor0, ArmatureTurnsPerCoil01)
    Wire_Diameter0, Copper_Diameter0 = round(wirediam, 3), round(wirediamco, 3)
    Wire_Diameter01 = 0
    while Wire_Diameter0 != Wire_Diameter01:
        mcad1.SetVariable('Wire_Diameter', Wire_Diameter0)
        Wire_Diameter01 = round(mcad1.GetVariable('Wire_Diameter')[1], 3)
    loggenera('线径' + ' : ' + str(Wire_Diameter0))

    Copper_Diameter01 = 0
    while Copper_Diameter0 != Copper_Diameter01:
        mcad1.SetVariable('Copper_Diameter', Copper_Diameter0)
        Copper_Diameter01 = round(mcad1.GetVariable('Copper_Diameter')[1], 3)
    loggenera('裸线径' + ' : ' + str(Copper_Diameter0))

    Tooth_Tip_Angle0 = round(Tooth_Tip_Angle0, 0)
    if Tooth_Tip_Angle0 != 0:
        Tooth_Tip_Angle01 = 0

        while Tooth_Tip_Angle0 != Tooth_Tip_Angle01:
            mcad1.SetVariable('Tooth_Tip_Angle', Tooth_Tip_Angle0)
            Tooth_Tip_Angle01 = round(mcad1.GetVariable('Tooth_Tip_Angle')[1], 0)

        loggenera('槽顶角度' + ' : ' + str(Tooth_Tip_Angle01))


def get_parameter(phasecheck, rownumber0, mcad1=None, num_list=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    if num_list is None:
        num_list = numpy.zeros((rownumbermax, 11))
    SystemEfficiency0 = mcad1.GetVariable('SystemEfficiency')[1]
    CoggingTorqueRippleVw0 = mcad1.GetVariable('CoggingTorqueRippleVw')[1]
    ShaftTorque0 = mcad1.GetVariable('ShaftTorque')[1]
    TorqueRippleMsVwPerCent0 = mcad1.GetVariable('TorqueRippleMsVwPerCent (MsVw)')[1]
    ConductorLoss0 = mcad1.GetVariable('ConductorLoss')[1]
    StatorIronLoss_Total0 = mcad1.GetVariable('StatorIronLoss_Total')[1]
    RotorBackIronLoss_Total0 = mcad1.GetVariable('RotorBackIronLoss_Total')[1]
    DCBusVoltage01 = 0
    if phasecheck == 3 or phasecheck == 4:
        PeakLineLineVoltage0 = mcad1.GetVariable('RmsLineLineVoltage')[1]
    else:
        PeakLineLineVoltage0 = mcad1.GetVariable('RmsBackEMFPhase')[1]
        DCBusVoltage01 = round(mcad1.GetVariable('DCBusVoltage')[1], 1)

    MeanDCSupplyCurrent0 = mcad1.GetVariable('MeanDCSupplyCurrent')[1]
    BMax_RotorBackIron0 = mcad1.GetVariable('BMax_RotorBackIron')[1]
    Copper_Diameter0 = mcad1.GetVariable('Copper_Diameter')[1]

    num_list[rownumber0][0] = round(SystemEfficiency0, 1)
    num_list[rownumber0][1] = round(CoggingTorqueRippleVw0 * 1000, 1)
    num_list[rownumber0][2] = round(ShaftTorque0 * 1000, 1)
    num_list[rownumber0][3] = round(TorqueRippleMsVwPerCent0, 1)
    num_list[rownumber0][4] = round(ConductorLoss0, 1)
    num_list[rownumber0][5] = round(StatorIronLoss_Total0, 1)
    num_list[rownumber0][6] = round(RotorBackIronLoss_Total0, 1)
    num_list[rownumber0][7] = round(PeakLineLineVoltage0, 1)
    num_list[rownumber0][8] = round(MeanDCSupplyCurrent0, 2)
    num_list[rownumber0][9] = round(BMax_RotorBackIron0, 1)
    num_list[rownumber0][10] = round(Copper_Diameter0, 3)
    if phasecheck == 1:
        if num_list[rownumber0][7] > DCBusVoltage01:
            for i in range(11):
                num_list[rownumber0][i] = 0.001
    for j in range(11):
        if num_list[rownumber0][j] <= 0:
            num_list[rownumber0][j] = 0.001
        if j == 0:
            if num_list[rownumber0][j] >= 100:
                num_list[rownumber0][j] = 0.001

    # main run function


def run_MotorCAD_Calcs(phasecheck, rownumber0, rownumber_, rownumberture, dexfileloca_, mcad1=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')

    if phasecheck == 3:
        read_parameter(rownumber0, rownumber_, rownumberture, mcad1, get_num_list)
    else:
        if phasecheck == 4:
            read_parameter_threeroundbo(rownumber0, rownumber_, rownumberture, dexfileloca_, mcad1, get_num_list)
        else:
            read_parameter_single(rownumber0, rownumber_, rownumberture, dexfileloca_, mcad1, get_num_list)
        mcad1.SetVariable('DXFFileName', dexfileloca_)
        mcad1.LoadDXFFile(dexfileloca_)
        mcad1.SetVariable('DXFImportType', 0)
        mcad1.SetVariable('UseDXFImportForFEA_Magnetic', True)
        mcad1.SaveToFile(MotorCAD_File)

    num_list1 = numpy.ones((1, 11))
    success = mcad1.DoMagneticCalculation()
    messagestr = mcad1.GetMessages(1)[1]
    return_num_list = numpy.zeros([rownumberture, columnnumber2])
    sheet = 'Sheet'
    if success == 0:
        get_parameter(phasecheck, rownumber0, mcad1, return_num_list)
        loggenera('计算结果：' + str(return_num_list[rownumber0]))
        write_data_to_excelholexc(rownumber0, rownumber0, rownumberture, columnnumber2, 6, rownumber_ + 3, sheet,
                                  pyseworkbookaddr, return_num_list)
        return 1
    elif success == -1 or 'fail' in messagestr or 'not' in messagestr or 'Unable' in messagestr or 'fatal' in messagestr or 'error' in messagestr:
        write_data_to_excelholexc(rownumber0, 0, 1, columnnumber2, 6, rownumber_ + 3, sheet, pyseworkbookaddr,
                                  num_list1)
        return 0


def Close_MotorCAD_Instances(mcad1=None):
    if mcad1 is None:
        mcad1 = win32com.client.Dispatch('motorcad.appautomation')
    mcad1.Quit()  # quits each instance of Motor-CAD


def maincal():
    phasecheck = numberget(fulldatanam, phasechecktabstr)
    if phasecheck == 3:
        rownumber_ = rownumber
        minitbookaddr_ = minitbookaddr
        dexfileloca_ = ''
    elif phasecheck == 4:
        rownumber_ = rownumbersthreerounbo
        minitbookaddr_ = minitbookroundboaddr
        dexfileloca_ = dexfilelocalist[sequence - 1]
    else:
        rownumber_ = rownumbersingle
        minitbookaddr_ = singminitbookaddr
        dexfileloca_ = dexfilelocalist[sequence - 1]

    global get_num_list
    get_num_list = numpy.zeros((rownumbermax, rownumber_))
    timecount1 = time.time()
    rownumberture = extract_data_from_excel(pyseworkbookaddr, rownumbermax, rownumber_, 2, get_num_list)
    if rownumberture != 0:
        mcad = openmcad(MotorCAD_File)
        mcad.DisplayScreen('Geometry')
        mcad.SetVariable('MessageDisplayState', 2)
        mcad.SetVariable('VerboseMessageOutput', True)
        mcad.SetVariable('VerboseFlag_Info', True)
        mcad.SetVariable('VerboseFlag_IOInfo', True)
        mcad.SetVariable('VerboseFlag_ScriptInfo', True)
        mcad.SetVariable('VerboseFlag_Results', True)

        if phasecheck == 3 or phasecheck == 4:
            quickcalcheck = numberget(fulldatanam, quickcalcheckstr)
            if quickcalcheck == 2:
                mcad.SetVariable('MagneticSolver', 2)

        loggenera('总计算次数：' + str(rownumberture))
        for rownumberture0 in range(rownumberture):
            timecount3 = time.time()
            try:
                run_MotorCAD_Calcs(phasecheck, rownumberture0, rownumber_, rownumberture, dexfileloca_, mcad)
                deltefile(dexfileloca_)
            except ValueError:
                loggenera(ValueError)
                pass

            timecount2 = time.time()
            loggenera(rownumberture0)

            loggenera('单次耗时：' + str(int(timecount2 - timecount3)) + '  ' + '总耗时：' + str(int(timecount2 - timecount1)))
        Close_MotorCAD_Instances(mcad)
        setdata(minitbookaddr_, sequence + 21, 9, 1)

    loggenera('完成')


def main():
    quickedit(0)
    paradifine()
    logconfi(logmotorcadaddrlist[sequence - 1])
    maincal()


if __name__ == '__main__':
    main()
