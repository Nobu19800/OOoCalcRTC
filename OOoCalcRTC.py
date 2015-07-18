# -*- coding: utf-8 -*-


##
#   @file OOoCalcRTC.py
#   @brief OOoCalcControl Component

import optparse
import sys,os,platform
import re


from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', home+'/OOoRTC/CalcIDL', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', '.\\OOoRTC\\CalcIDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages\\rtctree\\rtmidl']
    
    
    
    



import time
import random
import commands
import RTC
import OpenRTM_aist


from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming
from rtctree.utils import build_attr_string, dict_to_nvlist, nvlist_to_dict

import threading


import uno
import unohelper
from com.sun.star.awt import XActionListener

from com.sun.star.script.provider import XScriptContext

from com.sun.star.view import XSelectionChangeListener

from com.sun.star.awt import XTextListener


import SpreadSheet_idl



import OOoRTC
import CalcDataPort
from CalcControl import *
#from SpreadSheet_idl_example import *
from omniORB import PortableServer
import SpreadSheet, SpreadSheet__POA





#comp_num = random.randint(1,3000)
imp_id = "OOoCalcControl"# + str(comp_num)


##
# @class ControlName
# @brief ウィジェット名
#
class ControlName:
    NameServerFName = "nameserver"
    CreateBName = "CreateButton"
    DeleteBName = "DeleteButton"
    TextFName = "TextField1"
    SetColBName = "SetColButton"
    RowFName = "RowTextField"
    SheetCBName = "SheetComboBox"
    InfoTName = "InfoTextField"
    ColTName = "ColTextField"
    LenTName = "LenTextField"
    RTCTreeName = "RTCTreeControl"
    CreateTreeBName = "CreateRTCTreeButton"
    SetAllLineBName = "SetAllColButton"
    DetachBName = "DetachButton"
    AttachBName = "AttachButton"
    AttachCBName = "AttachComboBox"
    InPortCBName = "InPortComboBox"
    LCBName = "LowCheckBox"
    PortCBName = "PortComboBox"
    DataTypeCBName = "DataTypeComboBox"
    PortTypeCBName = "PortTypeComboBox"
    def __init__(self):
        pass


ooocalccontrol_spec = ["implementation_id", imp_id,
                  "type_name",         imp_id,
                  "description",       "Openoffice Calc Component",
                  "version",           "0.1.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "Office",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.actionLock", "1",
                  "conf.default.red", "255",
                  "conf.default.green", "255",
                  "conf.default.blue", "0",
                  "conf.default.stime", "0.05",
                  "conf.default.stCell_row", "A",
                  "conf.default.stCell_col", "1",
                  "conf.default.stCell_sheetname", "sheet1",
                  "conf.dataport0.port_type", "DataInPort",
                  "conf.dataport0.data_type", "TimedFloat",
                  "conf.dataport0.column", "1",
                  "conf.dataport0.start_row", "A",
                  "conf.dataport0.end_row", "A",
                  "conf.dataport0.sheetname", "sheet1",
                  "conf.dataport0.c_move", "1",
                  "conf.dataport0.Attach_Port", "None",
                  "conf.__widget__.actionLock", "radio",
                  "conf.__widget__.red", "spin",
                  "conf.__widget__.green", "spin",
                  "conf.__widget__.blue", "spin",
                  "conf.__widget__.file_path", "text",
                  "conf.__widget__.port_type", "radio",
                  "conf.__widget__.column", "spin",
                  "conf.__widget__.start_row", "text",
                  "conf.__widget__.end_row", "text",
                  "conf.__widget__.sheetname", "text",
                  "conf.__widget__.data_type", "radio",
                  "conf.__widget__.c_move", "radio",
                  "conf.__widget__.Attach_Port", "text",
                  "conf.__constraints__.actionLock", "(0,1,2)",
                  "conf.__constraints__.red", "0<=x<=255",
                  "conf.__constraints__.green", "0<=x<=255",
                  "conf.__constraints__.blue", "0<=x<=255",
                  "conf.__constraints__.column", "1<=x<=1000",
                  "conf.__constraints__.port_type", "(DataInPort,DataOutPort)",
                  "conf.__constraints__.data_type", """(TimedDouble,TimedLong,TimedFloat,TimedShort,TimedULong,TimedUShort,TimedChar,TimedWChar,
                                                    TimedBoolean,TimedOctet,TimedString,TimedWString,TimedDoubleSeq,TimedLongSeq,TimedFloatSeq,
                                                    TimedShortSeq,TimedULongSeq,TimedUShortSeq,TimedCharSeq,TimedWCharSeq,TimedOctetSeq,TimedStringSeq,
                                                    TimedWStringSeq,TimedRGBColour,TimedPoint2D,TimedVector2D,TimedPose2D,TimedVelocity2D,TimedAcceleration2D,
                                                    TimedPoseVel2D,TimedSize2D,TimedGeometry2D,TimedCovariance2D,TimedPointCovariance2D,TimedCarlike,TimedSpeedHeading2D,
                                                    TimedPoint3D,TimedVector3D,TimedOrientation3D,TimedPose3D,TimedVelocity3D,TimedAngularVelocity3D,TimedAcceleration3D,
                                                    TimedAngularAcceleration3D,TimedPoseVel3D,TimedSize3D,TimedGeometry3D,TimedCovariance3D,TimedSpeedHeading3D,TimedOAP)""",
                  "conf.__constraints__.c_move", "(0,1)",
                  ""]







##
# @class OOoCalcPortObject
# @brief 追加するポートのクラス
#


class OOoCalcPortObject(CalcDataPort.CalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    
    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def update_cellName(self, m_cal):
        cell, sheet, m_len = self.getCell2(m_cal,self._col,self._row,self._sn,self._length)
        self.update_cellNameSub(cell, m_len)


    
        
    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSingle(self, cell, m_len):
        cell.getCellByPosition(0, 0).String = self._name

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param m_len 行の範囲
    def update_cellNameSeq(self, cell, m_len):
        for j in range(0, m_len):
            cell.getCellByPosition(j, 0).String = self._name + ":" + str(j)

    

    ##
    # @brief 
    # @param self 
    # @param b データ名
    # @param count カウンター
    # @param m_len 行の範囲
    # @param cell セルオブジェクト
    # @return 
    def input_cellNameEx(self, b, count, m_len, cell):
        
        cell.getCellByPosition(count[0], 0).String = b
        
                    
        count[0] += 1
        if count[0] >= m_len:
            return False
        return True
            

    ##
    # @brief 
    # @param self
    # @param m_cal OOoCalcRTC
    # @param c 列番号
    # @param l 行番号
    # @param sn シート名
    # @param leng 行の範囲
    #
    def getCell2(self, m_cal, c, l, sn, leng):
        
        if m_cal.calc.sheets.hasByName(sn):
            sheet = m_cal.calc.sheets.getByName(sn)

            if leng == "":
                CN = l + str(c)
            else:
                CN = l + str(c) + ':' + leng + str(c)
            try:
                cell = sheet.getCellRangeByName(CN)
            except:
                pass
            m_len = cell.getRangeAddress().EndColumn - cell.getRangeAddress().StartColumn
            m_len = m_len + 1
            return cell, sheet, m_len
        else:
            return None, None, None

    ##
    # @brief 
    # @param self 
    # @param m_cal OOoCalcRTC
    def getCell(self, m_cal):
        return self.getCell2(m_cal,self._num,self._row,self._sn,self._length)

    

    
                    
    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param sheet シートオブジェクト
    # @param m_cal OOoCalcRTC
    def putOut(self, cell, sheet, m_cal):
        m_string = CalcDataPort.DataType.String
        m_value = CalcDataPort.DataType.Value

        cell.CellBackColor = OOoRTC.RGB(int(m_cal.red[0]), int(m_cal.green[0]), int(m_cal.blue[0]))
        
        if self._dataType[2] == m_string:
            if  self._length == "":
                val = cell.String
            else:
                val = []
                m_len = cell.getRangeAddress().EndColumn - cell.getRangeAddress().StartColumn
                for i in range(0, m_len+1):
                    val.append(cell.getCellByPosition(i, 0).String)
        elif self._dataType[2] == m_value:
            if self._length == "":
                val = cell.Value
            else:
                val = []
                m_len = cell.getRangeAddress().EndColumn - cell.getRangeAddress().StartColumn
                for i in range(0, m_len+1):
                    val.append(cell.getCellByPosition(i, 0).Value)
                    
        
                
        if self._num > 1 and self.state == True:
            t_n = self._num -1
            cell2, sheet2, m_len2 = self.getCell2(m_cal,t_n,self._row,self._sn,self._length)
            cell2.CellBackColor = OOoRTC.RGB(255, 255, 255)
            
            


        return val

##
# @class OOoCalcInPort
# @brief
#

class OOoCalcInPort(CalcDataPort.CalcInPort, OOoCalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcInPort.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        OOoCalcPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return OOoCalcPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return OOoCalcPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return OOoCalcPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcInPort.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcInPort.update_cellNameSub(self, cell, m_len)
        
    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        m_string = CalcDataPort.DataType.String
        m_value = CalcDataPort.DataType.Value
        
        cell, sheet, m_len = self.getCell(m_cal)
        

        if cell != None:
            
            if self._dataType[2] == m_string:
                cell.getCellByPosition(0, 0).String = b
            elif self._dataType[2] == m_value:
                if self.state:
                    pass
                else:
                    db = 0
                    val = cell.getCellByPosition(0, 0).Value
                    if self.count != 0:
                        db = (b - val)/m_cal.stime[0]
                    celld, sheetd, m_lend = self.getCell2(m_cal, self._num+1, self._row, self._sn, "")
                    celld.getCellByPosition(0, 0).Value = db

                    cellf, sheetf, m_lenf = self.getCell2(m_cal, self._num+2, self._row, self._sn, "")
                    fb = b*m_cal.stime[0]
              
                    if self.count != 0:
                        fb = cellf.getCellByPosition(0, 0).Value + b*m_cal.stime[0]

                    cellf.getCellByPosition(0, 0).Value = fb
                    
                cell.getCellByPosition(0, 0).Value = b
                self.count += 1
                    
            if self.state:
                self._num = self._num + 1


        
                    
##
# @class OOoCalcInPortSeq
# @brief 
class OOoCalcInPortSeq(CalcDataPort.CalcInPortSeq, OOoCalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcInPortSeq.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        OOoCalcPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return OOoCalcPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return OOoCalcPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return OOoCalcPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcInPortSeq.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcInPortSeq.update_cellNameSub(self, cell, m_len)

    ##
    # @brief 
    # @param self 
    # @param cell セルオブジェクト
    # @param b データ
    def updateIn(self, b, m_cal):
        m_string = CalcDataPort.DataType.String
        m_value = CalcDataPort.DataType.Value

        cell, sheet, m_len = self.getCell(m_cal)

        if cell != None:
            if self._dataType[2] == m_string:
                for j in range(0, len(b)):
                    if m_len > j:
                        cell.getCellByPosition(j, 0).String = b[j]
            elif self._dataType[2] == m_value:
                if self.state:
                    pass
                else:
                    celld, sheetd, m_lend = self.getCell2(m_cal, self._num+1, self._row, self._sn, self._length)
                    for j in range(0, len(b)):
                        if m_lend > j:
                            if self.count != 0:
                                val = cell.getCellByPosition(j, 0).Value
                                celld.getCellByPosition(j, 0).Value = (b[j] - val)/m_cal.stime[0]
                            else:
                                celld.getCellByPosition(j, 0).Value = 0
                    cellf, sheetf, m_lenf = self.getCell2(m_cal, self._num+2, self._row, self._sn, self._length)
                    for j in range(0, len(b)):
                        if m_lenf > j:
                            if self.count != 0:
                                val = cell.getCellByPosition(j, 0).Value
                                cellf.getCellByPosition(j, 0).Value = cellf.getCellByPosition(j, 0).Value + b[j]*m_cal.stime[0]
                            else:
                                cellf.getCellByPosition(j, 0).Value = b[j]*m_cal.stime[0]

                for j in range(0, len(b)):
                    if m_len > j:
                        cell.getCellByPosition(j, 0).Value = b[j]
                self.count += 1
                
            if self.state:
                self._num = self._num + 1
                
            

            
        
    

##
# @class OOoCalcInPortEx
# @brief 
class OOoCalcInPortEx(CalcDataPort.CalcInPortEx, OOoCalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcInPortEx.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

        
    def update_cellName(self, m_cal):
        OOoCalcPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return OOoCalcPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return OOoCalcPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return OOoCalcPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcInPortEx.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcInPortEx.update_cellNameSub(self, cell, m_len)
    

    ##
    # @brief
    # @param self 
    # @param b データ
    # @param count カウンター
    # @param m_len 行の範囲
    # @param cell セルオブジェクト
    # @param d_type データタイプ
    def putDataEx(self, b, count, m_len, cell, d_type):
        m_string = CalcDataPort.DataType.String
        m_value = CalcDataPort.DataType.Value

        if d_type == m_string:
            cell.getCellByPosition(count[0], 0).String = b
        elif d_type == m_value:
            cell.getCellByPosition(count[0], 0).Value = b
                    
        count[0] += 1
        if count[0] >= m_len:
            return False
        return True

        
##
# @class OOoCalcOutPort
# @brief 
class OOoCalcOutPort(CalcDataPort.CalcOutPort, OOoCalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcOutPort.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        OOoCalcPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return OOoCalcPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return OOoCalcPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return OOoCalcPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcOutPort.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcOutPort.update_cellNameSub(self, cell, m_len)
    

##
# @class OOoCalcOutPortSeq
# @brief 
class OOoCalcOutPortSeq(CalcDataPort.CalcOutPortSeq, OOoCalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcOutPortSeq.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        OOoCalcPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return OOoCalcPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return OOoCalcPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return OOoCalcPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcOutPortSeq.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcOutPortSeq.update_cellNameSub(self, cell, m_len)
    

##
# @class OOoCalcOutPortEx
# @brief 
#
class OOoCalcOutPortEx(CalcDataPort.CalcOutPortEx, OOoCalcPortObject):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param port データポート
    # @param data データオブジェクト
    # @param name データポート名
    # @param row 行番号
    # @param col 列番号
    # @param mlen 行の範囲
    # @param sn シート名
    # @param mstate 列を移動するか
    # @param port_a 接続するデータポート
    # @param m_dataType データ型
    # @param t_attachports 関連付けしたデータポート
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        CalcDataPort.CalcOutPortEx.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def update_cellName(self, m_cal):
        OOoCalcPortObject.update_cellName(self, m_cal)
    def update_cellNameSingle(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSingle(self, cell, m_len)
    def update_cellNameSeq(self, cell, m_len):
        OOoCalcPortObject.update_cellNameSeq(self, cell, m_len)
    def input_cellNameEx(self, b, count, m_len, cell):
        return OOoCalcPortObject.input_cellNameEx(self, b, count, m_len, cell)
    def getCell(self, m_cal):
        return OOoCalcPortObject.getCell(self, m_cal)
    def putOut(self, cell, sheet, m_cal):
        return OOoCalcPortObject.putOut(self, cell, sheet, m_cal)
    def putData(self, m_cal):
        CalcDataPort.CalcOutPortEx.putData(self, m_cal)
    def update_cellNameSub(self, cell, m_len):
        CalcDataPort.CalcOutPortEx.update_cellNameSub(self, cell, m_len)
    def putDataEx(self, count, val, d_type):
        return CalcDataPort.CalcOutPortEx.putDataEx(self, count, val, d_type)
    


##
# @class OOoCalcControl
# @brief OpenOffice Calcを操作するためのRTCのクラス
#

class OOoCalcControl(CalcControl):

    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    CalcControl.__init__(self, manager)
    

    try:
      self.calc = OOoCalc()
    except NotOOoCalcException:
      return

    self.m_CalcInPort = OOoCalcInPort
    self.m_CalcInPortSeq = OOoCalcInPortSeq
    self.m_CalcInPortEx = OOoCalcInPortEx

    self.m_CalcOutPort = OOoCalcOutPort
    self.m_CalcOutPortSeq = OOoCalcOutPortSeq
    self.m_CalcOutPortEx = OOoCalcOutPortEx

    
    
    return

  ##
  # @brief セルオブジェクト、シートオブジェクトの取得
  # @param self 
  # @param l 行番号
  # @param c 列番号
  # @param sn シート名
  # @return セルオブジェクト、シートオブジェクト
  #

  def getCell(self, l, c, sn):
    if self.m_comp.calc.sheets.hasByName(sn):
        sheet = self.m_comp.calc.sheets.getByName(sn)
        CN = l+c
        try:
            cell = sheet.getCellRangeByName(CN)
            return cell, sheet
        except:
            return None
    else:
        return None

    ##
    # @brief セルの文字列を取得
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @return セルの文字列
    #
  def get_string(self, l, c, sn):
    cell, sheet = self.getCell(l,c,sn)
    if cell:
        ans = str(cell.String)
        return ans
    return "error"

    ##
    # @brief セルの値を設定
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @param v 設定する値
    #
  def set_value(self, l, c, sn, v):
    cell, sheet = self.getCell(l,c,sn)
    if cell:
        cell.Value = v

    ##
    # @brief セルの文字列を設定
    # @param self 
    # @param l 行番号
    # @param c 列番号
    # @param sn シート名
    # @param v 設定する文字列
    #  
  def set_string(self, l, c, sn, v):
    cell, sheet = self.getCell(l,c,sn)
    if cell:
        cell.String = v

    ##
    # @brief 画面の更新停止
    # @param self 
    # 
  def addActionLock(self):
    self.calc.document.addActionLock()

    ##
    # @brief 画面の更新再開
    # @param self 
    # 
  def removeActionLock(self):
    self.calc.document.removeActionLock()

    ##
    # @brief データポートと関連付けしてあるセルの色を設定
    # @param self 
    # @param op データポートオブジェクト
    #  
  def setCellColor(self, op):
    t_n = op._num
    if op.state:
        t_n -= 1
    if op._length == "":
        CN = op._row + str(t_n)
    else:
        CN = op._row + str(t_n) + ':' + op._length + str(t_n)
    sheetname = op._sn
    if self.calc.sheets.hasByName(sheetname):
        sheet = self.calc.sheets.getByName(sheetname)
        try:
            cell = sheet.getCellRangeByName(CN)
            cell.CellBackColor = OOoRTC.RGB(255, 255, 255)
        except:
            pass

  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  #
  
  def onInitialize(self):
    CalcControl.onInitialize(self)
    OOoRTC.calc_comp = self
    
    
    return RTC.RTC_OK

  
  ##
  # @brief 不活性化時のコールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onDeactivated(self, ec_id):
    CalcControl.onDeactivated(self, ec_id)
    
    return RTC.RTC_OK

  def setTime(self):
    CN = self.stCell_row[0] + str(self.stCell_col[0])
    if self.calc.sheets.hasByName(self.stCell_sheetname[0]):
        sheet = self.calc.sheets.getByName(self.stCell_sheetname[0])
        try:
            cell = sheet.getCellRangeByName(CN)
            cell.getCellByPosition(0, 0).Value = self.m_time
        except:
            pass

  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    
    
    CalcControl.onExecute(self, ec_id)
    
    
    return RTC.RTC_OK

  

    
  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param 
  # @return RTC::ReturnCode_t
  def onFinalize(self):
      CalcControl.onFinalize(self)
      OOoRTC.calc_comp = None
      return RTC.RTC_OK





##
# @brief コンポーネントを活性化してCalcの操作を開始する関数
#

def Start():
    if OOoRTC.calc_comp:
        OOoRTC.calc_comp.mActivate()

##
# @brief コンポーネントを不活性化してCalcの操作を終了する関数
#
def Stop():
    if OOoRTC.calc_comp:
        OOoRTC.calc_comp.mDeactivate()
##
# @brief コンポーネントの実行周期を設定する関数
#

def Set_Rate():
    if OOoRTC.calc_comp:
      
      try:
        calc = OOoCalc()
      except NotOOoClacException:
          return
      
      
      
      
      
      
      
      for i in range(0, calc.sheets.Count):
          forms = calc.sheets.getByIndex(i).getDrawPage().getForms()
          for j in range(0, forms.Count):
              form = forms.getByIndex(j)
              st_control = form.getByName('Rate')
              if st_control:
                  try:
                      text = float(st_control.Text)
                  except:
                      return
                  
                  OOoRTC.calc_comp.mSetRate(text)


      
      
    




##
# @brief RTCをマネージャに登録する関数
# @param manager マネージャーオブジェクト
def OOoCalcControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=ooocalccontrol_spec)
  manager.registerFactory(profile,
                          OOoCalcControl,
                          OpenRTM_aist.Delete)
  


##
# @brief RTCをマネージャに登録する関数
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoCalcControlInit(manager)
  
  comp = manager.createComponent(imp_id)

  
        


##
# @brief アウトポートを追加する関数
# @param name データポート名
# @param i_port 接続するデータポート
# @param dlg_control ダイアログオブジェクト
def CompAddOutPort(name, i_port, dlg_control, autoCon = True):
    if OOoRTC.calc_comp != None:
        tfrow_control = dlg_control.getControl( ControlName.RowFName )
        tfc_control = dlg_control.getControl( ControlName.ColTName )
        mlen_control = dlg_control.getControl( ControlName.LenTName )
        tfst_control = dlg_control.getControl( ControlName.SheetCBName )
        cb_control = dlg_control.getControl( ControlName.LCBName )
        
        row = str(tfrow_control.Text)
        sn = str(tfst_control.Text)
        col = str(tfc_control.Text)
        
        mlen = str(mlen_control.Text)
        
        mstate = int(cb_control.State)
        mst = True
        if mstate == 0:
            mst = False

        
        
        tcomp = OOoRTC.calc_comp.mAddOutPort(name, i_port, row, col, mlen, sn, mst, {}, autoCon)
        if tcomp:
            tcomp.update_cellName(OOoRTC.calc_comp)

##
# @brief インポートを追加する関数
# @param name データポート名
# @param o_port 接続するデータポート
# @param dlg_control ダイアログオブジェクト
def CompAddInPort(name, o_port, dlg_control, autoCon = True):
    if OOoRTC.calc_comp != None:
        tfrow_control = dlg_control.getControl( ControlName.RowFName )
        tfc_control = dlg_control.getControl( ControlName.ColTName )
        mlen_control = dlg_control.getControl( ControlName.LenTName )
        tfst_control = dlg_control.getControl( ControlName.SheetCBName )
        cb_control = dlg_control.getControl( ControlName.LCBName )
        
        row = str(tfrow_control.Text)
        sn = str(tfst_control.Text)
        col = str(tfc_control.Text)
        mlen = str(mlen_control.Text)
        mstate = int(cb_control.State)
        mst = True
        if mstate == 0:
            mst = False
        tcomp = OOoRTC.calc_comp.mAddInPort(name, o_port, row, col, mlen, sn, mst, {}, autoCon)
        if tcomp:
            tcomp.update_cellName(OOoRTC.calc_comp)

##
# @brief RTC起動の関数
#

def createOOoCalcComp():
    if OOoRTC.calc_comp:
        MyMsgBox('',OOoRTC.SetCoding('RTCは起動済みです','utf-8'))
        return
    
    if OOoRTC.mgr == None:
        if os.name == 'posix':
            home = expanduser("~")
            OOoRTC.mgr = OpenRTM_aist.Manager.init([os.path.abspath(__file__), '-f', home+'/OOoRTC/rtc.conf'])
        elif os.name == 'nt':
            OOoRTC.mgr = OpenRTM_aist.Manager.init([os.path.abspath(__file__), '-f', '.\\rtc.conf'])
        else:
            return

      
        OOoRTC.mgr.setModuleInitProc(MyModuleInit)
        OOoRTC.mgr.activateManager()
        OOoRTC.mgr.runManager(True)
    else:
        MyModuleInit(OOoRTC.mgr)

    try:
      calc = OOoCalc()
    except NotOOoCalcException:
      return

    
    
    sheetname = OOoRTC.SetCoding('保存用','utf-8')
    if calc.sheets.hasByName(sheetname):
        pass
    else:
        try:
            cnt = calc.sheets.Count
            calc.sheets.insertNewByName(sheetname, cnt)
        except unohelper.RuntimeException:
            calc.run_errordialog(title='エラー', message='')
            return
        
    MyMsgBox('',OOoRTC.SetCoding('RTCを起動しました','utf-8'))

    
    
    LoadSheet()

    
    
    return




##
# @brief メッセージボックス表示の関数
# @param title ウインドウのタイトル
# @param message 表示する文章
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)


##
# @brief OpenOfficeを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

class Bridge(object):
  def __init__(self):
    self._desktop = XSCRIPTCONTEXT.getDesktop()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._frame = self._desktop.CurrentFrame
    self._window = self._frame.ContainerWindow
    self._toolkit = self._window.Toolkit
  def run_infodialog(self, title='', message=''):
    try:
        msgbox = self._toolkit.createMessageBox(self._window,uno.createUnoStruct('com.sun.star.awt.Rectangle'),'infobox',1,title,message)
        msgbox.execute()
        msgbox.dispose()
    except:
        msgbox = self._toolkit.createMessageBox(self._window,'infobox',1,title,message)
        msgbox.execute()
        msgbox.dispose()





##
# @brief ダイアログのツリーにネーミングサーバーのオブジェクトを登録する関数
# @param oTreeModel ツリーオブジェクト
# @param smgr UNOコンポーネントコンテキスト
# @param ctx UNOサービスマネージャー
# @param dlg_control ダイアログオブジェクト

def SetRTCTree(oTreeModel, smgr, ctx, dlg_control):
    oTree = dlg_control.getControl( ControlName.RTCTreeName )
    tfns_control = dlg_control.getControl( ControlName.NameServerFName )
    if OOoRTC.mgr != None:
        

               
        orb = OOoRTC.mgr._orb

        
       
        namingserver = OOoRTC.SetNamingServer(str(tfns_control.Text), orb, MyMsgBox)
        
        
         
        if namingserver:
            
            oTreeDataModel = smgr.createInstanceWithContext("com.sun.star.awt.tree.MutableTreeDataModel", ctx)
            root = oTreeDataModel.createNode("/", False)
            oTreeDataModel.setRoot(root)
            oChild = oTreeDataModel.createNode(str(tfns_control.Text),False)
            root.appendChild(oChild)

            
            
            _path = ['/', str(tfns_control.Text)]
            _paths = []
            OOoRTC.rtc_get_rtclist(namingserver, _paths, _path, oChild, oTreeDataModel)

            
                      
            
            oTreeModel.DataModel = oTreeDataModel

            tf1_control = dlg_control.getControl( ControlName.TextFName )
            tfrow_control = dlg_control.getControl( ControlName.RowFName )
            

            btn1_listener = CreatePortListener( dlg_control, _paths)
            cmdbtn1_control = dlg_control.getControl(ControlName.CreateBName)
            cmdbtn1_control.addActionListener(btn1_listener)

            delete_listener = DeleteListener(dlg_control, _paths)
            delete_control = dlg_control.getControl(ControlName.DeleteBName)
            delete_control.addActionListener(delete_listener)

            setCol_listener = SetColListener(dlg_control, _paths)
            setcol_control = dlg_control.getControl(ControlName.SetColBName)
            setcol_control.addActionListener(setCol_listener)

            attatch_listener = AttachListener( dlg_control, _paths)
            attatch_control = dlg_control.getControl(ControlName.AttachBName)
            attatch_control.addActionListener(attatch_listener)


            detatch_listener = DetachListener( dlg_control, _paths)
            detatch_control = dlg_control.getControl(ControlName.DetachBName)
            detatch_control.addActionListener(detatch_listener)

            
            
            

            oTree.addSelectionChangeListener(TreeSelectListener(dlg_control, _paths))



##
# @brief OpenOffice Calcを操作するためのクラス
# @class OOoCalc
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

class OOoCalc(Bridge):
  def __init__(self):
    Bridge.__init__(self)
    if not self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
      self.run_errordialog(title='エラー', message='このマクロはOpenOffice.org Calcの中で実行してください')
      raise NotOOoCalcException()
    self.__current_controller = self._document.CurrentController
    self.__sheets = self._document.Sheets
  def get_active_sheet(self):
    return self.__current_controller.ActiveSheet
  def set_active_sheet(self, value):
    self.__current_controller.ActiveSheet = value
  active_sheet = property(get_active_sheet, set_active_sheet)
  @property
  def sheets(self): return self.__sheets
  @property
  def document(self): return self._document




def LoadParam(count, sheet):
    CN = 'B' + str(count)
    cell = sheet.getCellRangeByName(CN)
    if cell.String == '':
        row = "A"
    row = cell.String

    CN = 'C' + str(count)
    cell = sheet.getCellRangeByName(CN)
    if cell.String == '':
        col = "1"
    col = cell.String

    CN = 'D' + str(count)
    cell = sheet.getCellRangeByName(CN)
                            
    mlen = cell.String
                                
                                
    CN = 'E' + str(count)
    cell = sheet.getCellRangeByName(CN)
    if cell.String == '':
       sn = "Sheet1"
    sn = cell.String

    CN = 'F' + str(count)
    cell = sheet.getCellRangeByName(CN)
    
    if str(cell.String) == "True":
        mstate = True
    else:
        mstate = False

    CN = 'G' + str(count)
    cell = sheet.getCellRangeByName(CN)
    tmp = re.split(':',cell.String)
    t_attachports = {}
    for pp in tmp:
        if pp != "":
            t_attachports[pp] = pp

    return row,col,mlen,sn,mstate,t_attachports


##
# @brief 読み込んだ保存用シートからポートを作成する関数
#

def LoadSheet():
    
    if OOoRTC.calc_comp:
        calc = OOoRTC.calc_comp.calc
        
        OOoRTC.calc_comp.mRemoveAllPort()
        sheetname = OOoRTC.SetCoding('保存用','utf-8')
        if calc.sheets.hasByName(sheetname):
            
            sheet = calc.sheets.getByName(sheetname)
            count = 1
            m_hostname = ''
            _path = []
            while True:
                
                CN = 'A' + str(count)
                try:
                    cell = sheet.getCellRangeByName(CN)
                    if cell.String == '':
                        return
                    
                    m_name = re.split(':',cell.String)
                    if len(m_name) < 2:
                        return
                    if len(m_name) == 2:
                        for dn in CalcDataPort.DataType.DataTypeList:
                            if m_name[1] == dn:
                                m_name[1] = dn
                        F_Name = m_name[1] + str(OpenRTM_aist.uuid1())
                        row,col,mlen,sn,mstate,t_attachports = LoadParam(count, sheet)
                        if m_name[0] == 'DataOutPort':
                            OOoRTC.calc_comp.mAddOutPort(F_Name, [[m_name[0],m_name[1]],m_name[1]], row, col, mlen, sn, mstate, t_attachports, False)
                        elif m_name[0] == 'DataInPort':
                            OOoRTC.calc_comp.mAddInPort(F_Name, [[m_name[0],m_name[1]],m_name[1]], row, col, mlen, sn, mstate, t_attachports, False)
                    #MyMsgBox('',str(m_name[1]))
                    else:
                        if m_hostname == m_name[1]:
                            pass
                        else:
                            
                            _paths = OOoRTC.GetPathList(m_name[1], OOoRTC.mgr, None)
                            
                            m_hostname = m_name[1]

                        
                        if _paths == None:
                            return
                        
                        for p in _paths:
                            if p[0] == m_name:
                                F_Name = p[0][-2] + p[0][-1]
                                profile = p[1].get_port_profile()
                                props = nvlist_to_dict(profile.properties)
                                
                                    
                                
                                row,col,mlen,sn,mstate,t_attachports = LoadParam(count, sheet)
                                    
                                
                                if props['port.port_type'] == 'DataInPort':
                                    OOoRTC.calc_comp.mAddOutPort(F_Name, p, row, col, mlen, sn, mstate, t_attachports)
                                elif props['port.port_type'] == 'DataOutPort':
                                    OOoRTC.calc_comp.mAddInPort(F_Name, p, row, col, mlen, sn, mstate, t_attachports)
                except:
                    pass
                count = count + 1


        
                
                


##
# @brief 作成したポートの設定を保存する関数
#
                

def UpdateSaveSheet():
    
    if OOoRTC.calc_comp:
        #OOoRTC.calc_comp.update_cellName()
        try:
          calc = OOoCalc()
        except NotOOoCalcException:
          return
        sheetname = OOoRTC.SetCoding('保存用','utf-8')
        if calc.sheets.hasByName(sheetname):
            sheet = calc.sheets.getByName(sheetname)
            for i in range(1, 30):
                try:
                    CN = 'A' + str(i)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = ''

                    CN = 'B' + str(i)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = ''

                    CN = 'C' + str(i)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = ''

                    CN = 'D' + str(i)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = ''

                    CN = 'E' + str(i)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = ''

                    CN = 'F' + str(i)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = ''

                    CN = 'G' + str(i)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = ''
                except:
                    pass
                
            count = 1
            PortList = []
            for n,o in OOoRTC.calc_comp.OutPorts.items():
                PortList.append(o)
            for n,i in OOoRTC.calc_comp.InPorts.items():
                PortList.append(i)

            for p in PortList:
                CN = 'A' + str(count)
                try:
                    cell = sheet.getCellRangeByName(CN)
                    pn = ''
                    for j in range(0, len(p._port_a[0])):
                        if j == 0:
                            pn = p._port_a[0][j]
                        else:
                            pn = pn + ':' + p._port_a[0][j]
                    cell.String = str(pn)

                    CN = 'B' + str(count)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = p._row

                    CN = 'C' + str(count)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = p._col

                    CN = 'D' + str(count)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = p._length

                    CN = 'E' + str(count)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = p._sn

                    CN = 'F' + str(count)
                    cell = sheet.getCellRangeByName(CN)
                    cell.String = str(p.state)

                    CN = 'G' + str(count)
                    cell = sheet.getCellRangeByName(CN)
                    pn = ''
                    tmp = 0
                    
                    for k,j in p.attachports.items():
                        if tmp == 0:
                            pn = j
                        else:
                            pn = pn + ':' + j
                        tmp += 1
                            
                        
                    cell.String = str(pn)
                except:
                    pass

                count = count + 1
            
        else:
            return

##
# @brief ツリーの選択位置が変わったときに各テキストボックスの内容を変更する関数
# @param dlg_control ダイアログオブジェクト
# @param m_port データポートオブジェクト
#

def UpdateTree(dlg_control, m_port):
    
    scb_control = dlg_control.getControl( ControlName.SheetCBName )
    scb_control.setText(m_port._sn)

    dtcb_control = dlg_control.getControl( ControlName.DataTypeCBName )
    ptcb_control = dlg_control.getControl( ControlName.PortTypeCBName )
    if len(m_port._port_a[0]) == 2:
        ptcb_control.setText(m_port._port_a[0][0])
        dtcb_control.setText(m_port._port_a[1])
    else:
        dtcb_control.setText("")
    
    tfrow_control = dlg_control.getControl( ControlName.RowFName )
    tfrow_control.setText(m_port._row)

    mlen_control = dlg_control.getControl( ControlName.LenTName )
    mlen_control.setText(m_port._length)
    

    ffcol_control = dlg_control.getControl( ControlName.InfoTName )
    ffcol_control.setText(u'作成済み')

    cfcol_control = dlg_control.getControl( ControlName.ColTName )
    cfcol_control.setText(str(m_port._col))

    cfcol_control = dlg_control.getControl( ControlName.LCBName )
    cfcol_control.enableTriState( True )
    if m_port.state:
        cfcol_control.setState(1)
    else:
        cfcol_control.setState(0)

    UpdateInPortList(dlg_control)
    UpdateAttachPort(dlg_control, m_port)

##
# @brief データポートのリストを更新する関数
# @param dlg_control ダイアログオブジェクト

def UpdateDataPortList(dlg_control):
    if OOoRTC.calc_comp:
        dpcb_control = dlg_control.getControl( ControlName.PortCBName )

        dpcb_control.removeItems(0,dpcb_control.ItemCount)
        dpcb_control.Text = ""
        
        for n,i in OOoRTC.calc_comp.InPorts.items():
            dpcb_control.addItem (i._name, dpcb_control.ItemCount)

        for n,i in OOoRTC.calc_comp.OutPorts.items():
            dpcb_control.addItem (i._name, dpcb_control.ItemCount)

            
##
# @brief インポートのリストを更新する関数
# @param dlg_control ダイアログオブジェクト
def UpdateInPortList(dlg_control):
    
    if OOoRTC.calc_comp:
        ipcb_control = dlg_control.getControl( ControlName.InPortCBName )
        ipcb_control.removeItems(0,ipcb_control.ItemCount)
        ipcb_control.Text = ""
        
        for n,i in OOoRTC.calc_comp.InPorts.items():
            
            
            ipcb_control.addItem (i._name, ipcb_control.ItemCount)
           

            
            
        


##
# @brief 関連付けしたポートのリストを更新する関数
# @param dlg_control ダイアログオブジェクト
# @param m_port データポートオブジェクト
def UpdateAttachPort(dlg_control, m_port):
    
    ipcb_control = dlg_control.getControl( ControlName.AttachCBName )
    ipcb_control.removeItems(0,ipcb_control.ItemCount)
    ipcb_control.Text = ""
    
    for n,i in m_port.attachports.items():
        
        ipcb_control.addItem (i, ipcb_control.ItemCount)
        
    

##
# @brief ポートを削除したときに各テキストボックスを変更する関数
# @param dlg_control ダイアログオブジェクト
def ClearInfo(dlg_control):
    
    ffcol_control = dlg_control.getControl( ControlName.InfoTName )
    ffcol_control.setText(u'未作成')

    cfcol_control = dlg_control.getControl( ControlName.ColTName )
    cfcol_control.setText("2")

    UpdateInPortList(dlg_control)
    UpdateDataPortList(dlg_control)



##
# @class PortListListener
# @brief データポートリストのコールバック
#
class PortListListener(unohelper.Base, XTextListener):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param dlg_control ダイアログオブジェクト
    def __init__(self, dlg_control):
        self.dlg_control = dlg_control
    
    ##
    # @brief 
    # @param self 
    # @param actionEvent 
    def textChanged(self, actionEvent):
        UpdateInPortList(self.dlg_control)
        if OOoRTC.calc_comp:
            ptlist_control = self.dlg_control.getControl( ControlName.PortCBName )
            
            
            
            if OOoRTC.calc_comp.InPorts.has_key(str(ptlist_control.Text)) == True:
                UpdateTree(self.dlg_control, OOoRTC.calc_comp.InPorts[str(ptlist_control.Text)])
            elif OOoRTC.calc_comp.OutPorts.has_key(str(ptlist_control.Text)) == True:
                UpdateTree(self.dlg_control, OOoRTC.calc_comp.OutPorts[str(ptlist_control.Text)])
        

##
# @brief ポート関連付けの関数
# @param dlg_control ダイアログオブジェクト
# @param m_port データポートオブジェクト
def AttachTC(dlg_control, m_port):
    
    tfcol_control = dlg_control.getControl( ControlName.InPortCBName )
    iname = str(tfcol_control.Text)
    
    if OOoRTC.calc_comp.InPorts.has_key(iname) == True:
                        
        m_port.attachports[iname] = iname
        OOoRTC.calc_comp.InPorts[iname].attachports[m_port._name] = m_port._name

        UpdateSaveSheet()
        UpdateAttachPort(dlg_control, m_port)

        MyMsgBox('',OOoRTC.SetCoding(m_port._name+"と"+iname+"を関連付けしました",'utf-8'))

        tfcol_control.Text = iname
                    
    else:
        MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('インポートの名前が正しくありません','utf-8'))
        return
        

##
# @class AttachListener
# @brief ポート関連付けボタンのコールバック
#
class AttachListener( unohelper.Base, XActionListener):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param dlg_control ダイアログオブジェクト
    # @param _paths データポートのパス
    def __init__(self, dlg_control, _paths):
        self._paths = _paths
        self.dlg_control = dlg_control

    ##
    # @brief 
    # @param self 
    # @param actionEvent 
    def actionPerformed(self, actionEvent):
        

        if OOoRTC.calc_comp:
            
            ptlist_control = self.dlg_control.getControl( ControlName.PortCBName )
            
            
            
            if OOoRTC.calc_comp.OutPorts.has_key(str(ptlist_control.Text)) == True:
                o = OOoRTC.calc_comp.OutPorts[str(ptlist_control.Text)]
                AttachTC(self.dlg_control, o)
                return

        objectTree = self.dlg_control.getControl( ControlName.RTCTreeName )
        t_comp, nd = OOoRTC.JudgePort(objectTree, self._paths)
            
        if t_comp:
            
            for n,o in OOoRTC.calc_comp.OutPorts.items():
                
                if o._port_a[0] == t_comp[0]:
                    
                    AttachTC(self.dlg_control, o)
                    return
                    
                    
            
        else:
            MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('アウトポートを選択してください','utf-8'))
            return
        
        MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('削除済みです','utf-8'))



##
# @brief ポート関連付け解除の関数
# @param dlg_control ダイアログオブジェクト
# @param m_port データポートオブジェクト

def DetachTC(dlg_control, m_port):
    tfcol_control = dlg_control.getControl( ControlName.AttachCBName )
    iname = str(tfcol_control.Text)
                    
    if m_port.attachports.has_key(iname) == True:
        del m_port.attachports[iname]
        if OOoRTC.calc_comp.InPorts[iname].attachports.has_key(m_port._name) == True:
            del OOoRTC.calc_comp.InPorts[iname].attachports[m_port._name]
            UpdateSaveSheet()  
            UpdateAttachPort(dlg_control, m_port)

            MyMsgBox('',OOoRTC.SetCoding(m_port._name+"と"+iname+"の関連付けを解除しました",'utf-8'))

                        
        else:
            MyMsgBox(OOoRTC.SetCoding('エラー'),OOoRTC.SetCoding('インポートの名前が正しくありません','utf-8'))
                    

##
# @class DetachListener
# @brief ポート関連付け解除ボタンのコールバック
#
class DetachListener( unohelper.Base, XActionListener):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param dlg_control ダイアログオブジェクト
    # @param _paths データポートのパス
    def __init__(self, dlg_control, _paths):
        
        self._paths = _paths
        self.dlg_control = dlg_control

    ##
    # @brief 
    # @param self 
    # @param actionEvent 
    def actionPerformed(self, actionEvent):

        if OOoRTC.calc_comp:
            
            ptlist_control = self.dlg_control.getControl( ControlName.PortCBName )
            
            
            
            if OOoRTC.calc_comp.OutPorts.has_key(str(ptlist_control.Text)) == True:
                o = OOoRTC.calc_comp.OutPorts[str(ptlist_control.Text)]
                DetachTC(self.dlg_control, o)
                return
            
        objectTree = self.dlg_control.getControl( ControlName.RTCTreeName )
        t_comp, nd = OOoRTC.JudgePort(objectTree, self._paths)
        if t_comp:
            
            for n,o in OOoRTC.calc_comp.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    DetachTC(self.dlg_control, o)
                    
                    
            
        else:
            MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('アウトポートを選択してください','utf-8'))
            return
        
        MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('削除済みです','utf-8'))

##
# @brief ポートのパラメータを設定する関数
# @param m_port データポートオブジェクト
# @param dlg_control ダイアログオブジェクト

def SetPortParam(m_port, dlg_control):
    objectControlRow = dlg_control.getControl( ControlName.RowFName )
    cfcol_control = dlg_control.getControl( ControlName.ColTName )
    cb_control = dlg_control.getControl( ControlName.LCBName )
    mlen_control = dlg_control.getControl( ControlName.LenTName )
    Stf = dlg_control.getControl( ControlName.SheetCBName )

    m_port._row = str(objectControlRow.Text)
    m_port._sn = str(Stf.Text)
    m_port._col = int(cfcol_control.Text)
    m_port._length = str(mlen_control.Text)
    m_port.update_cellName(OOoRTC.calc_comp)
    mstate = int(cb_control.State)
    if mstate == 0:
        m_port.state = False
    else:
        m_port.state = True
    UpdateSaveSheet()

##
# @class CreatePortListener
# @brief ポート作成ボタンのコールバック
#
class CreatePortListener( unohelper.Base, XActionListener):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param dlg_control ダイアログオブジェクト
    # @param _paths データポートのパス
    def __init__(self, dlg_control, _paths):
        self.nCount = 0
        
        self._paths = _paths
        self.dlg_control = dlg_control

    ##
    # @brief 
    # @param self 
    # @param actionEvent 
    def actionPerformed(self, actionEvent):
        objectControl = self.dlg_control.getControl( ControlName.TextFName )
        
        objectTree = self.dlg_control.getControl( ControlName.RTCTreeName )
        
        
        ffcol_control = self.dlg_control.getControl( ControlName.InfoTName )
                
        if OOoRTC.calc_comp:
            ptlist_control = self.dlg_control.getControl( ControlName.PortCBName )
            
            
            
            if OOoRTC.calc_comp.InPorts.has_key(str(ptlist_control.Text)) == True:
                SetPortParam(OOoRTC.calc_comp.InPorts[str(ptlist_control.Text)], self.dlg_control)
                return
            elif OOoRTC.calc_comp.OutPorts.has_key(str(ptlist_control.Text)) == True:
                SetPortParam(OOoRTC.calc_comp.OutPorts[str(ptlist_control.Text)], self.dlg_control)
                return

        dtcb_control = self.dlg_control.getControl( ControlName.DataTypeCBName )
        dt = str(dtcb_control.Text)
        if dt != "":
            F_Name = dt + str(OpenRTM_aist.uuid1())
            ptcb_control = self.dlg_control.getControl( ControlName.PortTypeCBName )
            pt = str(ptcb_control.Text)

            for d in CalcDataPort.DataType.DataTypeList:
                if dt == d:
                    dt = d
            if pt == "DataOutPort":
                CompAddOutPort(F_Name, [[pt,dt],dt], self.dlg_control, False)
                MyMsgBox('',OOoRTC.SetCoding(dt+"型のOutPortを作成しました。",'utf-8'))
            else:
                CompAddInPort(F_Name, [[pt,dt],dt], self.dlg_control, False)
                MyMsgBox('',OOoRTC.SetCoding(dt+"型のInPortを作成しました。",'utf-8'))

            UpdateSaveSheet()
            UpdateInPortList(self.dlg_control)
            UpdateDataPortList(self.dlg_control)
            return

        
        t_comp, nd = OOoRTC.JudgePort(objectTree, self._paths)
        if t_comp:
            
            for n,o in OOoRTC.calc_comp.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    SetPortParam(o, self.dlg_control)
                    
                    return
            for n,i in OOoRTC.calc_comp.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    SetPortParam(i, self.dlg_control)
                    
                    return

            
            
                                
            F_Name = t_comp[0][-2] + t_comp[0][-1]
            objectControl.setText(F_Name)
            
            profile = t_comp[1].get_port_profile()
            props = nvlist_to_dict(profile.properties)

            
            
            if props['port.port_type'] == 'DataInPort':
                CompAddOutPort(F_Name, t_comp, self.dlg_control)
            elif props['port.port_type'] == 'DataOutPort':
                CompAddInPort(F_Name, t_comp, self.dlg_control)

            MyMsgBox('',OOoRTC.SetCoding(t_comp[0][-2]+"の"+t_comp[0][-1]+"と通信するデータポートを作成しました。",'utf-8'))
            
            UpdateSaveSheet()
            
            
            ffcol_control.setText(u'作成済み')
            UpdateInPortList(self.dlg_control)
            UpdateDataPortList(self.dlg_control)

            #cfcol_control = self.dlg_control.getControl( ControlName.ColTName )
            #cfcol_control.setText(str(2))
        else:
            MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('データポートではありません','utf-8'))
        
##
# @class SetRTCTreeListener
# @brief ツリー作成ボタンのコールバック
#

class SetRTCTreeListener( unohelper.Base, XActionListener ):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param oTreeModel ツリーオブジェクト
    # @param smgr UNOコンポーネントコンテキスト
    # @param ctx UNOサービスマネージャー
    # @param dlg_control ダイアログオブジェクト
    def __init__(self, oTreeModel, smgr, ctx, dlg_control):
        
        self.oTreeModel = oTreeModel
        self.smgr = smgr
        self.ctx = ctx
        self.dlg_control = dlg_control

    ##
    # @brief 
    # @param self 
    # @param actionEvent 
    def actionPerformed(self, actionEvent):
        
        SetRTCTree(self.oTreeModel, self.smgr, self.ctx, self.dlg_control)



##
# @class TreeSelectListener
# @brief ツリーのマウスでの操作に対するコールバック
#

class TreeSelectListener( unohelper.Base, XSelectionChangeListener):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param dlg_control ダイアログオブジェクト
    # @param _paths データポートのパス
    def __init__(self, dlg_control, _paths):
        self.dlg_control = dlg_control
        self._paths = _paths

    ##
    # @brief 
    # @param self 
    # @param ev 
    def selectionChanged(self, ev):
        
        objectTree = self.dlg_control.getControl( ControlName.RTCTreeName )
        t_comp, nd = OOoRTC.JudgePort(objectTree, self._paths)

        
        ptlist_control = self.dlg_control.getControl( ControlName.PortCBName )
        ptlist_control.Text = ""

        dtcb_control = self.dlg_control.getControl( ControlName.DataTypeCBName )
        dtcb_control.setText("")
            
        if t_comp:
            for n,o in OOoRTC.calc_comp.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    UpdateTree(self.dlg_control, o)
                    return
            for n,i in OOoRTC.calc_comp.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    UpdateTree(self.dlg_control, i)
                    return
        else:
            return

        ffcol_control = self.dlg_control.getControl( ControlName.InfoTName )
        ffcol_control.setText(u'未作成')


##
# @brief ポートの削除の関数
# @param m_port データポートオブジェクト
# @param dlg_control ダイアログオブジェクト
def DelPortTC(m_port, dlg_control):
    ClearInfo(dlg_control)
    MyMsgBox('',OOoRTC.SetCoding('削除しました','utf-8'))
    UpdateSaveSheet()

    ptlist_control = self.dlg_control.getControl( ControlName.PortCBName )
    ptlist_control.Text = ""

##
# @class DeleteListener
# @brief ポート削除ボタンのコールバック
#
            
class DeleteListener( unohelper.Base, XActionListener ):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param dlg_control ダイアログオブジェクト
    # @param _paths データポートのパス
    def __init__(self, dlg_control, _paths):
        self._paths = _paths
        self.dlg_control = dlg_control

    ##
    # @brief 
    # @param self
    # @param actionEvent 
    def actionPerformed(self, actionEvent):
        objectTree = self.dlg_control.getControl( ControlName.RTCTreeName )
        

        if OOoRTC.calc_comp:
            ptlist_control = self.dlg_control.getControl( ControlName.PortCBName )
            
            
            if OOoRTC.calc_comp.InPorts.has_key(str(ptlist_control.Text)) == True:
                
                i = OOoRTC.calc_comp.InPorts[str(ptlist_control.Text)]
                
                OOoRTC.calc_comp.mRemoveInPort(i)
                DelPortTC(i, self.dlg_control)
                return
            elif OOoRTC.calc_comp.OutPorts.has_key(str(ptlist_control.Text)) == True:
                o = OOoRTC.calc_comp.OutPorts[str(ptlist_control.Text)]
                OOoRTC.calc_comp.mRemoveOutPort(o)
                DelPortTC(o, self.dlg_control)
                return

        t_comp, nd = OOoRTC.JudgePort(objectTree, self._paths)
        
        if t_comp:
            for n,o in OOoRTC.calc_comp.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    OOoRTC.calc_comp.mRemoveOutPort(o)
                    DelPortTC(o, self.dlg_control)
                    return
            for n,i in OOoRTC.calc_comp.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    OOoRTC.calc_comp.mRemoveInPort(i)
                    DelPortTC(i, self.dlg_control)
                    return
           
            
        else:
            MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('データポートを選択してください','utf-8'))
            return
        
        MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('削除済みです','utf-8'))

##
# @class SetColListener
# @brief データを書き込む列の初期化ボタンのコールバック
#

class SetColListener( unohelper.Base, XActionListener ):
    ##
    # @brief コンストラクタ
    # @param self
    # @param dlg_control
    # @param _paths
    def __init__(self, dlg_control, _paths):
        self._paths = _paths
        self.dlg_control = dlg_control

    ##
    # @brief
    # @param self
    # @param actionEvent
    def actionPerformed(self, actionEvent):
        objectTree = self.dlg_control.getControl( ControlName.RTCTreeName )
        t_comp, nd = OOoRTC.JudgePort(objectTree, self._paths)
        if t_comp:
            for n,o in OOoRTC.calc_comp.OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    o._num = int(o._col)
                    #tfcol_control = self.dlg_control.getControl( ControlName.ColTName )
                    #tfcol_control.setText(str(2))
                    return
            for n,i in OOoRTC.calc_comp.InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    i._num = int(i._col)
                    #tfcol_control = self.dlg_control.getControl( ControlName.ColTName )
                    #tfcol_control.setText(str(2))
                    return
        else:
            MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('データポートを選択してください','utf-8'))
            return
        
        MyMsgBox(OOoRTC.SetCoding('エラー','utf-8'),OOoRTC.SetCoding('削除済みです','utf-8'))

##
# @class SetColListener
# @brief データを書き込む列を全て初期化するボタンのコールバック
#

class SetAllColListener( unohelper.Base, XActionListener ):
    ##
    # @brief コンストラクタ
    # @param self
    # @param dlg_control
    def __init__(self, dlg_control):
        self.dlg_control = dlg_control

    ##
    # @brief
    # @param self
    # @param actionEvent
    def actionPerformed(self, actionEvent):
        #tfcol_control = self.dlg_control.getControl( ControlName.ColTName )
        #tfcol_control.setText(str(2))
        for n,o in OOoRTC.calc_comp.OutPorts.items():
            o._num = int(o._col)
        for n,i in OOoRTC.calc_comp.InPorts.items():
            i._num = int(i._col)
            
        
##
# @brief ダイアログ作成の関数
#
            
def SetDialog():
    dialog_name = "OOoCalcControlRTC.RTCTreeDialog"

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dp = smgr.createInstance("com.sun.star.awt.DialogProvider")
    dlg_control = dp.createDialog("vnd.sun.star.script:"+dialog_name+"?location=application")

    oTree = dlg_control.getControl(ControlName.RTCTreeName)
    

    
    
    LoadSheet()
    
    

    oTreeModel = oTree.getModel()
    
        
    
    
    SetRTCTree_listener = SetRTCTreeListener( oTreeModel, smgr, ctx, dlg_control )
    setrtctree_control = dlg_control.getControl(ControlName.CreateTreeBName)
    setrtctree_control.addActionListener(SetRTCTree_listener)

    setallcol_listener = SetAllColListener( dlg_control )
    setallcol_control = dlg_control.getControl(ControlName.SetAllLineBName)
    setallcol_control.addActionListener(setallcol_listener)
        

    tfns_control = dlg_control.getControl( ControlName.NameServerFName )
    tfns_control.setText('localhost')

    tccol_control = dlg_control.getControl( ControlName.ColTName )
    tccol_control.setText('2')

    tfcol_control = dlg_control.getControl( ControlName.RowFName )
    tfcol_control.setText('A')

    st_control = dlg_control.getControl( ControlName.SheetCBName )
    
    try:
      calc = OOoCalc()
    except NotOOoCalcException:
      return
    names = calc.sheets.getElementNames()

    for n in names:
        if n != OOoRTC.SetCoding('保存用','utf-8'):
            st_control.addItem (n, st_control.ItemCount)

    st_control.Text = names[0]


    dtcb_control = dlg_control.getControl( ControlName.DataTypeCBName )
    dtcb_control.addItem ("", dtcb_control.ItemCount)
    for n in CalcDataPort.DataType.DataTypeList:
        dtcb_control.addItem (n, dtcb_control.ItemCount)
    
    ptcb_control = dlg_control.getControl( ControlName.PortTypeCBName )
    ptcb_control.addItem ("DataInPort", ptcb_control.ItemCount)
    ptcb_control.addItem ("DataOutPort", ptcb_control.ItemCount)
    ptcb_control.setText('DataInPort')
    
    

    
    lcb_control = dlg_control.getControl( ControlName.LCBName )
    lcb_control.enableTriState( True )
    lcb_control.setState(1)



    dportl_listener = PortListListener( dlg_control )
    dportl_control = dlg_control.getControl( ControlName.PortCBName )
    dportl_control.addTextListener(dportl_listener)
    
    
    
    UpdateDataPortList(dlg_control)
    
    
    

    dlg_control.execute()
    dlg_control.dispose()




g_exportedScripts = (createOOoCalcComp, SetDialog, Start, Stop, Set_Rate)
