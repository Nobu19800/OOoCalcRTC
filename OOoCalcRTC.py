# -*- coding: utf-8 -*-

import optparse
import sys,os,platform
import re


if os.name == 'posix':
    sys.path += ['/usr/lib/openoffice/basis-link/program/CalcIDL', '/usr/lib/python2.6/dist-packages', '/usr/lib/python2.6/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['C:\\Program Files\\OpenOffice.org 3\\program\\CalcIDL', 'C:\\Python26\\lib\\site-packages', 'C:\\Python26\\lib\\site-packages\\rtctree\\rtmidl']



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
#from SpreadSheet_idl_example import *
from omniORB import PortableServer
import SpreadSheet, SpreadSheet__POA





#comp_num = random.randint(1,3000)
imp_id = "OOoCalcControl"# + str(comp_num)


class m_ControlName:
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
    def __init__(self):
        pass


ooocalccontrol_spec = ["implementation_id", imp_id,
                  "type_name",         imp_id,
                  "description",       "Openoffice Calc Component",
                  "version",           "0.1.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.actionLock", "1",
                  "conf.default.Red", "255",
                  "conf.default.Green", "255",
                  "conf.default.Blue", "0",
                  "conf.dataport0.port_type", "DataInPort",
                  "conf.dataport0.data_type", "TimedFloat",
                  "conf.dataport0.column", "1",
                  "conf.dataport0.start_row", "A",
                  "conf.dataport0.end_row", "A",
                  "conf.dataport0.sheetname", "sheet1",
                  "conf.__widget__.actionLock", "radio",
                  "conf.__widget__.Red", "spin",
                  "conf.__widget__.Green", "spin",
                  "conf.__widget__.Blue", "spin",
                  "conf.__widget__.file_path", "text",
                  "conf.__widget__.port_type", "radio",
                  "conf.__widget__.column", "spin",
                  "conf.__widget__.start_row", "text",
                  "conf.__widget__.end_row", "text",
                  "conf.__widget__.sheetname", "text",
                  "conf.__widget__.data_type", "radio",
                  "conf.__constraints__.actionLock", "(0,1)",
                  "conf.__constraints__.Red", "0<=x<=255",
                  "conf.__constraints__.Green", "0<=x<=255",
                  "conf.__constraints__.Blue", "0<=x<=255",
                  "conf.__constraints__.column", "1<=x<=1000",
                  "conf.__constraints__.port_type", "(DataInPort,DataOutPort)",
                  "conf.__constraints__.data_type", "(TimedDouble,TimedLong,TimedFloat,TimedShort,TimedULong,TimedUShort,TimedChar,TimedWChar,TimedBoolean,TimedOctet,TimedString,TimedWString,TimedDoubleSeq,TimedLongSeq,TimedFloatSeq,TimedShortSeq,TimedULongSeq,TimedUShortSeq,TimedCharSeq,TimedWCharSeq,TimedOctetSeq,TimedStringSeq,TimedWStringSeq)",
                  ""]

##
# コンフィギュレーションパラメータが更新されたときのコールバック
##

class MyConfigUpdateParam(OpenRTM_aist.ConfigurationSetListener):
   def __init__(self,e_rtc):
        self.m_rtc =  e_rtc
   def __call__(self, config_param_name):
        self.m_rtc.ConfigUpdate()

def SetCoding(m_str):
    if os.name == 'posix':
        return m_str
    elif os.name == 'nt':
        return m_str.decode('utf-8').encode('cp932')


##
# サービスポートSpreadSheet
##

class mSpreadSheet_i (SpreadSheet__POA.mSpreadSheet):


    def __init__(self):

        pass

    def GetCell(self, l, c, sn):
        if OOoRTC.calc_comp.calc.sheets.hasByName(sn):
            sheet = OOoRTC.calc_comp.calc.sheets.getByName(sn)
            CN = l+c
            try:
                cell = sheet.getCellRangeByName(CN)
                return cell, sheet
            except:
                pass
        else:
            return None



    # string get_string(in string l, in string c, in string sn)
    def get_string(self, l, c, sn):
        if OOoRTC.calc_comp:
            cell, sheet = self.GetCell(l,c,sn)
            if cell:
                
                return str(cell.String)

        return "error"
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # void set_value(in string l, in string c, in string sn, in float v)
    def set_value(self, l, c, sn, v):
        if OOoRTC.calc_comp:
            cell, sheet = self.GetCell(l,c,sn)
            if cell:
                cell.Value = v
                return
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None

    # StringList get_string_range(in string l1, in string c1, in string l2, in string c2, in string sn)
    def get_string_range(self, l1, c1, l2, c2, sn):
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: result

    # void set_value_range(in string l, in string c, in string sn, in ValueList v)
    def set_value_range(self, l, c, sn, v):
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None

    # void set_string(in string l, in string c, in string sn, in string v)
    def set_string(self, l, c, sn, v):
        if OOoRTC.calc_comp:
            cell, sheet = self.GetCell(l,c,sn)
            if cell:
                cell.String = v
                return
            
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None

    # void set_string_range(in string l, in string c, in string sn, in StringList v)
    def set_string_range(self, l, c, sn, v):
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        # *** Implement me
        # Must return: None





##
# OpenOffice Calcを操作するためのRTCのクラス
##

class OOoCalcControl(OpenRTM_aist.DataFlowComponentBase):
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    self.Num = 0
    self.Num2 = 0
    self._OutPorts = {}
    self._InPorts = {}
    self._ConfOutPorts = {}
    self._ConfInPorts = {}

    self._SpreadSheetPort = OpenRTM_aist.CorbaPort("SpreadSheet")
    self._spreadsheet = mSpreadSheet_i()

    try:
      self.calc = OOoCalc()
    except NotOOoCalcException:
      return


    self.conf_data_type = ["TimedFloat"]
    self.conf_port_type = ["DataInPort"]
    self.conf_column = [1]
    self.conf_start_row = ["A"]
    self.conf_end_row = ["A"]
    self.conf_sheetname = ["sheet1"]
    self.actionLock = [1]
    self.Red = [255]
    self.Green = [255]
    self.Blue = [0]
    
    
    return
  ##
  # 実行周期を設定する関数
  ##

  def m_setRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # 活性化するための関数
  ##    

  def m_activate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # 不活性化するための関数
  ##

  def m_deactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  ##
  # コンフィギュレーションパラメータによりアウトポートを追加する関数
  ##

  def m_addConfOutPort(self, name, data_type, row, col, mlen, sn, mstate, t_attachports):

    sig = m_DataType.Single
    sec = m_DataType.Sequence

    
    
    m_data_o, m_data_type =  GetDataSType(data_type)
    

    if m_data_o:
        
        
        m_outport = OpenRTM_aist.OutPort(name, m_data_o)
        self.addOutPort(name, m_outport)
        

        if m_data_type[1] == sig:
            self._ConfOutPorts[name] = MyOutPort(m_outport, m_data_o, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)
        else:
            self._ConfOutPorts[name] = MyOutPortSeq(m_outport, m_data_o, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)


  ##
  # コンフィギュレーションパラメータによりインポートを追加する関数
  ##
  
  def m_addConfInPort(self, name, data_type, row, col, mlen, sn, mstate, t_attachports):
    sig = m_DataType.Single
    sec = m_DataType.Sequence

    
    
    m_data_i, m_data_type =  GetDataSType(data_type)
    
    if m_data_i:
        
        
        m_inport = OpenRTM_aist.InPort(name, m_data_i)
        self.addInPort(name, m_inport)
        
        
        #self._InPorts[name] = MyPortObject(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)
        if m_data_type[1] == sig:
            self._ConfInPorts[name] = MyInPort(m_inport, m_data_i, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)
        else:
            self._ConfInPorts[name] = MyInPortSeq(m_inport, m_data_i, name, row, col, mlen, sn, mstate, None, m_data_type, t_attachports)


        
        m_inport.addConnectorDataListener(OpenRTM_aist.ConnectorDataListenerType.ON_BUFFER_WRITE,
                                          DataListener(self._ConfInPorts[name]))

  ##
  # アウトポート追加の関数
  # name：アウトポートの名前
  # m_inport：接続するインポート
  # row：データを書き込む行番号
  # sn；接続するインポートのパス
  ##
  def m_addOutPort(self, name, m_inport, row, col, mlen, sn, mstate, t_attachports):

    sig = m_DataType.Single
    sec = m_DataType.Sequence
    
    m_data_o, m_data_type =  GetDataType(m_inport[1])
    

    if m_data_o:
        
        m_outport = OpenRTM_aist.OutPort(name, m_data_o)
        self.addOutPort(name, m_outport)
        m_addport(m_inport[1], m_outport._objref, name)

        if m_data_type[1] == sig:
            self._OutPorts[name] = MyOutPort(m_outport, m_data_o, name, row, col, mlen, sn, mstate, m_inport, m_data_type, t_attachports)
        else:
            self._OutPorts[name] = MyOutPortSeq(m_outport, m_data_o, name, row, col, mlen, sn, mstate, m_inport, m_data_type, t_attachports)
        
        
        
                
    

            
        
  ##
  # インポート追加の関数
  # name：インポートの名前
  # m_inport：接続するアウトポート
  # row：データを書き込む行番号
  # sn；書き込むシート
  ##
        
  def m_addInPort(self, name, m_outport, row, col, mlen, sn, mstate, t_attachports):
    sig = m_DataType.Single
    sec = m_DataType.Sequence
    
    m_data_i, m_data_type =  GetDataType(m_outport[1])
    
    if m_data_i:
        m_inport = OpenRTM_aist.InPort(name, m_data_i)
        self.addInPort(name, m_inport)
        m_addport(m_inport._objref, m_outport[1], name)
        
        #self._InPorts[name] = MyPortObject(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)
        if m_data_type[1] == sig:
            self._InPorts[name] = MyInPort(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)
        else:
            self._InPorts[name] = MyInPortSeq(m_inport, m_data_i, name, row, col, mlen, sn, mstate, m_outport, m_data_type, t_attachports)


        
        m_inport.addConnectorDataListener(OpenRTM_aist.ConnectorDataListenerType.ON_BUFFER_WRITE,
                                          DataListener(self._InPorts[name]))
        
        

  ##
  # アウトポート削除の関数
  # outport：削除するアウトポート
  ##
  
  def m_removeOutComp(self, outport):
      outport._port.disconnect_all()
      self.removePort(outport._port)
      del self._OutPorts[outport._name]

  ##
  # インポート削除の関数
  # outport：削除するインポート
  ##

  def m_removeInComp(self, inport):
      inport._port.disconnect_all()
      self.removePort(inport._port)
      del self._InPorts[inport._name]


  ##
  # コンフィギュレーションパラメータが変更されたときに呼び出される関数
  ##
  
  def ConfigUpdate(self):
      
      for i in range(0, 100):
          dn = "dataport" + str(i+1)
          
          
          if self._configsets.haveConfig(dn):
              
              self._configsets.activateConfigurationSet(dn)
              self._configsets.update()

              
              tdt = ""
              tmp = None
              if self._ConfInPorts.has_key(dn):
                  tmp = self._ConfInPorts[dn]
                  tdt = "DataInPort"
              if self._ConfOutPorts.has_key(dn):
                  tmp = self._ConfOutPorts[dn]
                  tdt = "DataOutPort"

              data_type = ""
              if tmp != None:
                  profile = tmp._port.get_port_profile()
                  props = nvlist_to_dict(profile.properties)
                  data_type =  props['dataport.data_type']
                  if data_type.startswith('IDL:'):
                    data_type = data_type[4:]
                    colon = data_type.rfind(':')
                  if colon != -1:
                    data_type = data_type[:colon]

                    data_type = data_type.replace('RTC/','')
              

              if int(self.conf_column[0]) > 0 and len(self.conf_start_row[0]) > 0:
                  if tdt != None and data_type == self.conf_data_type[0]:# and self.conf_port_type[0] == tdt:
                      tmp._row = self.conf_start_row[0]
                      tmp._sn = self.conf_sheetname[0]
                      tmp._col = self.conf_column[0]
                      tmp._length = self.conf_end_row[0]
                      

                  else:
                      
                      if tmp != None:
                          tmp._port.disconnect_all()
                          self.removePort(tmp._port)

                      
                      
                      if self.conf_port_type[0] == "DataInPort":
                        self.m_addConfInPort(dn, self.conf_data_type[0], self.conf_start_row[0], int(self.conf_column[0]), self.conf_end_row[0], self.conf_sheetname[0], True, {})
                      elif self.conf_port_type[0] == "DataOutPort":
                        self.m_addConfOutPort(dn, self.conf_data_type[0], self.conf_start_row[0], int(self.conf_column[0]), self.conf_end_row[0], self.conf_sheetname[0], True, {})
                      

  ##
  # 初期化処理用コールバック関数
  ##
  
  def onInitialize(self):
    OOoRTC.calc_comp = self

    self._SpreadSheetPort.registerProvider("spreadsheet", "SpreadSheet::mSpreadSheet", self._spreadsheet)
    self.addPort(self._SpreadSheetPort)

    
    self.addConfigurationSetListener(OpenRTM_aist.ConfigurationSetListenerType.ON_SET_CONFIG_SET, MyConfigUpdateParam(self))

    self.bindParameter("data_type", self.conf_data_type, "TimedFloat")
    self.bindParameter("port_type", self.conf_port_type, "DataInPort")
    self.bindParameter("column", self.conf_column, "1")
    self.bindParameter("start_row", self.conf_start_row, "A")
    self.bindParameter("end_row", self.conf_end_row, "A")
    self.bindParameter("sheetname", self.conf_sheetname, "sheet1")
    self.bindParameter("actionLock", self.actionLock, "1")
    self.bindParameter("Red", self.Red, "255")
    self.bindParameter("Green", self.Green, "255")
    self.bindParameter("Blue", self.Blue, "0")
    
    
    
    
    
    return RTC.RTC_OK

  
  ##
  # 非活性化処理用コールバック関数
  ##
  
  def onDeactivated(self, ec_id):

    
        
    
    self.calc.document.addActionLock()
    
    for n,op in self._OutPorts.items():
        #m_row = re.split(':',op._row)
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
                cell.CellBackColor = RGB(255, 255, 255)
            except:
                pass

    for n,op in self._ConfOutPorts.items():
        #m_row = re.split(':',op._row)
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
                cell.CellBackColor = RGB(255, 255, 255)
            except:
                pass

    self.calc.document.removeActionLock()

    for n,op in self._ConfOutPorts.items():
        op._num = int(op._col)

    for n,ip in self._ConfInPorts.items():
        ip._num = int(ip._col)
    
    return RTC.RTC_OK


  ##
  # 周期処理用コールバック関数
  ##
  
  def onExecute(self, ec_id):
    
    
    

    if int(self.actionLock[0]) == 1:
        self.calc.document.addActionLock()



    for n,op in self._ConfOutPorts.items():
        op.putData(self)
            
    for n,ip in self._ConfInPorts.items():
        ip.putData(self)


            
    
    for n,op in self._OutPorts.items():
        
        if len(op.attachports) == 0:
            op.putData(self)
            
    for n,ip in self._InPorts.items():
        if len(ip.attachports) == 0:
            ip.putData(self)


            
            
            
        
            
    if int(self.actionLock[0]) == 1:
        self.calc.document.removeActionLock()

    for n,op in self._OutPorts.items():
        if len(op.attachports) != 0:
            Flag = True
            for i,j in op.attachports.items():
                if self._InPorts.has_key(j) == True:
                    #if len(self._InPorts[j].buffdata) == 0:
                    if self._InPorts[j]._port.isNew() != True:
                        Flag = False
                else:
                    Flag = False
            if Flag:
                for i,j in op.attachports.items():
                    self._InPorts[j].putData(self)
                    
                op.putData(self)
    
    return RTC.RTC_OK

  

    
  ##
  # 終了処理用コールバック関数
  ##
  def on_shutdown(self, ec_id):
      OOoRTC.calc_comp = None
      return RTC.RTC_OK


##
# 追加するポートのクラス
##


class MyPortObject:
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        
        self._port = port
        self._data = data
        self._name = name
        
        
        self._num = int(col)

        
        
        self._row = row
        self._length = mlen
        
        self._sn = sn
        self._port_a = port_a
        self._dataType = m_dataType
        self.buffdata = []
        self.attachports = t_attachports 
        self.state = mstate
        self._col = col

        self._mutex = threading.RLock()
        

    def putData(self, m_cal):
        pass

    def GetCell(self, m_cal):
        if m_cal.calc.sheets.hasByName(self._sn):
            sheet = m_cal.calc.sheets.getByName(self._sn)

            if self._length == "":
                CN = self._row + str(self._num)
            else:
                CN = self._row + str(self._num) + ':' + self._length + str(self._num)
            try:
                cell = sheet.getCellRangeByName(CN)
            except:
                pass
            
            return cell, sheet
        else:
            return None

    def updateIn(self, cell, b):
        pass

    def putIn(self, m_cal):
        m_string = m_DataType.String
        m_value = m_DataType.Value

        tmbd = []
        

        if len(self.attachports) != 0:
            self.buffdata = []
            if self._port.isNew():
                data = self._port.read()
                self.buffdata = [data.data]

        guard = OpenRTM_aist.ScopedLock(self._mutex)
        tmbd = self.buffdata[:]
        self.buffdata = []
        del guard
    
        tms = len(tmbd)-1
        if self.state:
            tms = 0
        for b in tmbd:
            sheetname = self._sn
            
            cell, sheet = self.GetCell(m_cal)

            if cell != None:
                self.updateIn(cell, b)

        
                    
                    
    def putOut(self, cell, sheet, m_cal):
        m_string = m_DataType.String
        m_value = m_DataType.Value

        cell.CellBackColor = RGB(int(m_cal.Red[0]), int(m_cal.Green[0]), int(m_cal.Blue[0]))
        
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
            if self._length == "":
                CN2 = self._row + str(t_n)
            else:
                CN2 = self._row + str(t_n) + ':' + self._length + str(t_n)

            try:
                cell2 = sheet.getCellRangeByName(CN2)
                cell2.CellBackColor = RGB(255, 255, 255)
            except:
                pass


        return val
        

class MyInPort(MyPortObject):
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        MyPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def putData(self, m_cal):
        self.putIn(m_cal)

    def updateIn(self, cell, b):
        m_string = m_DataType.String
        m_value = m_DataType.Value
        
        m_len = cell.getRangeAddress().EndColumn - cell.getRangeAddress().StartColumn
        m_len = m_len + 1
        if self._dataType[2] == m_string:
            cell.getCellByPosition(0, 0).String = b
        elif self._dataType[2] == m_value:
            cell.getCellByPosition(0, 0).Value = b
        if self.state:
            self._num = self._num + 1

        
                    

class MyInPortSeq(MyPortObject):
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        MyPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def putData(self, m_cal):
        self.putIn(m_cal)

    def updateIn(self, cell, b):
        m_string = m_DataType.String
        m_value = m_DataType.Value
        
        m_len = cell.getRangeAddress().EndColumn - cell.getRangeAddress().StartColumn
        m_len = m_len + 1

        for j in range(0, len(b)):
            if m_len > j:
                if self._dataType[2] == m_string:
                    cell.getCellByPosition(j, 0).String = b[j]
                elif self._dataType[2] == m_value:
                    cell.getCellByPosition(j, 0).Value = b[j]

        if self.state:
            self._num = self._num + 1


        

class MyOutPort(MyPortObject):
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        MyPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def putData(self, m_cal):
        cell, sheet = self.GetCell(m_cal)

        if cell != None:
            val = self.putOut(cell, sheet, m_cal)
            if self._length == "":
                if val != "":
                    self._data.data = self._dataType[0](val)
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1
            else:
                flag = True
                if val[0] != "":
                    self._data.data = self._dataType[0](val[0])
                else:
                    flag = False

                if flag:
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1

class MyOutPortSeq(MyPortObject):
    def __init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports):
        MyPortObject.__init__(self, port, data, name, row, col, mlen, sn, mstate, port_a, m_dataType, t_attachports)

    def putData(self, m_cal):
        cell, sheet = self.GetCell(m_cal)

        if cell != None:
            val = self.putOut(cell, sheet, m_cal)
            if self._length == "":
                if val != "":
                    self._data.data = self._dataType[0](val)
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1
            else:
                flag = True
                self._data.data = []
                for v in val:
                    if v != "":
                        self._data.data.append(self._dataType[0](v))
                    else:
                        flag = False

                if flag:
                    OpenRTM_aist.setTimestamp(self._data)
                    self._port.write()
                    if self.state:
                        self._num = self._num + 1

        
##
# データのタイプ
##

class m_DataType:
    Single = 0
    Sequence = 1

    String = 0
    Value = 1
    def __init__(self):
        pass

def GetDataSType(data_type):
    sig = m_DataType.Single
    sec = m_DataType.Sequence

    m_string = m_DataType.String
    m_value = m_DataType.Value
    
    if data_type == 'TimedDouble':
        dt = RTC.TimedDouble(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'TimedLong':
        dt = RTC.TimedLong(RTC.Time(0,0),0)
        return dt, [long, sig, m_value]
    elif data_type == 'TimedFloat':
        dt = RTC.TimedFloat(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'TimedInt':
        dt = RTC.TimedInt(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'TimedShort':
        dt = RTC.TimedShort(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'TimedUDouble':
        dt = RTC.TimedUDouble(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'TimedULong':
        dt = RTC.TimedULong(RTC.Time(0,0),0)
        return dt, [long, sig, m_value]
    elif data_type == 'TimedUFloat':
        dt = RTC.TimedUFloat(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'TimedUInt':
        dt = RTC.TimedUInt(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'TimedUShort':
        dt = RTC.TimedUShort(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'TimedChar':
        dt = RTC.TimedChar(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'TimedWChar':
        dt = RTC.TimedWChar(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'TimedBoolean':
        dt = RTC.TimedBoolean(RTC.Time(0,0),0)
        return dt, [bool, sig, m_value]
    elif data_type == 'TimedOctet':
        dt = RTC.TimedOctet(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'TimedString':
        dt = RTC.TimedString(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'TimedWString':
        dt = RTC.TimedWString(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'TimedDoubleSeq':
        dt = RTC.TimedDoubleSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'TimedLongSeq':
        dt = RTC.TimedLongSeq(RTC.Time(0,0),[])
        return dt, [long, sec, m_value]
    elif data_type == 'TimedFloatSeq':
        dt = RTC.TimedFloatSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'TimedIntSeq':
        dt = RTC.TimedIntSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'TimedShortSeq':
        dt = RTC.TimedShortSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'TimedUDoubleSeq':
        dt = RTC.TimedUDoubleSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'TimedULongSeq':
        dt = RTC.TimedULongSeq(RTC.Time(0,0),[])
        return dt, [long, sec, m_value]
    elif data_type == 'TimedUFloatSeq':
        dt = RTC.TimedUFloatSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'TimedUIntSeq':
        dt = RTC.TimedUIntSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'TimedUShortSeq':
        dt = RTC.TimedUShortSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'TimedCharSeq':
        dt = RTC.TimedCharSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    elif data_type == 'TimedWCharSeq':
        dt = RTC.TimedWCharSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    elif data_type == 'TimedBooleanSeq':
        dt = RTC.TimedBooleanSeq(RTC.Time(0,0),[])
        return dt, [bool, sec, m_value]
    elif data_type == 'TimedOctetSeq':
        dt = RTC.TimedOctetSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'TimedStringSeq':
        dt = RTC.TimedStringSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    elif data_type == 'TimedWStringSeq':
        dt = RTC.TimedWStringSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    else:
        return None
    

##
# データ型を返す関数
##
        
def GetDataType(m_port):
    sig = m_DataType.Single
    sec = m_DataType.Sequence

    m_string = m_DataType.String
    m_value = m_DataType.Value
    
    profile = m_port.get_port_profile()
    props = nvlist_to_dict(profile.properties)
    data_type =  props['dataport.data_type']
    if data_type.startswith('IDL:'):
        data_type = data_type[4:]
    colon = data_type.rfind(':')
    if colon != -1:
        data_type = data_type[:colon]

    data_type = data_type.replace('RTC/','')

    return GetDataSType(data_type)

    





##
# コンポーネントを活性化してCalcの操作を開始する関数
##

def Start():
    if OOoRTC.calc_comp:
        OOoRTC.calc_comp.m_activate()

##
# コンポーネントを不活性化してCalcの操作を終了する関数
##
def Stop():
    if OOoRTC.calc_comp:
        OOoRTC.calc_comp.m_deactivate()
##
# コンポーネントの実行周期を設定する関数
##

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
                  
                  OOoRTC.calc_comp.m_setRate(text)


      
      
      
##
# データが書き込まれたときに呼び出されるコールバック関数
##


class DataListener(OpenRTM_aist.ConnectorDataListenerT):
  def __init__(self, m_port):
    self.m_port = m_port

  def __del__(self):
    pass

  def __call__(self, info, cdrdata):
    data = OpenRTM_aist.ConnectorDataListenerT.__call__(self, info, cdrdata, self.m_port._data)

    guard = OpenRTM_aist.ScopedLock(self.m_port._mutex)
    self.m_port.buffdata.append(data.data)
    del guard




##
#RTCをマネージャに登録する関数
##
def OOoCalcControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=ooocalccontrol_spec)
  manager.registerFactory(profile,
                          OOoCalcControl,
                          OpenRTM_aist.Delete)
  



def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoCalcControlInit(manager)
  
  comp = manager.createComponent(imp_id)

  
        


##
# アウトポートを追加する関数
##
def CompAddOutPort(name, i_port, dlg_control):
    if OOoRTC.calc_comp != None:
        tfrow_control = dlg_control.getControl( m_ControlName.RowFName )
        tfc_control = dlg_control.getControl( m_ControlName.ColTName )
        mlen_control = dlg_control.getControl( m_ControlName.LenTName )
        tfst_control = dlg_control.getControl( m_ControlName.SheetCBName )
        cb_control = dlg_control.getControl( m_ControlName.LCBName )
        
        row = str(tfrow_control.Text)
        sn = str(tfst_control.Text)
        col = str(tfc_control.Text)
        
        mlen = str(mlen_control.Text)
        
        mstate = int(cb_control.State)
        mst = True
        if mstate == 0:
            mst = False

        
        
        OOoRTC.calc_comp.m_addOutPort(name, i_port, row, col, mlen, sn, mst, {})
        

##
# インポートを追加する関数
##

def CompAddInPort(name, o_port, dlg_control):
    if OOoRTC.calc_comp != None:
        tfrow_control = dlg_control.getControl( m_ControlName.RowFName )
        tfc_control = dlg_control.getControl( m_ControlName.ColTName )
        mlen_control = dlg_control.getControl( m_ControlName.LenTName )
        tfst_control = dlg_control.getControl( m_ControlName.SheetCBName )
        cb_control = dlg_control.getControl( m_ControlName.LCBName )
        
        row = str(tfrow_control.Text)
        sn = str(tfst_control.Text)
        col = str(tfc_control.Text)
        mlen = str(mlen_control.Text)
        mstate = int(cb_control.State)
        mst = True
        if mstate == 0:
            mst = False
        OOoRTC.calc_comp.m_addInPort(name, o_port, row, col, mlen, sn, mst, {})

##
# RTC起動の関数
##

def createOOoCalcComp():

    if OOoRTC.mgr == None:
      OOoRTC.mgr = OpenRTM_aist.Manager.init(['OOoCalcRTC.py'])
      OOoRTC.mgr.setModuleInitProc(MyModuleInit)
      OOoRTC.mgr.activateManager()
      OOoRTC.mgr.runManager(True)
    else:
      MyModuleInit(OOoRTC.mgr)

    try:
      calc = OOoCalc()
    except NotOOoCalcException:
      return

    
    
    sheetname = SetCoding('保存用')
    if calc.sheets.hasByName(sheetname):
        pass
    else:
        try:
            cnt = calc.sheets.Count
            calc.sheets.insertNewByName(sheetname, cnt)
        except unohelper.RuntimeException:
            calc.run_errordialog(title='エラー', message='')
            return
        
    MyMsgBox('',SetCoding('RTCを起動しました'))

    
    
    LoadSheet()

    
    
    return None

##
# ポートを接続する関数
##

def m_addport(obj1, obj2, c_name):

    subs_type = "Flush"

    obj1.disconnect_all()
    
    obj2.disconnect_all()

    # connect ports
    conprof = RTC.ConnectorProfile("connector0", "", [obj1,obj2], [])
    OpenRTM_aist.CORBA_SeqUtil.push_back(conprof.properties,
                                    OpenRTM_aist.NVUtil.newNV("dataport.interface_type",
                                                         "corba_cdr"))

    OpenRTM_aist.CORBA_SeqUtil.push_back(conprof.properties,
                                    OpenRTM_aist.NVUtil.newNV("dataport.dataflow_type",
                                                         "push"))

    OpenRTM_aist.CORBA_SeqUtil.push_back(conprof.properties,
                                    OpenRTM_aist.NVUtil.newNV("dataport.subscription_type",
                                                         subs_type))

    ret = obj2.connect(conprof)


##
# メッセージボックス表示の関数
# title：ウインドウのタイトル
# message：表示する文章
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
##

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)


##
# OpenOfficeを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
##

class Bridge(object):
  def __init__(self):
    self._desktop = XSCRIPTCONTEXT.getDesktop()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._frame = self._desktop.CurrentFrame
    self._window = self._frame.ContainerWindow
    self._toolkit = self._window.Toolkit
  def run_infodialog(self, title='', message=''):
    msgbox = self._toolkit.createMessageBox(self._window,uno.createUnoStruct('com.sun.star.awt.Rectangle'),'infobox',1,title,message)
    msgbox.execute()
    msgbox.dispose()

##
# ネーミングサービスへ接続する関数
##
def SetNamingServer(s_name, orb):
    
    try:
        namingserver = CorbaNaming(orb, s_name)
    except:
        MyMsgBox(SetCoding('エラー'),SetCoding('ネーミングサービスへの接続に失敗しました'))
        return None
    return namingserver

##
# ツリーで選択したアイテムがポートかどうか判定する関数
# objectTree：ダイアログのツリー
# _path：ポートのパスのリスト
##

def JudgePort(objectTree, _paths):
    m_list = []
        
    node = objectTree.getSelection()
    if node:
        parent = node.getParent()
        if parent:
            m_list.insert(0, node.getDisplayValue())
        else:
            return None
        if node.getChildCount() != 0:
            return None
    else:
        return None
            
    while(True):
        if node:
            node = node.getParent()
            if node:
                m_list.insert(0, node.getDisplayValue())
            else:
                break
        

    flag = False
    for t_comp in _paths:
        if t_comp[0] == m_list:
            return t_comp, node
            
            flag = True
            
                
    if flag == False:
        return None






##
# 各RTCのパスを取得する関数
##
def ListRecursive(context, rtclist, name, oParent, oTreeDataModel):
    
    m_blLength = 100
    
    bl = context.list(m_blLength)
    

    cont = True
    while cont:
        for i in bl[0]:
            if i.binding_type == CosNaming.ncontext:
                
                next_context = context.resolve(i.binding_name)
                name_buff = name[:]
                name.append(i.binding_name[0].id)

                if oTreeDataModel == None:
                    oChild = None
                else:
                    oChild = oTreeDataModel.createNode(i.binding_name[0].id,False)
                    oParent.appendChild(oChild)
                
                
                
                ListRecursive(next_context,rtclist,name, oChild, oTreeDataModel)
                

                name = name_buff
            elif i.binding_type == CosNaming.nobject:
                if oTreeDataModel == None:
                    oChild = None
                else:
                    oChild = oTreeDataModel.createNode(i.binding_name[0].id,False)
                    oParent.appendChild(oChild)
                
                if len(rtclist) > m_blLength:
                    break
                if i.binding_name[0].kind == 'rtc':
                    name_buff = name[:]
                    name_buff.append(i.binding_name[0].id)
                    
                    tkm = OpenRTM_aist.CorbaConsumer()
                    tkm.setObject(context.resolve(i.binding_name))
                    inobj = tkm.getObject()._narrow(RTC.RTObject)

                    try:
                        pin = inobj.get_ports()
                        for p in pin:
                            name_buff2 = name_buff[:]
                            profile = p.get_port_profile()
                            props = nvlist_to_dict(profile.properties)
                            tp_n = profile.name.split('.')[1]
                            name_buff2.append(tp_n)
                            if oTreeDataModel == None:
                                pass
                            else:
                                oChild_port = oTreeDataModel.createNode(tp_n,False)
                                oChild.appendChild(oChild_port)

                            rtclist.append([name_buff2,p])
                    except:
                        pass
                        
            else:
                pass
        if CORBA.is_nil(bl[1]):
            cont = False
        else:
            bl = i.next_n(m_blLength)


def rtc_get_rtclist(naming, rtclist, name, oParent, oTreeDataModel):  
    name_cxt = naming.getRootContext()
    ListRecursive(name_cxt,rtclist,name, oParent, oTreeDataModel)
    
    return 0







                       
##
# ポートのパスのリストを取得する関数
##
def getPathList(name):
    if OOoRTC.mgr != None:
        orb = OOoRTC.mgr._orb
        namingserver = SetNamingServer(str(name), orb)
        if namingserver:
            _path = ['/', name]
            _paths = []
            rtc_get_rtclist(namingserver, _paths, _path, None, None)
            return _paths
    return None

##
# ダイアログのツリーにネーミングサーバーのオブジェクトを登録する関数
##

def SetRTCTree(oTreeModel, smgr, ctx, dlg_control):
    oTree = dlg_control.getControl( m_ControlName.RTCTreeName )
    tfns_control = dlg_control.getControl( m_ControlName.NameServerFName )
    if OOoRTC.mgr != None:
        

               
        orb = OOoRTC.mgr._orb

        
       
        namingserver = SetNamingServer(str(tfns_control.Text), orb)
        
        
         
        if namingserver:
            
            oTreeDataModel = smgr.createInstanceWithContext("com.sun.star.awt.tree.MutableTreeDataModel", ctx)
            root = oTreeDataModel.createNode("/", False)
            oTreeDataModel.setRoot(root)
            oChild = oTreeDataModel.createNode(str(tfns_control.Text),False)
            root.appendChild(oChild)

            
            
            _path = ['/', str(tfns_control.Text)]
            _paths = []
            rtc_get_rtclist(namingserver, _paths, _path, oChild, oTreeDataModel)

            
                      
            
            oTreeModel.DataModel = oTreeDataModel

            tf1_control = dlg_control.getControl( m_ControlName.TextFName )
            tfrow_control = dlg_control.getControl( m_ControlName.RowFName )
            

            btn1_listener = CreatePortListener( dlg_control, _paths)
            cmdbtn1_control = dlg_control.getControl(m_ControlName.CreateBName)
            cmdbtn1_control.addActionListener(btn1_listener)

            delete_listener = DeleteListener(dlg_control, _paths)
            delete_control = dlg_control.getControl(m_ControlName.DeleteBName)
            delete_control.addActionListener(delete_listener)

            setCol_listener = SetColListener(dlg_control, _paths)
            setcol_control = dlg_control.getControl(m_ControlName.SetColBName)
            setcol_control.addActionListener(setCol_listener)

            attatch_listener = AttachListener( dlg_control, _paths)
            attatch_control = dlg_control.getControl(m_ControlName.AttachBName)
            attatch_control.addActionListener(attatch_listener)


            detatch_listener = DetachListener( dlg_control, _paths)
            detatch_control = dlg_control.getControl(m_ControlName.DetachBName)
            detatch_control.addActionListener(detatch_listener)

            
            
            

            oTree.addSelectionChangeListener(MySelectListener(dlg_control, _paths))



##
# OpenOffice Calcを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
##

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



##
# Cellの色の値を返すクラス
# red、green、blue：各色(0～255)
##

def RGB (red, green, blue):
    
    if red > 0xff:
      red = 0xff
    elif red < 0:
      red = 0
    if green > 0xff:
      green = 0xff
    elif green < 0:
      green = 0
    if blue > 0xff:
      blue = 0xff
    elif blue < 0:
      blue = 0
    return red * 0x010000 + green * 0x000100 + blue * 0x000001



##
# 読み込んだ保存用シートからポートを作成する関数
##

def LoadSheet():
    
    if OOoRTC.calc_comp:
        try:
          calc = OOoCalc()
        except NotOOoCalcException:
          return
        sheetname = SetCoding('保存用')
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
                    #MyMsgBox('',str(m_name[1]))
                    if m_hostname == m_name[1]:
                        pass
                    else:
                        _paths = getPathList(m_name[1])
                        m_hostname = m_name[1]
                    if _paths == None:
                        return
                    for p in _paths:
                        if p[0] == m_name:
                            F_Name = p[0][-2] + p[0][-1]
                            profile = p[1].get_port_profile()
                            props = nvlist_to_dict(profile.properties)
                            CN = 'B' + str(count)
                            cell = sheet.getCellRangeByName(CN)
                            if cell.String == '':
                                return
                            row = cell.String

                            CN = 'C' + str(count)
                            cell = sheet.getCellRangeByName(CN)
                            if cell.String == '':
                                return
                            col = cell.String

                            CN = 'D' + str(count)
                            cell = sheet.getCellRangeByName(CN)
                            #if cell.String == '':
                                #return
                            mlen = cell.String
                            
                            
                            CN = 'E' + str(count)
                            cell = sheet.getCellRangeByName(CN)
                            if cell.String == '':
                                return
                            sn = cell.String

                            CN = 'F' + str(count)
                            cell = sheet.getCellRangeByName(CN)
                            if cell.String == '':
                                return
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
                                
                            

                                
                            
                            if props['port.port_type'] == 'DataInPort':
                                OOoRTC.calc_comp.m_addOutPort(F_Name, p, row, col, mlen, sn, mstate, t_attachports)
                            elif props['port.port_type'] == 'DataOutPort':
                                OOoRTC.calc_comp.m_addInPort(F_Name, p, row, col, mlen, sn, mstate, t_attachports)
                except:
                    pass
                count = count + 1


        
                
                


##
# 作成したポートの設定を保存する関数
##
                

def UpdateSaveSheet():
    
    if OOoRTC.calc_comp:
        try:
          calc = OOoCalc()
        except NotOOoCalcException:
          return
        sheetname = SetCoding('保存用')
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
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                PortList.append(o)
            for n,i in OOoRTC.calc_comp._InPorts.items():
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
# ツリーの選択位置が変わったときに各テキストボックスの内容を変更する関数
##

def UpdateTree(dlg_control, m_port):
    
    scb_control = dlg_control.getControl( m_ControlName.SheetCBName )
    scb_control.setText(m_port._sn)
    
    tfrow_control = dlg_control.getControl( m_ControlName.RowFName )
    tfrow_control.setText(m_port._row)

    mlen_control = dlg_control.getControl( m_ControlName.LenTName )
    mlen_control.setText(m_port._length)
    

    ffcol_control = dlg_control.getControl( m_ControlName.InfoTName )
    ffcol_control.setText(u'作成済み')

    cfcol_control = dlg_control.getControl( m_ControlName.ColTName )
    cfcol_control.setText(str(m_port._col))

    cfcol_control = dlg_control.getControl( m_ControlName.LCBName )
    cfcol_control.enableTriState( True )
    if m_port.state:
        cfcol_control.setState(1)
    else:
        cfcol_control.setState(0)

    UpdateInPortList(dlg_control)
    UpdateAttachPort(dlg_control, m_port)

##
#データポートのリストを更新する関数
##

def UpdateDataPortList(dlg_control):
    if OOoRTC.calc_comp:
        dpcb_control = dlg_control.getControl( m_ControlName.PortCBName )

        dpcb_control.removeItems(0,dpcb_control.ItemCount)
        dpcb_control.Text = ""
        
        for n,i in OOoRTC.calc_comp._InPorts.items():
            dpcb_control.addItem (i._name, dpcb_control.ItemCount)

        for n,i in OOoRTC.calc_comp._OutPorts.items():
            dpcb_control.addItem (i._name, dpcb_control.ItemCount)

            
##
# インポートのリストを更新する関数
##
def UpdateInPortList(dlg_control):
    
    if OOoRTC.calc_comp:
        ipcb_control = dlg_control.getControl( m_ControlName.InPortCBName )
        ipcb_control.removeItems(0,ipcb_control.ItemCount)
        ipcb_control.Text = ""
        
        for n,i in OOoRTC.calc_comp._InPorts.items():
            
            
            ipcb_control.addItem (i._name, ipcb_control.ItemCount)
           

            
            
        


##
# 関連付けしたポートのリストを更新する関数
##
def UpdateAttachPort(dlg_control, m_port):
    
    ipcb_control = dlg_control.getControl( m_ControlName.AttachCBName )
    ipcb_control.removeItems(0,ipcb_control.ItemCount)
    ipcb_control.Text = ""
    
    for n,i in m_port.attachports.items():
        
        ipcb_control.addItem (i, ipcb_control.ItemCount)
        
    

##
# ポートを削除したときに各テキストボックスを変更する関数
##
def ClearInfo(dlg_control):
    
    ffcol_control = dlg_control.getControl( m_ControlName.InfoTName )
    ffcol_control.setText(u'未作成')

    cfcol_control = dlg_control.getControl( m_ControlName.ColTName )
    cfcol_control.setText("2")

    UpdateInPortList(dlg_control)
    UpdateDataPortList(dlg_control)



##
# データポートリストのコールバック
##
class PortListListener(unohelper.Base, XTextListener):
    def __init__(self, dlg_control):
        self.dlg_control = dlg_control
    
    def textChanged(self, actionEvent):
        UpdateInPortList(self.dlg_control)
        if OOoRTC.calc_comp:
            ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
            
            
            if OOoRTC.calc_comp._InPorts.has_key(str(ptlist_control.Text)) == True:
                UpdateTree(self.dlg_control, OOoRTC.calc_comp._InPorts[str(ptlist_control.Text)])
            elif OOoRTC.calc_comp._OutPorts.has_key(str(ptlist_control.Text)) == True:
                UpdateTree(self.dlg_control, OOoRTC.calc_comp._OutPorts[str(ptlist_control.Text)])
        

##
# ポート関連付けの関数
##
def AttachTC(dlg_control, m_port):
    
    tfcol_control = dlg_control.getControl( m_ControlName.InPortCBName )
    iname = str(tfcol_control.Text)
    
    if OOoRTC.calc_comp._InPorts.has_key(iname) == True:
                        
        m_port.attachports[iname] = iname
        OOoRTC.calc_comp._InPorts[iname].attachports[m_port._name] = m_port._name

        UpdateSaveSheet()
        UpdateAttachPort(dlg_control, m_port)

        MyMsgBox('',SetCoding(m_port._name+"と"+iname+"を関連付けしました"))

        tfcol_control.Text = iname
                    
    else:
        MyMsgBox(SetCoding('エラー'),SetCoding('インポートの名前が正しくありません'))
        return
        

##
# ポート関連付けボタンのコールバック
##
class AttachListener( unohelper.Base, XActionListener):
    def __init__(self, dlg_control, _paths):
        self._paths = _paths
        self.dlg_control = dlg_control
    def actionPerformed(self, actionEvent):
        

        if OOoRTC.calc_comp:
            
            ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
            
            
            
            if OOoRTC.calc_comp._OutPorts.has_key(str(ptlist_control.Text)) == True:
                o = OOoRTC.calc_comp._OutPorts[str(ptlist_control.Text)]
                AttachTC(self.dlg_control, o)
                return

        objectTree = self.dlg_control.getControl( m_ControlName.RTCTreeName )
        t_comp, nd = JudgePort(objectTree, self._paths)
            
        if t_comp:
            
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                
                if o._port_a[0] == t_comp[0]:
                    
                    AttachTC(self.dlg_control, o)
                    return
                    
                    
            
        else:
            MyMsgBox(SetCoding('エラー'),SetCoding('アウトポートを選択してください'))
            return
        
        MyMsgBox(SetCoding('エラー'),SetCoding('削除済みです'))



##
# ポート関連付けの関数
##
def DetachTC(dlg_control, m_port):
    tfcol_control = dlg_control.getControl( m_ControlName.AttachCBName )
    iname = str(tfcol_control.Text)
                    
    if m_port.attachports.has_key(iname) == True:
        del m_port.attachports[iname]
        if OOoRTC.calc_comp._InPorts[iname].attachports.has_key(m_port._name) == True:
            del OOoRTC.calc_comp._InPorts[iname].attachports[m_port._name]
            UpdateSaveSheet()  
            UpdateAttachPort(dlg_control, m_port)

            MyMsgBox('',SetCoding(m_port._name+"と"+iname+"の関連付けを解除しました"))

                        
        else:
            MyMsgBox(SetCoding('エラー'),SetCoding('インポートの名前が正しくありません'))
                    

##
# ポート関連付け解除ボタンのコールバック
##
class DetachListener( unohelper.Base, XActionListener):
    def __init__(self, dlg_control, _paths):
        
        self._paths = _paths
        self.dlg_control = dlg_control
    def actionPerformed(self, actionEvent):

        if OOoRTC.calc_comp:
            
            ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
            
            
            
            if OOoRTC.calc_comp._OutPorts.has_key(str(ptlist_control.Text)) == True:
                o = OOoRTC.calc_comp._OutPorts[str(ptlist_control.Text)]
                DetachTC(self.dlg_control, o)
                return
            
        objectTree = self.dlg_control.getControl( m_ControlName.RTCTreeName )
        t_comp, nd = JudgePort(objectTree, self._paths)
        if t_comp:
            
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    DetachTC(self.dlg_control, o)
                    
                    
            
        else:
            MyMsgBox(SetCoding('エラー'),SetCoding('アウトポートを選択してください'))
            return
        
        MyMsgBox(SetCoding('エラー'),SetCoding('削除済みです'))

##
# ポートのパラメータを設定する関数
##

def SetPortParam(m_port, dlg_control):
    objectControlRow = dlg_control.getControl( m_ControlName.RowFName )
    cfcol_control = dlg_control.getControl( m_ControlName.ColTName )
    cb_control = dlg_control.getControl( m_ControlName.LCBName )
    mlen_control = dlg_control.getControl( m_ControlName.LenTName )
    Stf = dlg_control.getControl( m_ControlName.SheetCBName )

    m_port._row = str(objectControlRow.Text)
    m_port._sn = str(Stf.Text)
    m_port._col = int(cfcol_control.Text)
    m_port._length = str(mlen_control.Text)
    mstate = int(cb_control.State)
    if mstate == 0:
        m_port.state = False
    else:
        m_port.state = True
    UpdateSaveSheet()

##
# ポート作成ボタンのコールバック
##
class CreatePortListener( unohelper.Base, XActionListener):
    def __init__(self, dlg_control, _paths):
        self.nCount = 0
        
        self._paths = _paths
        self.dlg_control = dlg_control

    def actionPerformed(self, actionEvent):
        objectControl = self.dlg_control.getControl( m_ControlName.TextFName )
        
        objectTree = self.dlg_control.getControl( m_ControlName.RTCTreeName )
        
        
        ffcol_control = self.dlg_control.getControl( m_ControlName.InfoTName )
                
        if OOoRTC.calc_comp:
            ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
            
            
            
            if OOoRTC.calc_comp._InPorts.has_key(str(ptlist_control.Text)) == True:
                SetPortParam(OOoRTC.calc_comp._InPorts[str(ptlist_control.Text)], self.dlg_control)
                return
            elif OOoRTC.calc_comp._OutPorts.has_key(str(ptlist_control.Text)) == True:
                SetPortParam(OOoRTC.calc_comp._OutPorts[str(ptlist_control.Text)], self.dlg_control)
                return

        
        t_comp, nd = JudgePort(objectTree, self._paths)
        if t_comp:
            
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    SetPortParam(o, self.dlg_control)
                    
                    return
            for n,i in OOoRTC.calc_comp._InPorts.items():
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

            MyMsgBox('',SetCoding(t_comp[0][-2]+"の"+t_comp[0][-1]+"と通信するデータポートを作成しました。"))
            
            UpdateSaveSheet()
            
            
            ffcol_control.setText(u'作成済み')
            UpdateInPortList(self.dlg_control)
            UpdateDataPortList(self.dlg_control)

            #cfcol_control = self.dlg_control.getControl( m_ControlName.ColTName )
            #cfcol_control.setText(str(2))
        else:
            MyMsgBox(SetCoding('エラー'),SetCoding('データポートではありません'))
        
##
# ツリー作成ボタンのコールバック
##

class SetRTCTreeListener( unohelper.Base, XActionListener ):
    def __init__(self, oTreeModel, smgr, ctx, dlg_control):
        
        self.oTreeModel = oTreeModel
        self.smgr = smgr
        self.ctx = ctx
        self.dlg_control = dlg_control

    def actionPerformed(self, actionEvent):
        
        SetRTCTree(self.oTreeModel, self.smgr, self.ctx, self.dlg_control)



##
# ツリーのマウスでの操作に対するコールバック
##

class MySelectListener( unohelper.Base, XSelectionChangeListener):
    def __init__(self, dlg_control, _paths):
        self.dlg_control = dlg_control
        self._paths = _paths
    def selectionChanged(self, ev):
        
        objectTree = self.dlg_control.getControl( m_ControlName.RTCTreeName )
        t_comp, nd = JudgePort(objectTree, self._paths)

        
        ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
        ptlist_control.Text = "" 
            
        if t_comp:
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    UpdateTree(self.dlg_control, o)
                    return
            for n,i in OOoRTC.calc_comp._InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    UpdateTree(self.dlg_control, i)
                    return
        else:
            return

        ffcol_control = self.dlg_control.getControl( m_ControlName.InfoTName )
        ffcol_control.setText(u'未作成')


##
# ポートの削除の関数
##
def DelPortTC(m_port, dlg_control):
    ClearInfo(dlg_control)
    MyMsgBox('',SetCoding('削除しました'))
    UpdateSaveSheet()

    ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
    ptlist_control.Text = ""

##
# ポート削除ボタンのコールバック
##
            
class DeleteListener( unohelper.Base, XActionListener ):
    def __init__(self, dlg_control, _paths):
        self._paths = _paths
        self.dlg_control = dlg_control

    def actionPerformed(self, actionEvent):
        objectTree = self.dlg_control.getControl( m_ControlName.RTCTreeName )
        t_comp, nd = JudgePort(objectTree, self._paths)

        if OOoRTC.calc_comp:
            ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
            
        
            if OOoRTC.calc_comp._InPorts.has_key(str(ptlist_control.Text)) == True:
                i = OOoRTC.calc_comp._InPorts[str(ptlist_control.Text)]
                OOoRTC.calc_comp.m_removeInComp(i)
                DelPortTC(i, self.dlg_control)
                return
            elif OOoRTC.calc_comp._OutPorts.has_key(str(ptlist_control.Text)) == True:
                o = OOoRTC.calc_comp._OutPorts[str(ptlist_control.Text)]
                OOoRTC.calc_comp.m_removeOutComp(o)
                DelPortTC(o, self.dlg_control)
                return
            
        if t_comp:
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    OOoRTC.calc_comp.m_removeOutComp(o)
                    DelPortTC(o, self.dlg_control)
                    return
            for n,i in OOoRTC.calc_comp._InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    OOoRTC.calc_comp.m_removeInComp(i)
                    DelPortTC(i, self.dlg_control)
                    return
           
            
        else:
            MyMsgBox(SetCoding('エラー'),SetCoding('データポートを選択してください'))
            return
        
        MyMsgBox(SetCoding('エラー'),SetCoding('削除済みです'))

##
# データを書き込む列の初期化ボタンのコールバック
##

class SetColListener( unohelper.Base, XActionListener ):
    def __init__(self, dlg_control, _paths):
        self._paths = _paths
        self.dlg_control = dlg_control

    def actionPerformed(self, actionEvent):
        objectTree = self.dlg_control.getControl( m_ControlName.RTCTreeName )
        t_comp, nd = JudgePort(objectTree, self._paths)
        if t_comp:
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                if o._port_a[0] == t_comp[0]:
                    o._num = int(o._col)
                    #tfcol_control = self.dlg_control.getControl( m_ControlName.ColTName )
                    #tfcol_control.setText(str(2))
                    return
            for n,i in OOoRTC.calc_comp._InPorts.items():
                if i._port_a[0] == t_comp[0]:
                    i._num = int(i._col)
                    #tfcol_control = self.dlg_control.getControl( m_ControlName.ColTName )
                    #tfcol_control.setText(str(2))
                    return
        else:
            MyMsgBox(SetCoding('エラー'),SetCoding('データポートを選択してください'))
            return
        
        MyMsgBox(SetCoding('エラー'),SetCoding('削除済みです'))

##
# データを書き込む列を全て初期化するボタンのコールバック
##

class SetAllColListener( unohelper.Base, XActionListener ):
    def __init__(self, dlg_control):
        self.dlg_control = dlg_control

    def actionPerformed(self, actionEvent):
        #tfcol_control = self.dlg_control.getControl( m_ControlName.ColTName )
        #tfcol_control.setText(str(2))
        for n,o in OOoRTC.calc_comp._OutPorts.items():
            o._num = int(o._col)
        for n,i in OOoRTC.calc_comp._InPorts.items():
            i._num = int(i._col)
            
        
##
# ダイアログ作成の関数
##
            
def SetDialog():
    dialog_name = "OOoCalcControlRTC.RTCTreeDialog"

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dp = smgr.createInstance("com.sun.star.awt.DialogProvider")
    dlg_control = dp.createDialog("vnd.sun.star.script:"+dialog_name+"?location=application")

    oTree = dlg_control.getControl(m_ControlName.RTCTreeName)
    

    
    
    
    
    

    oTreeModel = oTree.getModel()
    
        
    
    
    SetRTCTree_listener = SetRTCTreeListener( oTreeModel, smgr, ctx, dlg_control )
    setrtctree_control = dlg_control.getControl(m_ControlName.CreateTreeBName)
    setrtctree_control.addActionListener(SetRTCTree_listener)

    setallcol_listener = SetAllColListener( dlg_control )
    setallcol_control = dlg_control.getControl(m_ControlName.SetAllLineBName)
    setallcol_control.addActionListener(setallcol_listener)
        

    tfns_control = dlg_control.getControl( m_ControlName.NameServerFName )
    tfns_control.setText('localhost')

    tccol_control = dlg_control.getControl( m_ControlName.ColTName )
    tccol_control.setText('2')

    tfcol_control = dlg_control.getControl( m_ControlName.RowFName )
    tfcol_control.setText('A')

    st_control = dlg_control.getControl( m_ControlName.SheetCBName )
    
    try:
      calc = OOoCalc()
    except NotOOoCalcException:
      return
    names = calc.sheets.getElementNames()

    for n in names:
        if n != SetCoding('保存用'):
            st_control.addItem (n, st_control.ItemCount)
    
    
    st_control.Text = names[0]

    
    lcb_control = dlg_control.getControl( m_ControlName.LCBName )
    lcb_control.enableTriState( True )
    lcb_control.setState(1)



    dportl_listener = PortListListener( dlg_control )
    dportl_control = dlg_control.getControl( m_ControlName.PortCBName )
    dportl_control.addTextListener(dportl_listener)
    
    
    
    UpdateDataPortList(dlg_control)
    
    
    

    dlg_control.execute()
    dlg_control.dispose()




g_exportedScripts = (createOOoCalcComp, SetDialog, Start, Stop, Set_Rate)
