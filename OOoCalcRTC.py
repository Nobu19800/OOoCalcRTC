# -*- coding: cp932 -*-

import optparse
import sys,os,platform
import re

sys.path += ["C:\\Program Files\\OpenOffice.org 3\\program\\CalcIDL", 'C:\\Python26\\lib\\site-packages', 'C:\\Python26\\lib\\site-packages\\rtctree\\rtmidl']


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


import DataBase_idl



import OOoRTC
#from DataBase_idl_example import *
from omniORB import PortableServer
import DataBase, DataBase__POA





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
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  ""]




class mDataBase_i (DataBase__POA.mDataBase):


    def __init__(self):

        pass

    def GetCell(self, l, c, sn):
        if OOoRTC.calc_comp.calc.sheets.hasByName(sn):
            sheet = OOoRTC.calc_comp.calc.sheets.getByName(sn)
            CN = l+c
            cell = sheet.getCellRangeByName(CN)
            return cell, sheet
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
# OpenOffice Calc�𑀍삷�邽�߂�RTC�̃N���X
##

class OOoCalcControl(OpenRTM_aist.DataFlowComponentBase):
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    self.Num = 0
    self.Num2 = 0
    self._OutPorts = {}
    self._InPorts = {}

    self._DataBasePort = OpenRTM_aist.CorbaPort("DataBase")
    self._database = mDataBase_i()

    try:
      self.calc = OOoCalc()
    except NotOOoCalcException:
      return
    
    return
  ##
  # ���s������ݒ肷��֐�
  ##

  def m_setRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # ���������邽�߂̊֐�
  ##    

  def m_activate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # �s���������邽�߂̊֐�
  ##

  def m_deactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  ##
  # �A�E�g�|�[�g�ǉ��̊֐�
  # name�F�A�E�g�|�[�g�̖��O
  # m_inport�F�ڑ�����C���|�[�g
  # row�F�f�[�^���������ލs�ԍ�
  # sn�G�ڑ�����C���|�[�g�̃p�X
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
        
        
        sheetname = sn   
        if self.calc.sheets.hasByName(sheetname):
            sheet = self.calc.sheets.getByName(sheetname)
        
            
            
            if mlen == "":
                CN = row + str(1)
                cell = sheet.getCellRangeByName(CN)
                cell.String = str(m_inport[0])
            else:
                CN = row + str(1) + ':' + mlen + str(1)
                cell = sheet.getCellRangeByName(CN)
                cell.getCellByPosition(0, 0).String = str(m_inport[0])
                
    

            
        
  ##
  # �C���|�[�g�ǉ��̊֐�
  # name�F�C���|�[�g�̖��O
  # m_inport�F�ڑ�����A�E�g�|�[�g
  # row�F�f�[�^���������ލs�ԍ�
  # sn�G�������ރV�[�g
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
        
        sheetname = sn
        if self.calc.sheets.hasByName(sheetname):
            sheet = self.calc.sheets.getByName(sheetname)
        


            if mlen == "":
                CN = row + str(1)
                cell = sheet.getCellRangeByName(CN)
                cell.String = str(m_outport[0])
            else:
                CN = row + str(1) + ':' + mlen + str(1)
                cell = sheet.getCellRangeByName(CN)
                cell.getCellByPosition(0, 0).String = str(m_outport[0])

  ##
  # �A�E�g�|�[�g�폜�̊֐�
  # outport�F�폜����A�E�g�|�[�g
  ##
  
  def m_removeOutComp(self, outport):
      outport._port.disconnect_all()
      self.removePort(outport._port)
      del self._OutPorts[outport._name]

  ##
  # �C���|�[�g�폜�̊֐�
  # outport�F�폜����C���|�[�g
  ##

  def m_removeInComp(self, inport):
      inport._port.disconnect_all()
      self.removePort(inport._port)
      del self._InPorts[inport._name]

  ##
  # �����������p�R�[���o�b�N�֐�
  ##
  
  def onInitialize(self):
    OOoRTC.calc_comp = self

    self._DataBasePort.registerProvider("database", "DataBase::mDataBase", self._database)
    self.addPort(self._DataBasePort)
    
    
    return RTC.RTC_OK

  
  ##
  # �񊈐��������p�R�[���o�b�N�֐�
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
            cell = sheet.getCellRangeByName(CN)
            cell.CellBackColor = RGB(255, 255, 255)

    self.calc.document.removeActionLock()
    
    return RTC.RTC_OK


  ##
  # ���������p�R�[���o�b�N�֐�
  ##
  
  def onExecute(self, ec_id):
    
    
    

    
    self.calc.document.addActionLock()
    
    for n,op in self._OutPorts.items():
        
        if len(op.attachports) == 0:
            op.putData(self)
            
    for n,ip in self._InPorts.items():
        if len(ip.attachports) == 0:
            ip.putData(self)
            
            
        
            

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
  # �I�������p�R�[���o�b�N�֐�
  ##
  def on_shutdown(self, ec_id):
      OOoRTC.calc_comp = None
      return RTC.RTC_OK


##
# �ǉ�����|�[�g�̃N���X
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

            cell = sheet.getCellRangeByName(CN)
            
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

        
                    
                    
    def putOut(self, cell, sheet):
        m_string = m_DataType.String
        m_value = m_DataType.Value

        cell.CellBackColor = RGB(255, 255, 0)
        
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
            cell2 = sheet.getCellRangeByName(CN2)
            cell2.CellBackColor = RGB(255, 255, 255)


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
            val = self.putOut(cell, sheet)
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
            val = self.putOut(cell, sheet)
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
# �f�[�^�̃^�C�v
##

class m_DataType:
    Single = 0
    Sequence = 1

    String = 0
    Value = 1
    def __init__(self):
        pass

##
# �f�[�^�^��Ԃ��֐�
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

    if data_type == 'RTC/TimedDouble':
        dt = RTC.TimedDouble(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'RTC/TimedLong':
        dt = RTC.TimedLong(RTC.Time(0,0),0)
        return dt, [long, sig, m_value]
    elif data_type == 'RTC/TimedFloat':
        dt = RTC.TimedFloat(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'RTC/TimedInt':
        dt = RTC.TimedInt(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'RTC/TimedShort':
        dt = RTC.TimedShort(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'RTC/TimedUDouble':
        dt = RTC.TimedUDouble(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'RTC/TimedULong':
        dt = RTC.TimedULong(RTC.Time(0,0),0)
        return dt, [long, sig, m_value]
    elif data_type == 'RTC/TimedUFloat':
        dt = RTC.TimedUFloat(RTC.Time(0,0),0)
        return dt, [float, sig, m_value]
    elif data_type == 'RTC/TimedUInt':
        dt = RTC.TimedUInt(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'RTC/TimedUShort':
        dt = RTC.TimedUShort(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'RTC/TimedChar':
        dt = RTC.TimedChar(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'RTC/TimedWChar':
        dt = RTC.TimedWChar(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'RTC/TimedBoolean':
        dt = RTC.TimedBoolean(RTC.Time(0,0),0)
        return dt, [bool, sig, m_value]
    elif data_type == 'RTC/TimedOctet':
        dt = RTC.TimedOctet(RTC.Time(0,0),0)
        return dt, [int, sig, m_value]
    elif data_type == 'RTC/TimedString':
        dt = RTC.TimedString(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'RTC/TimedWString':
        dt = RTC.TimedWString(RTC.Time(0,0),0)
        return dt, [str, sig, m_string]
    elif data_type == 'RTC/TimedDoubleSeq':
        dt = RTC.TimedDoubleSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'RTC/TimedLongSeq':
        dt = RTC.TimedLongSeq(RTC.Time(0,0),[])
        return dt, [long, sec, m_value]
    elif data_type == 'RTC/TimedFloatSeq':
        dt = RTC.TimedFloatSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'RTC/TimedIntSeq':
        dt = RTC.TimedIntSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'RTC/TimedShortSeq':
        dt = RTC.TimedShortSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'RTC/TimedUDoubleSeq':
        dt = RTC.TimedUDoubleSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'RTC/TimedULongSeq':
        dt = RTC.TimedULongSeq(RTC.Time(0,0),[])
        return dt, [long, sec, m_value]
    elif data_type == 'RTC/TimedUFloatSeq':
        dt = RTC.TimedUFloatSeq(RTC.Time(0,0),[])
        return dt, [float, sec, m_value]
    elif data_type == 'RTC/TimedUIntSeq':
        dt = RTC.TimedUIntSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'RTC/TimedUShortSeq':
        dt = RTC.TimedUShortSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'RTC/TimedCharSeq':
        dt = RTC.TimedCharSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    elif data_type == 'RTC/TimedWCharSeq':
        dt = RTC.TimedWCharSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    elif data_type == 'RTC/TimedBooleanSeq':
        dt = RTC.TimedBooleanSeq(RTC.Time(0,0),[])
        return dt, [bool, sec, m_value]
    elif data_type == 'RTC/TimedOctetSeq':
        dt = RTC.TimedOctetSeq(RTC.Time(0,0),[])
        return dt, [int, sec, m_value]
    elif data_type == 'RTC/TimedStringSeq':
        dt = RTC.TimedStringSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    elif data_type == 'RTC/TimedWStringSeq':
        dt = RTC.TimedWStringSeq(RTC.Time(0,0),[])
        return dt, [str, sec, m_string]
    
    
    else:
        return None





##
# �R���|�[�l���g������������Calc�̑�����J�n����֐�
##

def Start():
    if OOoRTC.calc_comp:
        OOoRTC.calc_comp.m_activate()

##
# �R���|�[�l���g��s����������Calc�̑�����I������֐�
##
def Stop():
    if OOoRTC.calc_comp:
        OOoRTC.calc_comp.m_deactivate()
##
# �R���|�[�l���g�̎��s������ݒ肷��֐�
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
# �f�[�^���������܂ꂽ�Ƃ��ɌĂяo�����R�[���o�b�N�֐�
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
#RTC���}�l�[�W���ɓo�^����֐�
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
# �A�E�g�|�[�g��ǉ�����֐�
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
# �C���|�[�g��ǉ�����֐�
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
# RTC�N���̊֐�
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

    
    
    sheetname = '�ۑ��p'
    if calc.sheets.hasByName(sheetname):
        pass
    else:
        try:
            cnt = calc.sheets.Count
            calc.sheets.insertNewByName(sheetname, cnt)
        except unohelper.RuntimeException:
            calc.run_errordialog(title='�G���[', message='')
            return
        
    MyMsgBox('',u'RTC���N�����܂���')

    
    
    LoadSheet()

    
    
    return None

##
# �|�[�g��ڑ�����֐�
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
# ���b�Z�[�W�{�b�N�X�\���̊֐�
# title�F�E�C���h�E�̃^�C�g��
# message�F�\�����镶��
##

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)


##
# OpenOffice�𑀍삷�邽�߂̃N���X
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
# �l�[�~���O�T�[�r�X�֐ڑ�����֐�
##
def SetNamingServer(s_name, orb):
    
    try:
        namingserver = CorbaNaming(orb, s_name)
    except:
        MyMsgBox('�G���[',u'�l�[�~���O�T�[�r�X�ւ̐ڑ��Ɏ��s���܂���')
        return None
    return namingserver

##
# �c���[�őI�������A�C�e�����|�[�g���ǂ������肷��֐�
# objectTree�F�_�C�A���O�̃c���[
# _path�F�|�[�g�̃p�X�̃��X�g
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
# �eRTC�̃p�X���擾����֐�
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
# �|�[�g�̃p�X�̃��X�g���擾����֐�
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
# �_�C�A���O�̃c���[�Ƀl�[�~���O�T�[�o�[�̃I�u�W�F�N�g��o�^����֐�
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
# OpenOffice Calc�𑀍삷�邽�߂̃N���X
##

class OOoCalc(Bridge):
  def __init__(self):
    Bridge.__init__(self)
    if not self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
      self.run_errordialog(title='�G���[', message='���̃}�N����OpenOffice.org Calc�̒��Ŏ��s���Ă�������')
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
# Cell�̐F�̒l��Ԃ��N���X
# red�Agreen�Ablue�F�e�F(0�`255)
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
# �ǂݍ��񂾕ۑ��p�V�[�g����|�[�g���쐬����֐�
##

def LoadSheet():
    
    if OOoRTC.calc_comp:
        try:
          calc = OOoCalc()
        except NotOOoCalcException:
          return
        sheetname = '�ۑ��p'
        if calc.sheets.hasByName(sheetname):
            sheet = calc.sheets.getByName(sheetname)
            count = 1
            m_hostname = ''
            _path = []
            while True:
                CN = 'A' + str(count)
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
                count = count + 1


        
                
                


##
# �쐬�����|�[�g�̐ݒ��ۑ�����֐�
##
                

def UpdateSaveSheet():
    
    if OOoRTC.calc_comp:
        try:
          calc = OOoCalc()
        except NotOOoCalcException:
          return
        sheetname = '�ۑ��p'
        if calc.sheets.hasByName(sheetname):
            sheet = calc.sheets.getByName(sheetname)
            for i in range(1, 30):
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
                
            count = 1
            for n,o in OOoRTC.calc_comp._OutPorts.items():
                CN = 'A' + str(count)
                cell = sheet.getCellRangeByName(CN)
                pn = ''
                for j in range(0, len(o._port_a[0])):
                    if j == 0:
                        pn = o._port_a[0][j]
                    else:
                        pn = pn + ':' + o._port_a[0][j]
                cell.String = str(pn)

                CN = 'B' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = o._row

                CN = 'C' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = o._col

                CN = 'D' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = o._length

                CN = 'E' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = o._sn

                CN = 'F' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = str(o.state)

                CN = 'G' + str(count)
                cell = sheet.getCellRangeByName(CN)
                pn = ''
                tmp = 0
                
                for k,j in o.attachports.items():
                    if tmp == 0:
                        pn = j
                    else:
                        pn = pn + ':' + j
                    tmp += 1
                        
                    
                cell.String = str(pn)

                count = count + 1
            for n,i in OOoRTC.calc_comp._InPorts.items():
                CN = 'A' + str(count)
                cell = sheet.getCellRangeByName(CN)
                pn = ''
                for j in range(0, len(i._port_a[0])):
                    if j == 0:
                        pn = i._port_a[0][j]
                    else:
                        pn = pn + ':' + i._port_a[0][j]
                cell.String = str(pn)

                CN = 'B' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = i._row

                CN = 'C' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = i._col

                CN = 'D' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = i._length

                CN = 'E' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = i._sn

                CN = 'F' + str(count)
                cell = sheet.getCellRangeByName(CN)
                cell.String = str(i.state)

                CN = 'G' + str(count)
                cell = sheet.getCellRangeByName(CN)
                pn = ''
                tmp = 0
                
                for k,j in i.attachports.items():
                    if tmp == 0:
                        pn = j
                    else:
                        pn = pn + ':' + j
                    tmp += 1
                        
                    
                cell.String = str(pn)
                
                count = count + 1
        else:
            return

##
# �c���[�̑I���ʒu���ς�����Ƃ��Ɋe�e�L�X�g�{�b�N�X�̓��e��ύX����֐�
##

def UpdateTree(dlg_control, m_port):
    
    scb_control = dlg_control.getControl( m_ControlName.SheetCBName )
    scb_control.setText(m_port._sn)
    
    tfrow_control = dlg_control.getControl( m_ControlName.RowFName )
    tfrow_control.setText(m_port._row)

    mlen_control = dlg_control.getControl( m_ControlName.LenTName )
    mlen_control.setText(m_port._length)
    

    ffcol_control = dlg_control.getControl( m_ControlName.InfoTName )
    ffcol_control.setText(u'�쐬�ς�')

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
#�f�[�^�|�[�g�̃��X�g���X�V����֐�
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
# �C���|�[�g�̃��X�g���X�V����֐�
##
def UpdateInPortList(dlg_control):
    
    if OOoRTC.calc_comp:
        ipcb_control = dlg_control.getControl( m_ControlName.InPortCBName )
        ipcb_control.removeItems(0,ipcb_control.ItemCount)
        ipcb_control.Text = ""
        
        for n,i in OOoRTC.calc_comp._InPorts.items():
            
            
            ipcb_control.addItem (i._name, ipcb_control.ItemCount)
           

            
            
        


##
# �֘A�t�������|�[�g�̃��X�g���X�V����֐�
##
def UpdateAttachPort(dlg_control, m_port):
    
    ipcb_control = dlg_control.getControl( m_ControlName.AttachCBName )
    ipcb_control.removeItems(0,ipcb_control.ItemCount)
    ipcb_control.Text = ""
    
    for n,i in m_port.attachports.items():
        
        ipcb_control.addItem (i, ipcb_control.ItemCount)
        
    

##
# �|�[�g���폜�����Ƃ��Ɋe�e�L�X�g�{�b�N�X��ύX����֐�
##
def ClearInfo(dlg_control):
    
    ffcol_control = dlg_control.getControl( m_ControlName.InfoTName )
    ffcol_control.setText(u'���쐬')

    cfcol_control = dlg_control.getControl( m_ControlName.ColTName )
    cfcol_control.setText("2")

    UpdateInPortList(dlg_control)
    UpdateDataPortList(dlg_control)



##
# �f�[�^�|�[�g���X�g�̃R�[���o�b�N
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
# �|�[�g�֘A�t���̊֐�
##
def AttachTC(dlg_control, m_port):
    
    tfcol_control = dlg_control.getControl( m_ControlName.InPortCBName )
    iname = str(tfcol_control.Text)
    
    if OOoRTC.calc_comp._InPorts.has_key(iname) == True:
                        
        m_port.attachports[iname] = iname
        OOoRTC.calc_comp._InPorts[iname].attachports[m_port._name] = m_port._name

        UpdateSaveSheet()
        UpdateAttachPort(dlg_control, m_port)

        MyMsgBox('',m_port._name+"��"+iname+"���֘A�t�����܂���")

        tfcol_control.Text = iname
                    
    else:
        MyMsgBox('�G���[',u'�C���|�[�g�̖��O������������܂���')
        return
        

##
# �|�[�g�֘A�t���{�^���̃R�[���o�b�N
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
            MyMsgBox('�G���[',u'�A�E�g�|�[�g��I�����Ă�������')
            return
        
        MyMsgBox('�G���[',u'�폜�ς݂ł�')



##
# �|�[�g�֘A�t���̊֐�
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

            MyMsgBox('',m_port._name+"��"+iname+"�̊֘A�t�����������܂���")

                        
        else:
            MyMsgBox('�G���[',u'�C���|�[�g�̖��O������������܂���')
                    

##
# �|�[�g�֘A�t�������{�^���̃R�[���o�b�N
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
            MyMsgBox('�G���[',u'�A�E�g�|�[�g��I�����Ă�������')
            return
        
        MyMsgBox('�G���[',u'�폜�ς݂ł�')

##
# �|�[�g�̃p�����[�^��ݒ肷��֐�
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
# �|�[�g�쐬�{�^���̃R�[���o�b�N
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

            MyMsgBox('',t_comp[0][-2]+"��"+t_comp[0][-1]+"�ƒʐM����f�[�^�|�[�g���쐬���܂����B")
            
            UpdateSaveSheet()
            
            
            ffcol_control.setText(u'�쐬�ς�')
            UpdateInPortList(self.dlg_control)
            UpdateDataPortList(self.dlg_control)

            #cfcol_control = self.dlg_control.getControl( m_ControlName.ColTName )
            #cfcol_control.setText(str(2))
        else:
            MyMsgBox('�G���[',u'�f�[�^�|�[�g�ł͂���܂���')
        
##
# �c���[�쐬�{�^���̃R�[���o�b�N
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
# �c���[�̃}�E�X�ł̑���ɑ΂���R�[���o�b�N
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
        ffcol_control.setText(u'���쐬')


##
# �|�[�g�̍폜�̊֐�
##
def DelPortTC(m_port, dlg_control):
    ClearInfo(dlg_control)
    MyMsgBox('',u'�폜���܂���')
    UpdateSaveSheet()

    ptlist_control = self.dlg_control.getControl( m_ControlName.PortCBName )
    ptlist_control.Text = ""

##
# �|�[�g�폜�{�^���̃R�[���o�b�N
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
            MyMsgBox('�G���[',u'�f�[�^�|�[�g��I�����Ă�������')
            return
        
        MyMsgBox('�G���[',u'�폜�ς݂ł�')

##
# �f�[�^���������ޗ�̏������{�^���̃R�[���o�b�N
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
            MyMsgBox('�G���[',u'�f�[�^�|�[�g��I�����Ă�������')
            return
        
        MyMsgBox('�G���[',u'�폜�ς݂ł�')

##
# �f�[�^���������ޗ��S�ď���������{�^���̃R�[���o�b�N
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
# �_�C�A���O�쐬�̊֐�
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
        if n != u'�ۑ��p':
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
