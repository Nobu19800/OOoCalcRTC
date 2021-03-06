# Python stubs generated by omniidl from idl/SpreadSheet.idl

import omniORB, _omnipy
from omniORB import CORBA, PortableServer
_0_CORBA = CORBA

_omnipy.checkVersion(3,0, __file__)


#
# Start of module "_GlobalIDL"
#
__name__ = "_GlobalIDL"
_0__GlobalIDL = omniORB.openModule("_GlobalIDL", r"idl/SpreadSheet.idl")
_0__GlobalIDL__POA = omniORB.openModule("_GlobalIDL__POA", r"idl/SpreadSheet.idl")


# struct StringLine
_0__GlobalIDL.StringLine = omniORB.newEmptyClass()
class StringLine (omniORB.StructBase):
    _NP_RepositoryId = "IDL:StringLine:1.0"

    def __init__(self, value):
        self.value = value

_0__GlobalIDL.StringLine = StringLine
_0__GlobalIDL._d_StringLine  = (omniORB.tcInternal.tv_struct, StringLine, StringLine._NP_RepositoryId, "StringLine", "value", (omniORB.tcInternal.tv_sequence, (omniORB.tcInternal.tv_string,0), 0))
_0__GlobalIDL._tc_StringLine = omniORB.tcInternal.createTypeCode(_0__GlobalIDL._d_StringLine)
omniORB.registerType(StringLine._NP_RepositoryId, _0__GlobalIDL._d_StringLine, _0__GlobalIDL._tc_StringLine)
del StringLine

# typedef ... StringList
class StringList:
    _NP_RepositoryId = "IDL:StringList:1.0"
    def __init__(self, *args, **kw):
        raise RuntimeError("Cannot construct objects of this type.")
_0__GlobalIDL.StringList = StringList
_0__GlobalIDL._d_StringList  = (omniORB.tcInternal.tv_sequence, omniORB.typeMapping["IDL:StringLine:1.0"], 0)
_0__GlobalIDL._ad_StringList = (omniORB.tcInternal.tv_alias, StringList._NP_RepositoryId, "StringList", (omniORB.tcInternal.tv_sequence, omniORB.typeMapping["IDL:StringLine:1.0"], 0))
_0__GlobalIDL._tc_StringList = omniORB.tcInternal.createTypeCode(_0__GlobalIDL._ad_StringList)
omniORB.registerType(StringList._NP_RepositoryId, _0__GlobalIDL._ad_StringList, _0__GlobalIDL._tc_StringList)
del StringList

# struct ValueLine
_0__GlobalIDL.ValueLine = omniORB.newEmptyClass()
class ValueLine (omniORB.StructBase):
    _NP_RepositoryId = "IDL:ValueLine:1.0"

    def __init__(self, value):
        self.value = value

_0__GlobalIDL.ValueLine = ValueLine
_0__GlobalIDL._d_ValueLine  = (omniORB.tcInternal.tv_struct, ValueLine, ValueLine._NP_RepositoryId, "ValueLine", "value", (omniORB.tcInternal.tv_sequence, omniORB.tcInternal.tv_float, 0))
_0__GlobalIDL._tc_ValueLine = omniORB.tcInternal.createTypeCode(_0__GlobalIDL._d_ValueLine)
omniORB.registerType(ValueLine._NP_RepositoryId, _0__GlobalIDL._d_ValueLine, _0__GlobalIDL._tc_ValueLine)
del ValueLine

# typedef ... ValueList
class ValueList:
    _NP_RepositoryId = "IDL:ValueList:1.0"
    def __init__(self, *args, **kw):
        raise RuntimeError("Cannot construct objects of this type.")
_0__GlobalIDL.ValueList = ValueList
_0__GlobalIDL._d_ValueList  = (omniORB.tcInternal.tv_sequence, omniORB.typeMapping["IDL:ValueLine:1.0"], 0)
_0__GlobalIDL._ad_ValueList = (omniORB.tcInternal.tv_alias, ValueList._NP_RepositoryId, "ValueList", (omniORB.tcInternal.tv_sequence, omniORB.typeMapping["IDL:ValueLine:1.0"], 0))
_0__GlobalIDL._tc_ValueList = omniORB.tcInternal.createTypeCode(_0__GlobalIDL._ad_ValueList)
omniORB.registerType(ValueList._NP_RepositoryId, _0__GlobalIDL._ad_ValueList, _0__GlobalIDL._tc_ValueList)
del ValueList

#
# Start of module "SpreadSheet"
#
__name__ = "SpreadSheet"
_0_SpreadSheet = omniORB.openModule("SpreadSheet", r"idl/SpreadSheet.idl")
_0_SpreadSheet__POA = omniORB.openModule("SpreadSheet__POA", r"idl/SpreadSheet.idl")


# interface mSpreadSheet
_0_SpreadSheet._d_mSpreadSheet = (omniORB.tcInternal.tv_objref, "IDL:SpreadSheet/mSpreadSheet:1.0", "mSpreadSheet")
omniORB.typeMapping["IDL:SpreadSheet/mSpreadSheet:1.0"] = _0_SpreadSheet._d_mSpreadSheet
_0_SpreadSheet.mSpreadSheet = omniORB.newEmptyClass()
class mSpreadSheet :
    _NP_RepositoryId = _0_SpreadSheet._d_mSpreadSheet[1]

    def __init__(self, *args, **kw):
        raise RuntimeError("Cannot construct objects of this type.")

    _nil = CORBA.Object._nil


_0_SpreadSheet.mSpreadSheet = mSpreadSheet
_0_SpreadSheet._tc_mSpreadSheet = omniORB.tcInternal.createTypeCode(_0_SpreadSheet._d_mSpreadSheet)
omniORB.registerType(mSpreadSheet._NP_RepositoryId, _0_SpreadSheet._d_mSpreadSheet, _0_SpreadSheet._tc_mSpreadSheet)

# mSpreadSheet operations and attributes
mSpreadSheet._d_get_string = (((omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0)), ((omniORB.tcInternal.tv_string,0), ), None)
mSpreadSheet._d_set_value = (((omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), omniORB.tcInternal.tv_float), (), None)
mSpreadSheet._d_get_string_range = (((omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0)), (omniORB.typeMapping["IDL:StringList:1.0"], ), None)
mSpreadSheet._d_set_value_range = (((omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), omniORB.typeMapping["IDL:ValueList:1.0"]), (), None)
mSpreadSheet._d_set_string = (((omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0)), (), None)
mSpreadSheet._d_set_string_range = (((omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), (omniORB.tcInternal.tv_string,0), omniORB.typeMapping["IDL:StringList:1.0"]), (), None)

# mSpreadSheet object reference
class _objref_mSpreadSheet (CORBA.Object):
    _NP_RepositoryId = mSpreadSheet._NP_RepositoryId

    def __init__(self):
        CORBA.Object.__init__(self)

    def get_string(self, *args):
        return _omnipy.invoke(self, "get_string", _0_SpreadSheet.mSpreadSheet._d_get_string, args)

    def set_value(self, *args):
        return _omnipy.invoke(self, "set_value", _0_SpreadSheet.mSpreadSheet._d_set_value, args)

    def get_string_range(self, *args):
        return _omnipy.invoke(self, "get_string_range", _0_SpreadSheet.mSpreadSheet._d_get_string_range, args)

    def set_value_range(self, *args):
        return _omnipy.invoke(self, "set_value_range", _0_SpreadSheet.mSpreadSheet._d_set_value_range, args)

    def set_string(self, *args):
        return _omnipy.invoke(self, "set_string", _0_SpreadSheet.mSpreadSheet._d_set_string, args)

    def set_string_range(self, *args):
        return _omnipy.invoke(self, "set_string_range", _0_SpreadSheet.mSpreadSheet._d_set_string_range, args)

    __methods__ = ["get_string", "set_value", "get_string_range", "set_value_range", "set_string", "set_string_range"] + CORBA.Object.__methods__

omniORB.registerObjref(mSpreadSheet._NP_RepositoryId, _objref_mSpreadSheet)
_0_SpreadSheet._objref_mSpreadSheet = _objref_mSpreadSheet
del mSpreadSheet, _objref_mSpreadSheet

# mSpreadSheet skeleton
__name__ = "SpreadSheet__POA"
class mSpreadSheet (PortableServer.Servant):
    _NP_RepositoryId = _0_SpreadSheet.mSpreadSheet._NP_RepositoryId


    _omni_op_d = {"get_string": _0_SpreadSheet.mSpreadSheet._d_get_string, "set_value": _0_SpreadSheet.mSpreadSheet._d_set_value, "get_string_range": _0_SpreadSheet.mSpreadSheet._d_get_string_range, "set_value_range": _0_SpreadSheet.mSpreadSheet._d_set_value_range, "set_string": _0_SpreadSheet.mSpreadSheet._d_set_string, "set_string_range": _0_SpreadSheet.mSpreadSheet._d_set_string_range}

mSpreadSheet._omni_skeleton = mSpreadSheet
_0_SpreadSheet__POA.mSpreadSheet = mSpreadSheet
omniORB.registerSkeleton(mSpreadSheet._NP_RepositoryId, mSpreadSheet)
del mSpreadSheet
__name__ = "SpreadSheet"

#
# End of module "SpreadSheet"
#
__name__ = "_GlobalIDL"


#
# End of module "_GlobalIDL"
#
__name__ = "SpreadSheet_idl"

_exported_modules = ( "SpreadSheet", "_GlobalIDL")

# The end.
