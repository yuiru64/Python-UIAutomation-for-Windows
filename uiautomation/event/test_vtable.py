import sys
sys.coinit_flags = 0x0
import faulthandler
faulthandler.enable()
import os
from .. import uiautomation as automation
import ctypes
import comtypes

from ctypes import CFUNCTYPE, POINTER, WINFUNCTYPE,PYFUNCTYPE
aa=automation.UIAutomationClient().UIAutomationCore
IHInterface = aa.IUIAutomationEventHandler

ole32 = ctypes.windll.LoadLibrary('ole32.dll')
import ctypes.wintypes as wintypes

def readMem(addr, typ):
    p = ctypes.pointer(ctypes.pointer(typ(0)))
    p.contents = ctypes.c_longlong(addr)
    return p.contents.contents
def writeMem(addr, value):
    p = ctypes.pointer(ctypes.pointer(type(value)(0)))
    p.contents = ctypes.c_longlong(addr)
    p.contents.contents = value
    return p.contents.contents

class _Handler(ctypes.Structure):
	_fields = [
        ('vt', ctypes.c_void_p),
        ('QueryInterface', ctypes.c_void_p),
        ('AddRef', ctypes.c_void_p),
        ('Release', ctypes.c_void_p),
        ('Hanlder', ctypes.c_void_p),
    ]

class Handler(_Handler):
    @WINFUNCTYPE(ctypes.HRESULT, POINTER(comtypes.GUID))
    def QueryInterface(pSelf, pRIID, pObj):
        lplpsz = wintypes.LPOLESTR()
        ole32.StringFromIID(pRIID, ctypes.byref(lplpsz))
        sz = str(lplpsz.value)
        if sz == "00000000-0000-0000-C000-000000000046" \
            or sz == "c270f6b5-5c69-4290-9745-7a7f97169468":
              writeMem(pSelf.contents, ctypes.c_longlong(pObj.contents))
        else:
             return 0x80004002
    @WINFUNCTYPE(wintypes.ULONG, ctypes.c_void_p)
    def AddRef(pSelf):
        return 0
    @WINFUNCTYPE(wintypes.ULONG, ctypes.c_void_p)
    def Release(pSelf):
        return 0
    @WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p)
    def Handler(pSelf, pSender):
         print(pSender)


h = Handler(
    vt=None,
    QueryInterface=Handler.QueryInterface,
    AddRef=Handler.AddRef,
    Release=Handler.Release,
    Hanlder=Handler.Handler
)
h.vt = ctypes.c_void_p(ctypes.addressof(h) + 1)


automation.UIAutomationClient().IUIAutomation.AddFocusChangedEventHandler(
    None,
    h
)
input('Press key to stop')