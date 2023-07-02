import sys
sys.coinit_flags = 0x0
# import faulthandler
# faulthandler.enable()
import ctypes
import os
from .. import uiautomation as automation

from ctypes import CFUNCTYPE, POINTER
UIAutomationCore = automation.UIAutomationClient().UIAutomationCore
IUIAutomation = automation.UIAutomationClient().IUIAutomation

from .automationtypes import EventID, PropertyId, TreeScope
# from automationtypes import VARIANT # A toy implement
from comtypes.automation import VARIANT
from comtypes._safearray import SAFEARRAY

class HandlerType:
    HandlerType_AutomationEventHandler = 0
    HandlerType_FocusChangedEventHandler = 1
    HandlerType_PropertyChangedEventHandler = 2
    HandlerType_StructureChangedHandler = 3

EVENTID = ctypes.c_int
PROPERTYID = ctypes.c_int
StructureChangeType_T = ctypes.c_int

P_HandleAutomationEvent = CFUNCTYPE(None, POINTER(UIAutomationCore.IUIAutomationElement), EVENTID)
P_HandleFocusChangedEvent = CFUNCTYPE(None, POINTER(UIAutomationCore.IUIAutomationElement))
P_HandlePropertyChangedEvent = CFUNCTYPE(None, POINTER(UIAutomationCore.IUIAutomationElement), PROPERTYID, VARIANT)
P_HandleStructureChangedEvent = CFUNCTYPE(None, POINTER(UIAutomationCore.IUIAutomationElement), StructureChangeType_T, POINTER(SAFEARRAY))

binPath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin")
os.environ["PATH"] = binPath + os.pathsep + os.environ["PATH"]
os.add_dll_directory(binPath)

UIEventHandler = ctypes.cdll.LoadLibrary("UIEventHandler.dll")
UIEventHandler.NewHandler.restype = POINTER(UIAutomationCore.IUIAutomationEventHandler)
UIEventHandler.NewHandler.argtypes = [ctypes.c_int, ctypes.c_void_p]
UIEventHandler.DeleteHandler.restype = None
UIEventHandler.DeleteHandler.argtypes = [ctypes.c_void_p]
UIEventHandler.ReleaseHandler.restype = None
UIEventHandler.ReleaseHandler.argtypes = [ctypes.c_void_p]

_handlers = {}
_handler_types = {}

from types import FunctionType    
def RemoveEventHandler(callback: FunctionType):
    type = _handler_types[callback]
    if type == HandlerType.HandlerType_FocusChangedEventHandler:
        c_callback, h = _handlers[callback]
        IUIAutomation.RemoveFocusChangedEventHandler(h)
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
        del _handlers[callback]
    elif type == HandlerType.HandlerType_AutomationEventHandler:
        c_callback, eventId, element, h = _handlers[callback]
        """
        HRESULT RemoveAutomationEventHandler(
        [in] EVENTID                   eventId,
        [in] IUIAutomationElement      *element,
        [in] IUIAutomationEventHandler *handler
        );
        """
        IUIAutomation.RemoveAutomationEventHandler(eventId, element, h)
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
        del _handlers[callback]
    elif type == HandlerType.HandlerType_PropertyChangedEventHandler:
        c_callback, element, h = _handlers[callback]
        IUIAutomation.RemovePropertyChangedEventHandler(element, h)
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
        del _handlers[callback]
    elif type == HandlerType.HandlerType_StructureChangedHandler:
        c_callback, element, h = _handlers[callback]
        IUIAutomation. RemoveStructureChangedEventHandler(element, h)
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
        del _handlers[callback]
    else:
        raise RuntimeError
    print('Released event handler {}'.format(callback))

def RemoveAllEventHandler():
    keys = list(_handlers.keys())
    if len(keys) == 0:
        return
    for k in keys:
        RemoveEventHandler(k)
    print('Released all event handlers')
import atexit
atexit.register(RemoveAllEventHandler)

def AddFocusChangedEventHandler(callback):
    c_callback = P_HandleFocusChangedEvent(callback)
    h = UIEventHandler.NewHandler(HandlerType.HandlerType_FocusChangedEventHandler, c_callback)

    _handler_types[callback] = HandlerType.HandlerType_FocusChangedEventHandler
    _handlers[callback] = c_callback, h
    return IUIAutomation.AddFocusChangedEventHandler(None, h)

def AddAutomationEventHandler(callback, eventId, element: automation.Control, treeScope):
    """
    :param eventId: See :class:`EventId`
    :param treeScope: See :class:`TreeScope`
    """
    c_callback = P_HandleAutomationEvent(callback)
    h = UIEventHandler.NewHandler(HandlerType.HandlerType_AutomationEventHandler, c_callback)
   
    _handler_types[callback] = HandlerType.HandlerType_AutomationEventHandler
    _handlers[callback] = c_callback, eventId, element.Element, h
    return IUIAutomation.AddAutomationEventHandler(
        eventId, element.Element, treeScope, None, h
    )

def AddPropertyChangedEventHandler(callback, properties, element: automation.Control, treeScope):
        c_properties_type = ctypes.c_int * len(properties)
        c_properties = c_properties_type(*properties)
        c_callback = P_HandlePropertyChangedEvent(callback)
        h = UIEventHandler.NewHandler(HandlerType.HandlerType_PropertyChangedEventHandler, c_callback)

        _handler_types[callback] = HandlerType.HandlerType_PropertyChangedEventHandler
        _handlers[callback] = c_callback, element.Element, h
        return IUIAutomation.AddPropertyChangedEventHandlerNativeArray(
            element.Element, treeScope, None, h,
            ctypes.cast(ctypes.byref(c_properties), POINTER(ctypes.c_long)),
            len(properties)
        )

def AddStructureChangedEventHandler(callback, element: automation.Control, treeScope):
        c_callback = P_HandleStructureChangedEvent(callback)
        h = UIEventHandler.NewHandler(HandlerType.HandlerType_StructureChangedHandler, c_callback)

        _handler_types[callback] = HandlerType.HandlerType_StructureChangedHandler
        _handlers[callback] = c_callback, element.Element, h
        return IUIAutomation.AddStructureChangedEventHandler(
            element.Element, treeScope, None, h,
        )


if __name__ == '__main__':
    if len(sys.argv) <= 1:
        root, _ = os.path.splitext(os.path.basename(sys.argv[0]))
        print('python -m uiautomation.event.{} <focus|auto|prop|struc>'.format(root))
        exit(1)
    action = sys.argv[1]
    if action == 'focus':
        """
        HRESULT AddFocusChangedEventHandler(
        [in] IUIAutomationCacheRequest             *cacheRequest,
        [in] IUIAutomationFocusChangedEventHandler *handler
        );
        HRESULT HandleFocusChangedEvent(
        [in] IUIAutomationElement *sender
        );
        """
        def callback(pSender):
            print(pSender)

        c_callback = P_HandleFocusChangedEvent(callback)

        h = UIEventHandler.NewHandler(HandlerType.HandlerType_FocusChangedEventHandler, c_callback)
        IUIAutomation.AddFocusChangedEventHandler(None, h)
        input('Press key to stop')
        IUIAutomation.RemoveFocusChangedEventHandler(h)
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
    elif action == 'auto':
        """
        HRESULT AddAutomationEventHandler(
            [in] EVENTID                   eventId,
            [in] IUIAutomationElement      *element,
                    TreeScope                 scope,
            [in] IUIAutomationCacheRequest *cacheRequest,
            [in] IUIAutomationEventHandler *handler
        );
        HRESULT HandleAutomationEvent(
        [in] IUIAutomationElement *sender,
        [in] EVENTID              eventId
        );
        """
        def callback(pSender, eventID):
            print(pSender, eventID)

        c_callback = P_HandleAutomationEvent(callback)
        calcWindow = automation.WindowControl(searchDepth = 1, Name = '计算器')
        calcWindow.Maximize()
        calcWindow.SendKeys('{Alt}2')
        h = UIEventHandler.NewHandler(HandlerType.HandlerType_AutomationEventHandler, c_callback)
        IUIAutomation.AddAutomationEventHandler(
            EventID.UIA_StructureChangedEventId,
            calcWindow.Element,
            TreeScope.TreeScope_Subtree,
            None,
            h
        )
        input('Press key to stop')
        IUIAutomation.RemoveAllEventHandlers()
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
    elif action == 'prop':
        """
        HRESULT AddPropertyChangedEventHandlerNativeArray(
        [in] IUIAutomationElement                     *element,
            TreeScope                                scope,
        [in] IUIAutomationCacheRequest                *cacheRequest,
        [in] IUIAutomationPropertyChangedEventHandler *handler,
        [in] PROPERTYID                               *propertyArray,
        [in] int                                      propertyCount
        );
        HRESULT HandlePropertyChangedEvent(
        [in] IUIAutomationElement *sender,
        [in] PROPERTYID           propertyId,
        [in] VARIANT              newValue
        );
        """
        automation.SetDpiAwareness(True)
        import time
        time.sleep(3)
        calcWindow = automation.WindowControl(searchDepth = 1, Name = '计算器')

        point = ctypes.wintypes.POINT(0, 0)
        ctypes.windll.user32.GetPhysicalCursorPos(ctypes.byref(point))
        p = point.x, point.y
        el = automation.ControlFromPoint(*p)

        def callback(pSender, proprtId, newValue: VARIANT):
            print(pSender, proprtId, newValue)
            # print(newValue.__VARIANT_NAME_1.__VARIANT_NAME_2.__VARIANT_NAME_3.bstrval)

        c_callback = P_HandlePropertyChangedEvent(callback)


        arr_ty = ctypes.c_int * 2
        arr = arr_ty(PropertyId.UIA_ToggleToggleStatePropertyId, PropertyId.UIA_NamePropertyId)

        h = UIEventHandler.NewHandler(HandlerType.HandlerType_PropertyChangedEventHandler, c_callback)

        element = el.Element
        IUIAutomation.AddPropertyChangedEventHandlerNativeArray(
            element,
            TreeScope.TreeScope_Subtree,
            None,
            h,
            ctypes.cast(ctypes.byref(arr), POINTER(ctypes.c_long)),
            2
        )
        input('Press key to stop')
        IUIAutomation.RemovePropertyChangedEventHandler(element, h)
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
    elif action == 'struc':
        """
        HRESULT AddStructureChangedEventHandler(
        [in] IUIAutomationElement                      *element,
            TreeScope                                 scope,
        [in] IUIAutomationCacheRequest                 *cacheRequest,
        [in] IUIAutomationStructureChangedEventHandler *handler
        );
        HRESULT HandleStructureChangedEvent(
        [in] IUIAutomationElement           *sender,
        [in] StructureChangeType changeType unnamedParam2,
        [in] SAFEARRAY                      *runtimeId
        );
        """

        import time
        calcWindow = automation.WindowControl(searchDepth = 1, Name = '计算器')

        newValue_ = None
        def callback(pSender, changeType, runtimeId):
            global newValue_
            print(pSender, changeType, runtimeId)
            newValue_ = runtimeId

        c_callback = P_HandleStructureChangedEvent(callback)

        h = UIEventHandler.NewHandler(HandlerType.HandlerType_StructureChangedHandler, c_callback)
        element = calcWindow.Element
        IUIAutomation.AddStructureChangedEventHandler(
            element,
            TreeScope.TreeScope_Subtree,
            None,
            h,
        )
        input('Press key to stop')
        IUIAutomation. RemoveStructureChangedEventHandler(element, h)
        UIEventHandler.ReleaseHandler(ctypes.cast(h, ctypes.c_void_p))
    else:
         raise RuntimeError('Invalied argument')