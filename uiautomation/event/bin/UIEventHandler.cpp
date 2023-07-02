#include <windows.h>
#include <stdio.h>
#include <new.h>
#include <UIAutomation.h>
#pragma once

using namespace std;

#ifdef _MSC_VER
#define DLL_EXPORT __declspec(dllexport)
#else
#define DLL_EXPORT
#endif

#define DEBUG
#ifdef DEBUG
    #define LOG wprintf
#else
    #define LOG //
#endif

using P_HandleAutomationEvent = void (*)(IUIAutomationElement*, EVENTID);
using P_HandleFocusChangedEvent = void (*)(IUIAutomationElement*);
using P_HandlePropertyChangedEvent = void (*)(IUIAutomationElement*, PROPERTYID, VARIANT);
using P_HandleStructureChangedEvent = void (*)(IUIAutomationElement*, StructureChangeType, SAFEARRAY*);

// Defines an event handler for general UI Automation events. It listens for
// tooltip and window creation and destruction events. 
class AutomationEventHandler:
    public IUIAutomationEventHandler
{
private:
    LONG _refCount;

public:
    int _eventCount;
    P_HandleAutomationEvent callback;

    // Constructor.
    AutomationEventHandler(P_HandleAutomationEvent callback): _refCount(1), _eventCount(0), callback(callback) 
    {
        LOG(L">> AutomationEventHandler: new callback pointer %p\n", callback);
    }

    // IUnknown methods.
    ULONG STDMETHODCALLTYPE AddRef() 
    {
        ULONG ret = InterlockedIncrement(&_refCount);
        return ret;
    }

    ULONG STDMETHODCALLTYPE Release() 
    {
        ULONG ret = InterlockedDecrement(&_refCount);
        if (ret == 0) 
        {
            wprintf(L">> AutomationEventHandler: destruction\n");
            delete this;
            return 0;
        }
        LOG(L">> AutomationEventHandler: Released, ref count %d\n", _refCount);
        return ret;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface) 
    {
        if (riid == __uuidof(IUnknown)) 
            *ppInterface=static_cast<IUIAutomationEventHandler*>(this);
        else if (riid == __uuidof(IUIAutomationEventHandler)) 
            *ppInterface=static_cast<IUIAutomationEventHandler*>(this);
        else 
        {
            *ppInterface = NULL;
            return E_NOINTERFACE;
        }
        this->AddRef();
        return S_OK;
    }

    // IUIAutomationEventHandler methods
    HRESULT STDMETHODCALLTYPE HandleAutomationEvent(IUIAutomationElement * pSender, EVENTID eventID)
    {
        _eventCount++;
        switch (eventID) 
        {
            case UIA_ToolTipOpenedEventId:
                LOG(L">> AutomationEventHandler: ToolTipOpened Received! (count: %d)\n", _eventCount);
                break;
            case UIA_ToolTipClosedEventId:
                LOG(L">> AutomationEventHandler: ToolTipClosed Received! (count: %d)\n", _eventCount);
                break;
            case UIA_Window_WindowOpenedEventId:
                LOG(L">> AutomationEventHandler: WindowOpened Received! (count: %d)\n", _eventCount);
                break;
            case UIA_Window_WindowClosedEventId:
                LOG(L">> AutomationEventHandler: WindowClosed Received! (count: %d)\n", _eventCount);
                break;
            default:
                LOG(L">> AutomationEventHandler: (%d) Received! (count: %d)\n", eventID, _eventCount);
                break;
        }
        this->callback(pSender, eventID);
        return S_OK;
    }
};

// Defines an event handler for focus changed events and starts
// listening to them.
class FocusChangedEventHandler :
    public IUIAutomationFocusChangedEventHandler
{
private:
    LONG _refCount;

public:
    int _eventCount;
    P_HandleFocusChangedEvent callback;

    //Constructor.
    FocusChangedEventHandler(P_HandleFocusChangedEvent callback) :
        _refCount(1), _eventCount(0), callback(callback)
    {
        LOG(L">> FocusChangedEventHandler: new callback pointer %p\n", callback);
    }

    //IUnknown methods.
    ULONG STDMETHODCALLTYPE AddRef()
    {
        ULONG ret = InterlockedIncrement(&_refCount);
        return ret;
    }

    ULONG STDMETHODCALLTYPE Release()
    {
        ULONG ret = InterlockedDecrement(&_refCount);
        if (ret == 0)
        {
            wprintf(L">> FocusChangedEventHandler: destruction\n");
            delete this;
            return 0;
        }
        LOG(L">> FocusChangedEventHandler: Released, ref count %d\n", _refCount);
        return ret;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
    {
        if (riid == __uuidof(IUnknown))
            *ppInterface = static_cast<IUIAutomationFocusChangedEventHandler*>(this);
        else if (riid == __uuidof(IUIAutomationFocusChangedEventHandler))
            *ppInterface = static_cast<IUIAutomationFocusChangedEventHandler*>(this);
        else
        {
            *ppInterface = NULL;
            return E_NOINTERFACE;
        }
        this->AddRef();
        return S_OK;
    }

    // IUIAutomationFocusChangedEventHandler methods.
    HRESULT STDMETHODCALLTYPE HandleFocusChangedEvent(IUIAutomationElement * pSender)
    {
        _eventCount++;
        LOG(L">> FocusChangedEventHandler: Received! (count: %d)\n", _eventCount);
        this->callback(pSender);
        return S_OK;
    }
};

// Defines an event handler for property-changed events, and listens for
// ToggleState property changes on the element specifies by the user.

class PropertyChangedEventHandler:
    public IUIAutomationPropertyChangedEventHandler
{
private:
    LONG _refCount;

public:
    int _eventCount;
    P_HandlePropertyChangedEvent callback;

    //Constructor.
    PropertyChangedEventHandler(P_HandlePropertyChangedEvent callback): _refCount(1), _eventCount(0), callback(callback)
    {
        LOG(L">> PropertyChangedEventHandler: new callback pointer %p\n", callback);
    }

    //IUnknown methods.
    ULONG STDMETHODCALLTYPE AddRef() 
    {
        ULONG ret = InterlockedIncrement(&_refCount);
        return ret;
    }


    ULONG STDMETHODCALLTYPE Release()
    {
        ULONG ret = InterlockedDecrement(&_refCount);
        if (ret == 0)
        {
            LOG(L">> PropertyChangedEventHandler: destruction\n");
            delete this;
            return 0;
        }
        LOG(L">> PropertyChangedEventHandler: Released, ref count %d\n", _refCount);
        return ret;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface) 
    {
        if (riid == __uuidof(IUnknown))
            *ppInterface=static_cast<IUIAutomationPropertyChangedEventHandler*>(this);
        else if (riid == __uuidof(IUIAutomationPropertyChangedEventHandler)) 
            *ppInterface=static_cast<IUIAutomationPropertyChangedEventHandler*>(this);
        else 
        {
            *ppInterface = NULL;
            return E_NOINTERFACE;
        }
        this->AddRef();
        return S_OK;
    }

    // IUIAutomationPropertyChangedEventHandler methods.
    HRESULT STDMETHODCALLTYPE HandlePropertyChangedEvent(IUIAutomationElement* pSender, PROPERTYID propertyID, VARIANT newValue) 
    {
        _eventCount++;
        if (propertyID == UIA_ToggleToggleStatePropertyId) 
            LOG(L">> PropertyChangedEventHandler: Property ToggleState changed! ");
        else 
            LOG(L">> PropertyChangedEventHandler: Property (%d) changed! ", propertyID);

        if (newValue.vt == VT_I4) 
            LOG(L"(0x%0.8x) ", newValue.lVal);

        LOG(L"(count: %d)\n", _eventCount);
        this->callback(pSender, propertyID, newValue);
        return S_OK;
    }
};

class StructureChangedHandler:
    public IUIAutomationStructureChangedEventHandler
{
private:
    LONG _refCount;

public:
    int _eventCount;
    P_HandleStructureChangedEvent callback;
    // Constructor.
    StructureChangedHandler(P_HandleStructureChangedEvent callback): _refCount(1), _eventCount(0), callback(callback)
    {
        LOG(L">> StructureChangedHandler: new callback pointer %p\n", callback);
    }

    // IUnknown methods.
    ULONG STDMETHODCALLTYPE AddRef() 
    {
        ULONG ret = InterlockedIncrement(&_refCount);
        return ret;
    }

    ULONG STDMETHODCALLTYPE Release() 
    {
        ULONG ret = InterlockedDecrement(&_refCount);
        if (ret == 0) 
        {
              LOG(L">> StructureChangedHandler: destruction\n");
            delete this;
            return 0;
        }
        LOG(L">> StructureChangedHandler: Released, ref count %d\n", _refCount);
        return ret;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface) 
    {
        if (riid == __uuidof(IUnknown))
            *ppInterface=static_cast<IUIAutomationStructureChangedEventHandler*>(this);
        else if (riid == __uuidof(IUIAutomationStructureChangedEventHandler)) 
            *ppInterface=static_cast<IUIAutomationStructureChangedEventHandler*>(this);
        else 
        {
            *ppInterface = NULL;
            return E_NOINTERFACE;
        }
        this->AddRef();
        return S_OK;
    }

    // IUIAutomationStructureChangedEventHandler methods
    HRESULT STDMETHODCALLTYPE HandleStructureChangedEvent(IUIAutomationElement* pSender, StructureChangeType changeType, SAFEARRAY* pRuntimeID) {
        _eventCount++;
        switch (changeType) 
        {
            case StructureChangeType_ChildAdded:
                wprintf(L">> StructureChangedHandler: ChildAdded! (count: %d)\n", _eventCount);
                break;
            case StructureChangeType_ChildRemoved:
                wprintf(L">> StructureChangedHandler: ChildRemoved! (count: %d)\n", _eventCount);
                break;
            case StructureChangeType_ChildrenInvalidated:
                wprintf(L">> StructureChangedHandler: ChildrenInvalidated! (count: %d)\n", _eventCount);
                break;
            case StructureChangeType_ChildrenBulkAdded:
                wprintf(L">> StructureChangedHandler: ChildrenBulkAdded! (count: %d)\n", _eventCount);
                break;
            case StructureChangeType_ChildrenBulkRemoved:
                wprintf(L">> StructureChangedHandler: ChildrenBulkRemoved! (count: %d)\n", _eventCount);
                break;
            case StructureChangeType_ChildrenReordered:
                wprintf(L">> StructureChangedHandler: ChildrenReordered! (count: %d)\n", _eventCount);
                break;
        }
        this->callback(pSender, changeType, pRuntimeID);
        return S_OK;
    }
};


enum HandlerType {
    HandlerType_AutomationEventHandler = 0,
    HandlerType_FocusChangedEventHandler = 1,
    HandlerType_PropertyChangedEventHandler = 2,
    HandlerType_StructureChangedHandler = 3,
};

extern "C" {

    DLL_EXPORT void * NewHandler(HandlerType handlerType, void * callback) {
        // size_t memSize = sizeof(FocusChangedEventHandler);

        // HGLOBAL hMem = GlobalAlloc(0x40, memSize);
        // FocusChangedEventHandler* pInterface = static_cast<FocusChangedEventHandler*>(GlobalLock(hMem));
        // wprintf(L">> Test callback\n");
        // callback((IUIAutomationElement *) NULL);
        // wprintf(L">> Test callback end\n");
        // outercallback = callback;
        // return new(pInterface) FocusChangedEventHandler(&call);
        // wprintf(L">> Test callback\n");
        // callback((IUIAutomationElement *) NULL);
        // wprintf(L">> Test callback end\n");
        // outercallback = callback;
        switch (handlerType) {
            case HandlerType::HandlerType_AutomationEventHandler:
                return new(std::nothrow) AutomationEventHandler(
                    static_cast<P_HandleAutomationEvent>(callback)
                );
                break;
            case HandlerType::HandlerType_FocusChangedEventHandler:
                return new(std::nothrow) FocusChangedEventHandler(
                    static_cast<P_HandleFocusChangedEvent>(callback)
                );
                break;
            case HandlerType::HandlerType_PropertyChangedEventHandler:
                return new(std::nothrow) PropertyChangedEventHandler(
                    static_cast<P_HandlePropertyChangedEvent>(callback)
                );
                break;
            case HandlerType::HandlerType_StructureChangedHandler:
                return new(std::nothrow) StructureChangedHandler(
                    static_cast<P_HandleStructureChangedEvent>(callback)
                );
                break;
            default:
                LOG(L"Unknown handler type %d\n", handlerType);
                exit(255);
                break;
        }
        return nullptr;
    }

    DLL_EXPORT void ReleaseHandler(void * ptr) {
        auto ph = static_cast<IUnknown *>(ptr);
        ph->Release();
    }

    DLL_EXPORT void DeleteHandler (void * ptr) {
        LOG(L">> Delete Handler %p\n", ptr);
        delete ptr;
    }
}