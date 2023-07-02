import ctypes
from ctypes import POINTER
import ctypes.wintypes as wintypes
import comtypes
import enum

class EventID:
    UIA_ToolTipOpenedEventId	=	20000
    UIA_ToolTipClosedEventId	=	20001
    UIA_StructureChangedEventId	=	20002
    UIA_MenuOpenedEventId	=	20003
    UIA_AutomationPropertyChangedEventId	=	20004
    UIA_AutomationFocusChangedEventId	=	20005
    UIA_AsyncContentLoadedEventId	=	20006
    UIA_MenuClosedEventId	=	20007
    UIA_LayoutInvalidatedEventId	=	20008
    UIA_Invoke_InvokedEventId	=	20009
    UIA_SelectionItem_ElementAddedToSelectionEventId	=	20010
    UIA_SelectionItem_ElementRemovedFromSelectionEventId	=	20011
    UIA_SelectionItem_ElementSelectedEventId	=	20012
    UIA_Selection_InvalidatedEventId	=	20013
    UIA_Text_TextSelectionChangedEventId	=	20014
    UIA_Text_TextChangedEventId	=	20015
    UIA_Window_WindowOpenedEventId	=	20016
    UIA_Window_WindowClosedEventId	=	20017
    UIA_MenuModeStartEventId	=	20018
    UIA_MenuModeEndEventId	=	20019
    UIA_InputReachedTargetEventId	=	20020
    UIA_InputReachedOtherElementEventId	=	20021
    UIA_InputDiscardedEventId	=	20022
    UIA_SystemAlertEventId	=	20023
    UIA_LiveRegionChangedEventId	=	20024
    UIA_HostedFragmentRootsInvalidatedEventId	=	20025
    UIA_Drag_DragStartEventId	=	20026
    UIA_Drag_DragCancelEventId	=	20027
    UIA_Drag_DragCompleteEventId	=	20028
    UIA_DropTarget_DragEnterEventId	=	20029
    UIA_DropTarget_DragLeaveEventId	=	20030
    UIA_DropTarget_DroppedEventId	=	20031
    UIA_TextEdit_TextChangedEventId	=	20032
    UIA_TextEdit_ConversionTargetChangedEventId	=	20033
    UIA_ChangesEventId	=	20034
    UIA_NotificationEventId	=	20035
    UIA_ActiveTextPositionChangedEventId	=	20036

class PropertyId:
    UIA_RuntimeIdPropertyId	=	30000
    UIA_BoundingRectanglePropertyId	=	30001
    UIA_ProcessIdPropertyId	=	30002
    UIA_ControlTypePropertyId	=	30003
    UIA_LocalizedControlTypePropertyId	=	30004
    UIA_NamePropertyId	=	30005
    UIA_AcceleratorKeyPropertyId	=	30006
    UIA_AccessKeyPropertyId	=	30007
    UIA_HasKeyboardFocusPropertyId	=	30008
    UIA_IsKeyboardFocusablePropertyId	=	30009
    UIA_IsEnabledPropertyId	=	30010
    UIA_AutomationIdPropertyId	=	30011
    UIA_ClassNamePropertyId	=	30012
    UIA_HelpTextPropertyId	=	30013
    UIA_ClickablePointPropertyId	=	30014
    UIA_CulturePropertyId	=	30015
    UIA_IsControlElementPropertyId	=	30016
    UIA_IsContentElementPropertyId	=	30017
    UIA_LabeledByPropertyId	=	30018
    UIA_IsPasswordPropertyId	=	30019
    UIA_NativeWindowHandlePropertyId	=	30020
    UIA_ItemTypePropertyId	=	30021
    UIA_IsOffscreenPropertyId	=	30022
    UIA_OrientationPropertyId	=	30023
    UIA_FrameworkIdPropertyId	=	30024
    UIA_IsRequiredForFormPropertyId	=	30025
    UIA_ItemStatusPropertyId	=	30026
    UIA_IsDockPatternAvailablePropertyId	=	30027
    UIA_IsExpandCollapsePatternAvailablePropertyId	=	30028
    UIA_IsGridItemPatternAvailablePropertyId	=	30029
    UIA_IsGridPatternAvailablePropertyId	=	30030
    UIA_IsInvokePatternAvailablePropertyId	=	30031
    UIA_IsMultipleViewPatternAvailablePropertyId	=	30032
    UIA_IsRangeValuePatternAvailablePropertyId	=	30033
    UIA_IsScrollPatternAvailablePropertyId	=	30034
    UIA_IsScrollItemPatternAvailablePropertyId	=	30035
    UIA_IsSelectionItemPatternAvailablePropertyId	=	30036
    UIA_IsSelectionPatternAvailablePropertyId	=	30037
    UIA_IsTablePatternAvailablePropertyId	=	30038
    UIA_IsTableItemPatternAvailablePropertyId	=	30039
    UIA_IsTextPatternAvailablePropertyId	=	30040
    UIA_IsTogglePatternAvailablePropertyId	=	30041
    UIA_IsTransformPatternAvailablePropertyId	=	30042
    UIA_IsValuePatternAvailablePropertyId	=	30043
    UIA_IsWindowPatternAvailablePropertyId	=	30044
    UIA_ValueValuePropertyId	=	30045
    UIA_ValueIsReadOnlyPropertyId	=	30046
    UIA_RangeValueValuePropertyId	=	30047
    UIA_RangeValueIsReadOnlyPropertyId	=	30048
    UIA_RangeValueMinimumPropertyId	=	30049
    UIA_RangeValueMaximumPropertyId	=	30050
    UIA_RangeValueLargeChangePropertyId	=	30051
    UIA_RangeValueSmallChangePropertyId	=	30052
    UIA_ScrollHorizontalScrollPercentPropertyId	=	30053
    UIA_ScrollHorizontalViewSizePropertyId	=	30054
    UIA_ScrollVerticalScrollPercentPropertyId	=	30055
    UIA_ScrollVerticalViewSizePropertyId	=	30056
    UIA_ScrollHorizontallyScrollablePropertyId	=	30057
    UIA_ScrollVerticallyScrollablePropertyId	=	30058
    UIA_SelectionSelectionPropertyId	=	30059
    UIA_SelectionCanSelectMultiplePropertyId	=	30060
    UIA_SelectionIsSelectionRequiredPropertyId	=	30061
    UIA_GridRowCountPropertyId	=	30062
    UIA_GridColumnCountPropertyId	=	30063
    UIA_GridItemRowPropertyId	=	30064
    UIA_GridItemColumnPropertyId	=	30065
    UIA_GridItemRowSpanPropertyId	=	30066
    UIA_GridItemColumnSpanPropertyId	=	30067
    UIA_GridItemContainingGridPropertyId	=	30068
    UIA_DockDockPositionPropertyId	=	30069
    UIA_ExpandCollapseExpandCollapseStatePropertyId	=	30070
    UIA_MultipleViewCurrentViewPropertyId	=	30071
    UIA_MultipleViewSupportedViewsPropertyId	=	30072
    UIA_WindowCanMaximizePropertyId	=	30073
    UIA_WindowCanMinimizePropertyId	=	30074
    UIA_WindowWindowVisualStatePropertyId	=	30075
    UIA_WindowWindowInteractionStatePropertyId	=	30076
    UIA_WindowIsModalPropertyId	=	30077
    UIA_WindowIsTopmostPropertyId	=	30078
    UIA_SelectionItemIsSelectedPropertyId	=	30079
    UIA_SelectionItemSelectionContainerPropertyId	=	30080
    UIA_TableRowHeadersPropertyId	=	30081
    UIA_TableColumnHeadersPropertyId	=	30082
    UIA_TableRowOrColumnMajorPropertyId	=	30083
    UIA_TableItemRowHeaderItemsPropertyId	=	30084
    UIA_TableItemColumnHeaderItemsPropertyId	=	30085
    UIA_ToggleToggleStatePropertyId	=	30086
    UIA_TransformCanMovePropertyId	=	30087
    UIA_TransformCanResizePropertyId	=	30088
    UIA_TransformCanRotatePropertyId	=	30089
    UIA_IsLegacyIAccessiblePatternAvailablePropertyId	=	30090
    UIA_LegacyIAccessibleChildIdPropertyId	=	30091
    UIA_LegacyIAccessibleNamePropertyId	=	30092
    UIA_LegacyIAccessibleValuePropertyId	=	30093
    UIA_LegacyIAccessibleDescriptionPropertyId	=	30094
    UIA_LegacyIAccessibleRolePropertyId	=	30095
    UIA_LegacyIAccessibleStatePropertyId	=	30096
    UIA_LegacyIAccessibleHelpPropertyId	=	30097
    UIA_LegacyIAccessibleKeyboardShortcutPropertyId	=	30098
    UIA_LegacyIAccessibleSelectionPropertyId	=	30099
    UIA_LegacyIAccessibleDefaultActionPropertyId	=	30100
    UIA_AriaRolePropertyId	=	30101
    UIA_AriaPropertiesPropertyId	=	30102
    UIA_IsDataValidForFormPropertyId	=	30103
    UIA_ControllerForPropertyId	=	30104
    UIA_DescribedByPropertyId	=	30105
    UIA_FlowsToPropertyId	=	30106
    UIA_ProviderDescriptionPropertyId	=	30107
    UIA_IsItemContainerPatternAvailablePropertyId	=	30108
    UIA_IsVirtualizedItemPatternAvailablePropertyId	=	30109
    UIA_IsSynchronizedInputPatternAvailablePropertyId	=	30110
    UIA_OptimizeForVisualContentPropertyId	=	30111
    UIA_IsObjectModelPatternAvailablePropertyId	=	30112
    UIA_AnnotationAnnotationTypeIdPropertyId	=	30113
    UIA_AnnotationAnnotationTypeNamePropertyId	=	30114
    UIA_AnnotationAuthorPropertyId	=	30115
    UIA_AnnotationDateTimePropertyId	=	30116
    UIA_AnnotationTargetPropertyId	=	30117
    UIA_IsAnnotationPatternAvailablePropertyId	=	30118
    UIA_IsTextPattern2AvailablePropertyId	=	30119
    UIA_StylesStyleIdPropertyId	=	30120
    UIA_StylesStyleNamePropertyId	=	30121
    UIA_StylesFillColorPropertyId	=	30122
    UIA_StylesFillPatternStylePropertyId	=	30123
    UIA_StylesShapePropertyId	=	30124
    UIA_StylesFillPatternColorPropertyId	=	30125
    UIA_StylesExtendedPropertiesPropertyId	=	30126
    UIA_IsStylesPatternAvailablePropertyId	=	30127
    UIA_IsSpreadsheetPatternAvailablePropertyId	=	30128
    UIA_SpreadsheetItemFormulaPropertyId	=	30129
    UIA_SpreadsheetItemAnnotationObjectsPropertyId	=	30130
    UIA_SpreadsheetItemAnnotationTypesPropertyId	=	30131
    UIA_IsSpreadsheetItemPatternAvailablePropertyId	=	30132
    UIA_Transform2CanZoomPropertyId	=	30133
    UIA_IsTransformPattern2AvailablePropertyId	=	30134
    UIA_LiveSettingPropertyId	=	30135
    UIA_IsTextChildPatternAvailablePropertyId	=	30136
    UIA_IsDragPatternAvailablePropertyId	=	30137
    UIA_DragIsGrabbedPropertyId	=	30138
    UIA_DragDropEffectPropertyId	=	30139
    UIA_DragDropEffectsPropertyId	=	30140
    UIA_IsDropTargetPatternAvailablePropertyId	=	30141
    UIA_DropTargetDropTargetEffectPropertyId	=	30142
    UIA_DropTargetDropTargetEffectsPropertyId	=	30143
    UIA_DragGrabbedItemsPropertyId	=	30144
    UIA_Transform2ZoomLevelPropertyId	=	30145
    UIA_Transform2ZoomMinimumPropertyId	=	30146
    UIA_Transform2ZoomMaximumPropertyId	=	30147
    UIA_FlowsFromPropertyId	=	30148
    UIA_IsTextEditPatternAvailablePropertyId	=	30149
    UIA_IsPeripheralPropertyId	=	30150
    UIA_IsCustomNavigationPatternAvailablePropertyId	=	30151
    UIA_PositionInSetPropertyId	=	30152
    UIA_SizeOfSetPropertyId	=	30153
    UIA_LevelPropertyId	=	30154
    UIA_AnnotationTypesPropertyId	=	30155
    UIA_AnnotationObjectsPropertyId	=	30156
    UIA_LandmarkTypePropertyId	=	30157
    UIA_LocalizedLandmarkTypePropertyId	=	30158
    UIA_FullDescriptionPropertyId	=	30159
    UIA_FillColorPropertyId	=	30160
    UIA_OutlineColorPropertyId	=	30161
    UIA_FillTypePropertyId	=	30162
    UIA_VisualEffectsPropertyId	=	30163
    UIA_OutlineThicknessPropertyId	=	30164
    UIA_CenterPointPropertyId	=	30165
    UIA_RotationPropertyId	=	30166
    UIA_SizePropertyId	=	30167
    UIA_IsSelectionPattern2AvailablePropertyId	=	30168
    UIA_Selection2FirstSelectedItemPropertyId	=	30169
    UIA_Selection2LastSelectedItemPropertyId	=	30170
    UIA_Selection2CurrentSelectedItemPropertyId	=	30171
    UIA_Selection2ItemCountPropertyId	=	30172
    UIA_HeadingLevelPropertyId	=	30173
    UIA_IsDialogPropertyId	=	30174

class TreeScope:
    TreeScope_None	= 0
    TreeScope_Element	= 0x1
    TreeScope_Children	= 0x2
    TreeScope_Descendants	= 0x4
    TreeScope_Parent	= 0x8
    TreeScope_Ancestors	= 0x10
    TreeScope_Subtree	= ( ( TreeScope_Element | TreeScope_Children )  | TreeScope_Descendants ) 


class StructureChangeType:
    StructureChangeType_ChildAdded	= 0
    StructureChangeType_ChildRemoved	= ( StructureChangeType_ChildAdded + 1 )
    StructureChangeType_ChildrenInvalidated	= ( StructureChangeType_ChildRemoved + 1 )
    StructureChangeType_ChildrenBulkAdded	= ( StructureChangeType_ChildrenInvalidated + 1 )
    StructureChangeType_ChildrenBulkRemoved	= ( StructureChangeType_ChildrenBulkAdded + 1 )
    StructureChangeType_ChildrenReordered	= ( StructureChangeType_ChildrenBulkRemoved + 1 ) 

# /*
#  * VARENUM usage key,
#  *
#  * * [V] - may appear in a VARIANT
#  * * [T] - may appear in a TYPEDESC
#  * * [P] - may appear in an OLE property set
#  * * [S] - may appear in a Safe Array
#  *
#  *
#  *  VT_EMPTY            [V]   [P]     nothing
#  *  VT_NULL             [V]   [P]     SQL style Null
#  *  VT_I2               [V][T][P][S]  2 byte signed int
#  *  VT_I4               [V][T][P][S]  4 byte signed int
#  *  VT_R4               [V][T][P][S]  4 byte real
#  *  VT_R8               [V][T][P][S]  8 byte real
#  *  VT_CY               [V][T][P][S]  currency
#  *  VT_DATE             [V][T][P][S]  date
#  *  VT_BSTR             [V][T][P][S]  OLE Automation string
#  *  VT_DISPATCH         [V][T]   [S]  IDispatch *
#  *  VT_ERROR            [V][T][P][S]  SCODE
#  *  VT_BOOL             [V][T][P][S]  True=-1, False=0
#  *  VT_VARIANT          [V][T][P][S]  VARIANT *
#  *  VT_UNKNOWN          [V][T]   [S]  IUnknown *
#  *  VT_DECIMAL          [V][T]   [S]  16 byte fixed point
#  *  VT_RECORD           [V]   [P][S]  user defined type
#  *  VT_I1               [V][T][P][s]  signed char
#  *  VT_UI1              [V][T][P][S]  unsigned char
#  *  VT_UI2              [V][T][P][S]  unsigned short
#  *  VT_UI4              [V][T][P][S]  ULONG
#  *  VT_I8                  [T][P]     signed 64-bit int
#  *  VT_UI8                 [T][P]     unsigned 64-bit int
#  *  VT_INT              [V][T][P][S]  signed machine int
#  *  VT_UINT             [V][T]   [S]  unsigned machine int
#  *  VT_INT_PTR             [T]        signed machine register size width
#  *  VT_UINT_PTR            [T]        unsigned machine register size width
#  *  VT_VOID                [T]        C style void
#  *  VT_HRESULT             [T]        Standard return type
#  *  VT_PTR                 [T]        pointer type
#  *  VT_SAFEARRAY           [T]        (use VT_ARRAY in VARIANT)
#  *  VT_CARRAY              [T]        C style array
#  *  VT_USERDEFINED         [T]        user defined type
#  *  VT_LPSTR               [T][P]     null terminated string
#  *  VT_LPWSTR              [T][P]     wide null terminated string
#  *  VT_FILETIME               [P]     FILETIME
#  *  VT_BLOB                   [P]     Length prefixed bytes
#  *  VT_STREAM                 [P]     Name of the stream follows
#  *  VT_STORAGE                [P]     Name of the storage follows
#  *  VT_STREAMED_OBJECT        [P]     Stream contains an object
#  *  VT_STORED_OBJECT          [P]     Storage contains an object
#  *  VT_VERSIONED_STREAM       [P]     Stream with a GUID version
#  *  VT_BLOB_OBJECT            [P]     Blob contains an object 
#  *  VT_CF                     [P]     Clipboard format
#  *  VT_CLSID                  [P]     A Class ID
#  *  VT_VECTOR                 [P]     simple counted array
#  *  VT_ARRAY            [V]           SAFEARRAY*
#  *  VT_BYREF            [V]           void* for local use
#  *  VT_BSTR_BLOB                      Reserved for system use
#  */

# enum VARENUM
#     {
#         VT_EMPTY	= 0,
#         VT_NULL	= 1,
#         VT_I2	= 2,
#         VT_I4	= 3,
#         VT_R4	= 4,
#         VT_R8	= 5,
#         VT_CY	= 6,
#         VT_DATE	= 7,
#         VT_BSTR	= 8,
#         VT_DISPATCH	= 9,
#         VT_ERROR	= 10,
#         VT_BOOL	= 11,
#         VT_VARIANT	= 12,
#         VT_UNKNOWN	= 13,
#         VT_DECIMAL	= 14,
#         VT_I1	= 16,
#         VT_UI1	= 17,
#         VT_UI2	= 18,
#         VT_UI4	= 19,
#         VT_I8	= 20,
#         VT_UI8	= 21,
#         VT_INT	= 22,
#         VT_UINT	= 23,
#         VT_VOID	= 24,
#         VT_HRESULT	= 25,
#         VT_PTR	= 26,
#         VT_SAFEARRAY	= 27,
#         VT_CARRAY	= 28,
#         VT_USERDEFINED	= 29,
#         VT_LPSTR	= 30,
#         VT_LPWSTR	= 31,
#         VT_RECORD	= 36,
#         VT_INT_PTR	= 37,
#         VT_UINT_PTR	= 38,
#         VT_FILETIME	= 64,
#         VT_BLOB	= 65,
#         VT_STREAM	= 66,
#         VT_STORAGE	= 67,
#         VT_STREAMED_OBJECT	= 68,
#         VT_STORED_OBJECT	= 69,
#         VT_BLOB_OBJECT	= 70,
#         VT_CF	= 71,
#         VT_CLSID	= 72,
#         VT_VERSIONED_STREAM	= 73,
#         VT_BSTR_BLOB	= 0xfff,
#         VT_VECTOR	= 0x1000,
#         VT_ARRAY	= 0x2000,
#         VT_BYREF	= 0x4000,
#         VT_RESERVED	= 0x8000,
#         VT_ILLEGAL	= 0xffff,
#         VT_ILLEGALMASKED	= 0xfff,
#         VT_TYPEMASK	= 0xfff
#     } ;


class DECIMAL(ctypes.Structure):
  _fields_ = [("wReserved", ctypes.c_ushort),
                ("scale", ctypes.c_ubyte),
                ("sign", ctypes.c_ubyte),
                ("Hi32", ctypes.c_ulong),
                ("Lo64", ctypes.c_ulonglong)]

class SAFEARRAYBOUND(ctypes.Structure):
    _fields_ = [
        ('cElements', wintypes.ULONG),
        ('lLbound', wintypes.LONG),
    ]

class SAFEARRAY(ctypes.Structure):
    _fields_ = [
        ('cDims', wintypes.USHORT),
        ('fFeatures', wintypes.USHORT),
        ('cbElements', wintypes.ULONG),
        ('cLocks', wintypes.ULONG),
        ('pvData', ctypes.c_void_p),
        ('rgsabound', SAFEARRAYBOUND * 1),
    ]

class CY(ctypes.Structure):
    _fields_ = [("int64", ctypes.c_longlong)]

class VARIANT(ctypes.Structure):
    class VARIANT_NAME_1(ctypes.Union):
        class __tagVARIANT(ctypes.Structure):
            class __VARIANT_NAME_3(ctypes.Union):
                class __tagBRECORD(ctypes.Structure):
                    _fields_ = [
                        # IRecordInfo *pRecInfo;
                        ("pvRecord", ctypes.c_void_p),
                        ("pRecInfo", POINTER(comtypes.IUnknown))
                    ]
                _fields_ = [
                    ('llVal', ctypes.c_longlong),    
                    ('lVal', ctypes.c_long),
                    ('bVal', ctypes.c_byte),
                    ('iVal', ctypes.c_short),
                    ('fltVal', ctypes.c_float),
                    ('dblVal', ctypes.c_double),
                    ('dblVal', ctypes.c_double),
                    # VARIANT_BOOL boolVal; # 0 == FALSE, -1 == TRUE
                    # VARIANT_BOOL __OBSOLETE__VARIANT_BOOL;
                    ('boolVal', ctypes.c_short),
                    ('__OBSOLETE__VARIANT_BOOL', ctypes.c_short),
                    # SCODE scode;
                    ('scode', ctypes.c_long),
                    ('cyVal', CY),
                    # DATE date;
                    ('date', ctypes.c_double),
                    # BSTR bstrVal;
                    ('bstrval', ctypes.c_wchar_p),
                    ('punkVal', POINTER(comtypes.IUnknown)),
                    # IDispatch *pdispVal;
                    ('pdispVal', ctypes.c_void_p),
                    # SAFEARRAY *parray;
                    ('parray', POINTER(SAFEARRAY)),
                    # BYTE *pbVal
                    ('pbVal', POINTER(wintypes.BYTE)), #?
                    # SHORT *piVal;
                    ('piVal', POINTER(wintypes.SHORT)),
                    # LONG *plVal;
                    ('plVal', POINTER(wintypes.LONG)),
                    # LONGLONG *pllVal;
                    ('pllVal', POINTER(ctypes.c_longlong)),
                    # FLOAT *pfltVal;
                    ('pfltVal', POINTER(wintypes.FLOAT)),
                    # DOUBLE *pdblVal;
                    ('pdblVal', POINTER(wintypes.DOUBLE)),
                    # VARIANT_BOOL *pboolVal;
                    ('pboolVal', POINTER(ctypes.c_short)),
                    # VARIANT_BOOL *__OBSOLETE__VARIANT_PBOOL;
                    ('__OBSOLETE__VARIANT_PBOOL', POINTER(ctypes.c_short)),
                    # SCODE *pscode;
                    ('pscode', POINTER(ctypes.c_long)),
                    # CY *pcyVal;
                    ('pcyVal', POINTER(CY)),
                    # DATE *pdate;
                    ('pcyVal', POINTER(ctypes.c_double)),
                    # BSTR *pbstrVal;
                    ('pbstrVal', POINTER(ctypes.c_wchar_p)),
                    # IUnknown **ppunkVal;
                    ('ppunkVal', POINTER(POINTER(comtypes.IUnknown))),
                    # IDispatch **ppdispVal;
                    ('ppdispVal', POINTER(ctypes.c_void_p)),
                    # SAFEARRAY **ppdispVal;
                    ('ppdispVal', POINTER(POINTER(SAFEARRAY))),
                    # VARIANT *pvarVal; # self reference?
                    ('pvarVal', POINTER(ctypes.c_void_p)),
                    # PVOID byref;
                    ('byref', ctypes.c_void_p),
                    # CHAR cVal;
                    ('cVal', wintypes.CHAR),
                    # USHORT uiVal;
                    ('uiVal', wintypes.USHORT),
                    # ULONG ulVal;
                    ('ulVal', wintypes.ULONG),
                    # ULONGLONG ullVal;
                    ('ullVal', ctypes.c_ulonglong),
                    # INT intVal;
                    ('intVal', wintypes.INT),
                    # UINT uintVal;
                    ('uintVal', wintypes.UINT),
                    # DECIMAL *pdecVal;
                    ('pdecVal', ctypes.c_void_p),
                    # CHAR *pcVal;
                    ('pcVal', POINTER(wintypes.CHAR)),
                    # USHORT *puiVal;
                    ('puiVal', POINTER(wintypes.USHORT)),
                    # ULONG *pulVal;
                    ('pulVal', POINTER(wintypes.ULONG)),
                    # ULONGLONG *pullVal;
                    ('pullVal', POINTER(ctypes.c_ulonglong)),
                    # INT *pintVal;
                    ('pintVal', POINTER(wintypes.INT)),
                    # UINT *puintVal;
                    ('puintVal', POINTER(wintypes.UINT)),
                    ('__VARIANT_NAME_4', __tagBRECORD),
                ]
            _fields_ = [
                # VARTYPE vt;
                ('vt', ctypes.c_short),
                # WORD wReserved1;
                ('wReserved1', wintypes.WORD),
                # WORD wReserved2;
                ('wReserved2', wintypes.WORD),
                # WORD wReserved3;
                ('wReserved3', wintypes.WORD),
                ('__VARIANT_NAME_3', __VARIANT_NAME_3),
            ]
        _fields_ = [
            ('__VARIANT_NAME_2', __tagVARIANT),
            ('decVal', DECIMAL),
        ]
        _anonymous_ = ["__VARIANT_NAME_2"]
    _fields_ = [('__VARIANT_NAME_1', VARIANT_NAME_1)]
    _anonymous_ = ["__VARIANT_NAME_1"]
