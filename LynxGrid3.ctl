VERSION 5.00
Begin VB.UserControl LynxGrid3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   KeyPreview      =   -1  'True
   ScaleHeight     =   63
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   107
   ToolboxBitmap   =   "LynxGrid3.ctx":0000
End
Attribute VB_Name = "LynxGrid3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const DEBUG_MODE = False

'####################################################################################
'Title:     LynxGrid
'Function:  Owner-drawn editable Grid
'Author:    Richard Mewett
'Created:   01/08/05
'Version:   1.50 (19 June 2006)
'
'Copyright © 2005 Richard Mewett. All rights reserved.

'I created this control to provide a combination of MSFlexGrid and ListView (Report Style) functionality.

'I have had many requests for an "always-up" version of ComboView - so hopefully this provides
'the functionality that those people require.

'NOTES: This control was designed to be fairly lightweight - both in total code size and in
'resources required. Therefore Integers are used within UDT's/Arrays in place of Longs where
'appropriate - this is intentional!

'####################################################################################
'Credits:   Paul Caton - Subclassing
'           Gary Noble (Phantom Man)- API Scroll Bar Code
'           Heriberto Mantilla Santamaría - XP Theme API
'           Matthew R. Usner - DrawArrow + Beta testing
'           LaVolpe - Bug fixes & numerous suggestions
'           Riccardo Cohen - Bug reports & ownerdrawn XP/Office ThemeStyles

' - Bug reports/feature suggestions
'           Jeff Mayes, Gary Noble, Light Templer, Eric O'Sullivan

'Updates (dd/mm/yy):
'05/06/06   Added FindItem function, TopRow property
'           Added automatic Item searching when typing via SearchColumn property
'06/06/06   Added ColTag & ItemTag properties
'07/06/06   Added RemoveItem method & ItemCount property
'           Added BindControl method
'           Added Color Table (removed BackColor & ForeColor from Cell UDT)
'13/06/06   Added Boolean Column Type & CellChecked Property
'           Added ProgressBar Column Type & CellProgressValue Property
'14/06/06   Added Public Sort Method
'15/06/06   ItemImages check ScaleMode of ImageList
'16/06/06   Sort routines broken down per datatype
'           Redraw Tweaks
'19/06/06   Added ThemeColor & ThemeStyle properties
            
'####################################################################################
'This software is provided "as-is," without any express or implied warranty.
'In no event shall the author be held liable for any damages arising from the
'use of this software.
'If you do not agree with these terms, do not install "LynxGrid". Use of
'the program implicitly means you have agreed to these terms.
'
'Permission is granted to anyone to use this software for any purpose,
'including commercial use, and to alter and redistribute it, provided that
'the following conditions are met:
'
'1. All redistributions of source code files must retain all copyright
'   notices that are currently in place, and this list of conditions without
'   any modification.
'
'2. All redistributions in binary form must retain all occurrences of the
'   above copyright notice and web site addresses that are currently in
'   place (for example, in the About boxes).
'
'3. Modified versions in source or binary form must be plainly marked as
'   such, and must not be misrepresented as being the original software.

'################################################################
'API Declarations
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDc As Long, ByVal hRgn As Long) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hDc As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

'XP
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function DrawThemeEdge Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pDestRect As RECT, ByVal uEdge As Long, ByVal uFlags As Long, pContentRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long

Private Const CLR_INVALID = &HFFFF

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const DFC_BUTTON        As Long = &H4

Private Const DFCS_FLAT         As Long = &H4000
Private Const DFCS_BUTTONCHECK  As Long = &H0
Private Const DFCS_BUTTONPUSH   As Long = &H10
Private Const DFCS_CHECKED      As Long = &H400
Private Const DFCS_PUSHED = &H200
Private Const DFCS_TRANSPARENT = &H800 ' Win98/2000 only
Private Const DFCS_HOT = &H1000

Private Const VER_PLATFORM_WIN32_NT = 2

Private Const GRADIENT_FILL_RECT_H    As Long = &H0
Private Const GRADIENT_FILL_RECT_V    As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE  As Long = &H2
Private GRADIENT_FILL_RECT_DIRECTION  As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
   UPPERLEFT As Long
   LOWERRIGHT As Long
End Type

'################################################################
'Subclassing
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Const ALL_MESSAGES      As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED        As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC       As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04          As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05          As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08          As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09          As Long = 137                                      'Table A (after) entry count patch offset

Private Const WM_SETFOCUS       As Long = &H7
Private Const WM_KILLFOCUS      As Long = &H8
Private Const WM_MOUSELEAVE     As Long = &H2A3
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSEHOVER     As Long = &H2A1
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_VSCROLL        As Long = &H115
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_ACTIVATE       As Long = &H6
Private Const WM_ACTIVATEAPP    As Long = &H1C

Private Type tSubData                                                                   'Subclass data type
    hwnd          As Long                                            'Handle of the window being subclassed
    nAddrSub      As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig     As Long                                            'The address of the pre-existing WndProc
    nMsgCntA      As Long                                            'Msg after table entry count
    nMsgCntB      As Long                                            'Msg before table entry count
    aMsgTblA()    As Long                                            'Msg after table array
    aMsgTblB()    As Long                                            'Msg Before table array
End Type
                                
Private Type TRACKMOUSEEVENT_STRUCT
  cbSize          As Long
  dwFlags         As TRACKMOUSEEVENT_FLAGS
  hwndTrack       As Long
  dwHoverTime     As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean

Private sc_aSubData()                As tSubData                                        'Subclass data array

'################################################################
'API Scroll Bars
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long

Public Enum ScrollBarOrienationEnum
    Scroll_Horizontal
    Scroll_Vertical
    Scroll_Both
End Enum

Public Enum ScrollBarStyleEnum
    Style_Regular = 1& ' FSB_REGULAR_MODE
    Style_Flat = 0& 'FSB_FLAT_MODE
End Enum

Public Enum EFSScrollBarConstants
    efsHorizontal = 0 'SB_HORZ
    efsVertical = 1 'SB_VERT
End Enum

Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
Private Const MK_CONTROL = &H8
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const FSB_FLAT_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private m_bInitialised      As Boolean
Private m_eOrientation      As ScrollBarOrienationEnum
Private m_eStyle            As ScrollBarStyleEnum
Private m_hWnd              As Long
Private m_lSmallChangeHorz  As Long
Private m_lSmallChangeVert  As Long
Private m_bEnabledHorz      As Boolean
Private m_bEnabledVert      As Boolean
Private m_bVisibleHorz      As Boolean
Private m_bVisibleVert      As Boolean
Private m_bNoFlatScrollBars As Boolean

'################################################################
Private Enum lgFlagsEnum
    lgChecked = 2
    lgSelected = 4
    lgFontBold = 8
    lgFontItalic = 16
    lgFontUnderline = 32
End Enum

Private Enum lgCellFormatEnum
    lgCFBackColor = 2
    lgCFForeColor = 4
End Enum

Private Enum lgHeaderStateEnum
    lgNormal = 1
    lgHot = 2
    lgDown = 3
End Enum

Public Enum lgAllowUserResizingEnum
    lgResizeNone = 0
    lgResizeCol = 1
    'lgResizeRow = 2
    lgResizeBoth = 4
End Enum

Public Enum lgAlignmentEnum
    lgAlignLeftTop = DT_LEFT Or DT_TOP
    lgAlignLeftCenter = DT_LEFT Or DT_VCENTER
    lgAlignLeftBottom = DT_LEFT Or DT_BOTTOM
    lgAlignCenterTop = DT_CENTER Or DT_TOP
    lgAlignCenterCenter = DT_CENTER Or DT_VCENTER
    lgAlignCenterBottom = DT_CENTER Or DT_BOTTOM
    lgAlignRightTop = DT_RIGHT Or DT_TOP
    lgAlignRightCenter = DT_RIGHT Or DT_VCENTER
    lgAlignRightBottom = DT_RIGHT Or DT_BOTTOM
End Enum

Public Enum lgBorderStyleEnum
    lgNone = 0
    lgSingle = 1
End Enum

Public Enum lgDataTypeEnum
    lgString = 0
    lgNumeric = 1
    lgDate = 2
    lgBoolean = 3
    lgProgressBar = 4
    lgCustom = 5
End Enum

Public Enum lgEditTriggerEnum
    lgNone = 0
    lgEnterKey = 2
    lgF2Key = 4
    lgMouseClick = 8
    lgMouseDblClick = 16
End Enum

Public Enum lgFocusRectModeEnum
    lgNone = 0
    lgRow = 1
    lgCol = 2
End Enum

Public Enum lgFocusRectStyleEnum
    lgFRLight = 0
    lgFRHeavy = 1
End Enum

Public Enum lgMoveControlEnum
    lgBCNone = 0
    lgBCHeight = 1
    lgBCWidth = 2
    lgBCLeft = 4
    lgBCTop = 8
End Enum

Public Enum lgSearchModeEnum
    lgSMEqual = 0
    lgSMGreaterEqual = 1
    lgSMLike = 2
    lgSMNavigate = 4
    'Added By Vincent J. Jamero
    lgWith = 5
End Enum

Public Enum lgSortTypeEnum
    lgSTAscending = 0
    lgSTDescending = 1
End Enum

Public Enum lgThemeColorEnum
    lgTCCustom = 0
    lgTCDefault = 1
    lgTCBlue = 2
    lgTCGreen = 3
End Enum

Public Enum lgThemeStyleEnum
    lgTSWindows3D = 0
    lgTSWindowsFlat = 1
    lgTSWindowsXP = 2
    lgTSOfficeXP = 3
End Enum

#If False Then
    Private lgChecked, lgSelected, lgFontBold, lgFontItalic, lgFontUnderline
    Private lgNormal, lgHot, lgDown
    Private lgResizeNone, lgResizeCol, lgResizeRow, lgResizeBoth
    Private lgAlignLeftTop, lgAlignLeftCenter, lgAlignLeftBottom, lgAlignCenterTop, lgAlignCenterCenter, lgAlignCenterBottom, lgAlignRightTop, lgAlignRightCenter, lgAlignRightBottom
    Private lgNone, lgSingle
    Private lgString, lgNumeric, lgDate, lgBoolean, lgProgressBar, lgCustom
    Private lgNone, lgEnterKey, lgF2Key, lgMouseClick, lgMouseDblClick
    Private lgNone, lgRow, lgCol
    Private lgFRLight, lgFRHeavy
#End If

Private Const ROW_HEIGHT                As Long = 16

Private Const DEF_ALLOWUSERRESIZING     As Long = lgAllowUserResizingEnum.lgResizeNone
Private Const DEF_BACKCOLOR             As Long = vbWindowBackground
Private Const DEF_BACKCOLORBKG          As Long = &H808080
Private Const DEF_BACKCOLOREDIT         As Long = &HC0FFFF
Private Const DEF_BACKCOLORFIXED        As Long = vbButtonFace
Private Const DEF_BACKCOLORSEL          As Long = vbHighlight
Private Const DEF_BORDERSTYLE           As Long = lgBorderStyleEnum.lgSingle
Private Const DEF_CACHEINCREMENT        As Long = 10
Private Const DEF_CHECKBOXES            As Boolean = False
Private Const DEF_COLUMNHEADERS         As Boolean = True
Private Const DEF_COLUMNSORT            As Boolean = False
Private Const DEF_DISPLAYELLIPSIS       As Boolean = True
Private Const DEF_EDITABLE              As Boolean = False
Private Const DEF_EDITTRIGGER           As Long = lgEditTriggerEnum.lgEnterKey
Private Const DEF_ENABLED               As Boolean = True
Private Const DEF_FOCUSRECTCOLOR        As Long = &HFFFF&
Private Const DEF_FOCUSRECTMODE         As Long = lgFocusRectModeEnum.lgRow
Private Const DEF_FOCUSRECTSTYLE        As Long = lgFocusRectStyleEnum.lgFRHeavy
Private Const DEF_FORECOLOR             As Long = vbWindowText
Private Const DEF_FORECOLOREDIT         As Long = vbWindowText
Private Const DEF_FORECOLORFIXED        As Long = vbButtonText
Private Const DEF_FORECOLORSEL          As Long = vbHighlightText
Private Const DEF_FORECOLORTOTALS       As Long = vbRed
Private Const DEF_FORMATSTRING          As String = vbNullString
Private Const DEF_FULLROWSELECT         As Boolean = True
Private Const DEF_GRIDCOLOR             As Long = &HC0C0C0
Private Const DEF_GRIDLINES             As Boolean = True
Private Const DEF_GRIDLINEWIDTH         As Long = 1
Private Const DEF_HOTHEADERTRACKING     As Boolean = True
Private Const DEF_LOCKED                As Boolean = False
Private Const DEF_MULTISELECT           As Boolean = False
Private Const DEF_PROGRESSBARCOLOR      As Long = &H8080FF
Private Const DEF_REDRAW                As Boolean = True
Private Const DEF_ROWHEIGHTMIN          As Long = 0
Private Const DEF_SCALEUNITS            As Integer = vbPixels
Private Const DEF_SCROLLTRACK           As Boolean = True
Private Const DEF_SEARCHCOLUMN          As Long = 0
Private Const DEF_THEMECOLOR            As Long = lgThemeColorEnum.lgTCCustom
Private Const DEF_THEMESTYLE            As Long = lgThemeStyleEnum.lgTSWindowsXP

Private Const NULL_RESULT               As Long = -1
Private Const AUTOSCROLL_TIMEOUT        As Long = 25
Private Const SIZE_VARIANCE             As Long = 4

Private Const SCROLL_NONE               As Long = 0
Private Const SCROLL_UP                 As Long = 1
Private Const SCROLL_DOWN               As Long = 2

'##########################################
'For Rendering
Private Const MAX_CHECKBOXSIZE          As Long = 16
Private Const SIZE_SORTARROW            As Long = 8

Private Const HEADER_LEFT               As Long = 3
Private Const TEXT_SPACE                As Long = 3
Private Const ARROW_SPACE               As Long = 5

Private Const RIGHT_CHECKBOX            As Long = 15
'##########################################

Public Type udtColumn
    EditCtrl As Object
    dCustomWidth As Single
    nAlignment As Integer
    nSortOrder As lgSortTypeEnum
    nType As Integer
    lWidth As Long
    lX As Long
    MoveControl As Integer
    bVisible As Boolean
    sCaption As String
    sFormat As String
    sTag As String
End Type

Private Type udtCell
    nAlignment As Integer
    nFormat As Integer
    nFlags As Integer
    sValue As String
End Type

Private Type udtItem
    lHeight As Long
    lImage As Long
    lItemData As Long
    nFlags As Integer
    sTag As String
    Cell() As udtCell
End Type

Private Type udtFormat
    lBackColor As Long
    lForeColor As Long
    sFontName As String
    dFontSize As Single
    'Modified By: Vincent J. Jamero
    'Date: Agust 2, 2006
    'FROM: nCount As Integer
    nCount As Long
End Type

Private Type udtRender
    DTFlag As Long
    CheckBoxSize As Long
    ImageSpace As Long
    LeftImage As Long
    LeftText As Long
    HeaderHeight As Long
    TextHeight As Long
End Type

Private WithEvents txtEdit  As TextBox
Attribute txtEdit.VB_VarHelpID = -1

'################################################################
'Data & Columns
Private mCols() As udtColumn
Private mItems() As udtItem
Private mIX() As Long
Private mCF() As udtFormat

Private mItemCount As Long
Private mItemsVisible As Long
Private mSortColumn As Long
Private mSortSubColumn As Long

Private mEditCol As Long
Private mEditRow As Long
Private mCol As Long
Private mRow As Long
Private mMouseCol As Long
Private mMouseRow As Long
Private mMouseDownCol As Long
Private mMouseDownRow As Long

Private mSelectedRow As Long

Private mR As udtRender
Private mEditPending As Boolean
Private mMouseDown As Boolean
Private mResizeCol As Long
Private mEditParent As Long

'################################################################
'Appearance Properties
Private mBackColor As Long
Private mBackColorBkg As Long
Private mBackColorEdit As Long
Private mBackColorFixed As Long
Private mBackColorSel As Long
Private mForeColor As Long
Private mForeColorEdit As Long
Private mForeColorFixed As Long
Private mForeColorSel As Long
Private mForeColorTotals As Long

Private mFocusRectColor As Long
Private mGridColor As Long
Private mProgressBarColor As Long

Private mBorderStyle As lgBorderStyleEnum
Private mDisplayEllipsis As Boolean
Private mFocusRectMode As lgFocusRectModeEnum
Private mFocusRectStyle As lgFocusRectStyleEnum
Private mFont As Font
Private mGridLines As Boolean
Private mGridLineWidth As Long
Private mThemeColor As lgThemeColorEnum
Private mThemeStyle As lgThemeStyleEnum

'Added by:Vincent J. Jamero
'         July 2, 2006
'Stripes BackColor
Private Const DEF_Striped As Boolean = False
Private Const DEF_SBackColor1 = &HFFFFFF
Private Const DEF_SBackColor2 = &HF5F5F5
    
Private mStriped As Boolean
Private mSBackColor1 As Long
Private mSBackColor2 As Long

'the lRow variable is came from 'DrawGrid' function
'i make it global so we can used it on striping to the 'DrawRect' funciton
Dim lRow As Long

'################################################################
'Behaviour Properties
Private mAllowUserResizing As lgAllowUserResizingEnum
Private mCheckboxes As Boolean
Private mColumnHeaders As Boolean
Private mColumnSort As Boolean
Private mEditable As Boolean
Private mEditTrigger As lgEditTriggerEnum
Private mFullRowSelect As Boolean
Private mHotHeaderTracking As Boolean
Private mMultiSelect As Boolean
Private mRedraw As Boolean
Private mScrollTrack As Boolean

'################################################################
'Miscellaneous Properties
Private mCacheIncrement As Long
Private mEnabled As Boolean
Private mFormatString As String
Private mLocked As Boolean
Private mRowHeightMin As Long
Private mScaleUnits As ScaleModeConstants
Private mSearchColumn As Long

Private mImageList As Object

'################################################################
'Control State Variables
Private mInCtrl As Boolean
Private mInFocus As Boolean
Private mWindowsNT As Boolean
Private mWindowsXP As Boolean

Private mPendingRedraw As Boolean
Private mPendingScrollBar As Boolean

Private mClipRgn As Long
Private hTheme As Long
Private mScrollAction As Long
Private mScrollTick As Long
Private mHotColumn As Long
Private mIgnoreKeyPress As Boolean

'added by: Vincent J.Jamero
Private m_IgnoreEmpty As Boolean


'################################################################
'Events - Standard VB
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Events - Control Specific
Public Event ColumnClick(Col As Long)
Public Event ColumnSizeChanged(Col As Long, MoveControl As lgMoveControlEnum)
Public Event CustomSort(Ascending As Boolean, Col As Long, Value1 As String, Value2 As String, Swap As Boolean)
Public Event ItemChecked(Row As Long)
Public Event ItemCountChanged()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event RowColChanged()
Public Event Scroll()
Public Event SelectionChanged()
Public Event SortComplete()
Public Event ThemeChanged()

Public Event EnterCell()
Public Event RequestEdit(Row As Long, Col As Long, Cancel As Boolean)
Public Event RequestUpdate(Row As Long, Col As Long, NewValue As String, Cancel As Boolean)

'Added by: Vincent J. Jamero
Public Event BeforeDrawGrid(StartRow As Long, EndRow As Long)
Public Event BeforeDrawText(Row As Long, Col As Long, ByRef sNewValue As String)
Public Event AfterDrawGrid()

'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    Dim eBar As EFSScrollBarConstants
    Dim lV As Long, lSC As Long
    Dim lScrollCode As Long
    Dim tSI As SCROLLINFO
    Dim zDelta As Long
    Dim lHSB As Long
    Dim lVSB As Long
    Dim bRedraw As Boolean
    
    'Debug.Print "zSubclass_Proc " & Timer
    
    Select Case uMsg
        Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
            lScrollCode = (wParam And &HFFFF&)
            
            lHSB = SBValue(efsHorizontal)
            lVSB = SBValue(efsVertical)
    
            Select Case uMsg
            
                Case WM_HSCROLL ' Get the scrollbar type
                    eBar = efsHorizontal
                    
                Case WM_VSCROLL
                    eBar = efsVertical
                    
                Case Else     'WM_MOUSEWHEEL
                    eBar = IIf(lScrollCode And MK_CONTROL, efsHorizontal, efsVertical)
                    lScrollCode = IIf(wParam / 65536 < 0, SB_LINEDOWN, SB_LINEUP)
                    
            End Select
            
            bRedraw = True
    
            Select Case lScrollCode
            
                Case SB_THUMBTRACK
                    ' Is vertical/horizontal?
                    pSBGetSI eBar, tSI, SIF_TRACKPOS
                    SBValue(eBar) = tSI.nTrackPos
                    
                    bRedraw = mScrollTrack
    
                Case SB_LEFT, SB_BOTTOM
                     SBValue(eBar) = IIf(lScrollCode = 7, SBMax(eBar), SBMin(eBar))
    
                Case SB_RIGHT, SB_TOP
                     SBValue(eBar) = SBMin(eBar)
    
                Case SB_LINELEFT, SB_LINEUP
                
                    If SBVisible(eBar) Then
                    
                        lV = SBValue(eBar)
                        If (eBar = efsHorizontal) Then
                            lSC = m_lSmallChangeHorz
                        Else
                            lSC = m_lSmallChangeVert
                        End If
                        
                        If (lV - lSC < SBMin(eBar)) Then
                             SBValue(eBar) = SBMin(eBar)
                        Else
                             SBValue(eBar) = lV - lSC
                        End If
                        
                    End If
    
                Case SB_LINERIGHT, SB_LINEDOWN
                    If SBVisible(eBar) Then
            
                        lV = SBValue(eBar)
                        
                        If (eBar = efsHorizontal) Then
                            lSC = m_lSmallChangeHorz
                        Else
                            lSC = m_lSmallChangeVert
                        End If
                        
                        If (lV + lSC > SBMax(eBar)) Then
                             SBValue(eBar) = SBMax(eBar)
                        Else
                             SBValue(eBar) = lV + lSC
                        End If
                    End If
    
                Case SB_PAGELEFT, SB_PAGEUP
                     SBValue(eBar) = SBValue(eBar) - SBLargeChange(eBar)
    
                Case SB_PAGERIGHT, SB_PAGEDOWN
                     SBValue(eBar) = SBValue(eBar) + SBLargeChange(eBar)
    
                Case SB_ENDSCROLL
                    If Not mScrollTrack Then
                        DrawGrid True
                    End If
    
            End Select
            
            If (lHSB <> SBValue(efsHorizontal)) Or (lVSB <> SBValue(efsVertical)) Then
                UpdateCell
                
                If bRedraw Then
                    DrawGrid True
                End If
                
                RaiseEvent Scroll
            End If
        
        Case WM_MOUSEWHEEL
                
        Case WM_MOUSEMOVE
            If Not mInCtrl Then
                mInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                RaiseEvent MouseEnter
            End If
    
        Case WM_MOUSELEAVE
            mInCtrl = False
            DrawHeaderRow
            UserControl.Refresh
            RaiseEvent MouseLeave
            
        Case WM_SETFOCUS
             If mEnabled Then
                mInFocus = True
                DrawGrid True
             End If
    
        Case WM_KILLFOCUS
            If lng_hWnd = UserControl.hwnd Then
                If mEnabled Then
                   mInFocus = False
                   DrawGrid True
                End If
            ElseIf Not mInCtrl Then
                UpdateCell
            End If
    
        Case WM_THEMECHANGED
            DrawGrid True
            RaiseEvent ThemeChanged

    End Select
End Sub

Public Function AddColumn(Optional Caption As String, Optional Width As Single, Optional Alignment As lgAlignmentEnum = lgAlignLeftCenter, Optional DataType As lgDataTypeEnum = lgString, Optional Format As String) As Long
    '#############################################################################################################################
    'Purpose: Add a Column to the Grid
    
    'Caption    - The text that appears on the Header
    'Width      - The Width!
    'Alignment  - The Alignment!
    'DataType   - Allows the control to determine proper Sort Sequence when Sorting
    'Format     - Format Mask applied to Cell data before it is displayed (i.e. "#.00")
    '#############################################################################################################################
    
    Dim lNewCol As Long
    
    If mCols(0).nAlignment <> 0 Then
        lNewCol = UBound(mCols) + 1
        ReDim Preserve mCols(lNewCol)
    End If
 
    With mCols(lNewCol)
        .sCaption = Caption
        .dCustomWidth = Width
        
        'lWidth is always Pixels (because thats what API functions require) and
        'is calculated to prevent repeated Width Scaling calculations
        .lWidth = ScaleX(.dCustomWidth, mScaleUnits, vbPixels)
        
        .nAlignment = Alignment
        .nSortOrder = lgSTAscending
        .nType = DataType
        .sFormat = Format
        
        .bVisible = True
    End With
    
    DisplayChange
    
    AddColumn = lNewCol
End Function

Public Function AddItem(Optional ByVal Item As String, Optional Index As Long = 0, Optional Checked As Boolean) As Long
    '#############################################################################################################################
    'Purpose: Add an Item (new Row) to the Grid
    
    'Item       - This contains the data for the Cells in the new Row. You can pass multiple
    '           Cells by using a Delimiter between Cell data
    'Index      - Allows a new Item to be Inserted before an existing one
    'Checked    - Default Checked state of the new Item
    
    'mItems() is an array of the Items in the Grid
    'mIX() is used as an Index to the Items (a bit like an array of "pointers")
    
    'The Index technique is used to allow faster Inserts & Sorts since we only need to swap a Long (4 bytes)
    'rather than a large data structure (a UDT in this case)
    
    'The mItems() is resized incrementally to reduce the Redim Preserve overhead. The default mCacheIncrement
    'is 10 but this can be increased to a higher value to increase performance if adding thousands of Items
    '#############################################################################################################################

    Dim lCol As Long
    Dim lCount As Long
    Dim sText() As String
    
    mItemCount = mItemCount + 1
    If mItemCount > UBound(mItems) Then
        ReDim Preserve mItems(mItemCount + mCacheIncrement)
        ReDim Preserve mIX(mItemCount + mCacheIncrement)
    End If
    
    If (Index > 0) And (Index < mItemCount) Then
        If mItemCount > 1 Then
            For lCount = mItemCount To Index + 1 Step -1
                mIX(lCount) = mIX(lCount - 1)
            Next lCount
            mIX(Index) = mItemCount
        End If
        AddItem = Index
    Else
        mIX(mItemCount) = mItemCount
        AddItem = mItemCount
    End If
    
    If mRowHeightMin > 0 Then
        mItems(mItemCount).lHeight = ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
    Else
        mItems(mItemCount).lHeight = ROW_HEIGHT
    End If
    
    ReDim mItems(mItemCount).Cell(UBound(mCols))
        
    For lCount = LBound(mCols) To UBound(mCols)
        With mItems(mItemCount).Cell(lCount)
            .nAlignment = mCols(lCount).nAlignment
            .nFormat = -1
        End With
        
        ApplyCellFormat mItemCount, lCount, lgCFBackColor, mBackColor
        ApplyCellFormat mItemCount, lCount, lgCFForeColor, mForeColor
    Next lCount
    
    If UBound(mCols) > 0 Then
        lCol = 0
        sText() = Split(Item, vbTab)
        For lCount = LBound(sText) To UBound(sText)
            With mItems(mItemCount).Cell(lCol)
                .sValue = sText(lCount)
            End With
            
            lCol = lCol + 1
            If lCol > UBound(mCols) Then
                Exit For
            End If
        Next lCount
    Else
        mItems(mItemCount).Cell(0).sValue = Item
    End If
    
    If Checked Then
        SetFlag mItems(mItemCount).nFlags, lgChecked, True
    End If
    
    DisplayChange
    
    If mRow < 0 Then
        If mItemCount >= 0 Then
            Row = 0
        End If
    End If
    RaiseEvent ItemCountChanged
End Function

Public Property Get AllowUserResizing() As lgAllowUserResizingEnum
    AllowUserResizing = mAllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal NewValue As lgAllowUserResizingEnum)
    mAllowUserResizing = NewValue
    
    PropertyChanged "AllowUserResizing"
End Property

Private Sub ApplyCellFormat(Row As Long, Col As Long, Apply As lgCellFormatEnum, NewValue As Long)
    '#############################################################################################################################
    'Purpose: Apply formatting to a Cell. Attempts to find a matching entry in the
    'Format Table and creates a new entry if a match is not found.
    
    'In any "normal" use the grid will only have a few specifically formatted cells
    '(such as Red forecolor in a financial column to indicate negative). It is therefore
    'wasteful for each cell to store these properties. This system significantly reduces
    'the memory used by the cells in a large Grid at the cost of slightly reduced perfomance.
    
    'The Format element is an Integer allowing 32767 color combinations. It could be a
    'long for more combinations - however the aim is to keep the Cell UDT as small as possible!
    
    'This table may expand to have extra properties such as FontName & FontSize...
  
    Dim lBackColor As Long
    Dim lForeColor As Long
    Dim nCount As Integer
    Dim nIndex As Integer
    Dim nFreeIndex As Integer
    Dim nNewIndex As Integer
    Dim bMatch As Boolean
    
    nIndex = mItems(Row).Cell(Col).nFormat
    
    Select Case Apply
        Case lgCFBackColor
            lBackColor = NewValue
            
            If nIndex >= 0 Then
                lForeColor = mCF(nIndex).lForeColor
            Else
                lForeColor = mForeColor
            End If
        
        Case lgCFForeColor
            If nIndex >= 0 Then
                lBackColor = mCF(nIndex).lBackColor
            Else
                lBackColor = mBackColor
            End If
            
            lForeColor = NewValue
            
    End Select
    
    'Search Color Table for matching entry
    nFreeIndex = -1
    For nCount = 0 To UBound(mCF)
        'With mCF(nCount)
            If (mCF(nCount).lBackColor = lBackColor) And (mCF(nCount).lForeColor = lForeColor) Then
                'Existing Entry matches what we required
                bMatch = True
                nNewIndex = nCount
                Exit For
            ElseIf (mCF(nCount).nCount = 0) And (nFreeIndex = -1) Then
                'An unused entry
                nFreeIndex = nCount
            End If
        'End With
    Next nCount
    
    'No existing matches
    If Not bMatch Then
        'Is there an unused Entry?
        If nFreeIndex >= 0 Then
            nNewIndex = nFreeIndex
        Else
            nNewIndex = UBound(mCF) + 1
            ReDim Preserve mCF(nNewIndex + 9)
        End If
        
        With mCF(nNewIndex)
            .lBackColor = lBackColor
            .lForeColor = lForeColor
        End With
    End If
    
    'Has the Format Entry Index changed?
    If (nIndex <> nNewIndex) Then
        'Increment reference count for new entry
        mCF(nNewIndex).nCount = mCF(nNewIndex).nCount + 1
           
        If nIndex >= 0 Then
            'Decrement reference count for previous entry
            mCF(nIndex).nCount = mCF(nIndex).nCount - 1
        End If
    End If
        
    mItems(Row).Cell(Col).nFormat = nNewIndex
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    mBackColor = NewValue
    DrawGrid
    
    PropertyChanged "BackColor"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
    BackColorBkg = mBackColorBkg
End Property

Public Property Let BackColorBkg(ByVal NewValue As OLE_COLOR)
    mBackColorBkg = NewValue
    UserControl.BackColor = mBackColorBkg
    DisplayChange
    
    PropertyChanged "BackColorBkg"
End Property

Public Property Get BackColorEdit() As OLE_COLOR
    BackColorEdit = mBackColorEdit
End Property

Public Property Let BackColorEdit(ByVal lNewValue As OLE_COLOR)
    mBackColorEdit = lNewValue
    
    PropertyChanged "BackColorEdit"
End Property

Public Property Get BackColorFixed() As OLE_COLOR
    BackColorFixed = mBackColorFixed
End Property

Public Property Let BackColorFixed(ByVal NewValue As OLE_COLOR)
    mBackColorFixed = NewValue
    
    PropertyChanged "BackColorFixed"
End Property

Public Property Get BackColorSel() As OLE_COLOR
    BackColorSel = mBackColorSel
End Property

Public Property Let BackColorSel(ByVal NewValue As OLE_COLOR)
    mBackColorSel = NewValue
    DisplayChange
    
    PropertyChanged "BackColorSel"
End Property

Public Sub BindControl(ByVal Col As Long, Ctrl As Object, Optional MoveControl As lgMoveControlEnum = lgBCHeight Or lgBCLeft Or lgBCTop Or lgBCWidth)
    '#############################################################################################################################
    'Purpose: Bind an external Control to a Column
    
    'Col    - Column Index
    'Ctrl   - The Control!
    'Resize - Specify how the Control Size should be modified
    '#############################################################################################################################

    Set mCols(Col).EditCtrl = Ctrl
    mCols(Col).MoveControl = MoveControl
End Sub

Public Property Get BorderStyle() As lgBorderStyleEnum
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As lgBorderStyleEnum)
    mBorderStyle = NewValue
    UserControl.BorderStyle = mBorderStyle
    
    PropertyChanged "BorderStyle"
End Property

Public Property Get CacheIncrement() As Long
    CacheIncrement = mCacheIncrement
End Property

Public Property Let CacheIncrement(ByVal NewValue As Long)
    If NewValue < 0 Then
        mCacheIncrement = 1
    Else
        mCacheIncrement = NewValue
    End If
    
    PropertyChanged "CacheIncrement"
End Property

Public Property Let CellAlignment(Row As Long, Col As Long, NewValue As lgAlignmentEnum)
    mItems(mIX(Row)).Cell(Col).nAlignment = NewValue
    DrawGrid
End Property

Public Property Get CellAlignment(Row As Long, Col As Long) As lgAlignmentEnum
    CellAlignment = mItems(mIX(Row)).Cell(Col).nAlignment
End Property

Public Property Let CellBackColor(Row As Long, Col As Long, NewValue As Long)
    ApplyCellFormat Row, Col, lgCFBackColor, NewValue
    
    DrawGrid
End Property

Public Property Get CellBackColor(Row As Long, Col As Long) As Long
    CellBackColor = mCF(mItems(mIX(Row)).Cell(Col).nFormat).lBackColor
End Property

Public Property Let CellChecked(Row As Long, Col As Long, NewValue As Boolean)
    SetFlag mItems(mIX(Row)).Cell(Col).nFlags, lgChecked, NewValue
    DrawGrid
End Property

Public Property Get CellChecked(Row As Long, Col As Long) As Boolean
    CellChecked = mItems(mIX(Row)).Cell(Col).nFlags And lgChecked
End Property

Public Property Let CellFontBold(Row As Long, Col As Long, NewValue As Boolean)
    SetFlag mItems(mIX(Row)).Cell(Col).nFlags, lgFontBold, NewValue
    DrawGrid
End Property

Public Property Get CellFontBold(Row As Long, Col As Long) As Boolean
    CellFontBold = mItems(mIX(Row)).Cell(Col).nFlags And lgFontBold
End Property

Public Property Let CellFontItalic(Row As Long, Col As Long, NewValue As Boolean)
    SetFlag mItems(mIX(Row)).Cell(Col).nFlags, lgFontItalic, NewValue
    DrawGrid
End Property

Public Property Get CellFontItalic(Row As Long, Col As Long) As Boolean
    CellFontItalic = mItems(mIX(Row)).Cell(Col).nFlags And lgFontItalic
End Property

Public Property Let CellFontUnderline(Row As Long, Col As Long, NewValue As Boolean)
    SetFlag mItems(mIX(Row)).Cell(Col).nFlags, lgFontUnderline, NewValue
    DrawGrid
End Property

Public Property Get CellFontUnderline(Row As Long, Col As Long) As Boolean
    CellFontUnderline = mItems(mIX(Row)).Cell(Col).nFlags And lgFontUnderline
End Property

Public Property Let CellForeColor(Row As Long, Col As Long, NewValue As Long)
    ApplyCellFormat Row, Col, lgCFForeColor, NewValue
    DrawGrid
End Property

Public Property Get CellForeColor(Row As Long, Col As Long) As Long
    CellForeColor = mCF(mItems(mIX(Row)).Cell(Col).nFormat).lForeColor
End Property

Public Property Let CellProgressValue(Row As Long, Col As Long, NewValue As Integer)
    If mCols(Col).nType = lgProgressBar Then
        If NewValue > 100 Then
            NewValue = 100
        ElseIf NewValue < 0 Then
            NewValue = 0
        End If
        
        mItems(mIX(Row)).Cell(Col).nFlags = NewValue
        DrawGrid
    End If
End Property

Public Property Get CellProgressValue(Row As Long, Col As Long) As Integer
    If mCols(Col).nType = lgProgressBar Then
        CellProgressValue = mItems(mIX(Row)).Cell(Col).nFlags
    End If
End Property

Public Property Let CellText(Row As Long, Col As Long, NewValue As String)
    mItems(mIX(Row)).Cell(Col).sValue = NewValue
    DrawGrid
End Property

Public Property Get CellText(Row As Long, Col As Long) As String
    
    On Error GoTo Errh

    CellText = mItems(mIX(Row)).Cell(Col).sValue

    Exit Property
    
Errh:
    Err.Clear
    'temp
    'added by: Vincent J.Jamero
    CellText = mItems(mIX(0)).Cell(Col).sValue
End Property

Public Property Get CheckBoxes() As Boolean
    CheckBoxes = mCheckboxes
End Property

Public Property Let CheckBoxes(ByVal NewValue As Boolean)
    mCheckboxes = NewValue
    DisplayChange
    
    PropertyChanged "CheckBoxes"
End Property

Public Function CheckedCount() As Long
    '#############################################################################################################################
    'Purpose: Return Count of Checked Items
    '#############################################################################################################################
    
    Dim lCount As Long
    
    For lCount = LBound(mItems) To UBound(mItems)
        If mItems(lCount).nFlags And lgChecked Then
            CheckedCount = CheckedCount + 1
        End If
    Next lCount
End Function

Public Sub Clear()
    '#############################################################################################################################
    'Purpose: Remove all Items from the Grid. Does not affect Column Headers
    '#############################################################################################################################
  
    ReDim mItems(0)
    ReDim mIX(0)
    ReDim mCF(0)
    
    mMouseDownCol = NULL_RESULT
    mMouseDownRow = NULL_RESULT
    
    mCol = NULL_RESULT
    mRow = NULL_RESULT
    mSelectedRow = NULL_RESULT
    
    mHotColumn = NULL_RESULT
    mResizeCol = NULL_RESULT
    
    mSortColumn = NULL_RESULT
    mSortSubColumn = NULL_RESULT
    
    mScrollAction = SCROLL_NONE
    mItemCount = -1
End Sub

Public Property Get Col() As Long
    Col = mCol
End Property

Public Property Let Col(ByVal NewValue As Long)
    If SetRowCol(mRow, NewValue) Then
        DrawGrid
    End If
End Property

Public Property Get ColAlignment(ByVal Index As Long) As lgAlignmentEnum
    ColAlignment = mCols(Index).nAlignment
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal NewValue As lgAlignmentEnum)
    mCols(Index).nAlignment = NewValue
    
    DrawGrid
End Property

Public Property Get ColFormat(ByVal Index As Long) As String
    ColFormat = mCols(Index).sFormat
End Property

Public Property Let ColFormat(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sFormat = NewValue
    
    DrawGrid
End Property

Public Property Get ColHeading(ByVal Index As Long) As String
    ColHeading = mCols(Index).sCaption
End Property

Public Property Let ColHeading(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sCaption = NewValue
    
    DrawGrid
End Property

Public Function ColLeft(ByVal Index As Long) As Long
    Dim R As RECT
    
    SetColRect Index, R
    ColLeft = R.Left
End Function

Public Property Get Cols() As Long
    Cols = UBound(mCols)
End Property

Public Property Let Cols(ByVal NewValue As Long)
    ReDim mCols(NewValue)
End Property

Public Property Get ColType(ByVal Index As Long) As lgDataTypeEnum
    ColType = mCols(Index).nType
End Property

Public Property Let ColType(ByVal Index As Long, ByVal NewValue As lgDataTypeEnum)
    mCols(Index).nType = NewValue
End Property

Public Property Get ColumnSort() As Boolean
    ColumnSort = mColumnSort
End Property

Public Property Let ColumnSort(ByVal NewValue As Boolean)
    mColumnSort = NewValue
End Property

Public Property Get ColTag(ByVal Index As Long) As String
    ColTag = mCols(Index).sTag
End Property

Public Property Let ColTag(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sTag = NewValue
End Property

Public Property Get ColVisible(ByVal Index As Long) As Boolean
    ColVisible = mCols(Index).bVisible
End Property

Public Property Let ColVisible(ByVal Index As Long, ByVal NewValue As Boolean)
    mCols(Index).bVisible = NewValue
    
    DrawGrid
End Property

Public Property Get ColWidth(ByVal Index As Long) As Single
    ColWidth = mCols(Index).dCustomWidth
End Property

Public Property Let ColWidth(ByVal Index As Long, ByVal NewValue As Single)
    'dCustomWidth is in the Units the Control is operating in
    mCols(Index).dCustomWidth = NewValue
    mCols(Index).lWidth = ScaleX(NewValue, mScaleUnits, vbPixels)
       
    DrawGrid
End Property

Private Sub CreateRenderData()
    '#############################################################################################################################
    'Purpose: Calculates rendering parameters & sets display options. Used
    'to prevent unneccesary recalculations when redrawing the Grid
    '#############################################################################################################################
   
    Dim lSize As Long
    
    With mR
        lSize = ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
        If lSize > MAX_CHECKBOXSIZE Then
            .CheckBoxSize = MAX_CHECKBOXSIZE
        Else
            .CheckBoxSize = lSize - 4
        End If

        If mCheckboxes Then
            .LeftText = .CheckBoxSize
        Else
            .LeftImage = 0
            .LeftText = 3
        End If
        
        .LeftImage = .LeftText
        
        If mImageList Is Nothing Then
            .ImageSpace = 0
        Else
            .ImageSpace = ((GetRowHeight() - mImageList.ImageHeight) / 2)
            .LeftText = .LeftText + mImageList.ImageWidth + 2
        End If
        
        .HeaderHeight = GetColumnHeadingHeight()
        .TextHeight = UserControl.TextHeight("A")
        
        If mDisplayEllipsis Then
            .DTFlag = DT_SINGLELINE Or DT_WORD_ELLIPSIS
        Else
            .DTFlag = DT_SINGLELINE
        End If
    End With
End Sub

#If DEBUG_MODE Then
    Public Sub DebugFormatTable()
        Dim nCount As Integer
    
         For nCount = 0 To UBound(mCF)
            With mCF(nCount)
                Debug.Print Format$(nCount, "00000") & " " & .nCount & "," & .lBackColor & "," & .lForeColor
            End With
        Next nCount
    End Sub
#End If

Private Sub DisplayChange()
    If mRedraw Then
        Refresh
    Else
        mPendingRedraw = True
        mPendingScrollBar = True
    End If
End Sub

Public Property Get DisplayEllipsis() As Boolean
    DisplayEllipsis = mDisplayEllipsis
End Property

Public Property Let DisplayEllipsis(ByVal NewValue As Boolean)
    mDisplayEllipsis = NewValue
    DisplayChange
    
    PropertyChanged "DisplayEllipsis"
End Property

Public Sub Sort(Optional Sort As Long = -1, Optional SortType As lgSortTypeEnum = -1, Optional SubSort As Long = -1, Optional SubSortType As lgSortTypeEnum = -1)
    '#############################################################################################################################
    'Purpose: Sort Grid based on current Sort Columns.
    '#############################################################################################################################
    
    Dim lCount As Long
    Dim lRowIndex As Long
    
    If UpdateCell() Then
        'Set new Columns if specified
        If Sort <> -1 Then
            mSortColumn = Sort
        End If
        
        If SubSort <> -1 Then
            mSortSubColumn = SubSort
        End If
        
        'Validate Sort Columns
        If (mSortColumn = NULL_RESULT) And (mSortSubColumn <> NULL_RESULT) Then
            mSortColumn = mSortSubColumn
            mSortSubColumn = NULL_RESULT
        ElseIf mSortColumn = mSortSubColumn Then
            mSortSubColumn = NULL_RESULT
        End If
        
        'Set Sort Order if specified - otherwise inverse last Sort Order
        With mCols(mSortColumn)
            If SortType = -1 Then
                If .nSortOrder = lgSTAscending Then
                    .nSortOrder = lgSTDescending
                Else
                    .nSortOrder = lgSTAscending
                End If
            Else
                .nSortOrder = SortType
            End If
        End With
        
        If mSortSubColumn <> NULL_RESULT Then
            With mCols(mSortSubColumn)
                If SubSortType = -1 Then
                    If .nSortOrder = lgSTAscending Then
                        .nSortOrder = lgSTDescending
                    Else
                        .nSortOrder = lgSTAscending
                    End If
                Else
                    .nSortOrder = SubSortType
                End If
            End With
        End If
        
        'Note previously selected Row
        If mRow > NULL_RESULT Then
            lRowIndex = mIX(mRow)
        End If
        
        SortArray LBound(mItems), mItemCount, mSortColumn, mCols(mSortColumn).nSortOrder
        SortSubList
        
        For lCount = LBound(mIX) To mItemCount
            If mIX(lCount) = lRowIndex Then
                mRow = lCount
                Exit For
            End If
        Next lCount
        
        DrawGrid True
        
        RaiseEvent SortComplete
    End If
End Sub


Public Sub DrawGrid(Optional bForceRedraw As Boolean)
    '#############################################################################################################################
    'Purpose: The Primary Rendering routine. Draws Columns & Rows
    '#############################################################################################################################

    On Error GoTo Errh

    Dim R As RECT
    Dim lX As Long
    Dim lY As Long
    
    Dim lCol As Long
    
    Dim lMaxRow As Long
    Dim lStartCol As Long
    Dim lColumnsWidth As Long
    Dim lBottomEdge As Long
    Dim lGridColor As Long
    Dim lImageLeft As Long
    Dim lValue As Long
    Dim nImageListScaleMode As Integer
    Dim bLockColor As Boolean
    Dim sText As String
    Dim bBold As Boolean
    Dim bItalic As Boolean
    Dim bUnderLine As Boolean
    
    If mRedraw Or bForceRedraw Then
        lStartCol = SBValue(efsHorizontal)
        lGridColor = TranslateColor(mGridColor)
        
        lY = mR.HeaderHeight
        mItemsVisible = ItemsVisible()
    
        With UserControl
            .Cls
            
            bBold = .FontBold
            bItalic = .FontItalic
            bUnderLine = .FontUnderline
            
            lColumnsWidth = DrawHeaderRow()
            
            lMaxRow = (SBValue(efsVertical) + mItemsVisible)
            If lMaxRow > mItemCount Then
                lMaxRow = mItemCount
            End If
            
            RaiseEvent BeforeDrawGrid(SBValue(efsVertical), lMaxRow)
                
            For lRow = SBValue(efsVertical) To lMaxRow
                
               
                If (mMultiSelect Or mFullRowSelect) And (mItems(mIX(lRow)).nFlags And lgSelected) Then
                     'select item
                    bLockColor = True
                    If lStartCol = 0 Then ' ensure 1st column is visible
                        If mCols(0).lWidth < mR.LeftText Then
                            SetRect R, 0, lY + 1, mCols(0).lWidth, lY + (mItems(mIX(lRow)).lHeight)
                        Else
                            SetRect R, 0, lY + 1, mR.LeftText, lY + (mItems(mIX(lRow)).lHeight)
                        End If
                        DrawRect .hDc, R, TranslateColor(mBackColor), True
                    Else
                        R.Right = 0
                    End If
                    
                    SetRect R, R.Right, lY + 1, lColumnsWidth, lY + (mItems(mIX(lRow)).lHeight)
                    DrawRect .hDc, R, TranslateColor(mBackColorSel), True
                    .ForeColor = mForeColorSel
                Else
                    bLockColor = False
                    SetRect R, 0, lY + 1, lColumnsWidth, lY + (mItems(mIX(lRow)).lHeight)
                    DrawRect .hDc, R, TranslateColor(mBackColor), True, True
                End If
                
                lX = 0
                
                
                For lCol = lStartCol To UBound(mCols)
                    If mCols(lCol).bVisible Then
                        SetRectRgn mClipRgn, lX, lY, lX + mCols(lCol).lWidth, lY + mItems(mIX(lRow)).lHeight
                        SelectClipRgn .hDc, mClipRgn

                        Call SetRect(R, lX, lY, lX + mCols(lCol).lWidth, lY + mItems(mIX(lRow)).lHeight)
                        
                        If Not bLockColor Then
                            If mCF(mItems(mIX(lRow)).Cell(lCol).nFormat).lBackColor <> mBackColor Then
                                DrawRect .hDc, R, TranslateColor(mCF(mItems(mIX(lRow)).Cell(lCol).nFormat).lBackColor), True
                            End If
                            .ForeColor = mCF(mItems(mIX(lRow)).Cell(lCol).nFormat).lForeColor
                        End If
                        
                        If lCol = 0 Then
                            If mCheckboxes Then
                                Call SetRect(R, 3, lY, mR.CheckBoxSize, lY + mItems(mIX(lRow)).lHeight)
                                
                                If mItems(mIX(lRow)).nFlags And lgChecked Then
                                    Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                Else
                                    Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                End If
                            End If
                            
                            If mR.ImageSpace > 0 Then
                                'If we have an Image Index then Draw it
                                If mItems(mIX(lRow)).lImage <> 0 Then
                                    'Calculate Image offset (using ScaleMode of ImageList)
                                    If lImageLeft = 0 Then
                                        nImageListScaleMode = mImageList.Parent.ScaleMode
                                        lImageLeft = ScaleX(mR.LeftImage, vbPixels, nImageListScaleMode)
                                    End If
                                    
                                    If bLockColor Then
                                        mImageList.ListImages(Abs(mItems(mIX(lRow)).lImage)).Draw .hDc, lImageLeft, ScaleY(lY + mR.ImageSpace, vbPixels, nImageListScaleMode), 2
                                    Else
                                        mImageList.ListImages(Abs(mItems(mIX(lRow)).lImage)).Draw .hDc, lImageLeft, ScaleY(lY + mR.ImageSpace, vbPixels, nImageListScaleMode), 1
                                    End If
                                End If
                            End If
                            
                            Call SetRect(R, mR.LeftText + TEXT_SPACE, lY, (lX + mCols(lCol).lWidth) - TEXT_SPACE, lY + mItems(mIX(lRow)).lHeight)
                        Else
                            Call SetRect(R, lX + TEXT_SPACE, lY, (lX + mCols(lCol).lWidth) - TEXT_SPACE, lY + mItems(mIX(lRow)).lHeight)
                        End If
                       
                        Select Case mCols(lCol).nType
                            Case lgBoolean
                                SetCheckBoxRect mIX(lRow), lCol, lY, R
                                
                                If mItems(mIX(lRow)).Cell(lCol).nFlags And lgChecked Then
                                    Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                Else
                                    Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                End If
                                
                            Case lgProgressBar
                                If mItems(mIX(lRow)).Cell(lCol).nFlags > 0 Then
                                    lValue = ((mCols(lCol).lWidth - 2) / 100) * mItems(mIX(lRow)).Cell(lCol).nFlags
                                
                                    SetRect R, lX + 2, lY + 2, lX + lValue, (lY + mItems(mIX(lRow)).lHeight) - 2
                                    DrawRect .hDc, R, TranslateColor(mProgressBarColor), True
                                End If
                            
                            Case Else
                                With mItems(mIX(lRow)).Cell(lCol)
                                    UserControl.FontBold = .nFlags And lgFontBold
                                    UserControl.FontItalic = .nFlags And lgFontItalic
                                    UserControl.FontUnderline = .nFlags And lgFontUnderline
                                    
                                    If Len(mCols(lCol).sFormat) > 0 Then
                                        sText = Format$(.sValue, mCols(lCol).sFormat)
                                    Else
                                        sText = .sValue
                                    End If
                                    
                                    'added by: VincentJ. Jamero
                                    RaiseEvent BeforeDrawText(mIX(lRow), lCol, sText)
                                    .sValue = sText
                                    
                                    Call DrawText(UserControl.hDc, sText, -1, R, .nAlignment Or mR.DTFlag)
                                End With
                        End Select
                        
                        lX = lX + mCols(lCol).lWidth
                    End If
                Next lCol
                
                SelectClipRgn .hDc, 0&
                
                'Display Horizontal Lines
                If mGridLines Then
                    DrawLine .hDc, 0, lY, lColumnsWidth, lY, lGridColor, mGridLineWidth
                End If
                
                lY = lY + mItems(mIX(lRow)).lHeight
            Next lRow
            
            '#############################################################################################################################
            'Display Vertical Lines
            If mGridLines Then
                lBottomEdge = R.Bottom
            
                lX = 0
                For lCol = lStartCol To UBound(mCols)
                    If mCols(lCol).bVisible Then
                        DrawLine .hDc, lX, mR.HeaderHeight, lX, lBottomEdge, lGridColor, mGridLineWidth
                    
                        lX = lX + mCols(lCol).lWidth
                    End If
                Next lCol
            End If
            
            '#############################################################################################################################
            'Display Focus Rectangle
            If mInFocus And (mFocusRectMode <> lgFocusRectModeEnum.lgNone) And (mRow >= 0) Then
                lY = RowTop(mRow)
                R.Right = 0
                If mFocusRectMode = lgCol Then
                    If mCol >= 0 Then
                        lX = ColLeft(mCol)
                        SetRect R, lX, lY + 1, lX + mCols(mCol).lWidth, lY + mItems(mIX(mRow)).lHeight
                    End If
                ElseIf mFullRowSelect Then
                    SetRect R, mR.LeftText, lY + 1, lColumnsWidth, lY + mItems(mIX(mRow)).lHeight
                End If
                
                If R.Right > 0 Then
                    Select Case mFocusRectStyle
                        Case lgFRLight
                            Call DrawFocusRect(.hDc, R)
            
                        Case lgFRHeavy
                            DrawRect .hDc, R, TranslateColor(mFocusRectColor), False
                    End Select
                End If
            End If
            
            .Refresh
            
            .FontBold = bBold
            .FontItalic = bItalic
            .FontUnderline = bUnderLine
        End With
        
        'Debug.Print "DrawGrid " & Timer
        
        mPendingRedraw = False
    Else
        mPendingRedraw = True
    End If
    
    RaiseEvent AfterDrawGrid
    Exit Sub
    
Errh:

End Sub
Private Function LongToSignedShort(dwUnsigned As Long) As Integer
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
End Function
Private Sub FillGradient(lhDC As Long, rRect As RECT, ByVal clrFirst As OLE_COLOR, ByVal clrSecond As OLE_COLOR, Optional ByVal bVertical As Boolean)
    Dim pVert(0 To 1)   As TRIVERTEX
    Dim pGradRect       As GRADIENT_RECT
    
    With pVert(0)
        .X = rRect.Left
        .Y = rRect.Top
        .Red = LongToSignedShort((clrFirst And &HFF&) * 256)
        .Green = LongToSignedShort(((clrFirst And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((clrFirst And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With pVert(1)
        .X = rRect.Right
        .Y = rRect.Bottom
        .Red = LongToSignedShort((clrSecond And &HFF&) * 256)
        .Green = LongToSignedShort(((clrSecond And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((clrSecond And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With pGradRect
        .UPPERLEFT = 0
        .LOWERRIGHT = 1
    End With
    
    GradientFill lhDC, pVert(0), 2, pGradRect, 1, IIf(Not bVertical, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)
End Sub
    
Private Sub DrawHeader(lCol As Long, State As lgHeaderStateEnum)
    '#############################################################################################################################
    'Purpose: Renders a Column Header. This involves drawing the Border, displaying
    'the Caption and optionally Sort Arrows
    '#############################################################################################################################

    Dim R As RECT
    
    If lCol > NULL_RESULT Then
        With UserControl
            .ForeColor = mForeColor
            
            'Draw the Column Headers
            Call SetRect(R, mCols(lCol).lX, 0, mCols(lCol).lX + mCols(lCol).lWidth + 1, mR.HeaderHeight)
            DrawRect .hDc, R, TranslateColor(BackColorFixed), True
            
            Select Case mThemeStyle
                Case lgTSWindows3D
                    Select Case State
                         Case lgNormal
                             Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONPUSH)
                         Case lgHot
                             Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_HOT)
                         Case lgDown
                             Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
                     End Select
             
                Case lgTSWindowsFlat
                    Select Case State
                         Case lgNormal
                             Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_FLAT)
                         Case lgHot
                             Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_HOT)
                         Case lgDown
                             Call DrawFrameControl(.hDc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
                     End Select
                
                Case lgTSWindowsXP
                    'Try XP Theme API
                    If Not DrawTheme("Header", 1, State, R) Then
                        'Use XP emulation
                        DrawXPHeader .hDc, R, State
                    End If
                
                Case lgTSOfficeXP
                    DrawOfficeXPHeader .hDc, R, State
                   
            End Select
            
            'Render Sort Arrows
            If mCols(lCol).lWidth > SIZE_SORTARROW Then
                If lCol = mSortColumn Then
                    DrawSortArrow (mCols(lCol).lX + mCols(lCol).lWidth) - 12, 6, 9, 5, mCols(lCol).nSortOrder
                    
                    Call SetRect(R, mCols(lCol).lX + HEADER_LEFT, 0, (mCols(lCol).lX + mCols(lCol).lWidth) - (ARROW_SPACE + SIZE_SORTARROW), mR.HeaderHeight)
                ElseIf lCol = mSortSubColumn Then
                    DrawSortArrow (mCols(lCol).lX + mCols(lCol).lWidth) - 12, 6, 6, 3, mCols(lCol).nSortOrder
                    
                    Call SetRect(R, mCols(lCol).lX + HEADER_LEFT, 0, (mCols(lCol).lX + mCols(lCol).lWidth) - (ARROW_SPACE + SIZE_SORTARROW), mR.HeaderHeight)
                Else
                    Call SetRect(R, mCols(lCol).lX + HEADER_LEFT, 0, (mCols(lCol).lX + mCols(lCol).lWidth) - (HEADER_LEFT * 2), mR.HeaderHeight)
                End If
            Else
                Call SetRect(R, mCols(lCol).lX + HEADER_LEFT, 0, (mCols(lCol).lX + mCols(lCol).lWidth) - (HEADER_LEFT * 2), mR.HeaderHeight)
            End If
            
            Call DrawText(.hDc, mCols(lCol).sCaption, -1, R, mCols(lCol).nAlignment Or mR.DTFlag)
            
        End With
    End If
End Sub

Private Function DrawHeaderRow() As Long
    '#############################################################################################################################
    'Purpose: Renders all Column Headers
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim lX As Long
    
    mHotColumn = NULL_RESULT
    
    For lCol = SBValue(efsHorizontal) To UBound(mCols)
         If mCols(lCol).bVisible Then
            mCols(lCol).lX = lX
            DrawHeader lCol, lgNormal
            lX = lX + mCols(lCol).lWidth
        End If
    Next lCol
    
    DrawHeaderRow = lX
End Function

Private Function InvertThisColor(oInsColor As OLE_COLOR)
    '#############################################################################################################################
    'Source: Riccardo Cohen
    '#############################################################################################################################
    
    Dim lROut As Long, lGOut As Long, lBOut As Long
    Dim lRGB As Long
   
    lRGB = TranslateColor(oInsColor)
    
    lROut = (255 - (lRGB And &HFF&))
    lGOut = (255 - ((lRGB And &HFF00&) / &H100))
    lBOut = (255 - ((lRGB And &HFF0000) / &H10000))
    InvertThisColor = RGB(lROut, lGOut, lBOut)
End Function


Private Sub DrawLine(hDc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long, lWidth As Long)
    Dim PT As POINTAPI
    Dim hPen As Long
    Dim hPenOld As Long
    
    hPen = CreatePen(0, lWidth, lColor)
    hPenOld = SelectObject(hDc, hPen)
    MoveToEx hDc, X1, Y1, PT
    LineTo hDc, X2, Y2
    SelectObject hDc, hPenOld
    DeleteObject hPen
End Sub

Private Sub DrawOfficeXPHeader(lhDC As Long, rRect As RECT, State As lgHeaderStateEnum)
    '#############################################################################################################################
    'Purpose:   Draw a Column Header in Office XP Style
    'Notes:     Created from original source by Riccardo Cohen
    '#############################################################################################################################
    
    With rRect
        Select Case State
            Case lgNormal
                Call FillGradient(lhDC, rRect, &HFCE1CB, &HE0A57D, True)
                
                DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
                
                DrawLine lhDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, &HCB8C6A, 1
                DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1

            Case lgHot
                .Right = .Right - 1
                Call FillGradient(lhDC, rRect, &HDCFFFF, &H5BC0F7, True)
                
                DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
                
                DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1

            Case lgDown
                .Right = .Right - 1
                Call FillGradient(lhDC, rRect, &H87FE8, &H7CDAF7, True)
                
                DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
                
                DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1
                
        End Select
    End With
End Sub

Private Sub DrawXPHeader(lhDC As Long, rRect As RECT, State As lgHeaderStateEnum)
    '#############################################################################################################################
    'Purpose:   Draw a Column Header in XP Style
    'Notes:     Created from original source by Riccardo Cohen
    '#############################################################################################################################
    
    Dim TempColor As OLE_COLOR

    With rRect
        Select Case State
            Case lgNormal
                DrawRect lhDC, rRect, TranslateColor(vbButtonFace), True
        
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &HB2C2C5, 1
                DrawLine lhDC, .Left, .Bottom - 2, .Right, .Bottom - 2, &HBECFD2, 1
                DrawLine lhDC, .Left, .Bottom - 3, .Right, .Bottom - 3, &HC8D8DC, 1
                
                DrawLine lhDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, &H99A8AC, 1
                DrawLine lhDC, .Left, .Top + 2, .Left, .Bottom - 4, &HFFFFFF, 1
                
            Case lgHot
                DrawRect lhDC, rRect, &HF3F8FA, True
                
                DrawLine lhDC, .Left + 2, .Bottom - 1, .Right - 2, .Bottom - 1, &H19B1F9, 1
                DrawLine lhDC, .Left + 1, .Bottom - 2, .Right - 1, .Bottom - 2, &H47C2FC, 1
                DrawLine lhDC, .Left, .Bottom - 3, .Right, .Bottom - 3, 43512, 1

            Case lgDown
                TempColor = ForeColor
                
                UserControl.ForeColor = InvertThisColor(TempColor)
                .Bottom = .Bottom - 1
                DrawRect lhDC, rRect, &H0&, True
                
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, InvertThisColor(&HB2C2C5), 1
                DrawLine lhDC, .Left, .Bottom - 2, .Right, .Bottom - 2, InvertThisColor(&HBECFD2), 1
                DrawLine lhDC, .Left, .Bottom - 3, .Right, .Bottom - 3, InvertThisColor(&HC8D8DC), 1
                DrawLine lhDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, InvertThisColor(&H99A8AC), 1
                DrawLine lhDC, .Left, .Top + 2, .Left, .Bottom - 4, InvertThisColor(&HFFFFFF), 1
        End Select
    End With
End Sub


Private Sub DrawRect(hDc As Long, rc As RECT, lColor As Long, bFilled As Boolean, Optional bDrawStriped As Boolean = False)
    Dim lNewBrush As Long
  
    If mStriped = True And bDrawStriped = True Then
        If mBackColor = -2147483643 Or mBackColor = lColor Then
            If lRow Mod 2 = 0 Then
                lColor = mSBackColor1
            Else
                lColor = mSBackColor2
            End If
        End If
        
        lNewBrush = CreateSolidBrush(lColor)
        
    Else
        lNewBrush = CreateSolidBrush(lColor)
    End If
    
    If bFilled Then
        Call FillRect(hDc, rc, lNewBrush)
    Else
        Call FrameRect(hDc, rc, lNewBrush)
    End If

    Call DeleteObject(lNewBrush)
End Sub

Private Sub DrawSortArrow(lX As Long, lY As Long, lWidth As Long, lStep As Long, nOrientation As lgSortTypeEnum)
    '#############################################################################################################################
    'Purpose: Renders the Sort/Sub-Sort arrows
    '#############################################################################################################################
   
    Dim hPenOld As Long
    Dim hPen As Long
    Dim lCount As Long
    Dim lVerticalChange As Long
    Dim X1 As Long
    Dim X2 As Long
    Dim Y1 As Long
    
    hPen = CreatePen(0, 1, TranslateColor(vbButtonShadow))
    hPenOld = SelectObject(hDc, hPen)
    
    If nOrientation = lgSTDescending Then
        lVerticalChange = -1
        lY = lY + lStep - 1
    Else
        lVerticalChange = 1
    End If
    
    X1 = lX
    X2 = lWidth
    Y1 = lY
        
    MoveTo hDc, X1, Y1, ByVal 0&
    
    For lCount = 1 To lStep
        LineTo hDc, X1 + X2, Y1
        X1 = X1 + 1
        Y1 = Y1 + lVerticalChange
        X2 = X2 - 2
        MoveTo hDc, X1, Y1, ByVal 0&
    Next lCount
    
    Call SelectObject(hDc, hPenOld)
    Call DeleteObject(hPen)
End Sub

Private Sub DrawText(ByVal hDc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)
    '#############################################################################################################################
    'Purpose: Renders the Text for Column Headers & Cells. On Windows NT/2000/XP
    '(or better) the Control supports Unicode
    '#############################################################################################################################
   
    If mWindowsNT Then
        DrawTextW hDc, StrPtr(lpString), nCount, lpRect, wFormat
    Else
        DrawTextA hDc, lpString, nCount, lpRect, wFormat
    End If
End Sub

Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT, Optional ByVal CloseTheme As Boolean = False) As Boolean
    '#############################################################################################################################
    'Purpose: On Windows XP allows certain elements of the Grid to be drawn using
    'the current Windows Theme
    '#############################################################################################################################
    
    Dim lResult As Long
    
    On Error GoTo DrawThemeError
    
    If mWindowsXP Then
        hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
        If (hTheme) Then
            lResult = DrawThemeBackground(hTheme, UserControl.hDc, iPart, iState, rtRect, rtRect)
            DrawTheme = (lResult = 0)
        Else
            DrawTheme = False
        End If
        
        If CloseTheme Then
            Call CloseThemeData(hTheme)
        End If
    End If
    Exit Function

DrawThemeError:
    DrawTheme = False
End Function

Public Property Get Editable() As Boolean
    Editable = mEditable
End Property

Public Property Let Editable(ByVal NewValue As Boolean)
    mEditable = NewValue
    
    PropertyChanged "Editable"
End Property

Public Sub EditCell(ByVal Row As Long, ByVal Col As Long)
    '#############################################################################################################################
    'Purpose: Used to start an Edit. Note the RequestEdit event. This event allows
    'the Edit to be cancelled before anything visible occurs by setting the Cancel
    'flag.
    '#############################################################################################################################

    Dim R As RECT
    Dim bCancel As Boolean
    
    'Added by: Vincent J.Jamero
    If Row < 0 Or Row > RowCount Or Col < 0 Or Col > Cols Then
        Exit Sub
    End If
    
    If mEditPending Then
        If Not UpdateCell() Then
            Exit Sub
        End If
    End If
    
    If IsEditable() And (mCols(Col).nType <> lgBoolean) Then
        RaiseEvent RequestEdit(Row, Col, bCancel)
        If Not bCancel Then
            mEditCol = Col
            mEditRow = Row
            
            SetColRect mEditCol, R
            
            MoveEditControl mCols(mEditCol).MoveControl
            
            'Check if an external Control is used.
            If mCols(mEditCol).EditCtrl Is Nothing Then
                'Using internal TextBox
                With txtEdit
                    Select Case mItems(mIX(mEditRow)).Cell(mEditCol).nAlignment
                        Case lgAlignCenterBottom, lgAlignCenterCenter, lgAlignCenterTop
                            .Alignment = vbCenter
                        Case lgAlignLeftBottom, lgAlignLeftCenter, lgAlignLeftTop
                            .Alignment = vbLeftJustify
                        Case Else
                            .Alignment = vbRightJustify
                    End Select
                
                    .BackColor = mBackColorEdit
                    .FontBold = mItems(mIX(mEditRow)).Cell(mEditCol).nFlags And lgFontBold
                    .FontItalic = mItems(mIX(mEditRow)).Cell(mEditCol).nFlags And lgFontItalic
                    .FontUnderline = mItems(mIX(mEditRow)).Cell(mEditCol).nFlags And lgFontUnderline
                    .Text = mItems(mIX(mEditRow)).Cell(mEditCol).sValue
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .Visible = True
                    .SetFocus
                End With
            Else
                On Local Error Resume Next
                
                With mCols(mEditCol).EditCtrl
                    If UserControl.ContainerHwnd <> .Container.hwnd Then
                        mEditParent = UserControl.ContainerHwnd
                        SetParent .hwnd, UserControl.ContainerHwnd
                    Else
                        mEditParent = 0
                    End If
                    .Enabled = True
                    .Visible = True
                    .ZOrder
                    
                    Subclass_Start .hwnd
                    Call Subclass_AddMsg(.hwnd, WM_KILLFOCUS, MSG_AFTER)
                    
                    .SetFocus
                End With
                
                On Local Error GoTo 0
            End If
            
            mEditPending = True
        End If
    End If
End Sub

Public Property Get EditTrigger() As lgEditTriggerEnum
    EditTrigger = mEditTrigger
End Property

Public Property Let EditTrigger(ByVal NewValue As lgEditTriggerEnum)
    mEditTrigger = NewValue
    
    PropertyChanged "EditTrigger"
End Property

Public Function FindItem(ByVal SearchText As String, Optional ByVal SearchColumn As Long = -1, Optional SearchMode As lgSearchModeEnum = lgSMEqual, Optional MatchCase As Boolean) As Long
    '#############################################################################################################################
    'Purpose: Search the specified Column for a Cell that matches the search text
    
    'SearchText     - The text to look for
    'SearchColumn   - The Column to search in (defaults to the SearchColumn property if not specified)
    'SearchMode     - The type of search required. The lgSMNavigate mode is used by the Grid internally
    '               when searching for an entry that matches the keys the user is pressing.
    
    'MatchCase      - Specify a case sensitive or case insensitive search
    
    Dim lCount As Long
    Dim sCellText As String
    
    FindItem = NULL_RESULT
    
    If SearchColumn = -1 Then
        SearchColumn = mSearchColumn
    End If
    
    If (SearchColumn >= 0) And (Len(SearchText) > 0) Then
        If Not MatchCase Then
            SearchText = UCase$(SearchText)
        End If
        
        For lCount = LBound(mItems) To mItemCount
            If MatchCase Then
                sCellText = mItems(mIX(lCount)).Cell(SearchColumn).sValue
            Else
                sCellText = UCase$(mItems(mIX(lCount)).Cell(SearchColumn).sValue)
            End If
            
            Select Case SearchMode
                Case lgSMEqual
                    If sCellText = SearchText Then
                        FindItem = lCount
                        Exit For
                    End If
                
                Case lgSMGreaterEqual
                    If sCellText >= SearchText Then
                        FindItem = lCount
                        Exit For
                    End If
                
                Case lgSMLike
                    If sCellText Like SearchText & "*" Then
                        FindItem = lCount
                        Exit For
                    End If
                    
                Case lgWith
                
                    If InStr(1, sCellText, SearchText) > 0 Then
                        FindItem = lCount
                        Exit For
                    End If
                
                    
                
                Case lgSMNavigate
                    If Len(sCellText) > 0 Then
                        If (sCellText >= SearchText) And ((Mid$(sCellText, 1, 1)) = Mid$(SearchText, 1, 1)) Then
                            FindItem = lCount
                            Exit For
                        End If
                    End If
    
            End Select
            
        Next lCount
    End If
End Function

Public Property Let FocusRectColor(ByVal NewValue As OLE_COLOR)
    mFocusRectColor = NewValue
    
    PropertyChanged "FocusRectColor"
End Property

Public Property Get FocusRectColor() As OLE_COLOR
    FocusRectColor = mFocusRectColor
End Property

Public Property Get FocusRectMode() As lgFocusRectModeEnum
    FocusRectMode = mFocusRectMode
End Property

Public Property Let FocusRectMode(ByVal NewValue As lgFocusRectModeEnum)
    mFocusRectMode = NewValue
    DisplayChange
    
    PropertyChanged "FocusRectMode"
End Property

Public Property Get FocusRectStyle() As lgFocusRectStyleEnum
    FocusRectStyle = mFocusRectStyle
End Property

Public Property Let FocusRectStyle(ByVal NewValue As lgFocusRectStyleEnum)
    mFocusRectStyle = NewValue
    DisplayChange
    
    PropertyChanged "FocusRectStyle"
End Property

Public Property Get Font() As Font
   Set Font = mFont
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Set mFont = NewValue
    Set UserControl.Font = mFont
    
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    mForeColor = NewValue
    
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColorEdit() As OLE_COLOR
    ForeColorEdit = mForeColorEdit
End Property

Public Property Let ForeColorEdit(ByVal lNewValue As OLE_COLOR)
    mForeColorEdit = lNewValue
    
    PropertyChanged "ForeColorEdit"
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
    ForeColorFixed = mForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal lNewValue As OLE_COLOR)
    mForeColorFixed = lNewValue
    
    PropertyChanged "ForeColorFixed"
End Property

Public Property Get ForeColorSel() As OLE_COLOR
    ForeColorSel = mForeColorSel
End Property

Public Property Let ForeColorSel(ByVal lNewValue As OLE_COLOR)
    mForeColorSel = lNewValue
    DisplayChange
    
    PropertyChanged "ForeColorSel"
End Property

Public Property Get ForeColorTotals() As OLE_COLOR
    ForeColorTotals = mForeColorTotals
End Property

Public Property Let ForeColorTotals(ByVal NewValue As OLE_COLOR)
    mForeColorTotals = NewValue
    DisplayChange
    
    PropertyChanged "ForeColorTotals"
End Property

Public Property Get FormatString() As String
    FormatString = mFormatString
End Property

Public Property Let FormatString(ByVal NewValue As String)
    '#############################################################################################################################
    'Purpose: Used to create multiple Columns with one string
    
    'Each Column is seperated by a "|" char. The Alignment can be specified by
    'using "^" for Centre, "<" for right an ">" for left (default)
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim sCols() As String
    
    mFormatString = NewValue
    
    If Len(mFormatString) > 0 Then
        sCols() = Split(NewValue, "|")
        If UBound(sCols()) > UBound(mCols) Then
            Cols = UBound(sCols()) + 1
        End If
        
        For lCol = LBound(sCols) To UBound(sCols)
            Select Case Mid$(sCols(lCol), 1, 1)
                Case "^"
                    mCols(lCol).sCaption = Mid$(sCols(lCol), 2)
                    mCols(lCol).nAlignment = lgAlignCenterCenter
                Case "<"
                    mCols(lCol).sCaption = Mid$(sCols(lCol), 2)
                    mCols(lCol).nAlignment = lgAlignLeftCenter
                Case ">"
                    mCols(lCol).sCaption = Mid$(sCols(lCol), 2)
                    mCols(lCol).nAlignment = lgAlignRightCenter
                Case Else
                    mCols(lCol).sCaption = sCols(lCol)
            End Select
            
            mCols(lCol).dCustomWidth = 1000
            mCols(lCol).lWidth = ScaleX(mCols(lCol).dCustomWidth, mScaleUnits, vbPixels)
            mCols(lCol).bVisible = True
        Next lCol
    Else
        ReDim mCols(0)
        Clear
    End If
    
    DisplayChange
    
    PropertyChanged "FormatString"
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = mFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal NewValue As Boolean)
    mFullRowSelect = NewValue
    DisplayChange
    
    PropertyChanged "FullRowSelect"
End Property

Private Function GetColFromX(X As Single) As Long
    '#############################################################################################################################
    'Purpose: Return Column from mouse position
    '#############################################################################################################################
    
    Dim lX As Long
    Dim lCol As Long
    
    GetColFromX = -1
    
    For lCol = SBValue(efsHorizontal) To UBound(mCols)
        With mCols(lCol)
            If .bVisible Then
                If (X > lX) And (X <= lX + .lWidth) Then
                    GetColFromX = lCol
                    Exit For
                End If
                
                lX = lX + .lWidth
            End If
        End With
    Next lCol
End Function

Private Function GetColumnHeadingHeight() As Long
    '#############################################################################################################################
    'Purpose: Return Height of Header Row
    '#############################################################################################################################
    
    Dim lHeight As Long
    
    With UserControl
        lHeight = .TextHeight("A") + 4
        If GetRowHeight() > lHeight Then
            GetColumnHeadingHeight = GetRowHeight()
        Else
            GetColumnHeadingHeight = lHeight
        End If
    End With
End Function

Private Function GetFlag(ByVal nFlags As Integer, nFlag As lgFlagsEnum) As Boolean
    '#############################################################################################################################
    'Purpose: Gets information by bit flags
    '#############################################################################################################################
    
    If nFlags And nFlag Then
        GetFlag = True
    End If
End Function

Private Function GetRowFromY(Y As Single) As Long
    '#############################################################################################################################
    'Purpose: Return Row from mouse position
    '#############################################################################################################################

    Dim lColumnHeadingHeight As Long
    Dim lRow As Long
    
    If mColumnHeaders Then
        lColumnHeadingHeight = GetColumnHeadingHeight()
        
        If Y > lColumnHeadingHeight Then
            lRow = ((Y - lColumnHeadingHeight) \ GetRowHeight()) + SBValue(efsVertical)
        Else
            lRow = -1
        End If
    Else
        lRow = (Y \ GetRowHeight() - 1) + SBValue(efsVertical)
    End If
    
    If lRow <= mItemCount Then
        GetRowFromY = lRow
    Else
        GetRowFromY = -1
    End If
End Function

Private Function GetRowHeight() As Long
    '#############################################################################################################################
    'Purpose: Return Row Height
    '#############################################################################################################################
    
    With UserControl
        If mRowHeightMin > 0 Then
            GetRowHeight = .ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
        Else
            GetRowHeight = ROW_HEIGHT
        End If
    End With
End Function

Public Property Get GridColor() As OLE_COLOR
    GridColor = mGridColor
End Property

Public Property Let GridColor(ByVal NewValue As OLE_COLOR)
    mGridColor = NewValue
    DrawGrid
        
    PropertyChanged "GridColor"
End Property

Public Property Get ProgressBarColor() As OLE_COLOR
    ProgressBarColor = mProgressBarColor
End Property

Public Property Let ProgressBarColor(ByVal NewValue As OLE_COLOR)
    mProgressBarColor = NewValue
    DrawGrid
        
    PropertyChanged "ProgressBarColor"
End Property

Public Property Get GridLines() As Boolean
    GridLines = mGridLines
End Property

Public Property Let GridLines(ByVal NewValue As Boolean)
    mGridLines = NewValue
    DisplayChange
    
    PropertyChanged "GridLines"
End Property

Public Property Let GridLineWidth(NewValue As Long)
    mGridLineWidth = NewValue
    DrawGrid
    
    PropertyChanged "GridLineWidth"
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = mGridLineWidth
End Property

Public Property Get HotHeaderTracking() As Boolean
    HotHeaderTracking = mHotHeaderTracking
End Property

Public Property Let HotHeaderTracking(ByVal NewValue As Boolean)
    mHotHeaderTracking = NewValue
    
    If Not NewValue Then
        DrawHeaderRow
    End If
    
    PropertyChanged "HotHeaderTracking"
End Property

Public Property Get ImageList() As Object
    Set ImageList = mImageList
End Property

Public Property Let ImageList(ByVal NewValue As Object)
    Set mImageList = NewValue
    
    DisplayChange
End Property

Private Function IsEditable() As Boolean
    If Not mLocked And mEditable Then
        IsEditable = (mItemCount >= 0)
    End If
End Function

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    Call FreeLibrary(hMod)
  End If
End Function

Public Property Let ItemBackColor(ByVal Index As Long, ByVal NewValue As Long)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellBackColor(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid
End Property

Public Property Get ItemChecked(ByVal Index As Long) As Boolean
    ItemChecked = mItems(mIX(Index)).nFlags And lgChecked
End Property

Public Property Let ItemChecked(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mItems(mIX(Index)).nFlags, lgChecked, NewValue
    DrawGrid
End Property

Public Property Get ItemCount() As Long
    ItemCount = mItemCount + 1
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
    ItemData = mItems(mIX(Index)).lItemData
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal NewValue As Long)
    mItems(mIX(Index)).lItemData = NewValue
End Property

Public Property Let ItemFontBold(ByVal Index As Long, ByVal NewValue As Boolean)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellFontBold(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid
End Property

Public Property Let ItemForeColor(ByVal Index As Long, ByVal NewValue As Long)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellForeColor(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid
End Property

Public Property Let ItemImage(ByVal Index As Long, NewValue As Variant)
    On Local Error GoTo ItemImageError
    
    If IsNumeric(NewValue) Then
        mItems(mIX(Index)).lImage = NewValue
    Else
        mItems(mIX(Index)).lImage = -mImageList.ListImages(NewValue).Index
    End If
    
    DrawGrid
    Exit Property
    
ItemImageError:
    mItems(mIX(Index)).lImage = 0
End Property

Public Property Get ItemImage(ByVal Index As Long) As Variant
    If mItems(mIX(Index)).lImage >= 0 Then
        ItemImage = mItems(mIX(Index)).lImage
    Else
        ItemImage = mImageList.ListImages(Abs(mItems(mIX(Index)).lImage)).Key
    End If
End Property

Public Property Get ItemSelected(ByVal Index As Long) As Boolean
    ItemSelected = mItems(mIX(Index)).nFlags And lgSelected
End Property

Public Property Let ItemSelected(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mItems(mIX(Index)).nFlags, lgSelected, NewValue
    DrawGrid
End Property

Public Property Get ItemTag(ByVal Index As Long) As String
    ItemTag = mItems(mIX(Index)).sTag
End Property

Public Property Let ItemTag(ByVal Index As Long, ByVal NewValue As String)
    mItems(mIX(Index)).sTag = NewValue
End Property

Public Function ItemsVisible() As Long
    Dim lBorderWidth As Long
    
    If mBorderStyle = lgSingle Then
        lBorderWidth = 2
    End If

    With UserControl
        ItemsVisible = (.ScaleHeight - GetColumnHeadingHeight() - (lBorderWidth * 2)) / GetRowHeight()
    End With
End Function

Public Property Get MouseCol() As Long
    MouseCol = mMouseCol
End Property

Public Property Get MouseRow() As Long
    MouseRow = mMouseRow
End Property

Private Sub MoveEditControl(ByVal MoveControl As lgMoveControlEnum)
    '#############################################################################################################################
    'Purpose: Used to position and optionally resize the Edit control.
    '#############################################################################################################################
   
    Dim R As RECT
    Dim lBorderWidth As Long
    Dim nScaleMode As ScaleModeConstants
    
    SetColRect mEditCol, R
    
    On Local Error Resume Next
    
    'Check if an external Control is used.
    If mCols(mEditCol).EditCtrl Is Nothing Then
        'Using internal TextBox
        With txtEdit
            .Left = R.Left + mGridLineWidth
            .Top = RowTop(mEditRow) + mGridLineWidth
            .Height = mItems(mIX(mEditRow)).lHeight - mGridLineWidth
            .Width = R.Right - mGridLineWidth
        End With
    Else
        nScaleMode = UserControl.Parent.ScaleMode
        If mBorderStyle = lgSingle Then
            lBorderWidth = 2
        End If
                    
        With mCols(mEditCol).EditCtrl
            If mCols(mEditCol).MoveControl And lgBCLeft Then
                .Left = ScaleX(R.Left + mGridLineWidth + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Left
            End If
            If mCols(mEditCol).MoveControl And lgBCTop Then
                .Top = ScaleY(RowTop(mEditRow) + mGridLineWidth + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Top
            End If
            If mCols(mEditCol).MoveControl And lgBCHeight Then
                .Height = ScaleY(mItems(mIX(mEditRow)).lHeight - mGridLineWidth, vbPixels, nScaleMode)
            End If
            If mCols(mEditCol).MoveControl And lgBCWidth Then
                .Width = ScaleX(R.Right - mGridLineWidth, vbPixels, nScaleMode)
            End If
        End With
    End If
    
    On Local Error GoTo 0
End Sub

Public Property Get MultiSelect() As Boolean
    MultiSelect = mMultiSelect
End Property

Public Property Let MultiSelect(ByVal NewValue As Boolean)
    mMultiSelect = NewValue
    
    If Not NewValue Then
        SetSelection False
        DisplayChange
    End If
    
    PropertyChanged "MultiSelect"
End Property

Private Function NavigateDown() As Long
    If mRow < mItemCount Then
        NavigateDown = mRow + 1
    Else
        NavigateDown = mRow
    End If
End Function

Private Function NavigateLeft() As Long
    If mCol > 0 Then
        NavigateLeft = mCol - 1
    Else
        NavigateLeft = mCol
    End If
End Function

Private Function NavigateRight() As Long
    If mCol < UBound(mCols) Then
        NavigateRight = mCol + 1
    Else
        NavigateRight = mCol
    End If
End Function

Private Function NavigateUp() As Long
    If mRow > 0 Then
        NavigateUp = mRow - 1
    Else
        NavigateUp = mRow
    End If
End Function

Private Property Get Orientation() As ScrollBarOrienationEnum
    SBOrientation = m_eOrientation
End Property

Private Sub pSBClearUp()
    If m_hWnd <> 0 Then
        On Error Resume Next
        ' Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
            UninitializeFlatSB m_hWnd
        End If

        On Error GoTo 0
    End If
    m_hWnd = 0
    m_bInitialised = False
End Sub

Private Sub pSBCreateScrollBar()
    Dim lr As Long
    Dim hParent As Long

    On Error Resume Next
    lr = InitialiseFlatSB(m_hWnd)
    If (Err.Number <> 0) Then
        'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
        ' Means we have version prior to 4.71
        ' We get standard scroll bars.
        m_bNoFlatScrollBars = True
    Else
        SBStyle = m_eStyle
    End If
End Sub

Private Sub pSBGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, Lo, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, Lo, tSI
    End If

End Sub

Private Sub pSBLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, Lo, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, Lo, tSI, True
    End If

End Sub

Private Sub pSBSetOrientation()
    ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Horizontal))
    ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Vertical))
End Sub

Public Property Get Redraw() As Boolean
    Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal NewValue As Boolean)
    mRedraw = NewValue
    
    If mRedraw Then
        If mPendingScrollBar Then
            SetScrollBars
        End If
        If mPendingRedraw Then
            CreateRenderData
            DrawGrid
        End If
    Else
        mPendingScrollBar = False
        mPendingRedraw = False
    End If
    
    PropertyChanged "Redraw"
End Property

Public Sub Refresh()
    CreateRenderData
    SetScrollBars
    DrawGrid True
End Sub

Public Sub RemoveItem(ByVal Index As Long)
    Dim lCount As Long
    Dim lPosition As Long
    Dim bSelected As Boolean
   
    '#############################################################################################################################
    'See AddItem for details of the Arrays used
    '#############################################################################################################################
    
    'Note selected state before deletion
    bSelected = mItems(mIX(Index)).nFlags And lgSelected
    
    'Decrement the reference count on each cells format Entry
    If mItemCount >= 0 Then
        For lCount = 0 To UBound(mCols)
            If mItems(Index).Cell(Count).nFormat >= 0 Then
                mCF(mItems(Index).Cell(lCount).nFormat).nCount = mCF(mItems(Index).Cell(lCount).nFormat).nCount - 1
            End If
        Next lCount
    End If
    
    lPosition = mIX(Index)
    
    'Reset Item Data
    For lCount = mIX(Index) To mItemCount - 1
        mItems(lCount) = mItems(lCount + 1)
    Next lCount
    
    'Adjust Index
    For lCount = Index To mItemCount - 1
        mIX(lCount) = mIX(lCount + 1)
    Next lCount
    
    'Validate Indexes for Items after deleted Item
    For lCount = 0 To mItemCount - 1
        If mIX(lCount) > lPosition Then
            mIX(lCount) = mIX(lCount) - 1
        End If
    Next lCount
    
    mItemCount = mItemCount - 1
     
    If mItemCount < 0 Then
        Clear
    Else
        If (mItemCount + mCacheIncrement) < UBound(mItems) Then
            ReDim Preserve mItems(mItemCount)
            ReDim Preserve mIX(mItemCount)
        End If
 
        If bSelected Then
            If mMultiSelect Then
                RaiseEvent SelectionChanged
            ElseIf Index > mItemCount Then
                SetFlag mItems(mIX(mItemCount)).nFlags, lgSelected, True
            ElseIf mItemCount >= 0 Then
                SetFlag mItems(mIX(Index)).nFlags, lgSelected, True
            End If
        End If
        
        If Index > mItemCount Then
            SetRowCol mRow - 1, mCol
        End If
    End If
    
    DisplayChange
    
    If mRow < 0 Then
        If mItemCount >= 0 Then
            Row = 0
        End If
    End If
    RaiseEvent ItemCountChanged
End Sub

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal NewValue As Long)
    If SetRowCol(NewValue, mCol) Then
        DrawGrid
    End If
End Property

Public Property Get RowHeightMin() As Long
    RowHeightMin = mRowHeightMin
End Property

Public Property Let RowHeightMin(ByVal NewValue As Long)
    mRowHeightMin = NewValue
    DisplayChange
       
    PropertyChanged "RowHeightMin"
End Property

Public Function RowTop(Index As Long) As Long
    Dim lRow As Long
    Dim lY As Long
    
    lY = GetColumnHeadingHeight()
    For lRow = SBValue(efsVertical) To Index - 1
        lY = lY + mItems(mIX(lRow)).lHeight
    Next lRow
   
    RowTop = lY
End Function

Private Property Get SBCanBeFlat() As Boolean
    SBCanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub SBCreate(ByVal hWndA As Long)
    pSBClearUp
    m_hWnd = hWndA
    pSBCreateScrollBar
End Sub

Private Property Get SBEnabled(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBEnabled = m_bEnabledHorz
    Else
        SBEnabled = m_bEnabledVert
    End If
End Property

Private Property Let SBEnabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
    Dim Lo As Long
    Dim lF As Long

    Lo = eBar
    If (bEnabled) Then
        lF = ESB_ENABLE_BOTH
    Else
        lF = ESB_DISABLE_BOTH
    End If
    If (m_bNoFlatScrollBars) Then
        EnableScrollBar m_hWnd, Lo, lF
    Else
        FlatSB_EnableScrollBar m_hWnd, Lo, lF
    End If

End Property

Private Property Get SBLargeChange(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_PAGE
    SBLargeChange = tSI.nPage
End Property

Private Property Let SBLargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
    Dim tSI As SCROLLINFO

    pSBGetSI eBar, tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pSBLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property

Private Property Get SBMax(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    SBMax = tSI.nMax                                  ' - tSI.nPage
End Property

Private Property Let SBMax(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
    Dim tSI As SCROLLINFO
    tSI.nMax = iMax + SBLargeChange(eBar)
    tSI.nMin = SBMin(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Get SBMin(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE
    SBMin = tSI.nMin
End Property

Private Property Let SBMin(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
    Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = SBMax(eBar) + SBLargeChange(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Let SBOrientation(ByVal eOrientation As ScrollBarOrienationEnum)
    m_eOrientation = eOrientation
    pSBSetOrientation
End Property

Private Sub SBRefresh()
    EnableScrollBar m_hWnd, SB_VERT, ESB_ENABLE_BOTH
End Sub

Private Property Get SBSmallChange(ByVal eBar As EFSScrollBarConstants) As Long
    If (eBar = efsHorizontal) Then
        SBSmallChange = m_lSmallChangeHorz
    Else
        SBSmallChange = m_lSmallChangeVert
    End If
End Property

Private Property Let SBSmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
    If (eBar = efsHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Private Property Get SBStyle() As ScrollBarStyleEnum
    SBStyle = m_eStyle
End Property

Private Property Let SBStyle(ByVal eStyle As ScrollBarStyleEnum)
    Dim lr As Long
    If (m_bNoFlatScrollBars) Then
        ' can't do it..
        'Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
        Exit Property
    Else
        If (m_eOrientation = Scroll_Horizontal) Or (m_eOrientation = Scroll_Both) Then
            lr = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
        End If
        If (m_eOrientation = Scroll_Vertical) Or (m_eOrientation = Scroll_Both) Then
            lr = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
        End If
        'Debug.Print lR
        m_eStyle = eStyle
    End If

End Property

Private Property Get SBValue(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_POS
    SBValue = tSI.nPos
End Property

Private Property Let SBValue(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
    Dim tSI As SCROLLINFO
    If (iValue <> SBValue(eBar)) Then
        tSI.nPos = iValue
        pSBLetSI eBar, tSI, SIF_POS
        
    End If
End Property

Private Property Get SBVisible(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBVisible = m_bVisibleHorz
    Else
        SBVisible = m_bVisibleVert
    End If
End Property

Private Property Let SBVisible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
    If (eBar = efsHorizontal) Then
        m_bVisibleHorz = bState
    Else
        m_bVisibleVert = bState
    End If
    If (m_bNoFlatScrollBars) Then
        ShowScrollBar m_hWnd, eBar, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
    End If
End Property

Public Property Get ScaleUnits() As ScaleModeConstants
    ScaleUnits = mScaleUnits
End Property

Public Property Let ScaleUnits(ByVal NewValue As ScaleModeConstants)
    mScaleUnits = NewValue
    
    PropertyChanged "ScaleUnits"
End Property

Private Sub ScrollList(nDirection As Integer)
    '#############################################################################################################################
    'Purpose: Used to automatically scroll the list up or down when the mouse
    'is dragged out of the Control
    '#############################################################################################################################

    Dim lCount As Long
    Dim lItemsVisible As Long

    mScrollAction = nDirection
      
    Do While mScrollAction = nDirection
        mScrollTick = GetTickCount()
        
        If nDirection = SCROLL_UP Then
            If SBValue(efsVertical) > SBMin(efsVertical) Then
                SBValue(efsVertical) = SBValue(efsVertical) - 1
                If mMultiSelect Then
                    SetFlag mItems(mIX(SBValue(efsVertical))).nFlags, lgSelected, True
                Else
                    mRow = SBValue(efsVertical)
                    SetSelection False
                    SetSelection True, mRow, mRow
                End If
                
                RaiseEvent RowColChanged
            Else
                Exit Do
            End If
        Else
            If SBValue(efsVertical) < SBMax(efsVertical) Then
                lItemsVisible = ItemsVisible()
                
                SBValue(efsVertical) = SBValue(efsVertical) + 1
                If mMultiSelect Then
                    For lCount = SBValue(efsVertical) To SBValue(efsVertical) + lItemsVisible
                        If lCount > mItemCount Then
                            Exit For
                        Else
                            SetFlag mItems(mIX(lCount)).nFlags, lgSelected, True
                        End If
                    Next lCount
                Else
                    mRow = SBValue(efsVertical) + (lItemsVisible - 1)
                    If mRow > mItemCount Then
                        mRow = mItemCount
                    End If
                    SetSelection False
                    SetSelection True, mRow, mRow
                End If
                
                RaiseEvent RowColChanged
            Else
                Exit Do
            End If
        End If
        
        RaiseEvent SelectionChanged
        DrawGrid
        RaiseEvent Scroll
        
        Sleep AUTOSCROLL_TIMEOUT
        DoEvents
    Loop
End Sub

Public Property Get ScrollTrack() As Boolean
    ScrollTrack = mScrollTrack
End Property

Public Property Let ScrollTrack(ByVal NewValue As Boolean)
    mScrollTrack = NewValue
    
    PropertyChanged "ScrollTrack"
End Property

Public Property Get SearchColumn() As Long
    SearchColumn = mSearchColumn
End Property

Public Property Let SearchColumn(ByVal NewValue As Long)
    mSearchColumn = NewValue
    
    PropertyChanged "SearchColumn"
End Property

Public Function SelectedCount() As Long
    '#############################################################################################################################
    'Purpose: Return Count of Selected Items
    '#############################################################################################################################
    
    Dim lCount As Long
    
    For lCount = LBound(mItems) To UBound(mItems)
        If mItems(lCount).nFlags And lgSelected Then
            SelectedCount = SelectedCount + 1
        End If
    Next lCount
End Function

Private Function SetColRect(ByVal Index As Long, R As RECT)
    '#############################################################################################################################
    'Purpose: Set the drawing boundary for a Column
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim lCount As Long
    Dim lX As Long

    If Index < SBValue(efsHorizontal) Then
        R.Left = -1
    Else
        For lCol = SBValue(efsHorizontal) To Index - 1
            If mCols(lCol).bVisible Then
                lX = lX + mCols(lCol).lWidth
                lCount = lCount + 1
            End If
        Next lCol
        
        If mCheckboxes And (lCount = 0) Then
            R.Left = RIGHT_CHECKBOX
            R.Right = mCols(Index).lWidth - RIGHT_CHECKBOX
        Else
            R.Left = lX
            R.Right = mCols(Index).lWidth
        End If
    End If
End Function

Private Sub SetCheckBoxRect(ByVal Row As Long, ByVal Col As Long, lY As Long, R As RECT)
    Dim lLeft As Long
    
    lLeft = (mCols(Col).lX + (mCols(Col).lWidth) / 2) - (mR.CheckBoxSize / 2)
    Call SetRect(R, lLeft, lY, lLeft + mR.CheckBoxSize, lY + mItems(mIX(Row)).lHeight)
End Sub
Private Sub SetFlag(nFlags As Integer, nFlag As lgFlagsEnum, bValue As Boolean)
    If bValue Then
        nFlags = (nFlags Or nFlag)
    Else
        nFlags = (nFlags And Not (nFlag))
    End If
End Sub

Private Sub SetRedrawState(bState As Boolean)
    '#############################################################################################################################
    'Purpose: Used to prevent Internal Redraws while preserving User Controlled Redraw state
    '
    'bDrawLocked used to prevent nested Calls to Lock Redraw
    '#############################################################################################################################
   
    Static bDrawLocked As Boolean
    Static bOriginalRedraw As Boolean
    
    If bState Then
        bDrawLocked = False
        mRedraw = bOriginalRedraw
    ElseIf Not bDrawLocked Then
        bDrawLocked = True
        bOriginalRedraw = mRedraw
        mRedraw = False
    End If
End Sub


Private Function SetRowCol(lRow As Long, lCol As Long, Optional bSetScroll As Boolean) As Boolean
    '#############################################################################################################################
    'Purpose: To update current Row/Col and fire Events if necessary
    '#############################################################################################################################
    
    Dim R As RECT
    Dim lCount As Long
    
    
    'added By: Vincent j.Jamero
    If m_IgnoreEmpty = True Then
        If Len(Trim(CellText(lRow, 0))) < 1 Then
            Exit Function
        End If
    End If
    
    
    If (mCol <> lCol) Or (mRow <> lRow) Then
        mCol = lCol
        mRow = lRow
        
        RaiseEvent RowColChanged
        
        'Do we need to change Bars?
        If bSetScroll Then
            SetColRect mCol, R
            
            'Scroll to make Column visible
            If R.Left < 0 Then
                 For lCount = SBValue(efsHorizontal) To SBMin(efsHorizontal) Step -1
                    If R.Left > 0 Then
                        Exit For
                    End If
                    
                    SBValue(efsHorizontal) = SBValue(efsHorizontal) - 1
                    SetColRect mCol, R
                Next lCount
            Else
                For lCount = SBValue(efsHorizontal) To SBMax(efsHorizontal)
                    If R.Left + mCols(mCol).lWidth < UserControl.ScaleWidth Then
                        Exit For
                    End If
                    
                    SBValue(efsHorizontal) = SBValue(efsHorizontal) + 1
                    SetColRect mCol, R
                Next lCount
            End If
            
            If SBValue(efsHorizontal) = SBMin(efsHorizontal) Then
                SetScrollBars
            End If
            
            If mRow < SBValue(efsVertical) Then
                SBValue(efsVertical) = SBValue(efsVertical) - 1
            ElseIf mRow > SBValue(efsVertical) + (ItemsVisible() - 1) Then
                SBValue(efsVertical) = SBValue(efsVertical) + 1
            End If
            
            RaiseEvent Scroll
        End If
        
        SetRowCol = True
    End If
End Function

Private Sub SetScrollBars()
    '#############################################################################################################################
    'Purpose: Sets the visibilty of scroll bars and sets max scroll values
    '#############################################################################################################################

    Dim lCol As Long
    Dim lWidth As Long
    Dim bHVisible As Boolean
    Dim bVVisible As Boolean
    
    If m_hWnd <> 0 Then
        LockWindowUpdate UserControl.hwnd
        
        'Calculate total width of columns
        For lCol = LBound(mCols) To UBound(mCols)
            If mCols(lCol).bVisible Then
                lWidth = lWidth + mCols(lCol).lWidth
            End If
        Next lCol
        
        If (lWidth > UserControl.ScaleWidth) Then
            SBMax(efsHorizontal) = UBound(mCols) - 1
            bHVisible = True
        Else
            SBMax(efsHorizontal) = UBound(mCols)
            bHVisible = (SBValue(efsHorizontal) > SBMin(efsHorizontal))
        End If
        
        If ItemCount() > ItemsVisible() Then
            SBMax(efsVertical) = mItemCount - ItemsVisible()
            bVVisible = True
        Else
            SBMax(efsVertical) = mItemCount
        End If
        
        'If SBVisible(efsHorizontal) <> bHVisible Then
            SBVisible(efsHorizontal) = bHVisible
        'End If
        'If SBVisible(efsVertical) <> bVVisible Then
            SBVisible(efsVertical) = bVVisible
        'End If
        
        LockWindowUpdate 0
    End If
End Sub

Private Function SetSelection(bState As Boolean, Optional lFromRow As Long = -1, Optional lToRow As Long = -1) As Boolean
    Dim lCount As Long
    Dim lStep As Long
    Dim bSelectionChanged As Boolean
    
    If lFromRow = -1 Then
        lFromRow = LBound(mItems)
    End If
    
    If lToRow = -1 Then
        lToRow = UBound(mItems)
    End If
    
    If lFromRow >= lToRow Then
        lStep = -1
    Else
        lStep = 1
    End If
    
    For lCount = lFromRow To lToRow Step lStep
        If (mItems(mIX(lCount)).nFlags And lgSelected) <> bState Then
            SetFlag mItems(mIX(lCount)).nFlags, lgSelected, bState
            bSelectionChanged = True
        End If
    Next lCount
    
    SetSelection = bSelectionChanged
End Function

Private Sub SortArrayString(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mIX(lFirst), mIX((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = mItems(mIX(lIndex)).Cell(lSortColumn).sValue > mItems(mIX(lFirst)).Cell(lSortColumn).sValue
        Else
            bSwap = mItems(mIX(lIndex)).Cell(lSortColumn).sValue < mItems(mIX(lFirst)).Cell(lSortColumn).sValue
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mIX(lBoundary), mIX(lIndex)
        End If
    Next lIndex

    SwapLng mIX(lFirst), mIX(lBoundary)
    SortArrayString lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayString lBoundary + 1, lLast, lSortColumn, nSortType
End Sub

Private Sub SortArrayDate(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mIX(lFirst), mIX((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = CDate(mItems(mIX(lIndex)).Cell(lSortColumn).sValue) > CDate(mItems(mIX(lFirst)).Cell(lSortColumn).sValue)
        Else
            bSwap = CDate(mItems(mIX(lIndex)).Cell(lSortColumn).sValue) < CDate(mItems(mIX(lFirst)).Cell(lSortColumn).sValue)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mIX(lBoundary), mIX(lIndex)
        End If
    Next lIndex

    SwapLng mIX(lFirst), mIX(lBoundary)
    SortArrayDate lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayDate lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayNumeric(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mIX(lFirst), mIX((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = Val(mItems(mIX(lIndex)).Cell(lSortColumn).sValue) > Val(mItems(mIX(lFirst)).Cell(lSortColumn).sValue)
        Else
            bSwap = Val(mItems(mIX(lIndex)).Cell(lSortColumn).sValue) < Val(mItems(mIX(lFirst)).Cell(lSortColumn).sValue)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mIX(lBoundary), mIX(lIndex)
        End If
    Next lIndex

    SwapLng mIX(lFirst), mIX(lBoundary)
    SortArrayNumeric lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayNumeric lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayCustom(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mIX(lFirst), mIX((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            RaiseEvent CustomSort(True, lSortColumn, mItems(mIX(lIndex)).Cell(lSortColumn).sValue, mItems(mIX(lFirst)).Cell(lSortColumn).sValue, bSwap)
        Else
            RaiseEvent CustomSort(False, lSortColumn, mItems(mIX(lIndex)).Cell(lSortColumn).sValue, mItems(mIX(lFirst)).Cell(lSortColumn).sValue, bSwap)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mIX(lBoundary), mIX(lIndex)
        End If
    Next lIndex

    SwapLng mIX(lFirst), mIX(lBoundary)
    SortArrayCustom lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayCustom lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayBool(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mIX(lFirst), mIX((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = GetFlag(mItems(mIX(lIndex)).Cell(lSortColumn).nFlags, lgChecked) > GetFlag(mItems(mIX(lFirst)).Cell(lSortColumn).nFlags, lgChecked)
        Else
            bSwap = GetFlag(mItems(mIX(lIndex)).Cell(lSortColumn).nFlags, lgChecked) < GetFlag(mItems(mIX(lFirst)).Cell(lSortColumn).nFlags, lgChecked)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mIX(lBoundary), mIX(lIndex)
        End If
    Next lIndex

    SwapLng mIX(lFirst), mIX(lBoundary)
    SortArrayBool lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayBool lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArray(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################
    
    Select Case mCols(lSortColumn).nType
        Case lgBoolean
            SortArrayBool lFirst, lLast, lSortColumn, nSortType
        Case lgDate
            SortArrayDate lFirst, lLast, lSortColumn, nSortType
        Case lgNumeric
            SortArrayNumeric lFirst, lLast, lSortColumn, nSortType
        Case lgCustom
            SortArrayCustom lFirst, lLast, lSortColumn, nSortType
        Case Else
            SortArrayString lFirst, lLast, lSortColumn, nSortType
    End Select
End Sub
Private Sub SortSubList()
    '#############################################################################################################################
    'Purpose: Used to sort by a secondary Column after a Sort
    '#############################################################################################################################
    
    Dim lCount As Long
    Dim lStartSort As Long
    Dim bDifferent As Boolean
    Dim sMajorSort As String

    If mSortSubColumn > NULL_RESULT Then
        'Re-Sort the Items by a secondary column, preserving the sort sequence of the
        'primary sort
        
        lStartSort = LBound(mItems)
        For lCount = LBound(mItems) To mItemCount
            bDifferent = mItems(mIX(lCount)).Cell(mSortColumn).sValue <> sMajorSort
            If bDifferent Or lCount = mItemCount Then
                If lCount > 1 Then
                    If lCount - lStartSort > 1 Then
                        If lCount = mItemCount And Not bDifferent Then
                            SortArray lStartSort, lCount, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        Else
                            SortArray lStartSort, lCount - 1, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        End If
                    End If
                    lStartSort = lCount
                End If
                
                sMajorSort = mItems(mIX(lCount)).Cell(mSortColumn).sValue
            End If
        Next lCount
    End If
End Sub

'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    i = i - 1                                                                           'Next element
  Loop
Errs:
End Sub

Private Sub SwapLng(Value1 As Long, Value2 As Long)
    Static lTemp As Long

    lTemp = Value1
    Value1 = Value2
    Value2 = lTemp
End Sub

Public Function ToggleEdit() As Boolean
    '#############################################################################################################################
    'Purpose: Used to start a new Edit or commit a pending one
    '#############################################################################################################################
    
    If IsEditable() Then
        ToggleEdit = True
        
        If mEditPending Then
            UpdateCell
        ElseIf (mRow <> NULL_RESULT) And (mCol <> NULL_RESULT) Then
            EditCell mRow, mCol
        End If
    End If
End Function

Public Property Let TopRow(ByVal NewValue As Long)
    If NewValue > SBMax(efsVertical) Then
        SBValue(efsVertical) = SBMax(efsVertical)
    Else
        SBValue(efsVertical) = NewValue
    End If
    
    SetRowCol NewValue, mCol, True
    DrawGrid
End Property

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
        
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

Public Property Let ThemeColor(NewValue As lgThemeColorEnum)
    mThemeColor = NewValue
    SetColors
    DrawGrid True
    
    PropertyChanged "ThemeColor"
End Property

Public Property Get ThemeColor() As lgThemeColorEnum
    ThemeColor = mThemeColor
End Property

Private Sub SetColors()
    Select Case mThemeColor
        Case lgTCDefault
            mBackColor = DEF_BACKCOLOR
            mForeColor = DEF_FORECOLOR
            mBackColorSel = DEF_BACKCOLORSEL
            mForeColorSel = DEF_FORECOLORSEL
            
            mFocusRectColor = DEF_FOCUSRECTCOLOR
            mGridColor = DEF_GRIDCOLOR
            
        Case lgTCBlue
            mBackColor = DEF_BACKCOLOR
            mForeColor = DEF_FORECOLOR
            mBackColorSel = &HF1D8C9
            mForeColorSel = &H9C613B
            
            mFocusRectColor = &H9C613B
            mGridColor = &HEBEBEB
            
        Case lgTCGreen
            mBackColor = DEF_BACKCOLOR
            mForeColor = DEF_FORECOLOR
            mBackColorSel = &H8FC5B5
            mForeColorSel = &HE1F9F7
            
            mFocusRectColor = &H385D3F
            mGridColor = &HC0FFC0
           
    End Select
End Sub

Public Property Let ThemeStyle(NewValue As lgThemeStyleEnum)
    mThemeStyle = NewValue
    DrawGrid True
    
    PropertyChanged "ThemeStyle"
End Property

Public Property Get ThemeStyle() As lgThemeStyleEnum
    ThemeStyle = mThemeStyle
End Property


Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional hPalette As Long = 0) As Long
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Function UpdateCell() As Boolean
    '#############################################################################################################################
    'Purpose: Used to commit an Edit. Note the RequestUpate event. This event allows
    'the Upate to be cancelled by setting the Cancel flag.
    '#############################################################################################################################
   
    Dim bCancel As Boolean
    Dim bRequestUpdate As Boolean
    Dim sNewValue As String
    
    If mEditPending Then
        If mCols(mEditCol).EditCtrl Is Nothing Then
            bRequestUpdate = (mItems(mIX(mEditRow)).Cell(mEditCol).sValue <> txtEdit.Text)
            sNewValue = txtEdit.Text
        Else
            bRequestUpdate = True
        End If
        
        If bRequestUpdate Then
            RaiseEvent RequestUpdate(mEditRow, mEditCol, sNewValue, bCancel)
        End If
        
        If Not bCancel Then
            SetRedrawState False
        
            If mCols(mEditCol).EditCtrl Is Nothing Then
                txtEdit.Visible = False
            Else
                On Local Error Resume Next
                
                With mCols(mEditCol).EditCtrl
                    If mEditParent <> 0 Then
                        SetParent .hwnd, mEditParent
                    End If
                    
                    Subclass_Stop .hwnd
                    
                    .Visible = False
                End With
                
                On Local Error GoTo 0
            End If
            
            mEditPending = False
            
            mItems(mIX(mEditRow)).Cell(mEditCol).sValue = sNewValue
            DisplayChange
            
            SetRedrawState True
        End If
    End If
    
    UpdateCell = Not bCancel
End Function

Private Sub txtEdit_Validate(Cancel As Boolean)
    If Not Cancel Then
        UpdateCell
    End If
End Sub


Private Sub UserControl_Click()
    If (mEditTrigger And lgMouseClick) Then
        ToggleEdit
    End If
    
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If (mEditTrigger And lgMouseDblClick) And (mMouseRow > NULL_RESULT) Then
        ToggleEdit
    End If
    
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Dim OS As OSVERSIONINFO
    
    mClipRgn = CreateRectRgn(0, 0, 0, 0)
    
    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    
    mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    
    If (OS.dwMajorVersion > 5) Then
        mWindowsXP = True
    ElseIf (OS.dwMajorVersion = 5) And (OS.dwMinorVersion >= 1) Then
        mWindowsXP = True
    End If
    
    Set txtEdit = UserControl.Controls.Add("VB.TextBox", "txtEdit")
    With txtEdit
        .BorderStyle = 0
        .Visible = False
    End With

    ReDim mCols(0)
    Clear
End Sub

Private Sub UserControl_InitProperties()
    Set mFont = Ambient.Font

    '################################################################################
    'Appearance Properties
    mBackColor = DEF_BACKCOLOR
    mBackColorBkg = DEF_BACKCOLORBKG
    mBackColorEdit = DEF_BACKCOLOREDIT
    mBackColorFixed = DEF_BACKCOLORFIXED
    mBackColorSel = DEF_BACKCOLORSEL
    mForeColor = DEF_FORECOLOR
    mForeColorEdit = DEF_FORECOLOREDIT
    mForeColorFixed = DEF_FORECOLORFIXED
    mForeColorSel = DEF_FORECOLORSEL
    mForeColorTotals = DEF_FORECOLORTOTALS
    
    mFocusRectColor = DEF_FOCUSRECTCOLOR
    mGridColor = DEF_GRIDCOLOR
    mProgressBarColor = DEF_PROGRESSBARCOLOR
    
    mDisplayEllipsis = DEF_DISPLAYELLIPSIS
    mFocusRectMode = DEF_FOCUSRECTMODE
    mFocusRectStyle = DEF_FOCUSRECTSTYLE
    mGridLines = DEF_GRIDLINES
    mGridLineWidth = DEF_GRIDLINEWIDTH
    mThemeColor = DEF_THEMECOLOR
    mThemeStyle = DEF_THEMESTYLE
    
    '################################################################################
    'Behaviour Properties
    mAllowUserResizing = DEF_ALLOWUSERRESIZING
    mBorderStyle = DEF_BORDERSTYLE
    mCheckboxes = DEF_CHECKBOXES
    mColumnHeaders = DEF_COLUMNHEADERS
    mColumnSort = DEF_COLUMNSORT
    mEditable = DEF_EDITABLE
    mEditTrigger = DEF_EDITTRIGGER
    mFullRowSelect = DEF_FULLROWSELECT
    mHotHeaderTracking = DEF_HOTHEADERTRACKING
    mMultiSelect = DEF_MULTISELECT
    mRedraw = DEF_REDRAW
    mScrollTrack = DEF_SCROLLTRACK
    
    '################################################################################
    'Miscellaneous Properties
    mCacheIncrement = DEF_CACHEINCREMENT
    mEnabled = DEF_ENABLED
    mFormatString = DEF_FORMATSTRING
    mLocked = DEF_LOCKED
    mRowHeightMin = DEF_ROWHEIGHTMIN
    mScaleUnits = DEF_SCALEUNITS
    mSearchColumn = DEF_SEARCHCOLUMN
    
    '################################################################################
    'Apply Settings
    With UserControl
        .BackColor = mBackColorBkg
        .BorderStyle = mBorderStyle
    End With
    
    CreateRenderData
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lNewCol As Long
    Dim lNewRow As Long
    Dim bClearSelection As Boolean
    Dim bRedraw As Boolean

    lNewCol = mCol
    lNewRow = mRow
    
    SetRedrawState False
    
    'Used to determine if selected Items need to be cleared
    bClearSelection = True

    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape 'Allow escape to abort editing
           bClearSelection = False

           If (mEditTrigger And lgEnterKey) Then
               If KeyCode = vbKeyEscape Then
                   txtEdit.Visible = False
                   mEditPending = False
               Else
                   If ToggleEdit() Then KeyCode = 0
               End If
           End If
            
        Case vbKeyF2
            bClearSelection = False
            
            If (mEditTrigger And lgF2Key) Then
                If ToggleEdit() Then
                    KeyCode = 0
                End If
            End If
            
        Case vbKeySpace
            bClearSelection = False
        
            If mCheckboxes Then
                mIgnoreKeyPress = True
                
                SetFlag mItems(mIX(mRow)).nFlags, lgChecked, Not GetFlag(mItems(mIX(mRow)).nFlags, lgChecked)
                RaiseEvent ItemChecked(mRow)
                
                KeyCode = 0
            End If
            
        Case vbKeyA
            bClearSelection = False
            
            If (Shift And vbCtrlMask) And mMultiSelect Then
                mIgnoreKeyPress = True
                
                SetSelection True
                RaiseEvent SelectionChanged
                KeyCode = 0
            End If
    
        Case vbKeyUp
            If (Shift And vbShiftMask) And mMultiSelect Then
                bClearSelection = False
            End If
            
            If UpdateCell() Then
                lNewRow = NavigateUp()
                
                KeyCode = 0
            End If
            
        Case vbKeyDown
            If (Shift And vbShiftMask) And mMultiSelect Then
                bClearSelection = False
            End If
        
            If UpdateCell() Then
                lNewRow = NavigateDown()
                
                KeyCode = 0
            End If
            
        Case vbKeyLeft
            If Not mEditPending Then
                lNewCol = NavigateLeft()
                KeyCode = 0
            End If
            
        Case vbKeyRight
            If Not mEditPending Then
                lNewCol = NavigateRight()
                KeyCode = 0
            End If
            
        Case vbKeyPageUp
            If UpdateCell() Then
                If mRow > 0 Then
                    lNewRow = (mRow - ItemsVisible()) + 1
                    If lNewRow < 0 Then
                        lNewRow = 0
                    End If
                    
                    SBValue(efsVertical) = lNewRow
                End If
                
                KeyCode = 0
            End If
        
        Case vbKeyPageDown
            If UpdateCell() Then
                If mRow < mItemCount Then
                    lNewRow = (mRow + ItemsVisible()) - 1
                    If lNewRow > mItemCount Then
                        lNewRow = mItemCount
                    End If
         
                    SBValue(efsVertical) = lNewRow
                End If
                
                KeyCode = 0
            End If
        
        Case vbKeyHome
            If Shift And vbShiftMask Then
                If UpdateCell() Then
                    If mMultiSelect Then
                        bClearSelection = False
          
                        SetSelection False
                        SetSelection True, 1, mRow
                        RaiseEvent SelectionChanged
                    End If
                    
                    lNewRow = 0
                    
                    SBValue(efsVertical) = SBMin(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Shift And vbCtrlMask Then
                If UpdateCell() Then
                    lNewRow = 0
                    
                    SBValue(efsVertical) = SBMin(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Not mEditPending Then
                lNewCol = 0
                
                SBValue(efsHorizontal) = SBMin(efsHorizontal)
                KeyCode = 0
            End If
            
        Case vbKeyEnd
            If Shift And vbShiftMask Then
                If UpdateCell() Then
                    If mMultiSelect Then
                        bClearSelection = False
          
                        SetSelection False
                        SetSelection True, mRow, mItemCount
                        RaiseEvent SelectionChanged
                    End If
                    
                    lNewRow = mItemCount
                    
                    SBValue(efsVertical) = SBMax(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Shift And vbCtrlMask Then
                If UpdateCell() Then
                    lNewRow = mItemCount
                    
                    SBValue(efsVertical) = SBMax(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Not mEditPending Then
                lNewCol = UBound(mCols)
                
                SBValue(efsHorizontal) = SBMax(efsHorizontal)
                KeyCode = 0
            End If
            
    End Select
    
    SetRedrawState True
    
    If KeyCode = 0 Then
        'Do we want to clear selection?
        If bClearSelection And (mRow <> lNewRow) Then
            bRedraw = SetSelection(False)
        End If
        
        If Not mItems(mIX(lNewRow)).nFlags And lgSelected Then
            bRedraw = True
            SetFlag mItems(mIX(lNewRow)).nFlags, lgSelected, True
            RaiseEvent SelectionChanged
        End If
        
        If bRedraw Or SetRowCol(lNewRow, lNewCol, True) Then
            DrawGrid
        End If
    Else
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '#############################################################################################################################
    'Purpose: This will find the Item that contains a Cell with text that is >= to the text typed. Each
    'character entered is appended to the previous one if the time interval is less than 1 second.
    
    'Key searching is disabled if the Grid is Disabled, and Edit is in progress or the KeyPress event is
    'in an Ignore State (setting the SearchColumn to -1 will also prevent searches).
    
    Static lTime As Long
    Static sCode As String
    
    Dim lResult As Long
    Dim bEatKey As Boolean
   
    If mEnabled Then
        'Used to prevent a beep
        If (mEditTrigger And lgEnterKey) And (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
            KeyAscii = 0
            bEatKey = True
        ElseIf Not mIgnoreKeyPress And Not mEditPending Then
            If IsCharAlphaNumeric(KeyAscii) Then
                If (GetTickCount() - lTime) < 1000 Then
                    sCode = sCode & Chr$(KeyAscii)
                Else
                    sCode = Chr$(KeyAscii)
                End If
                
                lTime = GetTickCount()
                
                lResult = FindItem(sCode, mSearchColumn, lgSMNavigate)
                If lResult > NULL_RESULT Then
                    TopRow = lResult
                End If
            End If
        End If
        
        If Not bEatKey Then RaiseEvent KeyPress(KeyAscii)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    mIgnoreKeyPress = False
    
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As RECT
    Dim lCurrentMouseCol As Long
    Dim bCancel As Boolean
    Dim bProcessed As Boolean
    Dim bRedraw As Boolean
    Dim bSelectionChanged As Boolean
    Dim bState As Boolean
    
    If Not mLocked And (Button <> 0) And (mItemCount >= 0) Then
        mScrollAction = SCROLL_NONE
            
        lCurrentMouseCol = GetColFromX(X)
        mMouseDownRow = GetRowFromY(Y)
            
        If Button = vbLeftButton Then
        
            Call SetCapture(UserControl.hwnd)
            mMouseDown = True
            
            If Y < mR.HeaderHeight Then
                If (UserControl.MousePointer <> vbSizeWE) Then
                    mMouseDownCol = lCurrentMouseCol
                    If mMouseDownCol <> NULL_RESULT Then
                        With UserControl
                            DrawHeader mMouseCol, lgDown
                            .Refresh
                        End With
                    End If
                End If
            ElseIf mMouseDownRow > NULL_RESULT Then
                If UpdateCell() Then
                    If mCheckboxes And (X <= RIGHT_CHECKBOX) Then
                        bRedraw = True
                        mMouseDown = False
                        
                        SetFlag mItems(mIX(mMouseDownRow)).nFlags, lgChecked, Not GetFlag(mItems(mIX(mMouseDownRow)).nFlags, lgChecked)
                        
                        RaiseEvent ItemChecked(mMouseDownRow)
                    Else
                        If lCurrentMouseCol > NULL_RESULT Then
                            If mCols(lCurrentMouseCol).nType = lgBoolean Then
                                SetCheckBoxRect mIX(mMouseDownRow), lCurrentMouseCol, RowTop(mMouseDownRow), R
                                
                                If (X >= R.Left) And (Y >= R.Top) And (X <= R.Left + mR.CheckBoxSize) And (Y <= R.Top + mR.CheckBoxSize) Then
                                    bRedraw = True
                                    RaiseEvent RequestEdit(mMouseDownRow, lCurrentMouseCol, bCancel)
                                    
                                    If Not bCancel Then
                                        bState = (mItems(mIX(mMouseDownRow)).Cell(lCurrentMouseCol).nFlags And lgChecked)
                                        SetFlag mItems(mIX(mMouseDownRow)).Cell(lCurrentMouseCol).nFlags, lgChecked, Not bState
                                    End If
                                End If
                            End If
                        End If
                        
                        If Not bProcessed Then
                            bState = (mItems(mIX(mMouseDownRow)).nFlags And lgSelected)
                            
                            If (Shift And vbShiftMask) And mMultiSelect Then
                                bSelectionChanged = SetSelection(False) Or SetSelection(True, mRow, mMouseDownRow)
                            ElseIf Shift And vbCtrlMask Then
                                If Not mMultiSelect Then
                                    SetSelection False
                                End If
                                
                                SetFlag mItems(mIX(mMouseDownRow)).nFlags, lgSelected, Not bState
                                bSelectionChanged = True
                            ElseIf Not bState Then
                                SetSelection False
                                
                                SetFlag mItems(mIX(mMouseDownRow)).nFlags, lgSelected, True
                                bSelectionChanged = True
                            End If
                        End If
                        
                        bRedraw = bRedraw Or SetRowCol(mMouseDownRow, lCurrentMouseCol)
                    End If
                    
                    If bRedraw Then
                        DrawGrid
                    End If
                End If
            End If
        Else ' Right Button
            If mMouseDownRow > NULL_RESULT Then
                If UpdateCell() Then
                    SetRowCol mMouseDownRow, lCurrentMouseCol
                    bSelectionChanged = SetSelection(False) Or SetSelection(True, mMouseDownRow, mMouseDownRow)
                    DrawGrid
                End If
            End If
        End If
        
        If bSelectionChanged Then
            RaiseEvent SelectionChanged
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lResizeX As Long
    
    Dim lCount As Long
    Dim lWidth As Long
    Dim nMove As lgMoveControlEnum
    Dim nPointer As Integer
    Dim bSelectionChanged As Boolean
    
    If Not mLocked And (mItemCount > 0) Then
        mMouseCol = GetColFromX(X)
        mMouseRow = GetRowFromY(Y)
        RaiseEvent MouseMove(Button, Shift, X, Y)
        
        '####################################################################################
        'Header button tracking
        If mMouseDownCol <> NULL_RESULT Then
            If (mMouseDownCol = mMouseCol) And (MouseRow = NULL_RESULT) Then
                DrawHeader mMouseCol, lgDown
            Else
                DrawHeader mMouseDownCol, lgNormal
            End If
            UserControl.Refresh
        End If
        
        'Hot tracking
        If mHotHeaderTracking And (Button = 0) Then
            If Y < mR.HeaderHeight Then
                'Do we need to draw a new "hot" header?
                If (mMouseCol <> mHotColumn) Then
                    DrawHeaderRow
                    DrawHeader mMouseCol, lgHot
                    mHotColumn = mMouseCol
                End If
            ElseIf (mHotColumn <> NULL_RESULT) Then
                'We have a previous "hot" header to clear
                DrawHeaderRow
            End If
        End If
    
        '####################################################################################
        If (Button = vbLeftButton) Then
            If (mResizeCol >= 0) Then
                'We are resizing a Column
                lWidth = (X - lResizeX)
                If lWidth > 1 Then
                    mCols(mResizeCol).lWidth = lWidth
                    mCols(mResizeCol).dCustomWidth = ScaleX(mCols(mResizeCol).lWidth, vbPixels, mScaleUnits)
                    
                    DrawGrid
                    
                    nMove = mCols(mResizeCol).MoveControl
                    RaiseEvent ColumnSizeChanged(mResizeCol, nMove)
                    
                    If mEditPending Then
                        MoveEditControl nMove
                    End If
                End If
            ElseIf (mMouseDownRow > NULL_RESULT) Then
                If mMouseDown And Y < 0 Then
                    'Mouse has been dragged off off the control
                    ScrollList SCROLL_UP
                ElseIf mMouseDown And Y > UserControl.ScaleHeight Then
                    'Mouse has been dragged off off the control
                    ScrollList SCROLL_DOWN
                ElseIf mMouseDown And (Shift = 0) And (mMouseRow > NULL_RESULT) Then
                    If mScrollAction = SCROLL_NONE Then
                        bSelectionChanged = SetSelection(False)
                        
                        If mMultiSelect Then
                            SetSelection True, mMouseDownRow, mMouseRow
                        Else
                            SetSelection True, mMouseRow, mMouseRow
                        End If
                        
                        If SetRowCol(mMouseRow, mMouseCol) Then
                            RaiseEvent SelectionChanged
                            DrawGrid
                        End If
                    Else
                        mScrollAction = SCROLL_NONE
                    End If
                End If
            End If
        ElseIf (Button = 0) Then
            nPointer = vbDefault
                
            'Only check for resize cursor if no buttons depressed
            If (mMouseRow = NULL_RESULT) Then
                lResizeX = 0
                mResizeCol = NULL_RESULT
                
                If (mAllowUserResizing = lgResizeCol) Or (mAllowUserResizing = lgResizeBoth) Then
                     For lCount = SBValue(efsHorizontal) To UBound(mCols)
                        lWidth = lWidth + mCols(lCount).lWidth
                        
                        If (X < lWidth + SIZE_VARIANCE) And (X > lWidth - SIZE_VARIANCE) Then
                            nPointer = vbSizeWE
                            mResizeCol = lCount
                            Exit For
                        End If
                        
                        lResizeX = lResizeX + mCols(lCount).lWidth
                    Next lCount
                End If
            End If
        
            With UserControl
                If .MousePointer <> nPointer Then
                    .MousePointer = nPointer
                End If
            End With
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lCurrentMouseRow As Long
    
    If (Button = vbLeftButton) Then
        Call ReleaseCapture
        
        lCurrentMouseRow = GetRowFromY(Y)
        'added by: Vincent J. Jamero
        'get row
        mMouseDownRow = GetRowFromY(Y)
            
        If (mResizeCol >= 0) Then
            'We resized a Column so reset Scrollbars
            SetScrollBars
            DrawGrid
            
            UserControl.MousePointer = vbDefault
        ElseIf (lCurrentMouseRow = NULL_RESULT) Then
            
            'Sort requested from Column Header click
            If (GetColFromX(X) = mMouseDownCol) And (mMouseDownCol <> NULL_RESULT) Then
                If mColumnSort Then
                    If (Shift And vbCtrlMask) And (mSortColumn <> NULL_RESULT) Then
                        If mSortSubColumn <> mMouseDownCol Then
                            mCols(mMouseDownCol).nSortOrder = lgSTAscending
                        End If
                        mSortSubColumn = mMouseDownCol
                        
                        Sort , mCols(mSortColumn).nSortOrder
                    Else
                        If mSortColumn <> mMouseDownCol Then
                            mCols(mMouseDownCol).nSortOrder = lgSTAscending
                            mSortSubColumn = NULL_RESULT
                        End If
                        mSortColumn = mMouseDownCol
                        
                        If mSortSubColumn <> NULL_RESULT Then
                            Sort , , , mCols(mSortSubColumn).nSortOrder
                        Else
                            Sort
                        End If
                    End If
                Else
                    DrawHeaderRow
                    RaiseEvent ColumnClick(mMouseDownCol)
                End If
            End If
            
        ElseIf mMouseDownRow > NULL_RESULT Then
            mMouseCol = GetColFromX(X)
        
            If SetRowCol(mMouseRow, mMouseCol) Then
                DrawGrid
            End If
        Else
            DrawHeaderRow
        End If
    End If

    
    mMouseDown = False
    mMouseDownCol = NULL_RESULT
    mResizeCol = NULL_RESULT
    
    mScrollAction = SCROLL_NONE
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '################################################################################
    'Appearance Properties
    mBackColor = PropBag.ReadProperty("BackColor", DEF_BACKCOLOR)
    mBackColorBkg = PropBag.ReadProperty("BackColorBkg", DEF_BACKCOLORBKG)
    mBackColorEdit = PropBag.ReadProperty("BackColorEdit", DEF_BACKCOLOREDIT)
    mBackColorFixed = PropBag.ReadProperty("BackColorFixed", DEF_BACKCOLORFIXED)
    mBackColorSel = PropBag.ReadProperty("BackColorSel", DEF_BACKCOLORSEL)
    mForeColor = PropBag.ReadProperty("ForeColor", DEF_FORECOLOR)
    mForeColorEdit = PropBag.ReadProperty("ForeColorEdit", DEF_FORECOLOREDIT)
    mForeColorFixed = PropBag.ReadProperty("ForeColorFixed", DEF_FORECOLORFIXED)
    mForeColorSel = PropBag.ReadProperty("ForeColorSel", DEF_FORECOLORSEL)
    mForeColorTotals = PropBag.ReadProperty("ForeColorTotals", DEF_FORECOLORTOTALS)
    
    mGridColor = PropBag.ReadProperty("GridColor", DEF_GRIDCOLOR)
    mProgressBarColor = PropBag.ReadProperty("ProgressBarColor", DEF_PROGRESSBARCOLOR)
    
    mBorderStyle = PropBag.ReadProperty("BorderStyle", DEF_BORDERSTYLE)
    mDisplayEllipsis = PropBag.ReadProperty("DisplayEllipsis", DEF_DISPLAYELLIPSIS)
    mFocusRectColor = PropBag.ReadProperty("FocusRectColor", DEF_FOCUSRECTCOLOR)
    mFocusRectMode = PropBag.ReadProperty("FocusRectMode", DEF_FOCUSRECTMODE)
    mFocusRectStyle = PropBag.ReadProperty("FocusRectStyle", DEF_FOCUSRECTSTYLE)
    mGridLines = PropBag.ReadProperty("GridLines", DEF_GRIDLINES)
    mGridLineWidth = PropBag.ReadProperty("GridLineWidth", DEF_GRIDLINEWIDTH)
    mThemeColor = PropBag.ReadProperty("ThemeColor", DEF_THEMECOLOR)
    mThemeStyle = PropBag.ReadProperty("ThemeStyle", DEF_THEMESTYLE)
    
    '################################################################################
    'Behaviour Properties
    mAllowUserResizing = PropBag.ReadProperty("AllowUserResizing", DEF_ALLOWUSERRESIZING)
    mCheckboxes = PropBag.ReadProperty("Checkboxes", DEF_CHECKBOXES)
    mColumnHeaders = PropBag.ReadProperty("ColumnHeaders", DEF_COLUMNHEADERS)
    mColumnSort = PropBag.ReadProperty("ColumnSort", DEF_COLUMNSORT)
    mEditable = PropBag.ReadProperty("Editable", DEF_EDITABLE)
    mEditTrigger = PropBag.ReadProperty("EditTrigger", DEF_EDITTRIGGER)
    mFullRowSelect = PropBag.ReadProperty("FullRowSelect", DEF_FULLROWSELECT)
    mHotHeaderTracking = PropBag.ReadProperty("HotHeaderTracking", DEF_HOTHEADERTRACKING)
    mMultiSelect = PropBag.ReadProperty("MultiSelect", DEF_MULTISELECT)
    mRedraw = PropBag.ReadProperty("Redraw", DEF_REDRAW)
    mScrollTrack = PropBag.ReadProperty("ScrollTrack", DEF_SCROLLTRACK)
    
    '################################################################################
    'Miscellaneous Properties
    mCacheIncrement = PropBag.ReadProperty("CacheIncrement", DEF_CACHEINCREMENT)
    mEnabled = PropBag.ReadProperty("Enabled", DEF_ENABLED)
    mFormatString = PropBag.ReadProperty("FormatString", DEF_FORMATSTRING)
    mLocked = PropBag.ReadProperty("Locked", DEF_LOCKED)
    mRowHeightMin = PropBag.ReadProperty("RowHeightMin", DEF_ROWHEIGHTMIN)
    mScaleUnits = PropBag.ReadProperty("ScaleUnits", DEF_SCALEUNITS)
    mSearchColumn = PropBag.ReadProperty("SearchColumn", DEF_SEARCHCOLUMN)
    
    'added by: Vincent J. Jamero
    m_IgnoreEmpty = PropBag.ReadProperty("IgnoreEmpty", False)
    
    mStriped = PropBag.ReadProperty("Striped", DEF_Striped)
    mSBackColor1 = PropBag.ReadProperty("SBackColor1", DEF_SBackColor1)
    mSBackColor2 = PropBag.ReadProperty("SBackColor2", DEF_SBackColor2)
    
    
    '################################################################################
    'Apply Settings
    
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    With UserControl
        .BackColor = mBackColorBkg
        .BorderStyle = mBorderStyle
    End With
    
    FormatString = mFormatString
    CreateRenderData
    SetColors
    
    '#############################################################################################################################
    'Subclassing
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        
        With UserControl
            Call Subclass_Start(.hwnd)
            Call Subclass_AddMsg(.hwnd, WM_KILLFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_SETFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEHOVER, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_HSCROLL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_VSCROLL, MSG_AFTER)
            
            If mWindowsXP Then
                Call Subclass_AddMsg(.hwnd, WM_THEMECHANGED)
            End If
        End With
        
        SBCreate UserControl.hwnd
        SBStyle = Style_Regular
        
        SBLargeChange(efsHorizontal) = 5
        SBSmallChange(efsHorizontal) = 1
        
        SBLargeChange(efsVertical) = 5
        SBSmallChange(efsVertical) = 1
     End If
End Sub

Private Sub UserControl_Resize()
    SetScrollBars
End Sub

Private Sub UserControl_Terminate()
    On Local Error GoTo UserControl_TerminateError
    
    If Not mClipRgn = 0 Then DeleteObject mClipRgn
    
    pSBClearUp
    Call Subclass_Stop(UserControl.hwnd)

UserControl_TerminateError:
    Exit Sub
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
    
    '################################################################################
    'Appearance Properties
    Call PropBag.WriteProperty("BackColor", mBackColor, DEF_BACKCOLOR)
    Call PropBag.WriteProperty("BackColorBkg", mBackColorBkg, DEF_BACKCOLORBKG)
    Call PropBag.WriteProperty("BackColorEdit", mBackColorEdit, DEF_BACKCOLOREDIT)
    Call PropBag.WriteProperty("BackColorFixed", mBackColorFixed, DEF_BACKCOLORFIXED)
    Call PropBag.WriteProperty("BackColorSel", mBackColorSel, DEF_BACKCOLORSEL)
    Call PropBag.WriteProperty("ForeColor", mForeColor, DEF_FORECOLOR)
    Call PropBag.WriteProperty("ForeColorEdit", mForeColorEdit, DEF_FORECOLOREDIT)
    Call PropBag.WriteProperty("ForeColorFixed", mForeColorFixed, DEF_FORECOLORFIXED)
    Call PropBag.WriteProperty("ForeColorSel", mForeColorSel, DEF_FORECOLORSEL)
    Call PropBag.WriteProperty("ForeColorTotals", mForeColorTotals, DEF_FORECOLORTOTALS)
    
    Call PropBag.WriteProperty("GridColor", mGridColor, DEF_GRIDCOLOR)
    Call PropBag.WriteProperty("ProgressBarColor", mProgressBarColor, DEF_PROGRESSBARCOLOR)
    
    Call PropBag.WriteProperty("BorderStyle", mBorderStyle, DEF_BORDERSTYLE)
    Call PropBag.WriteProperty("DisplayEllipsis", mDisplayEllipsis, DEF_DISPLAYELLIPSIS)
    Call PropBag.WriteProperty("FocusRectMode", mFocusRectMode, DEF_FOCUSRECTMODE)
    Call PropBag.WriteProperty("FocusRectColor", mFocusRectColor, DEF_FOCUSRECTCOLOR)
    Call PropBag.WriteProperty("FocusRectStyle", mFocusRectStyle, DEF_FOCUSRECTSTYLE)
    Call PropBag.WriteProperty("GridLines", mGridLines, DEF_GRIDLINES)
    Call PropBag.WriteProperty("GridLineWidth", mGridLineWidth, DEF_GRIDLINEWIDTH)
    Call PropBag.WriteProperty("ThemeColor", mThemeColor, DEF_THEMECOLOR)
    Call PropBag.WriteProperty("ThemeStyle", mThemeStyle, DEF_THEMESTYLE)
    
    '################################################################################
    'Behaviour Properties
    Call PropBag.WriteProperty("AllowUserResizing", mAllowUserResizing, DEF_ALLOWUSERRESIZING)
    Call PropBag.WriteProperty("Checkboxes", mCheckboxes, DEF_CHECKBOXES)
    Call PropBag.WriteProperty("ColumnHeaders", mColumnHeaders, DEF_COLUMNHEADERS)
    Call PropBag.WriteProperty("ColumnSort", mColumnSort, DEF_COLUMNSORT)
    Call PropBag.WriteProperty("Editable", mEditable, DEF_EDITABLE)
    Call PropBag.WriteProperty("EditTrigger", mEditTrigger, DEF_EDITTRIGGER)
    Call PropBag.WriteProperty("FullRowSelect", mFullRowSelect, DEF_FULLROWSELECT)
    Call PropBag.WriteProperty("HotHeaderTracking", mHotHeaderTracking, DEF_HOTHEADERTRACKING)
    Call PropBag.WriteProperty("MultiSelect", mMultiSelect, DEF_MULTISELECT)
    Call PropBag.WriteProperty("Redraw", mRedraw, DEF_REDRAW)
    Call PropBag.WriteProperty("ScrollTrack", mScrollTrack, DEF_SCROLLTRACK)
    
    '################################################################################
    'Miscellaneous Properties
    Call PropBag.WriteProperty("CacheIncrement", mCacheIncrement, DEF_CACHEINCREMENT)
    Call PropBag.WriteProperty("Enabled", mEnabled, DEF_ENABLED)
    Call PropBag.WriteProperty("FormatString", mFormatString, DEF_FORMATSTRING)
    Call PropBag.WriteProperty("Locked", mLocked, DEF_LOCKED)
    Call PropBag.WriteProperty("RowHeightMin", mRowHeightMin, DEF_ROWHEIGHTMIN)
    Call PropBag.WriteProperty("ScaleUnits", mScaleUnits, DEF_SCALEUNITS)
    Call PropBag.WriteProperty("SearchColumn", mSearchColumn, DEF_SEARCHCOLUMN)

    'added by: Vincent J. Jamero
    Call PropBag.WriteProperty("IgnoreEmpty", m_IgnoreEmpty, False)
    
    Call PropBag.WriteProperty("Striped", mStriped, DEF_Striped)
    Call PropBag.WriteProperty("SBackColor1", mSBackColor1, DEF_SBackColor1)
    Call PropBag.WriteProperty("SBackColor2", mSBackColor2, DEF_SBackColor2)

End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function




Public Property Let IgnoreEmpty(ByVal New_IgnoreEmpty As Boolean)
    'added by: Vincent J. Jamero
    '          June 25, 2006
    m_IgnoreEmpty = New_IgnoreEmpty
    PropertyChanged "IgnoreEmpty"
End Property

Public Property Get IgnoreEmpty() As Boolean
    'added by: Vincent J. Jamero
    '          June 25, 2006
    IgnoreEmpty = m_IgnoreEmpty
End Property


Public Function RowCount() As Long
    'Return Row Count
    'added by: Vincent J. Jamero
    '          June 24, 2006
    RowCount = mItemCount + 1
End Function


Public Sub ColSort(ColIndex As Long, Optional SortASC As Boolean = True)
    'Column Sort
    'added by: Vincent J. Jamero
    '          June 24, 2006
    Dim oldmColumnSort As Boolean
    
    'hold mColumnSort
    oldmColumnSort = mColumnSort
    
    mColumnSort = True
    mSortColumn = ColIndex
                    
    If SortASC = True Then
        mCols(ColIndex).nSortOrder = 1
    Else
        mCols(ColIndex).nSortOrder = 2
    End If

    Sort , mCols(ColIndex).nSortOrder
    DrawGrid
    
    'restore mColumnSort
    mColumnSort = oldmColumnSort
    
End Sub

Public Function HitTest(X As Single, Y As Single, Optional lCol As Variant, Optional lRow As Variant) As String
    
    'added by: Vincent J. Jamero
    '          June 25, 2006
    Dim lCurrentMouseCol As Long
    Dim lCurrentMouseRow As Long
    
    HitTest = ""
    
    'default
    If IsMissing(lCol) = False Then
        lCol = -1
    End If
    If IsMissing(lRow) = False Then
        lRow = -1
    End If
    
    
    
    Call ReleaseCapture
            
    If Not mLocked Then

     
        lCurrentMouseCol = GetColFromX(X)
        lCurrentMouseRow = GetRowFromY(Y)
        
        If lCurrentMouseRow > NULL_RESULT And lCurrentMouseCol > NULL_RESULT Then
            If IsMissing(lCol) = False Then
                lCol = lCurrentMouseCol
            End If
            If IsMissing(lRow) = False Then
                lRow = lCurrentMouseRow
            End If
            
            HitTest = CellText(lCurrentMouseRow, lCurrentMouseCol)
        End If
        
    End If
    
    mMouseDown = False
    mScrollAction = SCROLL_NONE
    mResizeCol = NULL_RESULT

End Function


Public Property Get Striped() As Boolean
    Striped = mStriped
End Property

Public Property Let Striped(ByVal NewValue As Boolean)
    mStriped = NewValue
    PropertyChanged "Striped"
End Property

Public Property Get SBackColor1() As OLE_COLOR
    SBackColor1 = mSBackColor1
End Property

Public Property Let SBackColor1(ByVal NewValue As OLE_COLOR)
    mSBackColor1 = NewValue
    PropertyChanged "SBackColor1"
End Property

Public Property Get SBackColor2() As OLE_COLOR
    SBackColor2 = mSBackColor2
End Property

Public Property Let SBackColor2(ByVal NewValue As OLE_COLOR)
    mSBackColor2 = NewValue
    PropertyChanged "SBackColor2"
End Property


'Added By Vincent Jamero
Public Sub EnsureVisible(ByVal lNewRow As Long)
    Row = lNewRow
    SBValue(efsVertical) = lNewRow
End Sub

Public Sub FillBackColor(ByVal Row1 As Long, ByVal Row2 As Long, ByVal Col1 As Long, ByVal Col2 As Long, ByVal lColor As Long)

    Dim X As Long
    Dim Y As Long
    
    If Row1 < 0 Or _
        Row1 > Row2 Or _
        Col1 < 0 Or _
        Col1 > Col2 Then
        
        Exit Sub
        
    End If
    
        
    For Y = Row1 To Row2
        For X = Col1 To Col2
            
            ApplyCellFormat Y, X, lgCFBackColor, lColor
            
        Next
    Next

End Sub


