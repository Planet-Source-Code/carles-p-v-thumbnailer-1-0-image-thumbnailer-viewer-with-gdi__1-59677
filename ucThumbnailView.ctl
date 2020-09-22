VERSION 5.00
Begin VB.UserControl ucThumbnailView 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   ClipControls    =   0   'False
   FillColor       =   &H80000008&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
End
Attribute VB_Name = "ucThumbnailView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucThumbnailView.ctl
' Author:        Carles P.V. (*)
' Dependencies:  mIOIPAThumbnailView.bas -> OleGuids3.tlb (in IDE only)
'                mListViewEx.bas
' Last revision: 2004.10.14
' Version:       1.0.0
'----------------------------------------------------------------------------------------
'
' (*) Self-Subclassing UserControl template (IDE safe) by Paul Caton:
'
'     Self-subclassing Controls/Forms - NO dependencies
'     http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'----------------------------------------------------------------------------------------
'
' History:
'
'   * 1.0.0: - First release.
'========================================================================================

Option Explicit
Option Compare Text

'-- API

'= Misc =================================================================================

Private Const MAX_PATH                As Long = 260

Private Const SHGFI_SMALLICON         As Long = &H1
Private Const SHGFI_USEFILEATTRIBUTES As Long = &H10
Private Const SHGFI_TYPENAME          As Long = &H400
Private Const SHGFI_SYSICONINDEX      As Long = &H4000

Private Type SHFILEINFO
    hIcon         As Long
    iIcon         As Long
    dwAttributes  As Long
    szDisplayName As String * MAX_PATH
    szTypeName    As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'//

Private Const DIB_RGB_COLORS As Long = 0

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'//

Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

Private Const LOGPIXELSY             As Long = 90
Private Const FW_NORMAL              As Long = 400
Private Const FW_BOLD                As Long = 700

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'//

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'//

Private Const COLOR_WINDOW        As Long = 5
Private Const COLOR_WINDOWTEXT    As Long = 8
Private Const COLOR_HIGHLIGHT     As Long = 13
Private Const COLOR_HIGHLIGHTTEXT As Long = 14
Private Const COLOR_BTNFACE       As Long = 15
Private Const COLOR_BTNTEXT       As Long = 18

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

'//

Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BF_RECT         As Long = &HF

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long

'//

Private Const DT_CENTER       As Long = &H1
Private Const DT_VCENTER      As Long = &H4
Private Const DT_SINGLELINE   As Long = &H20
Private Const DT_NOCLIP       As Long = &H100
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_custom1      As Long = DT_CENTER Or DT_NOCLIP Or DT_END_ELLIPSIS
Private Const DT_custom2      As Long = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'= Window general =======================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_STYLE        As Long = (-16)
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_TABSTOP       As Long = &H10000
Private Const WS_CHILD         As Long = &H40000000

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

Private Const SW_SHOW  As Long = 5

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOOWNERZORDER As Long = &H200

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'= ListView =============================================================================

Private Const WC_LISTVIEW              As String = "SysListView32"

Private Const LVS_EX_GRIDLINES         As Long = &H1&
Private Const LVS_EX_FULLROWSELECT     As Long = &H20&
Private Const LVS_EX_INFOTIP           As Long = &H400&
Private Const LVS_EX_LABELTIP          As Long = &H4000&

Private Const LVS_ICON                 As Long = &H0
Private Const LVS_REPORT               As Long = &H1

Private Const LVS_SINGLESEL            As Long = &H4
Private Const LVS_SHOWSELALWAYS        As Long = &H8
Private Const LVS_SHAREIMAGELISTS      As Long = &H40
Private Const LVS_AUTOARRANGE          As Long = &H100

Private Const LVSCW_AUTOSIZE           As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

'//

Private Type LVITEM
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    stateMask  As Long
    pszText    As String
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type

Private Type LVITEM_lp
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    stateMask  As Long
    pszText    As Long
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type

Private Type LVFINDINFO
    Flags       As Long
    psz         As String
    lParam      As Long
    pt          As POINTAPI
    vkDirection As Long
End Type

Private Const LVIF_TEXT              As Long = &H1
Private Const LVIF_IMAGE             As Long = &H2
Private Const LVIF_PARAM             As Long = &H4
Private Const LVIF_STATE             As Long = &H8

Private Const LVIS_FOCUSED           As Long = &H1
Private Const LVIS_SELECTED          As Long = &H2

Private Const LVFI_STRING            As Long = &H2
Private Const LVFI_PARTIAL           As Long = &H8

Private Const LVSICF_NOINVALIDATEALL As Long = &H1
Private Const LVSICF_NOSCROLL        As Long = &H2

Private Type LVCOLUMN
    mask       As Long
    fmt        As Long
    cx         As Long
    pszText    As String
    cchTextMax As Long
    iSubItem   As Long
    iImage     As Long
    iOrder     As Long
End Type

Private Type LVCOLUMN_lp
    mask       As Long
    fmt        As Long
    cx         As Long
    pszText    As Long
    cchTextMax As Long
    iSubItem   As Long
    iImage     As Long
    iOrder     As Long
End Type

Private Const LVCF_FMT     As Long = &H1
Private Const LVCF_WIDTH   As Long = &H2
Private Const LVCF_TEXT    As Long = &H4
Private Const LVCF_IMAGE   As Long = &H10

'//

Private Type HDITEM
    mask       As Long
    cxy        As Long
    pszText    As String
    hbm        As Long
    cchTextMax As Long
    fmt        As Long
    lParam     As Long
    iImage     As Long
    iOrder     As Long
End Type

Private Const HDF_LEFT            As Long = &H0
Private Const HDF_RIGHT           As Long = &H1
Private Const HDF_CENTER          As Long = &H2
Private Const HDF_JUSTIFYMASK     As Long = &H3
Private Const HDF_IMAGE           As Long = &H800
Private Const HDF_STRING          As Long = &H4000
Private Const HDF_BITMAP_ON_RIGHT As Long = &H1000

Private Const HDI_WIDTH           As Long = &H1
Private Const HDI_HEIGHT          As Long = HDI_WIDTH
Private Const HDI_FORMAT          As Long = &H4
Private Const HDI_IMAGE           As Long = &H20

'//

Private Const WM_SIZE                      As Long = &H5
Private Const WM_SETFOCUS                  As Long = &H7
Private Const WM_SETREDRAW                 As Long = &HB
Private Const WM_MOUSEACTIVATE             As Long = &H21
Private Const WM_SETFONT                   As Long = &H30
Private Const WM_NOTIFY                    As Long = &H4E
Private Const WM_KEYDOWN                   As Long = &H100
Private Const WM_KEYUP                     As Long = &H101
Private Const WM_RBUTTONDOWN               As Long = &H204

Private Const LVM_FIRST                    As Long = &H1000
Private Const LVM_SETBKCOLOR               As Long = (LVM_FIRST + 1)
Private Const LVM_SETIMAGELIST             As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT             As Long = (LVM_FIRST + 4)
Private Const LVM_SETITEM                  As Long = (LVM_FIRST + 6)
Private Const LVM_INSERTITEM               As Long = (LVM_FIRST + 7)
Private Const LVM_DELETEITEM               As Long = (LVM_FIRST + 8)
Private Const LVM_DELETEALLITEMS           As Long = (LVM_FIRST + 9)
Private Const LVM_GETNEXTITEM              As Long = (LVM_FIRST + 12)
Private Const LVM_FINDITEM                 As Long = (LVM_FIRST + 13)
Private Const LVM_GETITEMRECT              As Long = (LVM_FIRST + 14)
Private Const LVM_HITTEST                  As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE            As Long = (LVM_FIRST + 19)
Private Const LVM_REDRAWITEMS              As Long = (LVM_FIRST + 21)
Private Const LVM_GETCOLUMN                As Long = (LVM_FIRST + 25)
Private Const LVM_INSERTCOLUMN             As Long = (LVM_FIRST + 27)
Private Const LVM_GETCOLUMNWIDTH           As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH           As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER                As Long = (LVM_FIRST + 31)
Private Const LVM_SETTEXTCOLOR             As Long = (LVM_FIRST + 36)
Private Const LVM_SETTEXTBKCOLOR           As Long = (LVM_FIRST + 38)
Private Const LVM_SETITEMSTATE             As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE             As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT              As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMCOUNT             As Long = (LVM_FIRST + 47)
Private Const LVM_SETICONSPACING           As Long = (LVM_FIRST + 53)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)

Private Const HDM_FIRST                    As Long = &H1200
Private Const HDM_GETITEM                  As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEM                  As Long = (HDM_FIRST + 4)
Private Const HDM_HITTEST                  As Long = (HDM_FIRST + 6)
Private Const HDM_SETIMAGELIST             As Long = (HDM_FIRST + 8)

'//

Private Type NMHDR
    hWndFrom As Long
    idfrom   As Long
    code     As Long
End Type

Private Type NMLISTVIEW
    hdr       As NMHDR
    iItem     As Long
    iSubItem  As Long
    uNewState As Long
    uOldState As Long
    uChanged  As Long
    ptAction  As POINTAPI
    lParam    As Long
End Type

Private Type LVHITTESTINFO
    pt       As POINTAPI
    Flags    As Long
    iItem    As Long
    iSubItem As Long
End Type

Private Type NMHEADER_short
    hdr     As NMHDR
    iItem   As Long
    iButton As Long
    hbm     As Long
End Type

Private Type NMLVGETINFOTIP_lp
   hdr        As NMHDR
   dwFlags    As Long
   pszText    As Long
   cchTextMax As Long
   iItem      As Long
   iSubItem   As Long
   lParam     As Long
End Type

Private Type HDHITTESTINFO
    pt    As POINTAPI
    Flags As Long
    iItem As Long
End Type

Private Const NM_FIRST             As Long = 0
Private Const NM_DBLCLK            As Long = (NM_FIRST - 3)
Private Const NM_RDBLCLK           As Long = (NM_FIRST - 6)
Private Const NM_SETFOCUS          As Long = (NM_FIRST - 7)
Private Const NM_CUSTOMDRAW        As Long = (NM_FIRST - 12)

Private Const LVN_FIRST            As Long = -100
Private Const LVN_ITEMCHANGED      As Long = (LVN_FIRST - 1)
Private Const LVN_GETINFOTIP       As Long = (LVN_FIRST - 57)

Private Const HDN_FIRST            As Long = -300
Private Const HDN_ITEMCHANGED      As Long = (HDN_FIRST - 1)
Private Const HDN_ITEMCLICK        As Long = (HDN_FIRST - 2)

Private Const LVNI_FOCUSED         As Long = &H1
Private Const LVNI_SELECTED        As Long = &H2

Private Const LVHT_NOWHERE         As Long = &H1
Private Const LVHT_ONITEMICON      As Long = &H2
Private Const LVHT_ONITEMLABEL     As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM          As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

Private Const LVIR_ICON            As Long = 1

'= Custom draw ==========================================================================

Private Type NMCUSTOMDRAW
    hdr         As NMHDR
    dwDrawStage As Long
    hDC         As Long
    rc          As RECT2
    dwItemSpec  As Long
    uItemState  As Long
    lItemlParam As Long
End Type

Private Type NMLVCUSTOMDRAW
    nmcd      As NMCUSTOMDRAW
    clrText   As Long
    clrTextBk As Long
    iSubItem  As Long
End Type

Private Const CDDS_PREPAINT          As Long = &H1
Private Const CDDS_ITEM              As Long = &H10000
Private Const CDDS_ITEMPREPAINT      As Long = (CDDS_ITEM Or CDDS_PREPAINT)
Private Const CDIS_FOCUS             As Long = &H10
Private Const CDRF_SKIPDEFAULT       As Long = &H4
Private Const CDRF_NOTIFYITEMDRAW    As Long = &H20

'= Image list ===========================================================================

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImageList As Long) As Long

Private Const LVSIL_NORMAL    As Long = 0
Private Const LVSIL_SMALL     As Long = 1

Private Const ILD_TRANSPARENT As Long = 1&
Private Const ILD_MASK        As Long = &H10&

'//

'-- Public enums.:

Public Enum tvViewModeConstants
    [tvThumbnail] = LVS_ICON
    [tvDetails] = LVS_REPORT
End Enum

Public Enum tvColumnAlignConstants
    [tvLeft] = HDF_LEFT
    [tvRight] = HDF_RIGHT
    [tvCenter] = HDF_CENTER
End Enum

Public Enum tvColumnIDConstants
    [tvFileName] = 0
    [tvFileType]
    [tvFileSize]
    [tvFileDate]
End Enum

Public Enum tvColumnAutosizeConstants
    [tvContents] = LVSCW_AUTOSIZE
    [caHeader] = LVSCW_AUTOSIZE_USEHEADER
End Enum

Public Enum tvCoincidenceConstants
    [tvWholeWord] = LVFI_STRING
    [tvPartial] = LVFI_PARTIAL
End Enum

Public Enum eStateConstants
    [tvSelected] = LVNI_SELECTED
    [tvFocused] = LVNI_FOCUSED
End Enum

Public Enum tvSortOrderConstants
    [tvDefault] = 0
    [tvAscending] = 1
    [tvDescending] = -1
End Enum

Public Enum tvSortTypeConstants
    [tvString] = 0
    [tvNumeric] = 1
    [tvDate] = 2
End Enum

Public Enum tvBorderStyleConstants
    [tvNone] = 0
    [tvFixedSingle]
End Enum

'-- Private types:

Private Type THUMBNAIL_DATA
    TBIH() As Byte
    Data() As Byte
    Info   As String
End Type

Private Type IMAGETYPE_INFO
    Extension As String
    Name      As String
    IconIndex As Long
End Type

'-- Property variables:

Private m_HideSelection        As Boolean
Private m_ViewMode             As tvViewModeConstants
Private m_ThumbnailWidth       As Long
Private m_ThumbnailHeight      As Long

'-- Private constants:

Private Const THUMBNAIL_BORDER As Long = 4
 
'-- Private variables:

Private m_bInitialized         As Boolean
Private m_hListView            As Long
Private m_hHeader              As Long
Private m_uImageTypeInfo()     As IMAGETYPE_INFO
Private m_uThumbnailInfo()     As THUMBNAIL_DATA
Private m_hILLarge             As Long
Private m_hILHeader            As Long
Private m_hFont                As Long
Private m_lFontHeight          As Long
Private m_eColumnSortOrder(3)  As tvSortOrderConstants
Private m_lColumnSortIndex     As Long
Private m_lColumnSortIndexPrev As Long
Private m_uIPAO                As IPAOHookStructThumbnailView

'-- Event declarations:

Public Event ItemClick(ByVal Item As Long)
Public Event ItemDblClick(ByVal Item As Long)
Public Event ItemRightClick(ByVal Item As Long)
Public Event ColumnResize(ByVal ColumnID As tvColumnIDConstants)
Public Event Resize()



'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                  'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                 'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES     As Long = -1          'All messages added or deleted
Private Const CODE_LEN         As Long = 200         'Length of the machine code in bytes
Private Const GWL_WNDPROC      As Long = -4          'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04         As Long = 88          'Table B (before) address patch offset
Private Const PATCH_05         As Long = 93          'Table B (before) entry count patch offset
Private Const PATCH_08         As Long = 132         'Table A (after) address patch offset
Private Const PATCH_09         As Long = 137         'Table A (after) entry count patch offset

Private Type tSubData                                'Subclass data type
    hWnd                       As Long               'Handle of the window being subclassed
    nAddrSub                   As Long               'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long               'The address of the pre-existing WndProc
    nMsgCntA                   As Long               'Msg after table entry count
    nMsgCntB                   As Long               'Msg before table entry count
    aMsgTblA()                 As Long               'Msg after table array
    aMsgTblB()                 As Long               'Msg Before table array
End Type

Private sc_aSubData()          As tSubData           'Subclass data array
Private sc_aBuf(1 To CODE_LEN) As Byte               'Code buffer byte array
Private sc_pCWP                As Long               'Address of the CallWindowsProc
Private sc_pEbMode             As Long               'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long               'Address of the SetWindowsLong function
  
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long



'========================================================================================
' Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"

'Parameters:
'   bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'   bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'   lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'   lng_hWnd - The window handle
'   uMsg     - The message number
'   wParam   - Message related data
'   lParam   - Message related data
'
'Notes:
'   If you really know what you're doing, it's possible to change the values of the
'   hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'   values get passed to the default handler.. and optionaly, the 'after' callback

  Dim uNMH      As NMHDR
  Dim uNMHH     As NMHEADER_short
  Dim uNMLV     As NMLISTVIEW
  Dim uLVHTI    As LVHITTESTINFO
  Dim uHDHTI    As HDHITTESTINFO
  Dim uNMLVIT   As NMLVGETINFOTIP_lp
  Dim bDblClick As Boolean
  Dim lItm      As Long
  Dim sTip      As String
  
    Select Case lng_hWnd
    
        Case UserControl.hWnd
        
            Select Case uMsg
                
                Case WM_SETFOCUS
                    
                    Call SetFocus(m_hListView)
                
                Case WM_MOUSEACTIVATE
                    
                    Call pvSetIPAO
                                    
                Case WM_SIZE
  
                    Call MoveWindow(m_hListView, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1)
                
                Case WM_NOTIFY

                    Call CopyMemory(uNMH, ByVal lParam, Len(uNMH))
                    
                    If (uNMH.hWndFrom = m_hHeader) Then
                    
                        Select Case uNMH.code
                        
                            Case HDN_ITEMCLICK
                                
                                If (SendMessageLong(m_hListView, LVM_GETITEMCOUNT, 0, 0) > 1) Then
                                    With uHDHTI
                                        Call GetCursorPos(.pt)
                                        Call ScreenToClient(m_hHeader, .pt)
                                        Call SendMessage(m_hHeader, HDM_HITTEST, 0, uHDHTI)
                                        Call pvSort(.iItem): m_lColumnSortIndexPrev = -1
                                    End With
                                End If
                                
                            Case HDN_ITEMCHANGED
                                
                                If (m_ViewMode = [tvDetails]) Then
                                    Call CopyMemory(uNMHH, ByVal lParam, Len(uNMHH))
                                    RaiseEvent ColumnResize(uNMHH.iItem)
                                End If
                        End Select
                        
                    ElseIf (uNMH.hWndFrom = m_hListView) Then
                    
                        Select Case uNMH.code
                        
                            Case NM_SETFOCUS
                                
                                Call pvSetIPAO
                            
                            Case NM_DBLCLK, NM_RDBLCLK
                                
                                Call CopyMemory(uNMLV, ByVal lParam, Len(uNMLV))
                                
                                With uLVHTI
                                    Call pvUCCoords(.pt)
                                    Call SendMessage(m_hListView, LVM_HITTEST, 0, uLVHTI)
                                    If (.Flags <> LVHT_NOWHERE) Then
                                        If ((.Flags = LVHT_ONITEMICON) Or _
                                            (.Flags = LVHT_ONITEMLABEL) Or _
                                            (.Flags = LVHT_ONITEM)) Then
                                            bDblClick = True
                                        End If
                                    End If
                                End With
                                
                                If (bDblClick) Then
                                    lItm = pvItemHitTest()
                                    If (lItm <> -1) Then
                                        RaiseEvent ItemDblClick(lItm)
                                    End If
                                End If
                    
                            Case LVN_ITEMCHANGED
                            
                                Call CopyMemory(uNMLV, ByVal lParam, Len(uNMLV))
                                
                                With uNMLV
                                    If (.uOldState = 0) Then
                                        If ((.uNewState And LVIS_SELECTED)) Then
                                            RaiseEvent ItemClick(.iItem)
                                        End If
                                    End If
                                End With
                               
                            Case LVN_GETINFOTIP
                                    
                                Call CopyMemory(uNMLVIT, ByVal lParam, Len(uNMLVIT))
                                
                                sTip = Me.ItemText(uNMLVIT.iItem, 0) & vbCrLf & _
                                       Me.ItemText(uNMLVIT.iItem, 1) & vbCrLf & _
                                       Me.ItemText(uNMLVIT.iItem, 2) & " bytes" & vbCrLf & _
                                       Me.ItemText(uNMLVIT.iItem, 3)

                                If (Len(sTip)) Then
                                    uNMLVIT.cchTextMax = Len(sTip)
                                    Call CopyMemory(ByVal uNMLVIT.pszText, ByVal sTip, Len(sTip))
                                End If
                                bHandled = True
                                
                            Case NM_CUSTOMDRAW
                                
                                If (m_ViewMode = [tvThumbnail]) Then
                                    bHandled = True
                                    lReturn = pvCustomDrawThumbnail(lParam)
                                End If
                        End Select
                    End If
            End Select
            
        Case m_hListView
        
            If (uMsg = WM_RBUTTONDOWN) Then
            
                lItm = pvItemHitTest()
                If (lItm <> -1) Then
                    RaiseEvent ItemRightClick(lItm)
                End If
            End If
    End Select
End Sub



'========================================================================================
' Usercontrol
'========================================================================================

Private Sub UserControl_Initialize()
    
    '-- Defaults
    m_ThumbnailWidth = 80
    m_ThumbnailHeight = 80
    
    '-- Private
    m_lColumnSortIndex = -1
    m_lColumnSortIndexPrev = -1
End Sub

Private Sub UserControl_Terminate()
  
  On Error GoTo errH
  
    If (m_bInitialized) Then
        
        Call mIOIPAThumbnailView.TerminateIPAO(m_uIPAO)
        Call Subclass_StopAll
        Call pvDestroyImageListThumbnail
        Call pvDestroyImageListHeader
        Call pvDestroyFont
        Call pvDestroyListView
    End If
errH:
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function Initialize(ByVal FormatsMask As String, _
                           ByVal MaskSeparator As String, _
                           Optional ByVal ViewMode As tvViewModeConstants = [tvThumbnail], _
                           Optional ByVal NameColumnWidth As Long = 100, _
                           Optional ByVal TypeColumnWidth As Long = 100, _
                           Optional ByVal SizeColumnWidth As Long = 100, _
                           Optional ByVal DateColumnWidth As Long = 100) As Boolean
    
    If (m_bInitialized = False) Then
    
        Initialize = pvCreate(FormatsMask, MaskSeparator, ViewMode, NameColumnWidth, TypeColumnWidth, SizeColumnWidth, DateColumnWidth)
        
        If (m_hListView) Then
        
            m_bInitialized = True
            
            '-- Subclass UserControl (parent) and ListView (child)
            Call Subclass_Start(UserControl.hWnd)
            Call Subclass_Start(m_hListView)
            
            '-- Check next messages...
            Call Subclass_AddMsg(UserControl.hWnd, WM_SETFOCUS)
            Call Subclass_AddMsg(UserControl.hWnd, WM_MOUSEACTIVATE)
            Call Subclass_AddMsg(UserControl.hWnd, WM_NOTIFY)
            Call Subclass_AddMsg(UserControl.hWnd, WM_SIZE)
            Call Subclass_AddMsg(m_hListView, WM_RBUTTONDOWN)
            
            '-- Initialize IOLEInPlaceActiveObject
            Call mIOIPAThumbnailView.InitIPAO(m_uIPAO, Me)
        End If
    End If
End Function

Public Sub SetRedraw(ByVal bRedraw As Boolean)

    If (m_hListView) Then
        '-- Enable/disable redraw mode
        Call SendMessageLong(m_hListView, WM_SETREDRAW, -bRedraw, 0)
    End If
End Sub

Public Sub RefreshItems(ByVal ItemFirst As Long, ByVal ItemLast As Long)

    If (m_hListView) Then
        '-- Repaint a range of items
        Call SendMessageLong(m_hListView, LVM_REDRAWITEMS, ItemFirst, ItemLast)
    End If
End Sub

'//

Public Function Clear( _
                ) As Boolean

  Dim lCol As Long
    
    If (m_hListView) Then
        
        '-- Clear all items (and DIBs collection)
        Clear = CBool(SendMessageLong(m_hListView, LVM_DELETEALLITEMS, 0, 0))
        Erase m_uThumbnailInfo()
        
        '-- Reset column icons
        For lCol = 0 To 3
            m_eColumnSortOrder(lCol) = [tvDefault]
            Me.ColumnIcon(lCol) = -1
        Next lCol
        m_lColumnSortIndex = -1
        m_lColumnSortIndexPrev = -1
    End If
End Function

Public Function Sort(ByVal ColumnID As tvColumnIDConstants) As Boolean
    
    '-- Sort column
    Sort = pvSort(ColumnID)
End Function

Public Function ItemSetCount(ByVal FinalCount As Long) As Boolean

    If (m_hListView) Then
        
        Call SendMessageLong(m_hListView, LVM_SETITEMCOUNT, FinalCount, LVSICF_NOINVALIDATEALL Or LVSICF_NOSCROLL)
        ReDim m_uThumbnailInfo(FinalCount - 1)
    End If
End Function

Public Function ItemAdd( _
                ByVal Item As Long, _
                ByVal Filename As String, _
                ByVal FileSize As String, _
                ByVal FileDate As String _
                ) As Boolean
   
  Dim uLVI As LVITEM_lp
  Dim uITI As IMAGETYPE_INFO
  
    If (m_hListView) Then
    
        uITI = pvGetImageTypeInfo(Filename)
        
        With uLVI
            .iItem = Item
            .lParam = Item
            .pszText = StrPtr(StrConv(Filename, vbFromUnicode))
            .cchTextMax = Len(Filename)
            .iImage = uITI.IconIndex
            .mask = LVIF_TEXT Or LVIF_IMAGE Or LVIF_PARAM
        End With
        ItemAdd = (SendMessage(m_hListView, LVM_INSERTITEM, 0, uLVI) > -1)
        
        Call pvSubItemSet(Item, 1, uITI.Name)
        Call pvSubItemSet(Item, 2, Format$(FileSize, "#,0"))
        Call pvSubItemSet(Item, 3, FileDate)
    End If
End Function

Public Function ItemRemove( _
                ByVal Item As Long _
                ) As Boolean
    
    If (m_hListView) Then
    
        ItemRemove = CBool(SendMessageLong(m_hListView, LVM_DELETEITEM, Item, 0))
    End If
End Function

Public Function ItemEnsureVisible( _
                ByVal Item As Long _
                ) As Boolean

    If (m_hListView) Then
        
        ItemEnsureVisible = CBool(SendMessageLong(m_hListView, LVM_ENSUREVISIBLE, Item, 0))
    End If
End Function
 
Public Function ItemFindText( _
                ByVal Text As String, _
                Optional ByVal StartItem As Long = -1, _
                Optional ByVal Coincidence As tvCoincidenceConstants = [tvWholeWord] _
                ) As Long
  
  Dim uLVFI As LVFINDINFO
    
    If (m_hListView) Then
    
        With uLVFI
            .psz = Text + vbNullChar
            .Flags = Coincidence
        End With
        
        ItemFindText = SendMessage(m_hListView, LVM_FINDITEM, StartItem, uLVFI)
    End If
End Function

Public Function ItemFindState( _
                Optional ByVal StartItem As Long = -1, _
                Optional ByVal State As eStateConstants = [tvSelected] _
                ) As Long

    If (m_hListView) Then
        
        ItemFindState = SendMessageLong(m_hListView, LVM_GETNEXTITEM, StartItem, State)
    End If
End Function

Public Function ItemHitTest( _
                ByVal x As Long, _
                ByVal y As Long _
                ) As Long

  Dim uLVHI As LVHITTESTINFO
    
    If (m_hListView) Then
    
        With uLVHI.pt
            .x = ScaleX(x, UserControl.ScaleMode, vbPixels)
            .y = ScaleY(y, UserControl.ScaleMode, vbPixels)
        End With
        
        ItemHitTest = SendMessage(m_hListView, LVM_HITTEST, 0, uLVHI)
    End If
End Function

'//

Public Sub ThumbnailInfo_SetTBIH(ByVal Index As Long, aTBIH() As Byte)
On Error Resume Next
    m_uThumbnailInfo(Index).TBIH() = aTBIH()
End Sub
Public Sub ThumbnailInfo_GetTBIH(ByVal Index As Long, aTBIH() As Byte)
    aTBIH() = m_uThumbnailInfo(Index).TBIH()
End Sub

Public Sub ThumbnailInfo_SetData(ByVal Index As Long, aData() As Byte)
On Error Resume Next
    m_uThumbnailInfo(Index).Data() = aData()
End Sub
Public Sub ThumbnailInfo_GetData(ByVal Index As Long, aData() As Byte)
    aData() = m_uThumbnailInfo(Index).Data()
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Get ThumbnailWidth() As Long
    ThumbnailWidth = m_ThumbnailWidth
End Property

Public Property Get ThumbnailHeight() As Long
    ThumbnailHeight = m_ThumbnailHeight
End Property

Public Sub SetThumbnailSize(ByVal New_Width As Long, ByVal New_Height As Long)
    
    m_ThumbnailWidth = New_Width
    m_ThumbnailHeight = New_Height
    
    If (m_ThumbnailWidth < 32) Then m_ThumbnailWidth = 32
    If (m_ThumbnailWidth > 125) Then m_ThumbnailWidth = 125
    If (m_ThumbnailHeight < 32) Then m_ThumbnailHeight = 32
    If (m_ThumbnailHeight > 125) Then m_ThumbnailHeight = 125
    
    Call pvDestroyImageListThumbnail
    Call pvInitializeImageListThumbnail
End Sub

'//

Public Property Get ColumnIcon(ByVal ColumnID As tvColumnIDConstants) As Long

  Dim uLVC As LVCOLUMN
  
    If (m_hListView) Then
        
        With uLVC
            .mask = LVCF_IMAGE
        End With
        Call SendMessage(m_hListView, LVM_GETCOLUMN, ColumnID, uLVC)
        
        ColumnIcon = uLVC.iImage
    End If
End Property

Public Property Let ColumnIcon(ByVal ColumnID As tvColumnIDConstants, ByVal Icon As Long)
  
  Dim uHDI   As HDITEM
  Dim lAlign As Long
  
    If (m_hListView) Then
                
        With uHDI
            .mask = HDI_FORMAT
            Call SendMessage(m_hHeader, HDM_GETITEM, ColumnID, uHDI): lAlign = HDF_JUSTIFYMASK And .fmt
            .iImage = Icon
            .fmt = HDF_STRING Or lAlign Or HDF_IMAGE * -(Icon > -1 And m_hILHeader <> 0) Or HDF_BITMAP_ON_RIGHT
            .mask = HDI_IMAGE * -(Icon > -1) Or HDI_FORMAT
        End With
        Call SendMessage(m_hHeader, HDM_SETITEM, ColumnID, uHDI)
    End If
End Property

Public Property Get ColumnWidth(ByVal ColumnID As tvColumnIDConstants) As Long
Attribute ColumnWidth.VB_MemberFlags = "400"

    If (m_hListView) Then

        ColumnWidth = SendMessageLong(m_hListView, LVM_GETCOLUMNWIDTH, ColumnID, 0)
    End If
End Property
Public Property Let ColumnWidth(ByVal ColumnID As tvColumnIDConstants, ByVal Width As Long)

    If (m_hListView) Then

        Call SendMessageLong(m_hListView, LVM_SETCOLUMNWIDTH, ColumnID, Width)
    End If
End Property

'//

Public Property Get ItemText(ByVal Item As Long, ByVal ColumnID As tvColumnIDConstants) As String
  
  Dim uLVI   As LVITEM_lp
  Dim A(260) As Byte
  Dim lLen   As Long
    
    If (m_hListView) Then
        
        With uLVI
            .iSubItem = ColumnID
            .pszText = VarPtr(A(0))
            .cchTextMax = 261
            .mask = LVIF_TEXT
        End With
        lLen = SendMessage(m_hListView, LVM_GETITEMTEXT, Item, uLVI)
        
        ItemText = Left$(StrConv(A(), vbUnicode), lLen)
    End If
End Property

Public Property Get ItemSelected(ByVal Item As Long) As Boolean
Attribute ItemSelected.VB_MemberFlags = "400"
  
    If (m_hListView) Then
    
        ItemSelected = CBool(SendMessageLong(m_hListView, LVM_GETITEMSTATE, Item, LVIS_SELECTED))
    End If
End Property
Public Property Let ItemSelected(ByVal Item As Long, ByVal Selected As Boolean)
  
  Dim uLVI As LVITEM
    
    If (m_hListView) Then
        
        With uLVI
            .stateMask = LVIS_SELECTED Or LVIS_FOCUSED
            .State = -Selected * (LVIS_SELECTED Or LVIS_FOCUSED)
            .mask = LVIF_STATE
        End With
        Call SendMessage(m_hListView, LVM_SETITEMSTATE, Item, uLVI)
    End If
End Property

Public Property Get ItemFocused(ByVal Item As Long) As Boolean
Attribute ItemFocused.VB_MemberFlags = "400"
  
    If (m_hListView) Then
    
        ItemFocused = CBool(SendMessageLong(m_hListView, LVM_GETITEMSTATE, Item, LVIS_FOCUSED))
    End If
End Property
Public Property Let ItemFocused(ByVal Item As Long, ByVal Focused As Boolean)
  
  Dim uLVI As LVITEM
    
    If (m_hListView) Then
        
        With uLVI
            .stateMask = LVIS_FOCUSED
            .State = -Focused * LVIS_FOCUSED
            .mask = LVIF_STATE
        End With
        Call SendMessage(m_hListView, LVM_SETITEMSTATE, Item, uLVI)
    End If
End Property

'//

Public Property Get BorderStyle() As fvBorderStyleConstants
    If (m_hListView) Then
        BorderStyle = -((GetWindowLong(m_hListView, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE)
    End If
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As fvBorderStyleConstants)
    If (m_hListView) Then
        Select Case New_BorderStyle
            Case [tvNone]
                Call SetWindowLong(m_hListView, GWL_EXSTYLE, 0)
            Case [tvFixedSingle]
                Call SetWindowLong(m_hListView, GWL_EXSTYLE, WS_EX_CLIENTEDGE)
        End Select
    End If
End Property
Public Property Get Count() As Long
Attribute Count.VB_MemberFlags = "400"
    If (m_hListView) Then
        Count = SendMessageLong(m_hListView, LVM_GETITEMCOUNT, 0, 0)
    End If
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_MemberFlags = "400"
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If (m_hListView) Then
        UserControl.Enabled = New_Enabled
        Call EnableWindow(m_hListView, New_Enabled)
    End If
End Property

Public Property Get HideSelection() As Boolean
    HideSelection = m_HideSelection
End Property
Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    If (m_hListView) Then
        m_HideSelection = New_HideSelection
        If (m_HideSelection) Then
            Call pvSetStyle(0, LVS_SHOWSELALWAYS)
          Else
            Call pvSetStyle(LVS_SHOWSELALWAYS, 0)
        End If
    End If
End Property

Public Property Let ViewMode(ByVal New_ViewMode As tvViewModeConstants)
    
    If (m_hListView) Then
        
        If (m_ViewMode <> New_ViewMode) Then
            m_ViewMode = New_ViewMode
            
            '-- Change colors/styles
            Select Case m_ViewMode
                
                Case [tvThumbnail]
                        
                    Call SendMessageLong(m_hListView, LVM_SETBKCOLOR, 0, GetSysColor(COLOR_BTNFACE))
                    Call SendMessageLong(m_hListView, LVM_SETTEXTBKCOLOR, 0, GetSysColor(COLOR_BTNFACE))
                    Call SendMessageLong(m_hListView, LVM_SETTEXTCOLOR, 0, GetSysColor(COLOR_BTNTEXT))
                    Call pvSetExStyle(LVS_EX_INFOTIP, 0)
                    Call pvSetExStyle(0, LVS_EX_LABELTIP)
                
                Case [tvDetails]
                    
                    Call SendMessageLong(m_hListView, LVM_SETBKCOLOR, 0, GetSysColor(COLOR_WINDOW))
                    Call SendMessageLong(m_hListView, LVM_SETTEXTBKCOLOR, 0, GetSysColor(COLOR_WINDOW))
                    Call SendMessageLong(m_hListView, LVM_SETTEXTCOLOR, 0, GetSysColor(COLOR_WINDOWTEXT))
                    Call pvSetExStyle(0, LVS_EX_INFOTIP)
                    Call pvSetExStyle(LVS_EX_LABELTIP, 0)
            End Select
            
            '-- Change view
            Call pvSetStyle(m_ViewMode, IIf(m_ViewMode = [tvThumbnail], LVS_REPORT, LVS_ICON))
            
            '-- Re-sort ?
            If (m_lColumnSortIndexPrev <> m_lColumnSortIndex And m_lColumnSortIndex <> -1) Then
                m_lColumnSortIndexPrev = m_lColumnSortIndex
                Call pvSort(m_lColumnSortIndex, True)
            End If
        End If
    End If
End Property
Public Property Get ViewMode() As tvViewModeConstants
Attribute ViewMode.VB_MemberFlags = "400"
    ViewMode = m_ViewMode
End Property

'//

Public Property Get hWnd() As Long
    hWnd = m_hListView
End Property



'========================================================================================
' Private
'========================================================================================

Private Function pvCreate(ByVal FormatsMask As String, _
                          ByVal MaskSeparator As String, _
                          ByVal ViewMode As tvViewModeConstants, _
                          ByVal NameColumnWidth As Long, _
                          ByVal TypeColumnWidth As Long, _
                          ByVal SizeColumnWidth As Long, _
                          ByVal DateColumnWidth As Long) As Boolean
    
  Dim lExStyle As Long
  Dim lLVStyle As Long
  
    '-- Define window style
    lExStyle = WS_EX_CLIENTEDGE
    lLVStyle = WS_CHILD Or WS_TABSTOP Or ViewMode Or LVS_AUTOARRANGE Or LVS_SINGLESEL Or LVS_SHOWSELALWAYS Or LVS_SHAREIMAGELISTS

    '-- Create ListView window
    m_hListView = CreateWindowEx(lExStyle, WC_LISTVIEW, vbNullString, lLVStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0)
    
    '-- Success [?]
    If (m_hListView) Then
    
        '-- Initialize font
        Call pvSetFont(Ambient.Font)
        
        '-- Initialize columns
        Call pvColumnAdd(0, "Name", NameColumnWidth)
        Call pvColumnAdd(1, "Type", TypeColumnWidth)
        Call pvColumnAdd(2, "Size", SizeColumnWidth, [tvRight])
        Call pvColumnAdd(3, "Date", DateColumnWidth)
        m_hHeader = SendMessageLong(m_hListView, LVM_GETHEADER, 0, 0)
    
        '-- Initialize image lists
        Call pvInitializeImageListThumbnail
        Call pvInitializeImageListImageType(FormatsMask, MaskSeparator)
        Call pvInitializeImageListHeader
        
        '-- Initial/fixed features and colors
        Select Case ViewMode
            
            Case [tvThumbnail]
                m_ViewMode = [tvThumbnail]
                Call SendMessageLong(m_hListView, LVM_SETBKCOLOR, 0, GetSysColor(COLOR_BTNFACE))
                Call SendMessageLong(m_hListView, LVM_SETTEXTBKCOLOR, 0, GetSysColor(COLOR_BTNFACE))
                Call SendMessageLong(m_hListView, LVM_SETTEXTCOLOR, 0, GetSysColor(COLOR_BTNTEXT))
                Call pvSetExStyle(LVS_EX_INFOTIP, 0)
                Call pvSetExStyle(0, LVS_EX_LABELTIP)
            
            Case [tvDetails]
                m_ViewMode = [tvDetails]
                Call SendMessageLong(m_hListView, LVM_SETBKCOLOR, 0, GetSysColor(COLOR_WINDOW))
                Call SendMessageLong(m_hListView, LVM_SETTEXTBKCOLOR, 0, GetSysColor(COLOR_WINDOW))
                Call SendMessageLong(m_hListView, LVM_SETTEXTCOLOR, 0, GetSysColor(COLOR_WINDOWTEXT))
                Call pvSetExStyle(0, LVS_EX_INFOTIP)
                Call pvSetExStyle(LVS_EX_LABELTIP, 0)
        End Select
        Call pvSetExStyle(LVS_EX_GRIDLINES, 0)
        Call pvSetExStyle(LVS_EX_FULLROWSELECT, 0)
        
        '-- Show ListView
        Call ShowWindow(m_hListView, SW_SHOW)
        pvCreate = True
    End If
End Function

Private Function pvDestroyListView() As Boolean
    
    If (m_hListView) Then
        If (DestroyWindow(m_hListView)) Then
            pvDestroyListView = True
            m_hListView = 0
        End If
    End If
End Function

'//

Private Sub pvInitializeImageListThumbnail()
    
  Dim lcx As Long
  Dim lcy As Long
    
    '-- Create dummy large imagelist
    m_hILLarge = ImageList_Create(m_ThumbnailWidth + 2 * THUMBNAIL_BORDER, m_ThumbnailHeight + 2 * THUMBNAIL_BORDER, 0, 0, 0)
    Call SendMessageLong(m_hListView, LVM_SETIMAGELIST, LVSIL_NORMAL, m_hILLarge)
    
    '-- Set new icon spacing
    lcx = m_ThumbnailWidth + 2 * THUMBNAIL_BORDER + 10
    lcy = m_ThumbnailHeight + 2 * THUMBNAIL_BORDER + (m_lFontHeight + 4) + 20
    Call SendMessageLong(m_hListView, LVM_SETICONSPACING, 0, lcx + (lcy * &H10000))
End Sub

Private Function pvDestroyImageListThumbnail() As Boolean

    If (m_hILLarge) Then
        If (ImageList_Destroy(m_hILLarge)) Then
            pvDestroyImageListThumbnail = True
            m_hILLarge = 0
        End If
    End If
End Function

Private Sub pvInitializeImageListImageType(ByVal FormatsMask As String, ByVal MaskSeparator As String)

  Dim lc1    As Long
  Dim lc2    As Long
  Dim sExt() As String
    
    '-- Create small imagelist
    Call SendMessageLong(m_hListView, LVM_SETIMAGELIST, LVSIL_SMALL, pvGetSystemImageList(SHGFI_SMALLICON))
    
    '-- Get supported image types
    sExt() = Split(FormatsMask, MaskSeparator)
    
    '-- Get info for each type
    ReDim Preserve m_uImageTypeInfo(UBound(sExt()))
    For lc1 = 0 To UBound(sExt())
        If (Len(sExt(lc1))) Then
            With m_uImageTypeInfo(lc2)
                .Extension = sExt(lc1)
                .Name = pvGetImageTypeName("." & sExt(lc1))
                .IconIndex = pvGetImageTypeIconIndex("." & sExt(lc1))
            End With
            lc2 = lc2 + 1
        End If
    Next lc1
    ReDim Preserve m_uImageTypeInfo(lc2 - -(lc2 > 0))
End Sub

Private Sub pvInitializeImageListHeader()
        
    '-- Create header imagelist
    m_hILHeader = ImageList_Create(16, 16, ILD_MASK Or ILD_TRANSPARENT, 0, 0)
    Call SendMessageLong(m_hHeader, HDM_SETIMAGELIST, 0, m_hILHeader)
    
    Call ImageList_AddMasked(m_hILHeader, LoadResPicture("IMAGELIST_HEADER", vbResBitmap), vbMagenta)
End Sub

Private Function pvDestroyImageListHeader() As Boolean

    If (m_hILHeader) Then
        If (ImageList_Destroy(m_hILHeader)) Then
            pvDestroyImageListHeader = True
            m_hILHeader = 0
        End If
    End If
End Function

Private Sub pvSetFont(oFont As StdFont)

  Dim uLF   As LOGFONT
  Dim lChar As Long
        
    '-- Create logic font from standart font
    With uLF
         
        For lChar = 1 To Len(oFont.Name)
            .lfFaceName(lChar - 1) = CByte(Asc(Mid$(oFont.Name, lChar, 1)))
        Next lChar
        
        .lfHeight = -MulDiv(oFont.Size, GetDeviceCaps(UserControl.hDC, LOGPIXELSY), 72)
        .lfItalic = oFont.Italic
        .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
        .lfUnderline = oFont.Underline
        .lfStrikeOut = oFont.Strikethrough
        .lfCharSet = oFont.Charset
        
        m_lFontHeight = -.lfHeight
    End With
    m_hFont = CreateFontIndirect(uLF)
    
    Call SendMessageLong(m_hListView, WM_SETFONT, m_hFont, 0)
End Sub

Private Function pvDestroyFont() As Boolean

    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            pvDestroyFont = True
            m_hFont = 0
        End If
    End If
End Function

'//

Private Function pvGetSystemImageList(ByVal uSize As Long) As Long

  Dim uSHFI As SHFILEINFO
    
    pvGetSystemImageList = SHGetFileInfo("C:\", 0, uSHFI, Len(uSHFI), SHGFI_SYSICONINDEX Or uSize)
End Function

Private Function pvGetImageTypeIconIndex(sFile As String) As Long
  
  Dim uSHFI As SHFILEINFO
    
    '-- Get system icon index
    If (SHGetFileInfo(sFile, 0, uSHFI, Len(uSHFI), SHGFI_SMALLICON Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)) Then
        pvGetImageTypeIconIndex = uSHFI.iIcon
    End If
End Function

Private Function pvGetImageTypeName(sFile As String) As String
  
  Dim uSHFI As SHFILEINFO
    
    '-- Get file type name
    If (SHGetFileInfo(sFile, 0, uSHFI, Len(uSHFI), SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES)) Then
        pvGetImageTypeName = pvStripNulls(uSHFI.szTypeName)
    End If
End Function

Private Function pvGetImageTypeInfo(ByVal sFile As String) As IMAGETYPE_INFO

  Dim lc   As Long
  Dim sExt As String
    
    sExt = Mid$(sFile, InStrRev(sFile, ".") + 1)
    
    For lc = 0 To UBound(m_uImageTypeInfo())
        
        If (sExt = m_uImageTypeInfo(lc).Extension) Then
            pvGetImageTypeInfo = m_uImageTypeInfo(lc)
            Exit For
        End If
    Next lc
End Function

'//

Private Function pvColumnAdd( _
                 ByVal lColumn As Long, _
                 ByVal sText As String, _
                 ByVal lWidth As Long, _
                 Optional ByVal lAlign As tvColumnAlignConstants = [tvLeft] _
                 ) As Boolean

  Dim uLVC As LVCOLUMN_lp
    
    With uLVC
        .pszText = StrPtr(StrConv(sText, vbFromUnicode))
        .cchTextMax = Len(sText)
        .cx = lWidth
        .fmt = lAlign
        .mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_FMT
    End With
    pvColumnAdd = (SendMessage(m_hListView, LVM_INSERTCOLUMN, lColumn, uLVC) > -1)
End Function

Private Function pvColumnAutosize( _
                 ByVal lColumn As Long, _
                 Optional ByVal AutosizeType As tvColumnAutosizeConstants = [tvContents] _
                 ) As Boolean

    If (m_hListView) Then
        
        pvColumnAutosize = CBool(SendMessageLong(m_hListView, LVM_SETCOLUMNWIDTH, lColumn, AutosizeType))
    End If
End Function

Private Function pvSubItemSet( _
                 ByVal Item As Long, _
                 ByVal SubItem As Long, _
                 ByVal Text As String _
                 ) As Boolean
   
  Dim uLVI As LVITEM_lp

    With uLVI
        .iItem = Item
        .iSubItem = SubItem
        .pszText = StrPtr(StrConv(Text, vbFromUnicode))
        .cchTextMax = Len(Text)
        .mask = LVIF_TEXT
    End With
    pvSubItemSet = CBool(SendMessage(m_hListView, LVM_SETITEM, 0, uLVI))
End Function

Private Function pvSort( _
                 Optional ByVal lColumnID As tvColumnIDConstants = [tvFileName], _
                 Optional ByVal bPreserveSortOrder As Boolean = False _
                 ) As Boolean
  
  Dim lCol As Long

    If (m_hListView) Then
    
        m_lColumnSortIndex = lColumnID
    
        For lCol = 0 To 3
            If (lCol <> lColumnID) Then
                m_eColumnSortOrder(lCol) = [tvDefault]
                Me.ColumnIcon(lCol) = -1
            End If
        Next lCol
        
        If (bPreserveSortOrder = False) Then
            If (m_eColumnSortOrder(lColumnID) = [tvAscending]) Then
                m_eColumnSortOrder(lColumnID) = [tvDescending]
                Me.ColumnIcon(lColumnID) = 1
              Else
                m_eColumnSortOrder(lColumnID) = [tvAscending]
                Me.ColumnIcon(lColumnID) = 0
            End If
        End If
        
        Select Case lColumnID
            Case 0: pvSort = mListViewSort.Sort(m_hListView, 0, m_eColumnSortOrder(0), [tvString])
            Case 1: pvSort = mListViewSort.Sort(m_hListView, 1, m_eColumnSortOrder(1), [tvString])
            Case 2: pvSort = mListViewSort.Sort(m_hListView, 2, m_eColumnSortOrder(2), [tvNumeric])
            Case 3: pvSort = mListViewSort.Sort(m_hListView, 3, m_eColumnSortOrder(3), [tvDate])
        End Select
    End If
End Function

'//

Private Sub pvUCCoords(uPoint As POINTAPI)

    Call GetCursorPos(uPoint)
    Call ScreenToClient(m_hListView, uPoint)
End Sub

Private Function pvItemHitTest() As Long

  Dim uLVHI As LVHITTESTINFO
   
    Call pvUCCoords(uLVHI.pt)
    pvItemHitTest = SendMessage(m_hListView, LVM_HITTEST, 0, uLVHI)
End Function

Private Function pvShiftState() As Integer

  Dim lS As Integer
   
    If (GetAsyncKeyState(vbKeyShift) < 0) Then
        lS = lS Or vbShiftMask
    End If
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then
        lS = lS Or vbAltMask
    End If
    If (GetAsyncKeyState(vbKeyControl) < 0) Then
        lS = lS Or vbCtrlMask
    End If
    pvShiftState = lS
End Function

Private Sub pvSetStyle(ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
    
    lS = GetWindowLong(m_hListView, GWL_STYLE)
    lS = (lS And Not lStyleNot) Or lStyle
    Call SetWindowLong(m_hListView, GWL_STYLE, lS)
    Call SetWindowPos(m_hListView, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

Private Sub pvSetExStyle(ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
   
    lS = SendMessageLong(m_hListView, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lS = lS And Not lStyleNot
    lS = lS Or lStyle
    Call SendMessageLong(m_hListView, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lS)
End Sub

Private Function pvStripNulls(ByVal sString As String) As String
    
  Dim lPos As Long

    lPos = InStr(sString, vbNullChar)
    
    If (lPos = 1) Then
        pvStripNulls = vbNullString
    ElseIf (lPos > 1) Then
        pvStripNulls = Left$(sString, lPos - 1)
        Exit Function
    End If
    
    pvStripNulls = sString
End Function



'========================================================================================
' Custom draw routines
'========================================================================================

Private Function pvCustomDrawThumbnail(ByVal lParam As Long) As Long
  
  Dim lItm        As Long
  Dim lPrm        As Long
  Dim lhDC        As Long
  Dim bSelected   As Boolean
  Dim bHasFocus   As Boolean
  Dim uNMLVCD     As NMLVCUSTOMDRAW
  Dim rcItem      As RECT2
  Dim rcThumbnail As RECT2
  Dim uTBIH       As BITMAPINFOHEADER
  
    On Error Resume Next
  
    Call CopyMemory(uNMLVCD, ByVal lParam, Len(uNMLVCD))
    
    Select Case uNMLVCD.nmcd.dwDrawStage
      
        Case CDDS_PREPAINT
        
            pvCustomDrawThumbnail = CDRF_NOTIFYITEMDRAW
        
        Case CDDS_ITEMPREPAINT
            
            lItm = uNMLVCD.nmcd.dwItemSpec
            
            With rcItem
                
                '-- Get icon rect.
                .x1 = LVIR_ICON
                Call SendMessage(m_hListView, LVM_GETITEMRECT, lItm, rcItem)
                
                '-- Prepare thumbnail rect.
                .x1 = .x1 + 2 * THUMBNAIL_BORDER
                .y1 = .y1 + THUMBNAIL_BORDER
                .x2 = .x1 + m_ThumbnailWidth + 2 * THUMBNAIL_BORDER
                .y2 = .y1 + m_ThumbnailHeight + 2 * THUMBNAIL_BORDER
            
                '-- Bug?: Listview seems to process all items...
                If (.y2 + m_lFontHeight > -4) Then
                        
                    '-- Extratc item paint info
                    lPrm = uNMLVCD.nmcd.lItemlParam
                    lhDC = uNMLVCD.nmcd.hDC
                    bSelected = (Me.ItemFindState(lItm - 1, [tvSelected]) = lItm)
                    bHasFocus = (uNMLVCD.nmcd.uItemState And CDIS_FOCUS)
                    
                    '-- Draw background
                    If (bSelected And bHasFocus) Then
                        Call FillRect(lhDC, rcItem, GetSysColorBrush(COLOR_HIGHLIGHT))
                      Else
                        Call FillRect(lhDC, rcItem, GetSysColorBrush(COLOR_BTNFACE))
                    End If
                    Call DrawEdge(lhDC, rcItem, BDR_RAISEDINNER, BF_RECT)
                    
                    '-- Get thumbnail info
                    Call CopyMemory(uTBIH, m_uThumbnailInfo(lPrm).TBIH(0), Len(uTBIH))
                    
                    '-- Valid thumbnail data ?
                    If (uTBIH.biWidth Or uTBIH.biHeight) Then
                            
                        '-- Paint thumbnail
                        rcThumbnail.x1 = .x1 + (m_ThumbnailWidth - uTBIH.biWidth) \ 2 + THUMBNAIL_BORDER
                        rcThumbnail.y1 = .y1 + (m_ThumbnailHeight - uTBIH.biHeight) \ 2 + THUMBNAIL_BORDER
                        rcThumbnail.x2 = rcThumbnail.x1 + uTBIH.biWidth
                        rcThumbnail.y2 = rcThumbnail.y1 + uTBIH.biHeight
                        Call StretchDIBits(lhDC, rcThumbnail.x1, rcThumbnail.y1, uTBIH.biWidth, uTBIH.biHeight, 0, 0, uTBIH.biWidth, uTBIH.biHeight, m_uThumbnailInfo(lPrm).Data(0), m_uThumbnailInfo(lPrm).TBIH(0), DIB_RGB_COLORS, vbSrcCopy)
                        
                        '-- Paint edge
                        Call InflateRect(rcThumbnail, 1, 1)
                        Call DrawEdge(lhDC, rcThumbnail, BDR_SUNKENOUTER, BF_RECT)
                      
                      Else
                        '-- No valid thumbnail data...
                        If (Err = 0) Then
                            rcThumbnail = rcItem
                            If (bSelected) Then
                                If (bHasFocus) Then
                                    Call SetTextColor(lhDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
                                  Else
                                    Call SetTextColor(lhDC, GetSysColor(COLOR_BTNTEXT))
                                End If
                              Else
                                Call SetTextColor(lhDC, GetSysColor(COLOR_BTNTEXT))
                            End If
                            Call DrawText(lhDC, "Error!", -1, rcThumbnail, DT_custom2)
                        End If
                    End If
                    
                    '-- Prepare label background rect.
                    .y1 = .y2 + 2
                    .y2 = .y1 + m_lFontHeight + 4
                    
                    '-- Paint background
                    If (bSelected) Then
                        If (bHasFocus) Then
                            Call FillRect(lhDC, rcItem, GetSysColorBrush(COLOR_HIGHLIGHT))
                          Else
                            Call FillRect(lhDC, rcItem, GetSysColorBrush(COLOR_BTNFACE))
                        End If
                      Else
                        Call FillRect(lhDC, rcItem, GetSysColorBrush(COLOR_WINDOW))
                    End If
                    If (bHasFocus) Then
                        Call SetTextColor(lhDC, 0)
                        Call DrawFocusRect(lhDC, rcItem)
                      Else
                        Call DrawEdge(lhDC, rcItem, BDR_SUNKENOUTER, BF_RECT)
                    End If
                    
                    '-- Prepare label rect.
                    .x1 = .x1 + 2
                    .x2 = .x2 - 2
                    
                    '-- Paint text
                    If (bSelected) Then
                        If (bHasFocus) Then
                            Call SetTextColor(lhDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
                          Else
                            Call SetTextColor(lhDC, GetSysColor(COLOR_BTNTEXT))
                        End If
                      Else
                        Call SetTextColor(lhDC, GetSysColor(COLOR_WINDOWTEXT))
                    End If
                    Call DrawText(lhDC, Me.ItemText(lItm, 0), -1, rcItem, DT_custom1)
                End If
            End With
            
            '-- Skip default paints
            pvCustomDrawThumbnail = CDRF_SKIPDEFAULT
    End Select
    
    On Error GoTo 0
End Function



'========================================================================================
' OLEInPlaceActiveObject interface
'========================================================================================

Private Sub pvSetIPAO()

  Dim pOleObject          As IOleObject
  Dim pOleInPlaceSite     As IOleInPlaceSite
  Dim pOleInPlaceFrame    As IOleInPlaceFrame
  Dim pOleInPlaceUIWindow As IOleInPlaceUIWindow
  Dim rcPos               As RECT2
  Dim rcClip              As RECT2
  Dim uFrameInfo          As OLEINPLACEFRAMEINFO
       
    On Error Resume Next
    
    Set pOleObject = Me
    Set pOleInPlaceSite = pOleObject.GetClientSite
    
    If (Not pOleInPlaceSite Is Nothing) Then
        Call pOleInPlaceSite.GetWindowContext(pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo))
        If (Not pOleInPlaceFrame Is Nothing) Then
            Call pOleInPlaceFrame.SetActiveObject(m_uIPAO.ThisPointer, vbNullString)
        End If
        If (Not pOleInPlaceUIWindow Is Nothing) Then '-- And Not m_bMouseActivate
            Call pOleInPlaceUIWindow.SetActiveObject(m_uIPAO.ThisPointer, vbNullString)
          Else
            Call pOleObject.DoVerb(OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, UserControl.hWnd, VarPtr(rcPos))
        End If
    End If
    
    On Error GoTo 0
End Sub

Friend Function frTranslateAccel(pMsg As MSG) As Boolean
    
  Dim pOleObject      As IOleObject
  Dim pOleControlSite As IOleControlSite
  Dim hEdit           As Long
  
    On Error Resume Next
    
    Select Case pMsg.message
    
        Case WM_KEYDOWN, WM_KEYUP
        
            Select Case pMsg.wParam
            
                Case vbKeyTab
                    
                    If (pvShiftState() And vbCtrlMask) Then
                        Set pOleObject = Me
                        Set pOleControlSite = pOleObject.GetClientSite
                        If (Not pOleControlSite Is Nothing) Then
                            Call pOleControlSite.TranslateAccelerator(VarPtr(pMsg), pvShiftState() And vbShiftMask)
                        End If
                    End If
                    frTranslateAccel = False
                    
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                     
                    Call SendMessageLong(m_hListView, pMsg.message, pMsg.wParam, pMsg.lParam)
                    frTranslateAccel = True
            End Select
    End Select
    
    On Error GoTo 0
End Function



'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be added to the callback table
'   uMsg     - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
''Delete a message from the table of those that will invoke a callback.
''Parameters:
''   lng_hWnd - The handle of the window for which the uMsg is to be removed from the callback table
''   uMsg     - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
''   When     - Whether the msg is to be removed from the before, after or both callback tables
'
'    With sc_aSubData(zIdx(lng_hWnd))
'        If (When And eMsgWhen.MSG_BEFORE) Then
'            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
'        End If
'        If (When And eMsgWhen.MSG_AFTER) Then
'            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
'        End If
'    End With
'End Sub

Private Function Subclass_InIDE() As Boolean
'Return whether we're running in the IDE.
    
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Start subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to be subclassed
'Returns;
'   The sc_aSubData() index

  Dim i                        As Long                       'Loop index
  Dim j                        As Long                       'Loop index
  Dim nSubIdx                  As Long                       'Subclass data index
  Dim sSubCode                 As String                     'Subclass code string
  
  Const PUB_CLASSES            As Long = 0                   'The number of UserControl public classes
  Const GMEM_FIXED             As Long = 0                   'Fixed memory GlobalAlloc flag
  Const PAGE_EXECUTE_READWRITE As Long = &H40&               'Allow memory to execute without violating XP SP2 Data Execution Prevention
  Const PATCH_01               As Long = 18                  'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02               As Long = 68                  'Address of the previous WndProc
  Const PATCH_03               As Long = 78                  'Relative address of SetWindowsLong
  Const PATCH_06               As Long = 116                 'Address of the previous WndProc
  Const PATCH_07               As Long = 121                 'Relative address of CallWindowProc
  Const PATCH_0A               As Long = 186                 'Address of the owner object
  Const FUNC_CWP               As String = "CallWindowProcA" 'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM               As String = "EbMode"          'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL               As String = "SetWindowLongA"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER               As String = "user32"          'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5               As String = "vba5"            'Location of the EbMode function if running VB5
  Const MOD_VBA6               As String = "vba6"            'Location of the EbMode function if running VB6

    'If it's the first time through here..
    If (sc_aBuf(1) = 0) Then

        'Build the hex pair subclass string
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
                   "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
                   "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
                   "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90" & _
                   Hex$(&HA4 + (PUB_CLASSES * 12)) & "070000C3"
    
        'Convert the string from hex pairs to bytes and store in the machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                       'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                      'Next pair of hex characters
    
        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(16) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_aBuf(17) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found?
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If
    
        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                  'Patch the address of this object instance into the static machine code buffer
    
        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                   'Get the address of the CallWindowsProc function
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                   'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                     'Create the first sc_aSubData element
    
      Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then                                                    'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                   'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                  'Create a new sc_aSubData element
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                             'Allocate memory for the machine code WndProc
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i) 'Mark memory as executable
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                 'Copy the machine code from the static byte array to the code array in sc_aSubData
    
        .hWnd = lng_hWnd                                                          'Store the hWnd
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                'Set our WndProc in place
    
        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                           'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                           'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                              'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                           'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                              'Patch the relative address of the CallWindowProc api function
    End With
End Function

Private Sub Subclass_StopAll()
'Stop all subclassing
  
  Dim i As Long
  
    i = UBound(sc_aSubData())                                                     'Get the upper bound of the subclass data array
    Do While i >= 0                                                               'Iterate through each element
        With sc_aSubData(i)
            If (.hWnd <> 0) Then                                                  'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)                                         'Subclass_Stop
            End If
        End With
    
        i = i - 1                                                                 'Next element
    Loop
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Stop subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to stop being subclassed
  
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                       'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                    'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                    'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                'Release the machine code memory
        .hWnd = 0                                                                 'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                             'Clear the before table
        .nMsgCntA = 0                                                             'Clear the after table
        Erase .aMsgTblB                                                           'Erase the before table
        Erase .aMsgTblA                                                           'Erase the after table
    End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'Worker sub for Subclass_AddMsg
  
  Dim nEntry  As Long                                                             'Message table entry index
  Dim nOff1   As Long                                                             'Machine code buffer offset 1
  Dim nOff2   As Long                                                             'Machine code buffer offset 2
  
    If (uMsg = ALL_MESSAGES) Then                                                 'If all messages
        nMsgCnt = ALL_MESSAGES                                                    'Indicates that all messages will callback
      Else                                                                        'Else a specific message number
        Do While nEntry < nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
        
            If (aMsgTbl(nEntry) = 0) Then                                         'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                            'Re-use this entry
                Exit Sub                                                          'Bail
            ElseIf (aMsgTbl(nEntry) = uMsg) Then                                  'The msg is already in the table!
                Exit Sub                                                          'Bail
            End If
        Loop                                                                      'Next entry

        nMsgCnt = nMsgCnt + 1                                                     'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                              'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                   'Store the message number in the table
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then                                          'If before
        nOff1 = PATCH_04                                                          'Offset to the Before table
        nOff2 = PATCH_05                                                          'Offset to the Before table entry count
      Else                                                                        'Else after
        nOff1 = PATCH_08                                                          'Offset to the After table
        nOff2 = PATCH_09                                                          'Offset to the After table entry count
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                          'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                         'Patch the appropriate table entry count
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the memory address of the passed function in the passed dll
    
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                        'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
''Worker sub for Subclass_DelMsg
'
'  Dim nEntry As Long
'
'    If (uMsg = ALL_MESSAGES) Then                                                 'If deleting all messages
'        nMsgCnt = 0                                                               'Message count is now zero
'        If When = eMsgWhen.MSG_BEFORE Then                                        'If before
'            nEntry = PATCH_05                                                     'Patch the before table message count location
'          Else                                                                    'Else after
'            nEntry = PATCH_09                                                     'Patch the after table message count location
'        End If
'        Call zPatchVal(nAddr, nEntry, 0)                                          'Patch the table message count to zero
'      Else                                                                        'Else deleteting a specific message
'        Do While nEntry < nMsgCnt                                                 'For each table entry
'            nEntry = nEntry + 1
'            If (aMsgTbl(nEntry) = uMsg) Then                                      'If this entry is the message we wish to delete
'                aMsgTbl(nEntry) = 0                                               'Mark the table slot as available
'                Exit Do                                                           'Bail
'            End If
'        Loop                                                                      'Next entry
'    End If
'End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the sc_aSubData() array index of the passed hWnd
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                            'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If (.hWnd = lng_hWnd) Then                                            'If the hWnd of this element is the one we're looking for
                If (Not bAdd) Then                                                'If we're searching not adding
                    Exit Function                                                 'Found
                End If
            ElseIf (.hWnd = 0) Then                                               'If this an element marked for reuse.
                If (bAdd) Then                                                    'If we're adding
                    Exit Function                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                           'Decrement the index
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False                                                        'hWnd not found, programmer error
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
'Patch the machine code buffer at the indicated offset with the relative address to the target address.
    
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
'Patch the machine code buffer at the indicated offset with the passed value
    
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
'Worker function for Subclass_InIDE
    
    zSetTrue = True
    bValue = True
End Function
