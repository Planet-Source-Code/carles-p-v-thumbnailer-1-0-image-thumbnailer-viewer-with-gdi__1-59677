VERSION 5.00
Begin VB.UserControl ucFolderView 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
End
Attribute VB_Name = "ucFolderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucFolderView.ctl (basic: standard folders)
' Author:        Carles P.V. (*)
' Dependencies:
' Last revision: 2004.11.30
' Version:       1.0.4
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
'   * 1.0.1: - LockWindowUpdate -> WM_SETREDRAW.
'   * 1.0.2: - File attributes directly from WIN32_FIND_DATA struct.
'   * 1.0.3: - WM_SETREDRAW not used
'            - Added PathIsRoot() property.
'            - Added PathParentIsRoot() property.
'            - Added PathIsValid() property.
'   * 1.0.4: - Added ChangeBefore() event -> allows cancel
'            - Changed Change() to ChangeAfter() event.
'========================================================================================

Option Explicit
Option Compare Text

'== Folder/File/Resource

Private Const MAX_PATH                   As Long = 260
Private Const INVALID_HANDLE_VALUE       As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY   As Long = &H10
Private Const FILE_ATTRIBUTE_READONLY    As Long = &H1
Private Const FILE_ATTRIBUTE_RODIRECTORY As Long = FILE_ATTRIBUTE_DIRECTORY + FILE_ATTRIBUTE_READONLY

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long

'//

Private Const CSIDL_DRIVES                  As Long = &H11

Private Const SHGFI_SMALLICON               As Long = &H1
Private Const SHGFI_PIDL                    As Long = &H8
Private Const SHGFI_DISPLAYNAME             As Long = &H200
Private Const SHGFI_SYSICONINDEX            As Long = &H4000

Private Const SI_FOLDER_CLOSED              As Long = &H3

Private Type SHFILEINFO
    hIcon         As Long
    iIcon         As Long
    dwAttributes  As Long
    szDisplayName As String * MAX_PATH
    szTypeName    As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef PIDL As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

'== TreeView

Private Const WC_TREEVIEW  As String = "SysTreeView32"

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

Private Type TVITEM
    mask           As Long
    hItem          As Long
    State          As Long
    stateMask      As Long
    pszText        As Long
    cchTextMax     As Long
    iImage         As Long
    iSelectedImage As Long
    cChildren      As Long
    lParam         As Long
End Type

Private Type TVINSERTSTRUCT
    hParent      As Long
    hInsertAfter As Long
    Item         As TVITEM
End Type

Private Type NMHDR
    hWndFrom As Long
    idfrom   As Long
    code     As Long
End Type

Private Type NMTREEVIEW
    hdr     As NMHDR
    action  As Long
    itemOld As TVITEM
    itemNew As TVITEM
    ptDrag  As POINTAPI
End Type

Private Const GWL_STYLE           As Long = (-16)
Private Const GWL_EXSTYLE         As Long = (-20)

Private Const WS_TABSTOP          As Long = &H10000
Private Const WS_CHILD            As Long = &H40000000
Private Const WS_EX_CLIENTEDGE    As Long = &H200

Private Const WM_SIZE             As Long = &H5
Private Const WM_SETFOCUS         As Long = &H7
Private Const WM_SETFONT          As Long = &H30
Private Const WM_MOUSEACTIVATE    As Long = &H21
Private Const WM_NOTIFY           As Long = &H4E
Private Const WM_KEYDOWN          As Long = &H100
Private Const WM_KEYUP            As Long = &H101

Private Const NM_FIRST            As Long = 0
Private Const NM_SETFOCUS         As Long = (NM_FIRST - 7)

'//

Private Const TVS_HASBUTTONS      As Long = &H1
Private Const TVS_HASLINES        As Long = &H2
Private Const TVS_DISABLEDRAGDROP As Long = &H10
Private Const TVS_SHOWSELALWAYS   As Long = &H20
Private Const TVS_SINGLEEXPAND    As Long = &H400
Private Const TVS_TRACKSELECT     As Long = &H200
Private Const TVSIL_NORMAL        As Long = &H0

Private Const TVE_EXPAND          As Long = &H2

Private Const TVGN_ROOT           As Long = &H0
Private Const TVGN_PARENT         As Long = &H3
Private Const TVGN_CARET          As Long = &H9

Private Const TVI_ROOT            As Long = &HFFFF0000

Private Const TVIF_TEXT           As Long = &H1
Private Const TVIF_IMAGE          As Long = &H2
Private Const TVIF_PARAM          As Long = &H4
Private Const TVIF_STATE          As Long = &H8
Private Const TVIF_HANDLE         As Long = &H10
Private Const TVIF_SELECTEDIMAGE  As Long = &H20
Private Const TVIF_CHILDREN       As Long = &H40

Private Const TVIS_DROPHILITED    As Long = &H8
Private Const TVIS_EXPANDED       As Long = &H20
Private Const TVIS_EXPANDEDONCE   As Long = &H40

Private Const TV_FIRST            As Long = &H1100
Private Const TVM_INSERTITEM      As Long = (TV_FIRST + 0)
Private Const TVM_EXPAND          As Long = (TV_FIRST + 2)
Private Const TVM_SETIMAGELIST    As Long = (TV_FIRST + 9)
Private Const TVM_GETNEXTITEM     As Long = (TV_FIRST + 10)
Private Const TVM_SELECTITEM      As Long = (TV_FIRST + 11)
Private Const TVM_GETITEM         As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM         As Long = (TV_FIRST + 13)
Private Const TVM_SORTCHILDREN    As Long = (TV_FIRST + 19)
Private Const TVM_ENSUREVISIBLE   As Long = (TV_FIRST + 20)
Private Const TVM_SETBKCOLOR      As Long = (TV_FIRST + 29)
Private Const TVM_SETTEXTCOLOR    As Long = (TV_FIRST + 30)

Private Const TVN_FIRST           As Long = -400
Private Const TVN_SELCHANGING     As Long = (TVN_FIRST - 1)
Private Const TVN_SELCHANGED      As Long = (TVN_FIRST - 2)
Private Const TVN_ITEMEXPANDING   As Long = (TVN_FIRST - 5)

'== Misc.

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

Private Const SW_SHOW                As Long = 5

Private Const SWP_NOMOVE             As Long = &H2
Private Const SWP_NOSIZE             As Long = &H1
Private Const SWP_NOOWNERZORDER      As Long = &H200
Private Const SWP_NOZORDER           As Long = &H4
Private Const SWP_FRAMECHANGED       As Long = &H20

Private Const COLOR_WINDOW           As Long = 5
Private Const COLOR_WINDOWTEXT       As Long = 8

Private Declare Function FileIconInit Lib "shell32" Alias "#660" (ByVal cmd As Boolean) As Boolean

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'//

'-- Public enums.:

Public Enum fvBorderStyleConstants
    [fvNone] = 0
    [fvFixedSingle]
End Enum

Public Enum fvGoConstants
    [fvGoUp] = 0
End Enum

'-- Property variables:

Private WithEvents m_oFont As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_HasButtons       As Boolean
Private m_HasLines         As Boolean
Private m_HideSelection    As Boolean
Private m_SingleExpand     As Boolean
Private m_TrackSelect      As Boolean

'-- Private enums.:

Private Enum eStateConstants
    [eStateDropHilited] = TVIS_DROPHILITED
    [eStateExpanded] = TVIS_EXPANDED
    [eStateExpandedOnce] = TVIS_EXPANDEDONCE
End Enum

'-- Private constants:

'-- Private types:

Private Type NODEDATA
    hNode As Long
    sPath As String
End Type

'-- Private variables:

Private m_bInitialized     As Boolean
Private m_hTreeView        As Long
Private m_hFont            As Long
Private uNodeData()        As NODEDATA
Private m_uIPAO            As IPAOHookStructFolderView

'-- Event declarations:

Public Event ChangeAfter(ByVal OldPath As String)
Public Event ChangeBefore(ByVal NewPath As String, Cancel As Boolean)



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
  
  Dim uNMH    As NMHDR
  Dim uNMTV   As NMTREEVIEW
  Dim bCancel As Boolean
  
    Select Case lng_hWnd
    
        Case UserControl.hWnd
            
            Select Case uMsg
            
                Case WM_SIZE
                    
                    Call MoveWindow(m_hTreeView, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1)
                
                Case WM_SETFOCUS
                
                    Call SetFocus(m_hTreeView)
                
                Case WM_MOUSEACTIVATE
                    
                    Call pvSetIPAO
                                    
                Case WM_NOTIFY
                    
                    Call CopyMemory(uNMH, ByVal lParam, Len(uNMH))
                    
                    If (uNMH.hWndFrom = m_hTreeView) Then
                    
                        Select Case uNMH.code
                        
                            Case NM_SETFOCUS
                            
                                Call pvSetIPAO
                    
                            Case TVN_ITEMEXPANDING
                            
                                Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                                
                                If (uNMTV.action = TVE_EXPAND) Then
                                    If ((uNMTV.itemNew.State And TVIS_EXPANDEDONCE) = 0) Then
                                        Call pvAddFolders(uNMTV.itemNew.hItem)
                                    End If
                                End If
                                
                            Case TVN_SELCHANGING
                                
                                Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                                
                                If (pvTVGetRoot() <> uNMTV.itemNew.hItem) Then
                                    
                                    Call pvTVSetState(uNMTV.itemOld.hItem, [eStateDropHilited], False)
                                    Call pvTVSetState(uNMTV.itemNew.hItem, [eStateDropHilited], True)
                                    
                                    DoEvents
                                    RaiseEvent ChangeBefore(uNodeData(uNMTV.itemNew.lParam).sPath, bCancel)
                                    
                                    Call pvTVSetState(uNMTV.itemNew.hItem, [eStateDropHilited], False)
                                    
                                    If (bCancel) Then
                                        Call pvTVEnsureVisible(uNMTV.itemOld.hItem)
                                        bHandled = True
                                        lReturn = 1
                                    End If
                                End If
                            
                            Case TVN_SELCHANGED
                                
                                Call CopyMemory(uNMTV, ByVal lParam, Len(uNMTV))
                                
                                If (pvTVGetRoot() <> uNMTV.itemNew.hItem) Then
                                    DoEvents
                                    RaiseEvent ChangeAfter(uNodeData(uNMTV.itemOld.lParam).sPath)
                                End If
                        End Select
                    End If
            End Select
    End Select
End Sub



'========================================================================================
' Usercontrol
'========================================================================================

Private Sub UserControl_Initialize()

    On Error Resume Next
       Call FileIconInit(cmd:=True) 'NT only
    On Error GoTo 0

    Set m_oFont = New StdFont
    Let m_HasButtons = True
    Let m_HasLines = True
End Sub

Private Sub UserControl_Terminate()
  
  On Error GoTo errH
  
    If (m_hTreeView) Then
    
        Call Subclass_StopAll
        Call mIOIPAFolderView.TerminateIPAO(m_uIPAO)
        Call pvTVSetImageList(hImageList:=0)
        Call pvDestroyFont
        Call pvDestroyTreeView
        
        Erase uNodeData()
    End If
errH:
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function Initialize() As Boolean

    If (m_bInitialized = False) Then
    
        Initialize = pvCreate()
        
        If (m_hTreeView) Then

            '-- Subclass UserControl (parent) and TreeView (child)
            Call Subclass_Start(UserControl.hWnd)
            Call Subclass_AddMsg(UserControl.hWnd, WM_SETFOCUS)
            Call Subclass_AddMsg(UserControl.hWnd, WM_SIZE)
            Call Subclass_AddMsg(UserControl.hWnd, WM_MOUSEACTIVATE)
            Call Subclass_AddMsg(UserControl.hWnd, WM_NOTIFY)
            
            '-- Initialize IOLEInPlaceActiveObject
            Call mIOIPAFolderView.InitIPAO(m_uIPAO, Me)
            
            m_bInitialized = True
        End If
    End If
End Function

Public Function Go(ByVal Where As fvGoConstants) As Boolean
    
  Dim hNode As Long
  
    Select Case Where
    
        Case [fvGoUp]
            
            hNode = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, pvTVGetSelected())
            If (hNode And pvTVGetRoot() <> hNode) Then
                Call pvTVSetSelected(hNode)
            End If
    End Select
End Function



'========================================================================================
' Properties
'========================================================================================

Public Property Get BorderStyle() As fvBorderStyleConstants
    If (m_hTreeView) Then
        BorderStyle = -((GetWindowLong(m_hTreeView, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE)
    End If
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As fvBorderStyleConstants)
    If (m_hTreeView) Then
        Select Case New_BorderStyle
            Case [fvNone]
                Call SetWindowLong(m_hTreeView, GWL_EXSTYLE, 0)
            Case [fvFixedSingle]
                Call SetWindowLong(m_hTreeView, GWL_EXSTYLE, WS_EX_CLIENTEDGE)
        End Select
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If (m_hTreeView) Then
        UserControl.Enabled = New_Enabled
        Call EnableWindow(m_hTreeView, New_Enabled)
    End If
End Property

Public Property Get Font() As StdFont
    Set Font = m_oFont
End Property
Public Property Set Font(ByVal New_Font As StdFont)

  Dim uLF   As LOGFONT
  Dim lChar As Long

    If (m_hTreeView) Then
    
         With m_oFont
             For lChar = 1 To Len(.Name)
                 uLF.lfFaceName(lChar - 1) = CByte(Asc(Mid$(.Name, lChar, 1)))
             Next lChar
             uLF.lfHeight = -MulDiv(.Size, GetDeviceCaps(UserControl.hDC, LOGPIXELSY), 72)
             uLF.lfItalic = .Italic
             uLF.lfWeight = IIf(.Bold, FW_BOLD, FW_NORMAL)
             uLF.lfUnderline = .Underline
             uLF.lfStrikeOut = .Strikethrough
             uLF.lfCharSet = .Charset
        End With
        Call pvDestroyFont: m_hFont = CreateFontIndirect(uLF)
        
        Call SendMessageLong(m_hTreeView, WM_SETFONT, m_hFont, True)
    End If
End Property
Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set Font = m_oFont
End Sub

Public Property Get HasButtons() As Boolean
    HasButtons = m_HasButtons
End Property
Public Property Let HasButtons(ByVal New_HasButtons As Boolean)
    If (m_hTreeView) Then
        m_HasButtons = New_HasButtons
        If (m_HasButtons) Then
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_HASBUTTONS, 0)
          Else
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_HASBUTTONS)
        End If
    End If
End Property

Public Property Get HasLines() As Boolean
    HasLines = m_HasLines
End Property
Public Property Let HasLines(ByVal New_HasLines As Boolean)
    If (m_hTreeView) Then
        m_HasLines = New_HasLines
        If (m_HasLines) Then
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_HASLINES, 0)
          Else
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_HASLINES)
        End If
    End If
End Property

Public Property Get HideSelection() As Boolean
    HideSelection = m_HideSelection
End Property
Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    If (m_hTreeView) Then
        m_HideSelection = New_HideSelection
        If (m_HideSelection) Then
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_SHOWSELALWAYS)
          Else
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_SHOWSELALWAYS, 0)
        End If
    End If
End Property

Public Property Get Path() As String
  
  Dim lParam As Long
    
    If (m_hTreeView) Then
        
        Call pvTVGetlParam(pvTVGetSelected(), lParam)
        Path = uNodeData(lParam).sPath
    End If
End Property
Public Property Let Path(ByVal New_Path As String)
    
  Dim lPos  As Long
  Dim hNode As Long
    
    If (m_hTreeView) Then
        
        '-- Back-slash ending path
        New_Path = New_Path & IIf(Right$(New_Path, 1) <> "\", "\", vbNullString)
        
        '-- Search/expand until last folder
        Do: lPos = InStr(lPos + 1, New_Path, "\")
            If (lPos) Then
                hNode = pvGetNodeFromPath(Left$(New_Path, lPos))
                If (hNode) Then
                    If (InStr(lPos + 1, New_Path, "\")) Then
                        Call pvTVExpand(hNode)
                    End If
                  Else
                    Exit Do
                End If
            End If
        Loop Until (lPos < 1)
        
        '-- Select and ensure visible
        Call pvTVSetSelected(hNode)
    End If
End Property

Public Property Get PathIsRoot() As Boolean
    If (m_hTreeView) Then
        PathIsRoot = (pvTVGetSelected() = pvTVGetRoot())
    End If
End Property

Public Property Get PathParentIsRoot() As Boolean
    If (m_hTreeView) Then
        PathParentIsRoot = (SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_PARENT, pvTVGetSelected()) = pvTVGetRoot())
    End If
End Property

Public Property Get PathIsValid(ByVal Path As String) As Boolean
  
  Dim sPath   As String
  Dim uWFD    As WIN32_FIND_DATA
  Dim hSearch As Long
    
    If (m_hTreeView) Then
        
        sPath = Path & "*.*" & vbNullChar
        
        hSearch = FindFirstFile(sPath, uWFD)
        PathIsValid = (hSearch <> INVALID_HANDLE_VALUE)
        
        Call FindClose(hSearch)
    End If
End Property

Public Property Get SingleExpand() As Boolean
    SingleExpand = m_SingleExpand
End Property
Public Property Let SingleExpand(ByVal New_SingleExpand As Boolean)
    If (m_hTreeView) Then
        m_SingleExpand = New_SingleExpand
        If (m_SingleExpand) Then
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_SINGLEEXPAND, 0)
          Else
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_SINGLEEXPAND)
        End If
    End If
End Property

Public Property Get TrackSelect() As Boolean
    TrackSelect = m_TrackSelect
End Property
Public Property Let TrackSelect(ByVal New_TrackSelect As Boolean)
    If (m_hTreeView) Then
        m_TrackSelect = New_TrackSelect
        If (m_TrackSelect) Then
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, TVS_TRACKSELECT, 0)
          Else
            Call pvSetWndStyle(m_hTreeView, GWL_STYLE, 0, TVS_TRACKSELECT)
        End If
    End If
End Property



'========================================================================================
' Private
'========================================================================================

Private Function pvCreate() As Boolean

  Dim lExStyle As Long
  Dim lTVStyle   As Long
    
    '-- Define window style
    lExStyle = WS_EX_CLIENTEDGE
    lTVStyle = WS_CHILD Or WS_TABSTOP Or TVS_HASBUTTONS Or TVS_HASLINES Or TVS_SHOWSELALWAYS Or TVS_DISABLEDRAGDROP
    
    '-- Create TreeView window
    m_hTreeView = CreateWindowEx(lExStyle, WC_TREEVIEW, vbNullString, lTVStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0)
    
    '-- Success [?]
    If (m_hTreeView) Then
        
        '-- Set TreeView image list (system)
        Call pvTVSetImageList(pvGetSystemImageList(SHGFI_SMALLICON))
        '-- System background and foreground colors
        Call SendMessageLong(m_hTreeView, TVM_SETBKCOLOR, 0, GetSysColor(COLOR_WINDOW))
        Call SendMessageLong(m_hTreeView, TVM_SETTEXTCOLOR, 0, GetSysColor(COLOR_WINDOWTEXT))
        '-- Initialize font
        Set m_oFont = Ambient.Font: Call m_oFont_FontChanged(vbNullString)
        
        '-- Add drives and expand root
        Call pvAddDrives
        Call pvTVExpand(pvTVGetRoot())
        
        '-- Show TreeView
        Call ShowWindow(m_hTreeView, SW_SHOW)
        pvCreate = True
   End If
End Function

Private Sub pvDestroyTreeView()
    
    If (m_hTreeView) Then
        If (DestroyWindow(m_hTreeView)) Then
            m_hTreeView = 0
        End If
    End If
End Sub

Private Sub pvDestroyFont()

    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            m_hFont = 0
        End If
    End If
End Sub

Private Sub pvSetWndStyle(ByVal hWnd As Long, ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
    
    lS = GetWindowLong(hWnd, lType)
    lS = (lS And Not lStyleNot) Or lStyle
    Call SetWindowLong(hWnd, lType, lS)
    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

'//

Private Sub pvAddDrives()

  Dim uSHFI          As SHFILEINFO
  Dim lPIDL          As Long
  Dim sBuffer        As String * MAX_PATH
  Dim lDrivesBitMask As Long
  Dim lMaxPwr        As Long
  Dim lPwr           As Long
  Dim hNodeRoot      As Long
  Dim hNode          As Long
  Dim sText          As String
    
    '-- Add root node ('My Computer') (PIDL)
  
    Call SHGetSpecialFolderLocation(0, CSIDL_DRIVES, lPIDL)
    Call SHGetFileInfo(ByVal lPIDL, 0, uSHFI, Len(uSHFI), SHGFI_PIDL Or SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX)
    Call CoTaskMemFree(lPIDL)
    hNodeRoot = pvTVAdd(, , uSHFI.szDisplayName, 0, uSHFI.iIcon, uSHFI.iIcon)
    
    Call SendMessageLong(m_hTreeView, TVM_SELECTITEM, TVGN_CARET, hNodeRoot)
    
    '-- Add drives
    
    lDrivesBitMask = GetLogicalDrives()

    If (lDrivesBitMask) Then
      
        lMaxPwr = Int(Log(lDrivesBitMask) / Log(2))
        ReDim Preserve uNodeData(lMaxPwr)

        For lPwr = 0 To lMaxPwr
            
            If (2 ^ lPwr And lDrivesBitMask) Then
            
                sText = Chr$(65 + lPwr) & ":\"
                
                Call SHGetFileInfo(sText, 0, uSHFI, Len(uSHFI), SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX)
                hNode = pvTVAdd(hNodeRoot, , uSHFI.szDisplayName, lPwr, uSHFI.iIcon, uSHFI.iIcon, bForcePlusButton:=True)
                
                With uNodeData(lPwr)
                    .hNode = hNode
                    .sPath = sText
                End With
            End If
        Next lPwr
    End If
End Sub

Private Sub pvAddFolders(ByVal hNode As Long)
  
  Dim uSHFI       As SHFILEINFO
  Dim uWFD        As WIN32_FIND_DATA
  Dim lParam      As Long
  Dim lNextIdx    As Long
  Dim sPath       As String
  Dim sFolderName As String
  Dim lFolders    As Long
  Dim hSearch     As Long
  Dim hNext       As Long
  Dim lRet        As Long
  
    If (pvTVGetRoot() <> hNode) Then
    
        Screen.MousePointer = vbHourglass
        
        '-- Get node full path
        Call pvTVGetlParam(hNode, lParam)
        sPath = uNodeData(lParam).sPath
        
        '-- Start searching
        hNext = 1
        hSearch = FindFirstFile(sPath & "*." & vbNullChar, uWFD)
        
        If (hSearch <> INVALID_HANDLE_VALUE) Then
            
            Do While hNext
                
                '-- Get file [folder] name
                sFolderName = pvStripNulls(uWFD.cFileName)
                If (sFolderName <> "." And sFolderName <> "..") Then
                    
                    '-- Only standard folders
                    If (uWFD.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Or _
                        uWFD.dwFileAttributes = FILE_ATTRIBUTE_RODIRECTORY) Then
                        
                        '-- Get info (name and image list index)
                        lNextIdx = UBound(uNodeData()) + 1
                        ReDim Preserve uNodeData(lNextIdx)
                        lRet = SHGetFileInfo(sPath & sFolderName & "\", 0, uSHFI, Len(uSHFI), SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX)
                        
                        '-- Add node
                        lRet = pvTVAdd(hNode, , uSHFI.szDisplayName, lNextIdx, uSHFI.iIcon, uSHFI.iIcon + -(uSHFI.iIcon = SI_FOLDER_CLOSED), -pvHasSubFolders(sPath & sFolderName & "\"))
                        With uNodeData(lNextIdx)
                            .sPath = sPath & sFolderName & "\"
                            .hNode = lRet
                        End With
                        
                        '-- Count folders (-> hide 'plus button')
                        lFolders = lFolders + 1
                    End If
                End If
                hNext = FindNextFile(hSearch, uWFD)
            Loop
            hNext = FindClose(hSearch)
            
            '-- Sort added folders and ensure visible parent one
            Call pvTVSortChildren(hNode)
            Call pvTVEnsureVisible(hNode)
        End If
        
        '-- Hide 'plus button' ?
        If (lFolders = 0) Then
            Call pvTVSetcChildren(hNode, 0)
        End If
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Function pvHasSubFolders(ByVal sPath As String) As Boolean

  Dim uWFD        As WIN32_FIND_DATA
  Dim sFolderName As String
  Dim hSearch     As Long
  Dim hNext       As Long
    
    '-- Start searching
    hNext = 1
    hSearch = FindFirstFile(sPath & "*." & vbNullChar, uWFD)
    
    If (hSearch <> INVALID_HANDLE_VALUE) Then
        
        Do While hNext
        
            sFolderName = pvStripNulls(uWFD.cFileName)
            If (sFolderName <> "." And sFolderName <> "..") Then
                
                If (uWFD.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Or _
                    uWFD.dwFileAttributes = FILE_ATTRIBUTE_RODIRECTORY) Then
                    
                    '-- Found one: enough
                    pvHasSubFolders = True
                    Exit Do
                End If
            End If
            hNext = FindNextFile(hSearch, uWFD)
        Loop
        hNext = FindClose(hSearch)
    End If
End Function

Private Function pvGetNodeFromPath(ByVal sPath As String) As Long
    
  Dim lc As Long
  
    For lc = 0 To UBound(uNodeData())
        If (sPath = uNodeData(lc).sPath) Then
            pvGetNodeFromPath = uNodeData(lc).hNode
            Exit For
        End If
    Next lc
End Function

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

'//

Private Function pvTVAdd( _
                 Optional ByVal hParent As Long = -1, _
                 Optional ByVal hInsertAfter As Long = -1, _
                 Optional ByVal sText As String = vbNullString, _
                 Optional ByVal lParam As Long = -1, _
                 Optional ByVal lImage As Long = -1, _
                 Optional ByVal lSelectedImage As Long = -1, _
                 Optional ByVal bForcePlusButton As Boolean = False _
                 ) As Long
                 
  Dim uTVIS As TVINSERTSTRUCT
  Dim uTVI  As TVITEM
  Dim lMask As Long
    
    With uTVI
        If (LenB(sText) > 0) Then
            .mask = TVIF_TEXT
            .cchTextMax = Len(sText)
            .pszText = StrPtr(StrConv(sText, vbFromUnicode))
        End If
        If (lParam > 0) Then
            .mask = .mask Or TVIF_PARAM
            .lParam = lParam
        End If
        If (lImage > 0) Then
            .mask = .mask Or TVIF_IMAGE
            .iImage = lImage
        End If
        If (lSelectedImage > 0) Then
            .mask = .mask Or TVIF_SELECTEDIMAGE
            .iSelectedImage = lSelectedImage
        End If
        If (bForcePlusButton) Then
            .mask = .mask Or TVIF_CHILDREN
            .cChildren = 1 'I_CHILDRENCALLBACK
        End If
    End With
    
    With uTVIS
        If (hParent > 0) Then
            .hParent = hParent
          Else
            .hParent = TVI_ROOT
        End If
        If (hInsertAfter > 0) Then
            .hInsertAfter = hInsertAfter
          Else
            .hInsertAfter = TVI_ROOT
        End If
        .Item = uTVI
    End With
    
    pvTVAdd = SendMessage(m_hTreeView, TVM_INSERTITEM, 0, uTVIS)
End Function

Private Function pvGetSystemImageList(ByVal uSize As Long) As Long
Dim uSHFI As SHFILEINFO
    pvGetSystemImageList = SHGetFileInfo("C:\", 0, uSHFI, Len(uSHFI), SHGFI_SYSICONINDEX Or uSize)
End Function

Private Function pvTVSetImageList(ByVal hImageList As Long) As Long
    pvTVSetImageList = SendMessageLong(m_hTreeView, TVM_SETIMAGELIST, TVSIL_NORMAL, hImageList)
End Function

Private Function pvTVGetRoot() As Long
    pvTVGetRoot = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_ROOT, 0)
End Function

Private Function pvTVExpand(ByVal hNode As Long) As Boolean
    pvTVExpand = SendMessageLong(m_hTreeView, TVM_EXPAND, TVE_EXPAND, hNode)
End Function

Private Function pvTVEnsureVisible(ByVal hNode As Long) As Long
    pvTVEnsureVisible = SendMessageLong(m_hTreeView, TVM_ENSUREVISIBLE, 0, hNode)
End Function

Private Function pvTVSortChildren(ByVal hNode As Long, Optional ByVal fRecurse As Long = 1) As Boolean
    pvTVSortChildren = SendMessageLong(m_hTreeView, TVM_SORTCHILDREN, fRecurse, hNode)
End Function

Private Function pvTVSetSelected(ByVal hNode As Long) As Boolean
    pvTVSetSelected = SendMessageLong(m_hTreeView, TVM_SELECTITEM, TVGN_CARET, hNode)
End Function
Private Function pvTVGetSelected() As Long
    pvTVGetSelected = SendMessageLong(m_hTreeView, TVM_GETNEXTITEM, TVGN_CARET, 0)
End Function

Private Function pvTVSetcChildren(ByVal hNode As Long, ByVal cChildren As Long) As Boolean

  Dim uTVI As TVITEM
    
    With uTVI
        .hItem = hNode
        .mask = TVIF_CHILDREN
        .cChildren = cChildren
    End With
    
    pvTVSetcChildren = SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
End Function
Private Function pvTVGetcChildren(ByVal hNode As Long, cChildren As Long) As Boolean
  
  Dim uTVI As TVITEM
    
    With uTVI
        .hItem = hNode
        .mask = TVIF_CHILDREN
    End With
    
    pvTVGetcChildren = SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)
    cChildren = uTVI.cChildren
End Function

Private Function pvTVSetlParam(ByVal hNode As Long, ByVal lParam As Long) As Boolean

  Dim uTVI As TVITEM
    
    With uTVI
        .hItem = hNode
        .mask = TVIF_PARAM
        .lParam = lParam
    End With
    
    pvTVSetlParam = SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
End Function
Private Function pvTVGetlParam(ByVal hNode As Long, lParam As Long) As Boolean
  
  Dim uTVI As TVITEM
    
    With uTVI
        .hItem = hNode
        .mask = TVIF_PARAM
    End With
    
    pvTVGetlParam = SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)
    lParam = uTVI.lParam
End Function

Private Sub pvTVSetState(hNode As Long, lState As eStateConstants, fAdd As Boolean)
  
  Dim uTVI As TVITEM

    With uTVI
        .hItem = hNode
        .mask = TVIF_HANDLE Or TVIF_STATE
        .stateMask = lState
        .State = fAdd And lState
    End With
    
    Call SendMessage(m_hTreeView, TVM_SETITEM, 0, uTVI)
End Sub
Private Function pvTVGetState(hNode As Long, lState As eStateConstants) As Boolean
  
  Dim uTVI As TVITEM

    With uTVI
        .hItem = hNode
        .mask = TVIF_HANDLE Or TVIF_STATE
    End With
    
    If (SendMessage(m_hTreeView, TVM_GETITEM, 0, uTVI)) Then
        pvTVGetState = (uTVI.State And lState)
    End If
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
        If (Not pOleInPlaceUIWindow Is Nothing) Then 'And Not m_bMouseActivate
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
                    
                    With pMsg
                        Call SendMessageLong(m_hTreeView, .message, .wParam, .lParam)
                    End With
                    frTranslateAccel = True
            End Select
    End Select
    
    On Error GoTo 0
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
