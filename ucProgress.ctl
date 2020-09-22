VERSION 5.00
Begin VB.UserControl ucProgress 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ClipControls    =   0   'False
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
End
Attribute VB_Name = "ucProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucProgress.ctl
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2003.05.25
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function TranslateColor Lib "olepro32" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal Palette As Long, Col As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE         As Long = (-16)
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_BORDER         As Long = &H800000
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_WINDOWEDGE  As Long = &H100&
Private Const WS_EX_CLIENTEDGE  As Long = &H200&
Private Const WS_EX_STATICEDGE  As Long = &H20000

'//

'-- Public Enums.:

Public Enum pbBorderStyleConstants
    [pbNone] = 0
    [pbThin]
    [pbThick]
End Enum

Public Enum pbOrientationConstants
    [pbHorizontal] = 0
    [pbVertical]
End Enum

'-- Default Property Values:
Private Const m_def_Orientation = [pbHorizontal]
Private Const m_def_BorderStyle = [pbThick]
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_ForeColor = vbHighlight
Private Const m_def_Max = 100

'-- Property Variables:
Private m_Orientation As pbOrientationConstants
Private m_BorderStyle As pbBorderStyleConstants
Private m_BackColor   As OLE_COLOR
Private m_ForeColor   As OLE_COLOR
Private m_Max         As Long

'-- Private Variables:
Private m_lValue     As Long
Private m_rcFore     As RECT2
Private m_rcBack     As RECT2
Private m_lPos       As Long
Private m_lLastPos   As Long
Private m_hForeBrush As Long
Private m_hBackBrush As Long

'-- Event Declarations:
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()
    If (m_Max = 0) Then m_Max = 1
End Sub

Private Sub UserControl_Terminate()
    If (m_hForeBrush) Then Call DeleteObject(m_hForeBrush)
    If (m_hBackBrush) Then Call DeleteObject(m_hBackBrush)
End Sub

Private Sub UserControl_Resize()
    Call pvGetProgress
    Call pvCalcRects
    Call UserControl_Paint
End Sub

Private Sub UserControl_Paint()
    Call FillRect(UserControl.hDC, m_rcFore, m_hForeBrush)
    Call FillRect(UserControl.hDC, m_rcBack, m_hBackBrush)
End Sub



'========================================================================================
' Events
'========================================================================================

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Get BorderStyle() As pbBorderStyleConstants
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As pbBorderStyleConstants)
    m_BorderStyle = New_BorderStyle
    Call pvSetBorder
    Call pvGetProgress
    Call pvCalcRects
    Call UserControl_Paint
End Property

Public Property Get Orientation() As pbOrientationConstants
    Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal New_Orientation As pbOrientationConstants)

    m_Orientation = New_Orientation

    With Extender
        Call .Move(.Left, .Top, .Height, .Width)
    End With
    Call pvGetProgress
    Call pvCalcRects
    Call UserControl_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    Call pvCreateBackBrush
    Call UserControl_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    Call pvCreateForeBrush
    Call UserControl_Paint
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < 1) Then New_Max = 1
    m_Max = New_Max
End Property

Public Property Get Value() As Long
    Value = m_lValue
End Property
Public Property Let Value(ByVal New_Value As Long)

    m_lValue = New_Value
    
    Call pvGetProgress
    If (m_lPos <> m_lLastPos) Then
        m_lLastPos = m_lPos
        Call pvCalcRects
        Call UserControl_Paint
    End If
End Property

'*

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'//

Private Sub UserControl_InitProperties()

    m_BorderStyle = m_def_BorderStyle
    m_Orientation = m_def_Orientation
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Max = m_def_Max
    
    Call pvSetBorder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        m_BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        m_Orientation = .ReadProperty("Orientation", m_def_Orientation)
        m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        m_Max = .ReadProperty("Max", m_def_Max)
        UserControl.Enabled = .ReadProperty("Enabled", True)
    End With

    Call pvSetBorder
    Call pvCalcRects
    Call pvCreateForeBrush
    Call pvCreateBackBrush
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
        Call .WriteProperty("Orientation", m_Orientation, m_def_Orientation)
        Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
        Call .WriteProperty("Max", m_Max, m_def_Max)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
    End With
End Sub



'========================================================================================
' Private
'========================================================================================

Private Sub pvCreateForeBrush()
    
  Dim lColor As Long
    
    If (m_hForeBrush) Then
        Call DeleteObject(m_hForeBrush)
        m_hForeBrush = 0
    End If
    Call TranslateColor(ForeColor, 0, lColor)
    m_hForeBrush = CreateSolidBrush(lColor)
End Sub

Private Sub pvCreateBackBrush()

  Dim lColor As Long
  
    If (m_hBackBrush) Then
        Call DeleteObject(m_hBackBrush)
        m_hBackBrush = 0
    End If
    Call TranslateColor(BackColor, 0, lColor)
    m_hBackBrush = CreateSolidBrush(lColor)
End Sub

Private Sub pvGetProgress()
    
    Select Case m_Orientation
        Case [pbHorizontal]
            m_lPos = (m_lValue * UserControl.ScaleWidth) \ m_Max
        Case [pbVertical]
            m_lPos = (m_lValue * UserControl.ScaleHeight) \ m_Max
    End Select
End Sub

Private Sub pvCalcRects()
    
    Select Case m_Orientation
        Case [pbHorizontal]
            Call SetRect(m_rcFore, 0, 0, m_lPos, UserControl.ScaleHeight)
            Call SetRect(m_rcBack, m_lPos, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
        Case [pbVertical]
            Call SetRect(m_rcFore, 0, UserControl.ScaleHeight - m_lPos, UserControl.ScaleWidth, UserControl.ScaleHeight)
            Call SetRect(m_rcBack, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - m_lPos)
    End Select
End Sub

Private Sub pvSetBorder()

    Select Case m_BorderStyle
        Case [pbNone]
            Call pvSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [pbThin]
            Call pvSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [pbThick]
            Call pvSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE)
    End Select
End Sub

Private Sub pvSetWinStyle(ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
    
    lS = GetWindowLong(hWnd, lType)
    lS = (lS And Not lStyleNot) Or lStyle
    Call SetWindowLong(hWnd, lType, lS)
    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

