VERSION 5.00
Begin VB.UserControl ucSplitter 
   CanGetFocus     =   0   'False
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ClipControls    =   0   'False
   MousePointer    =   9  'Size W E
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   6
End
Attribute VB_Name = "ucSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucSplitter.ctl
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2003.03.28
'================================================

Option Explicit

'-- API:

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'//

'-- Public Enums.:

Public Enum SplitterOrientationCts
    [splHorizontal] = 0
    [splVertical]
End Enum

'-- Property Variables:
Private m_Orientation As SplitterOrientationCts
Private m_xMax        As Single
Private m_xMin        As Single
Private m_yMax        As Single
Private m_yMin        As Single

'-- Private Variables:
Private WithEvents m_ParentForm As Form
Attribute m_ParentForm.VB_VarHelpID = -1
Private m_Initialized           As Boolean
Private m_Hooked                As Boolean
Private m_HookOffsetX           As Long
Private m_HookOffsetY           As Long

'-- Event Declarations:
Public Event Move(x As Single, y As Single)
Public Event Release()



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Splitter hooked
    m_Hooked = True
    
    '-- Get hook offsets
    m_HookOffsetX = x
    m_HookOffsetY = y
    
    '-- Pass mouse capture to parent form
    Call pvUserControlPaint
    Call ReleaseCapture
    Call SetCapture(m_ParentForm.hWnd)
End Sub

Private Sub UserControl_Paint()

    If (Not Ambient.UserMode) Then
        Call pvUserControlPaint
    End If
End Sub

Private Sub UserControl_Terminate()

    If (Not m_ParentForm Is Nothing) Then
        Set m_ParentForm = Nothing
    End If
End Sub

'//

Private Sub pvUserControlPaint()

  Dim lW As Long
  Dim lH As Long
  
    lW = UserControl.ScaleWidth - 1
    lH = UserControl.ScaleHeight - 1
    UserControl.Line (0, 0)-(lW, lH), vbHighlight, BF
End Sub



'========================================================================================
' Parent form
'========================================================================================

Private Sub m_ParentForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (m_Hooked And m_Initialized) Then
    
        '-- Move splitter
        Select Case m_Orientation
            
            Case [splHorizontal]
                x = x - m_HookOffsetX
                If (x < m_xMin) Then x = m_xMin
                If (x > m_xMax) Then x = m_xMax
                Extender.Move x, Extender.Top
            
            Case [splVertical]
                y = y - m_HookOffsetY
                If (y < m_yMin) Then y = m_yMin
                If (y > m_yMax) Then y = m_yMax
                Extender.Move Extender.Left, y
        End Select
        
        RaiseEvent Move(x, y)
    End If
End Sub

Private Sub m_ParentForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '-- Splitter released
    If (m_Hooked And m_Initialized) Then
        m_Hooked = False
        
        Call UserControl.Cls
        RaiseEvent Release
    End If
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Sub Initialize(FormParent As Form)

    '-- Set parent form
    Set m_ParentForm = FormParent
    
    '-- Splitter initialized flag
    m_Initialized = True
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Get Orientation() As SplitterOrientationCts
    Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal New_Orientation As SplitterOrientationCts)

    m_Orientation = New_Orientation
    
    '-- Set mouse pointer
    Select Case m_Orientation
        Case [splHorizontal]
            UserControl.MousePointer = vbSizeWE
        Case [splVertical]
            UserControl.MousePointer = vbSizeNS
    End Select
End Property

Public Property Get xMax() As Single
    xMax = m_xMax
End Property
Public Property Let xMax(ByVal New_xMax As Single)
    If (New_xMax < m_xMin) Then m_xMax = m_xMin Else m_xMax = New_xMax
End Property

Public Property Get xMin() As Single
    xMin = m_xMin
End Property
Public Property Let xMin(ByVal New_xMin As Single)
    If (New_xMin > m_xMax) Then m_xMin = m_xMax Else m_xMin = New_xMin
End Property

Public Property Get yMax() As Single
    yMax = m_yMax
End Property
Public Property Let yMax(ByVal New_yMax As Single)
    If (New_yMax < m_yMin) Then m_yMax = m_yMin Else m_yMax = New_yMax
End Property

Public Property Get yMin() As Single
    yMin = m_yMin
End Property
Public Property Let yMin(ByVal New_yMin As Single)
    If (New_yMin > m_yMax) Then m_yMin = m_yMax Else m_yMin = New_yMin
End Property

'//

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        Orientation = .ReadProperty("Orientation", 0)
        m_xMax = .ReadProperty("xMax", 0)
        m_xMin = .ReadProperty("xMin", 0)
        m_yMax = .ReadProperty("yMax", 0)
        m_yMin = .ReadProperty("yMin", 0)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Orientation", m_Orientation, 0)
        Call .WriteProperty("xMax", m_xMax, 0)
        Call .WriteProperty("xMin", m_xMin, 0)
        Call .WriteProperty("yMax", m_yMax, 0)
        Call .WriteProperty("yMin", m_yMin, 0)
    End With
End Sub
