VERSION 5.00
Begin VB.UserControl ucPlayer 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   HitBehavior     =   2  'Use Paint
   LockControls    =   -1  'True
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   Begin VB.Timer tmrGIF 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ucPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucPlayer.ctl
' Author:        Carles P.V.
' Dependencies:  mGDIplus.bas (->gdiplus.dll)
' Last revision: 2004.11.24
'================================================

Option Explicit
Option Compare Text

'-- API:

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Const RGN_DIFF As Long = 4

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TranslateColor Lib "olepro32" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal Palette As Long, Col As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

'//

'-- Property Variables:

Private m_BackColor       As OLE_COLOR
Private m_BestFitMode     As Boolean
Private m_MaxRenderWidth  As Long
Private m_MaxRenderHeight As Long
Private m_Zoom            As Long

'-- Private Objects:

Private m_oDIB            As cDIB

'-- Private Variables:

Private m_bIsNT           As Boolean

Private m_sImageTypesMask As String

Private m_hImage          As Long
Private m_lImageWidth     As Long
Private m_lImageHeight    As Long
Private m_uCLSID          As CLSID
Private m_lFrame          As Long
Private m_lFrames         As Long
Private m_sTime           As String
Private m_lDelay()        As Long

Private m_lWidth          As Long
Private m_lHeight         As Long
Private m_lLeft           As Long
Private m_lTop            As Long
Private m_lHPos           As Long
Private m_lHMax           As Long
Private m_lVPos           As Long
Private m_lVMax           As Long
Private m_lLastHPos       As Long
Private m_lLastVPos       As Long
Private m_lLastHMax       As Long
Private m_lLastVMax       As Long
Private m_bMouseDown      As Boolean
Private m_ptCurr          As POINTAPI

'-- Event Declarations:

Public Event Click()
Public Event DblClick()
Public Event RightClick()
Public Event Scroll()
Public Event Resize()



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Default values
    m_MaxRenderWidth = Screen.Width \ Screen.TwipsPerPixelX
    m_MaxRenderHeight = m_MaxRenderWidth
    m_Zoom = 1
    
    '-- Initialize DIB section
    Set m_oDIB = New cDIB
    
    '-- NT system ? (Halfone stretching)
    m_bIsNT = pvIsNT
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy GDI/GDI+ objects
    Call DestroyImage
End Sub

Private Sub UserControl_Resize()
    
    '-- Resize and refresh
    Call pvResizeCanvas
    Call pvRefreshCanvas
    
    RaiseEvent Resize
End Sub

Private Sub UserControl_Paint()

    '-- Refresh Canvas
    Call pvRefreshCanvas
End Sub

'//

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Mouse down flag / Store values
    m_bMouseDown = (Button = vbLeftButton)
    m_ptCurr.x = x
    m_ptCurr.y = y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If (m_bMouseDown And Not m_BestFitMode) Then
    
        '-- Apply offsets
        m_lHPos = m_lHPos + (m_ptCurr.x - x)
        m_lVPos = m_lVPos + (m_ptCurr.y - y)
        
        '-- Check margins
        If (m_lHPos < 0) Then m_lHPos = 0 Else If (m_lHPos > m_lHMax) Then m_lHPos = m_lHMax
        If (m_lVPos < 0) Then m_lVPos = 0 Else If (m_lVPos > m_lVMax) Then m_lVPos = m_lVMax
        
        '-- Get current position
        m_ptCurr.x = x
        m_ptCurr.y = y
        
        '-- Srolled ?
        If (m_lLastHPos <> m_lHPos Or m_lLastVPos <> m_lVPos) Then
            Call pvRefreshCanvas
            RaiseEvent Scroll
        End If
        
        '-- Store current position
        m_lLastHPos = m_lHPos
        m_lLastVPos = m_lVPos
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '-- Mouse down flag
    m_bMouseDown = False
    
    If (Button = vbRightButton) Then
        RaiseEvent RightClick
    End If
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeTypes(ByVal ImageTypesMask As String)
    
    m_sImageTypesMask = ImageTypesMask
End Sub
      
Public Function ImportImage(ByVal Filename As String) As Boolean
  
  Dim aItem()     As mGDIplus.PropertyItem
  Dim lBufferSize As Long
  Dim lc          As Long
  Dim sngTime     As Single
   
    '-- Destroy previous
    Call Me.DestroyImage
    
    '-- Try load
    If (mGDIplus.GdipLoadImageFromFile(StrConv(Filename, vbUnicode), m_hImage) = [Ok]) Then
        
        '-- Success: build DIB buffer
        If (pvBuildDIBBuffer(m_hImage)) Then
            
            '-- Update canvas screen
            Call pvResizeCanvas
            Call pvEraseBackground
        
            '-- Animation [?]
            Call mGDIplus.DEFINE_GUID(mGDIplus.FrameDimensionTime, m_uCLSID)
            Call mGDIplus.GdipImageGetFrameCount(m_hImage, m_uCLSID, m_lFrames)
            
            If (m_lFrames > 1) Then
                    
                If (mGDIplus.GdipGetPropertyItemSize(m_hImage, [PropertyTagFrameDelay], lBufferSize) = [Ok]) Then
                  
                    '-- Get animation delays
                    ReDim aItem(0 To lBufferSize / Len(aItem(0)))
                    Call mGDIplus.GdipGetPropertyItem(m_hImage, [PropertyTagFrameDelay], lBufferSize, aItem(0))
                    m_lDelay() = mGDIplus.GetPropertyValue(aItem(0))
                    
                    '-- Adjust for timing
                    For lc = 1 To UBound(m_lDelay())
                        Select Case m_lDelay(lc)
                            Case Is > 6000
                                m_lDelay(lc) = 60000 ' Max.: 1 min.
                            Case Is < 5
                                m_lDelay(lc) = 50   ' Min.: 0.05 sec.
                            Case Else
                                m_lDelay(lc) = m_lDelay(lc) * 10
                        End Select
                        sngTime = sngTime + m_lDelay(lc) / 1000
                    Next lc
                    m_sTime = Format$(sngTime \ 60, "~ 00' ") & Format$(sngTime - (sngTime \ 60) * 60, "00''")
                    
                    '-- Start animation
                    m_lFrame = 0
                    With tmrGIF
                        .Interval = 1
                        .Enabled = True
                    End With
                End If
              
              Else
                '-- Render single frame
                Call pvRenderFrame
                Call pvRefreshCanvas(bEraseBackground:=False)
            End If
        End If
        
        '-- Success
        ImportImage = True
    End If
            
    Call pvUpdatePointer
End Function

Public Function DestroyImage()
    
    '-- Stop GIF timer
    tmrGIF.Enabled = False
    
    '-- Dispose GDI+ image ?
    If (m_hImage) Then
        Call mGDIplus.GdipDisposeImage(m_hImage)
        m_hImage = 0
    End If
    
    '-- Destroy DIB ?
    If (m_oDIB.hDIB) Then
        Call m_oDIB.Destroy
    End If
    
    '-- Reset image variables
    m_lWidth = 0
    m_lHeight = 0
    m_lImageWidth = 0
    m_lImageHeight = 0
    m_lFrames = 0
    m_sTime = vbNullString
End Function

Public Sub Clear()

    '-- Destroy GDI/GDI+ objects
    Call DestroyImage
    
    '-- Update canvas
    Call pvResizeCanvas
    Call pvRefreshCanvas
End Sub

Public Sub Refresh()
    
    '-- Refresh canvas
    Call pvRefreshCanvas
End Sub

Public Function Scroll(ByVal x As Long, ByVal y As Long) As Boolean

    '-- Apply offsets
    m_lHPos = m_lHPos - x
    m_lVPos = m_lVPos - y
    
    '-- Check margins
    If (m_lHPos < 0) Then m_lHPos = 0 Else If (m_lHPos > m_lHMax) Then m_lHPos = m_lHMax
    If (m_lVPos < 0) Then m_lVPos = 0 Else If (m_lVPos > m_lVMax) Then m_lVPos = m_lVMax
    
    '-- Scolled ?
    If (m_lLastHPos <> m_lHPos Or m_lLastVPos <> m_lVPos) Then
        Call pvRefreshCanvas: Scroll = True
        RaiseEvent Scroll
    End If
    
    '-- Store current position
    m_lLastHPos = m_lHPos
    m_lLastVPos = m_lVPos
End Function

Public Sub UpdatePointer()
    
    '-- Updatemouse pointer
    Call pvUpdatePointer
End Sub

'//

Public Sub PauseAnimation()
    
    If (m_lFrames > 1) Then
        tmrGIF.Enabled = False
    End If
End Sub

Public Sub ResumeAnimation()
    
    If (m_lFrames > 1) Then
        tmrGIF.Enabled = True
    End If
End Sub

Public Sub Rotate90CW()

    If (m_hImage) Then
        
        Screen.MousePointer = vbHourglass
        
        Call mGDIplus.GdipImageRotateFlip(m_hImage, [Rotate90FlipNone])
        Call pvBuildDIBBuffer(m_hImage)
        Call pvRenderFrame
        Call pvResizeCanvas
        Call pvUpdatePointer
        Call pvRefreshCanvas
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub Rotate90ACW()

    If (m_hImage) Then
    
        Screen.MousePointer = vbHourglass
        
        Call mGDIplus.GdipImageRotateFlip(m_hImage, [Rotate270FlipNone])
        Call pvBuildDIBBuffer(m_hImage)
        Call pvRenderFrame
        Call pvResizeCanvas
        Call pvUpdatePointer
        Call pvRefreshCanvas
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub CopyImage()
  
  Dim oDIB24 As cDIB
    
    If (m_hImage) Then
        
        Screen.MousePointer = vbHourglass
        
        '-- Create buffer DIB
        Set oDIB24 = New cDIB
        
        '-- Get image (use current buffer)
        Call oDIB24.Create(m_oDIB.Width, m_oDIB.Height, [24_bpp])
        m_oDIB.Paint (oDIB24.hDC)
        
        '-- Copy to clipboard
        Call oDIB24.CopyToClipboard
        Set oDIB24 = Nothing
        
        Screen.MousePointer = vbDefault
    End If
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    If (m_hImage) Then
        Call pvRenderFrame
    End If
    Call pvRefreshCanvas
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BestFitMode(ByVal New_BestFitMode As Boolean)
    m_BestFitMode = New_BestFitMode
    Call pvResizeCanvas
    Call pvUpdatePointer
End Property
Public Property Get BestFitMode() As Boolean
    BestFitMode = m_BestFitMode
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    UserControl.BorderStyle = -New_Border
End Property
Public Property Get Border() As Boolean
    Border = (UserControl.BorderStyle = 1)
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Attribute Enabled.VB_MemberFlags = "400"
    UserControl.Enabled = New_Enabled
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Get HasImage() As Boolean
    HasImage = (m_oDIB.hDIB <> 0)
End Property

Public Property Get ImageWidth() As Long
    ImageWidth = m_lImageWidth
End Property

Public Property Get ImageHeight() As Long
    ImageHeight = m_lImageHeight
End Property

Public Property Get ImageFrames() As Long
    ImageFrames = m_lFrames
End Property

Public Property Get ImageTimeString() As String
    ImageTimeString = m_sTime
End Property

Public Property Get IsPlaying() As Boolean
    If (m_lFrames > 1) Then
        IsPlaying = tmrGIF.Enabled
    End If
End Property

Public Property Get MaxRenderWidth() As Long
    MaxRenderWidth = m_MaxRenderWidth
End Property
Public Property Let MaxRenderWidth(ByVal New_MaxRenderWidth As Long)
    m_MaxRenderWidth = New_MaxRenderWidth
End Property

Public Property Get MaxRenderHeight() As Long
    MaxRenderHeight = m_MaxRenderHeight
End Property
Public Property Let MaxRenderHeight(ByVal New_MaxRenderHeight As Long)
    m_MaxRenderHeight = New_MaxRenderHeight
End Property

Public Property Let Zoom(ByVal New_Zoom As Long)
Attribute Zoom.VB_MemberFlags = "400"
    m_Zoom = IIf(New_Zoom < 1, 1, New_Zoom)
    Call pvResizeCanvas
    Call pvUpdatePointer
End Property
Public Property Get Zoom() As Long
    Zoom = m_Zoom
End Property

'//

Public Property Get ScrollHMax() As Long
    ScrollHMax = m_lHMax
End Property
Public Property Get ScrollVMax() As Long
    ScrollVMax = m_lVMax
End Property
Public Property Get ScrollHPos() As Long
    ScrollHPos = m_lHPos
End Property
Public Property Get ScrollVPos() As Long
    ScrollVPos = m_lVPos
End Property

'//

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'//

Private Sub UserControl_InitProperties()
    m_BackColor = vbApplicationWorkspace
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, vbApplicationWorkspace)
End Sub



'========================================================================================
' Private
'========================================================================================

Private Function pvBuildDIBBuffer(ByVal hImage As Long) As Boolean
    
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
 
    '-- Get image width and height
    Call mGDIplus.GdipGetImageWidth(hImage, m_lImageWidth)
    Call mGDIplus.GdipGetImageHeight(hImage, m_lImageHeight)
    
    '-- Best fit to maximum render size
    Call m_oDIB.GetBestFitInfo(m_lImageWidth, m_lImageHeight, m_MaxRenderWidth, m_MaxRenderHeight, bfx, bfy, bfW, bfH)
    
    '-- Success
    pvBuildDIBBuffer = (m_oDIB.Create(bfW, bfH, [32_bpp]) <> 0)
End Function

Private Sub pvRenderFrame()

  Dim hGraphics As Long
  Dim lColor    As Long
    
    '-- Erase buffer background
    Call TranslateColor(m_BackColor, 0, lColor)
    Call m_oDIB.Cls(lColor)
    
    '-- Select frame / prepare render surface
    If (m_lFrames > 1) Then
        Call mGDIplus.GdipImageSelectActiveFrame(m_hImage, m_uCLSID, m_lFrame)
    End If
    Call mGDIplus.GdipCreateFromHDC(m_oDIB.hDC, hGraphics)
    
    '-- Render frame
    Call mGDIplus.GdipDrawImageRectI(hGraphics, m_hImage, 0, 0, m_oDIB.Width, m_oDIB.Height)
    
    '-- Clean up
    Call mGDIplus.GdipDeleteGraphics(hGraphics)
End Sub

Private Sub pvEraseBackground()

  Dim hRgn_1 As Long
  Dim hRgn_2 As Long
  Dim lColor As Long
  Dim hBrush As Long
    
    '-- Create brush (background)
    Call TranslateColor(m_BackColor, 0, lColor)
    hBrush = CreateSolidBrush(lColor)

    '-- Create Cls region (Control Rect. - Canvas Rect.)
    hRgn_1 = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
    hRgn_2 = CreateRectRgn(m_lLeft, m_lTop, m_lLeft + m_lWidth, m_lTop + m_lHeight)
    Call CombineRgn(hRgn_1, hRgn_1, hRgn_2, RGN_DIFF)
    
    '-- Fill it
    Call FillRgn(UserControl.hDC, hRgn_1, hBrush)
    
    '-- Clear
    Call DeleteObject(hBrush)
    Call DeleteObject(hRgn_1)
    Call DeleteObject(hRgn_2)
End Sub

Private Sub pvRefreshCanvas(Optional ByVal bEraseBackground As Boolean = True)
  
  Dim xOff As Long, yOff As Long
  Dim wDst As Long, hDst As Long
  Dim xSrc As Long, ySrc As Long
  Dim wSrc As Long, hSrc As Long
    
    If (m_oDIB.hDIB) Then
        
        '-- Get Left and Width of source image rectangle
        If (m_lHMax And Not m_BestFitMode) Then
            
            xOff = -m_lHPos Mod m_Zoom
            wDst = (m_lWidth \ m_Zoom) * m_Zoom + 2 * m_Zoom
            xSrc = m_lHPos \ m_Zoom
            wSrc = m_lWidth \ m_Zoom + 2
          
          Else
            xOff = m_lLeft
            wDst = m_lWidth
            xSrc = 0
            wSrc = m_oDIB.Width
        End If
        
        '-- Get Top and Height of source image rectangle
        If (m_lVMax And Not m_BestFitMode) Then
            
            yOff = -m_lVPos Mod m_Zoom
            hDst = (m_lHeight \ m_Zoom) * m_Zoom + 2 * m_Zoom
            ySrc = m_lVPos \ m_Zoom
            hSrc = m_lHeight \ m_Zoom + 2
          
          Else
            yOff = m_lTop
            hDst = m_lHeight
            ySrc = 0
            hSrc = m_oDIB.Height
        End If
        
        '-- Paint visible source rectangle
        If (m_BestFitMode And m_bIsNT) Then
            Call m_oDIB.Stretch(UserControl.hDC, xOff, yOff, wDst, hDst, xSrc, ySrc, wSrc, hSrc, , [sbmHalftone])
          Else
            Call m_oDIB.Stretch(UserControl.hDC, xOff, yOff, wDst, hDst, xSrc, ySrc, wSrc, hSrc)
        End If
    End If
    
    '-- Erase background
    If (bEraseBackground) Then
        Call pvEraseBackground
    End If
End Sub

Private Sub pvResizeCanvas()
  
    If (m_oDIB.hDIB) Then
    
        If (m_BestFitMode = False) Then
        
            '-- Get new Width
            If (m_oDIB.Width * m_Zoom > UserControl.ScaleWidth) Then
                m_lHMax = m_oDIB.Width * m_Zoom - UserControl.ScaleWidth
                m_lWidth = UserControl.ScaleWidth
              Else
                m_lHMax = 0
                m_lWidth = m_oDIB.Width * m_Zoom
            End If
            
            '-- Get new Height
            If (m_oDIB.Height * m_Zoom > UserControl.ScaleHeight) Then
                m_lVMax = m_oDIB.Height * m_Zoom - UserControl.ScaleHeight
                m_lHeight = UserControl.ScaleHeight
              Else
                m_lVMax = 0
                m_lHeight = m_oDIB.Height * m_Zoom
            End If
            
            '-- Offsets
            m_lLeft = (UserControl.ScaleWidth - m_lWidth) \ 2
            m_lTop = (UserControl.ScaleHeight - m_lHeight) \ 2
          
          Else
            '-- Get best fit dimensions
            Call m_oDIB.GetBestFitInfo(m_oDIB.Width, m_oDIB.Height, UserControl.ScaleWidth, UserControl.ScaleHeight, m_lLeft, m_lTop, m_lWidth, m_lHeight)
        End If
                            
        '-- Memorize position
        If (m_lLastHMax) Then
            m_lHPos = (m_lLastHPos * m_lHMax) \ m_lLastHMax
          Else
            m_lHPos = m_lHMax \ 2
        End If
        If (m_lLastVMax) Then
            m_lVPos = (m_lLastVPos * m_lVMax) \ m_lLastVMax
          Else
            m_lVPos = m_lVMax \ 2
        End If
        m_lLastHPos = m_lHPos: m_lLastVPos = m_lVPos
        m_lLastHMax = m_lHMax: m_lLastVMax = m_lVMax
      
      Else
        '-- *Hide* canvas
        m_lWidth = 0
        m_lHeight = 0
    End If
End Sub

Private Sub tmrGIF_Timer()
    
    '-- Render current frame
    Call pvRenderFrame
    
    '-- Refresh canvas
    Call pvRefreshCanvas(bEraseBackground:=False)
    
    '-- Set delay
    tmrGIF.Interval = m_lDelay(m_lFrame + 1)
    
    '-- Next frame
    m_lFrame = m_lFrame + 1
    If (m_lFrame = m_lFrames) Then
        m_lFrame = 0
    End If
End Sub

Private Sub pvUpdatePointer()

    If ((m_lHMax Or m_lVMax) And (Not m_BestFitMode)) Then
        UserControl.MousePointer = vbSizeAll
      Else
        UserControl.MousePointer = vbDefault
    End If
End Sub

'*

Private Function pvIsNT() As Boolean

  Dim uOSVI As OSVERSIONINFO
  
    uOSVI.dwOSVersionInfoSize = Len(uOSVI)
    If (GetVersionEx(uOSVI)) Then
        pvIsNT = (uOSVI.dwPlatformId = 2)
    End If
End Function
