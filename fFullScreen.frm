VERSION 5.00
Begin VB.Form fFullScreen 
   BorderStyle     =   0  'None
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "fFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private m_bLoaded As Boolean

Private Sub Form_Load()

    With fMain
        
        Call SetParent(.ucPlayer.hWnd, Me.hWnd)
        
        .ucPlayer.Border = False
        .ucPlayer.BestFitMode = uAPP_SETTINGS.FullScreenBestFit
        .ucPlayer.Zoom = uAPP_SETTINGS.FullScreenZoom
    End With
    
    m_bLoaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Const DSEP As Long = 2

    With fMain
        
        Call SetParent(.ucPlayer.hWnd, .hWnd)
        
        On Error Resume Next
            
            .ucPlayer.Border = True
            .ucPlayer.BestFitMode = uAPP_SETTINGS.PreviewBestFit
            .ucPlayer.Zoom = uAPP_SETTINGS.PreviewZoom
            
            Call .ucPlayer.Move(DSEP, .ucSplitterV.Top + .ucSplitterV.Height, .ucSplitterH.Left - DSEP, .ScaleHeight - .ucToolbar.Height - .cbPath.Height - .ucStatusbar.Height - .ucSplitterV.Height - .ucFolderView.Height - 3 * DSEP)
            Call .ucPlayer.UpdatePointer
            Call .ucSplitterV.ZOrder
            Call .ucSplitterH.ZOrder
        
        On Error GoTo 0
    End With
    
    Set fFullScreen = Nothing
    m_bLoaded = False
End Sub

Private Sub Form_Resize()
    Call fMain.ucPlayer.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
    Call fMain.ucPlayer.UpdatePointer
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call fMain.Form_KeyDown(KeyCode, Shift)
End Sub

Public Property Get Loaded() As Boolean
    Loaded = m_bLoaded
End Property
