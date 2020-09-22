VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Thumbnailer"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   Begin VB.Timer tmrExploreFolder 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3780
      Top             =   855
   End
   Begin Thumbnailer.ucStatusbar ucStatusbar 
      Height          =   285
      Left            =   30
      Top             =   7680
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   503
   End
   Begin Thumbnailer.ucToolbar ucToolbar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   688
   End
   Begin Thumbnailer.ucSplitter ucSplitterH 
      Height          =   6735
      Left            =   4320
      Top             =   840
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   11880
   End
   Begin Thumbnailer.ucSplitter ucSplitterV 
      Height          =   60
      Left            =   240
      Top             =   4080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   106
      Orientation     =   1
   End
   Begin VB.ComboBox cbPath 
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin Thumbnailer.ucProgress ucProgress 
      Height          =   270
      Left            =   7755
      Top             =   7695
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   476
      BorderStyle     =   0
   End
   Begin Thumbnailer.ucFolderView ucFolderView 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5530
   End
   Begin Thumbnailer.ucThumbnailView ucThumbnailView 
      Height          =   6735
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11880
   End
   Begin Thumbnailer.ucPlayer ucPlayer 
      Height          =   3255
      Left            =   240
      Top             =   4320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5741
      BackColor       =   0
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuGoTop 
      Caption         =   "&Go"
      Begin VB.Menu mnuGo 
         Caption         =   "&Back"
         Index           =   0
      End
      Begin VB.Menu mnuGo 
         Caption         =   "&Forward"
         Index           =   1
      End
      Begin VB.Menu mnuGo 
         Caption         =   "&Up"
         Index           =   2
      End
   End
   Begin VB.Menu mnuViewTop 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Refresh"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Thumbnails"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Details"
         Index           =   3
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuDatabaseTop 
      Caption         =   "&Database"
      Begin VB.Menu mnuDatabase 
         Caption         =   "&Maintenance..."
         Index           =   0
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
   Begin VB.Menu mnuViewModeTop 
      Caption         =   "View mode"
      Visible         =   0   'False
      Begin VB.Menu mnuViewMode 
         Caption         =   "View &thumbnails"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "View &details"
         Index           =   1
      End
   End
   Begin VB.Menu mnuContextThumbnailTop 
      Caption         =   "Context thumbnail"
      Visible         =   0   'False
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Properties"
         Index           =   0
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Shell open..."
         Index           =   2
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Shell edit..."
         Index           =   3
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Explore folder..."
         Index           =   4
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Update item"
         Index           =   6
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Update folder"
         Index           =   7
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Cancel"
         Index           =   9
      End
   End
   Begin VB.Menu mnuContextPreviewTop 
      Caption         =   "Context preview"
      Visible         =   0   'False
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Background color..."
         Index           =   0
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Pause/Resume"
         Index           =   2
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Rotate +90º"
         Index           =   4
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Rotate -90º"
         Index           =   5
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Copy image"
         Index           =   6
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuContextPreview 
         Caption         =   "Cancel"
         Index           =   8
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Application:   Thumbnailer.exe
' Version:       1.0.0
' Last revision: 2004.11.29
' Dependencies:  gdiplus.dll (place in application folder)
'
' Author:        Carles P.V. - ©2004
'========================================================================================



Option Explicit

'-- A little bit of API

Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const WM_SETICON              As Long = &H80
Private Const LR_SHARED               As Long = &H8000&
Private Const ICON_SMALL              As Long = 0
Private Const IMAGE_ICON              As Long = 1

Private Const CB_ERR                  As Long = (-1)
Private Const CB_GETCURSEL            As Long = &H147
Private Const CB_SETCURSEL            As Long = &H14E
Private Const CB_SHOWDROPDOWN         As Long = &H14F
Private Const CB_GETDROPPEDSTATE      As Long = &H157

Private Const SEM_NOGPFAULTERRORBOX   As Long = &H2&

Private Const SEE_MASK_INVOKEIDLIST   As Long = &HC
Private Const SEE_MASK_FLAG_NO_UI     As Long = &H400
Private Const SW_NORMAL               As Long = 1

Private Type SHELLEXECUTEINFO
    cbSize       As Long
    fMask        As Long
    hWnd         As Long
    lpVerb       As String
    lpFile       As String
    lpParameters As String
    lpDirectory  As String
    nShow        As Long
    hInstApp     As Long
    lpIDList     As Long
    lpClass      As String
    hkeyClass    As Long
    dwHotKey     As Long
    hIcon        As Long
    hProcess     As Long
End Type

'-- Private variables

Private m_bInIDE           As Boolean
Private m_GDIplusToken     As Long
Private m_bLoaded          As Boolean
Private m_bEnding          As Boolean
Private m_bComboHasFocus   As Boolean

Private Const m_PathLevels As Long = 100
Private m_Paths()          As String
Private m_PathsPos         As Long
Private m_PathsMax         As Long
Private m_bSkipPath        As Boolean



'========================================================================================
' Initializing / Terminating
'========================================================================================

Private Sub Form_Initialize()

    If (App.PrevInstance) Then End
   
    '-- Initialize common controls
    Call InitCommonControls
    
    '-- Load the GDI+ library
    Dim uGpSI As mGDIplus.GdiplusStartupInput
    Let uGpSI.GdiplusVersion = 1
    If (mGDIplus.GdiplusStartup(m_GDIplusToken, uGpSI) <> [Ok]) Then
        Call MsgBox("Error initializing application!", vbCritical)
        End
    End If
End Sub

Private Sub Form_Load()

    If (m_bLoaded = False) Then
        m_bLoaded = True
        
        '-- Small icon
        Call SendMessage(Me.hWnd, WM_SETICON, ICON_SMALL, ByVal LoadImageAsString(App.hInstance, ByVal "SMALL_ICON", IMAGE_ICON, 16, 16, LR_SHARED))
        
        '-- Initialize database-thumbnail module / Load settings
        Call mThumbnail.InitializeModule
        Call mSettings.LoadSettings

        '-- Modify some menus
        mnuGo(0).Caption = mnuGo(0).Caption & vbTab & "Alt+Left"
        mnuGo(1).Caption = mnuGo(1).Caption & vbTab & "Alt+Right"
        mnuGo(2).Caption = mnuGo(2).Caption & vbTab & "Alt+Up"
        mnuContextPreview(2).Caption = mnuContextPreview(2).Caption & vbTab & "Ctrl+P"
        mnuContextPreview(4).Caption = mnuContextPreview(4).Caption & vbTab & "Ctrl+[+]"
        mnuContextPreview(5).Caption = mnuContextPreview(5).Caption & vbTab & "Ctrl+[-]"
        mnuContextPreview(6).Caption = mnuContextPreview(6).Caption & vbTab & "Ctrl+C"
        
        '-- Initialize toolbar
        With ucToolbar
        
            Call .Initialize(16, FlatStyle:=True, ListStyle:=False, Divider:=True)
            Call .AddBitmap(LoadResPicture("TOOLBAR", vbResBitmap), vbMagenta)
            
            Call .AddButton("Back", 0, , , False)
            Call .AddButton("Forward", 1, , , False)
            Call .AddButton("Up", 2, , , False)
            Call .AddButton(, , , [eSeparator])
            Call .AddButton("Refresh", 3, , , False)
            Call .AddButton(, , , [eSeparator])
            Call .AddButton("View", 4, , [eDropDown], False)
            Call .AddButton("Full screen", 6, , , False)
            Call .AddButton(, , , [eSeparator])
'           Call .AddButton("Preferences", 7, , , False)
'           Call .AddButton(, , , [eSeparator])
            Call .AddButton("Maintenance", 8, , , False)
            
            .Height = .ToolbarHeight
        End With
        
        '-- Initialize paths list
        Call pvChangeDropDownListHeight(cbPath, 400)

        '-- Initialize folder view
        With ucFolderView
            Call .Initialize
            .HasLines = False
        End With
        
        '-- Initialize thumbnail view
        With ucThumbnailView
            Call .Initialize(IMAGETYPES_MASK, "|", _
                             uAPP_SETTINGS.ViewMode, _
                             uAPP_SETTINGS.ViewColumnWidth(0), _
                             uAPP_SETTINGS.ViewColumnWidth(1), _
                             uAPP_SETTINGS.ViewColumnWidth(2), _
                             uAPP_SETTINGS.ViewColumnWidth(3))
            Call .SetThumbnailSize(uAPP_SETTINGS.ThumbnailWidth, uAPP_SETTINGS.ThumbnailHeight)
        End With
        
        '-- Initialize player
        With ucPlayer
            Call .InitializeTypes(IMAGETYPES_MASK)
            .BackColor = uAPP_SETTINGS.PreviewBackColor
            .BestFitMode = uAPP_SETTINGS.PreviewBestFit
            .Zoom = uAPP_SETTINGS.PreviewZoom
        End With
        
        '-- Initialize status bar
        With ucStatusbar
            Call .Initialize(SizeGrip:=True)
            Call .AddPanel(, 150, , [sbSpring])
            Call .AddPanel(, 150)
            Call .AddPanel(, 150)
        End With
        
        '-- Initialize splitters
        Call ucSplitterH.Initialize(Me)
        Call ucSplitterV.Initialize(Me)
        
        '-- Show form
        Call Me.Show: Me.Refresh: Call VBA.DoEvents
        
        '-- Initialize Back/Forward paths list / Go to last recent path
        ReDim m_Paths(0 To m_PathLevels)
        If (cbPath.List(0) <> vbNullString) Then
            m_bSkipPath = True
            cbPath.ListIndex = 0
            m_Paths(1) = cbPath.List(0)
            m_PathsPos = 1
          Else
            Call pvCheckNavigationButtons
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If (m_bLoaded) Then
        m_bEnding = True
        
        '-- Save all settings
        Call mSettings.SaveSettings
        
        '-- Terminate all
        Call mThumbnail.Cancel 'Fix this termination! (-> independent thread: ActiveX EXE ?)
        Call mThumbnail.TerminateModule
        Call ucPlayer.DestroyImage
        
        '-- Shut down gdiplus session
        If (m_GDIplusToken) Then
            Call mGDIplus.GdiplusShutdown(m_GDIplusToken)
        End If
    End If
End Sub

Private Sub Form_Terminate()

    If (Not InIDE()) Then
        Call SetErrorMode(SEM_NOGPFAULTERRORBOX) '(*)
    End If
    End
    
'(*) From vbAccelerator
'    http://www.vbaccelerator.com/home/VB/Code/Libraries/XP_Visual_Styles/Preventing_Crashes_at_Shutdown/article.asp
'    KBID 309366 (http://support.microsoft.com/default.aspx?scid=kb;en-us;309366)
End Sub




'========================================================================================
' Resizing
'========================================================================================

Private Sub Form_Resize()
  
  Const DXMIN As Long = 200
  Const DXMAX As Long = 225
  Const DYMIN As Long = 200
  Const DYMAX As Long = 200
  Const DSEP  As Long = 2
    
    On Error Resume Next
    
    '-- Resize splitters
    Call ucSplitterH.Move(ucSplitterH.Left, ucToolbar.Height + cbPath.Height + 2 * DSEP, ucSplitterH.Width, Me.ScaleHeight - ucToolbar.Height - cbPath.Height - ucStatusbar.Height - 3 * DSEP)
    Call ucSplitterV.Move(DSEP, ucSplitterV.Top, ucSplitterH.Left, ucSplitterV.Height)
    
    '-- Update their min/max pos.
    ucSplitterH.xMax = Me.ScaleWidth - DXMAX
    ucSplitterH.xMin = DXMIN
    ucSplitterV.yMax = Me.ScaleHeight - DYMAX
    ucSplitterV.yMin = DYMIN
    
    '-- Relocate splitters
    If (Me.WindowState = vbNormal) Then
        If (ucSplitterH.Left < ucSplitterH.xMin) Then ucSplitterH.Left = ucSplitterH.xMin
        If (ucSplitterV.Top < ucSplitterV.yMin) Then ucSplitterV.Top = ucSplitterV.yMin
        If (ucSplitterH.Left > ucSplitterH.xMax) Then ucSplitterH.Left = ucSplitterH.xMax
        If (ucSplitterV.Top > ucSplitterV.yMax) Then ucSplitterV.Top = ucSplitterV.yMax
    End If
    
    '-- Status bar size-grip?
    Call SetParent(ucProgress.hWnd, Me.hWnd)
    ucStatusbar.SizeGrip = Not (Me.WindowState = vbMaximized)
    Call SetParent(ucProgress.hWnd, ucStatusbar.hWnd)
    Call ucStatusbar_Resize
    
    '-- Relocate controls
    Call cbPath.Move(DSEP, ucToolbar.Height + DSEP, Me.ScaleWidth - 2 * DSEP)
    Call ucFolderView.Move(DSEP, ucToolbar.Height + cbPath.Height + 2 * DSEP, ucSplitterH.Left - DSEP, ucSplitterV.Top - ucToolbar.Height - cbPath.Height - 2 * DSEP)
    Call ucThumbnailView.Move(ucSplitterH.Left + ucSplitterH.Width, ucToolbar.Height + cbPath.Height + 2 * DSEP, Me.ScaleWidth - ucSplitterH.Left - ucSplitterH.Width - DSEP, Me.ScaleHeight - cbPath.Height - ucToolbar.Height - ucStatusbar.Height - 3 * DSEP)
    Call ucPlayer.Move(DSEP, ucSplitterV.Top + ucSplitterV.Height, ucSplitterH.Left - DSEP, Me.ScaleHeight - ucToolbar.Height - cbPath.Height - ucStatusbar.Height - ucSplitterV.Height - ucFolderView.Height - 3 * DSEP)
    
    On Error GoTo 0
End Sub

Private Sub ucStatusbar_Resize()

  Dim x1 As Long, y1 As Long
  Dim x2 As Long, y2 As Long
    
    '-- Relocate progress bar
    If (ucStatusbar.hWnd) Then
        Call ucStatusbar.GetPanelRect(3, x1, y1, x2, y2)
        Call MoveWindow(ucProgress.hWnd, x1 + 2, y1 + 2, x2 - x1 - 4, y2 - y1 - 4, 0)
    End If
End Sub

Private Sub ucSplitterH_Release()
    Call Form_Resize
End Sub

Private Sub ucSplitterV_Release()
    Call Form_Resize
End Sub



'========================================================================================
' Menus
'========================================================================================

Private Sub mnuFile_Click(Index As Integer)
    
    '-- Exit
    Call Unload(Me)
End Sub

Private Sub mnuGo_Click(Index As Integer)
    
    Select Case Index
        
        Case 0 '-- Back
            Call pvUndoPath
            
        Case 1 '-- Forward
            Call pvRedoPath
            
        Case 2 '-- Up
            Call ucFolderView.Go([fvGoUp])
            Call pvCheckNavigationButtons
    End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
  
    Select Case Index
        
        Case 0    '-- Refresh
            
            If (Not ucFolderView.PathIsRoot) Then
                Call ucPlayer.Clear
                Call ucThumbnailView.Clear
                m_bSkipPath = True
                Call ucFolderView_ChangeAfter(vbNullString)
            End If
        
        Case Else '-- View mode changed
            
            Screen.MousePointer = vbArrowHourglass
            ucThumbnailView.Visible = False
            
            '-- Modify main menu and change view mode
            Select Case Index
                
                Case 2 '-- Thumbnails
                    mnuView(3).Checked = False
                    mnuView(2).Checked = True
                    mnuViewMode(1).Checked = False
                    mnuViewMode(0).Checked = True
                    ucThumbnailView.ViewMode = [tvThumbnail]
                
                Case 3 '-- Details
                    mnuView(2).Checked = False
                    mnuView(3).Checked = True
                    mnuViewMode(0).Checked = False
                    mnuViewMode(1).Checked = True
                    ucThumbnailView.ViewMode = [tvDetails]
            End Select
            
            '-- Modify toolbar icon
            ucToolbar.ButtonImage(7) = 4 + -(ucThumbnailView.ViewMode = [tvDetails])
            
            '-- Store
            uAPP_SETTINGS.ViewMode = ucThumbnailView.ViewMode
            
            ucThumbnailView.Visible = True
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub mnuViewMode_Click(Index As Integer)
    
    Call mnuView_Click(Index + 2)
End Sub

Private Sub mnuDatabase_Click(Index As Integer)
    
    Call fMaintenance.Show(vbModal, Me)
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    
    Call MsgBox("Thumbnailer v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
                "Carles P.V. - 2004" & Space$(15), _
                vbInformation, "About")
End Sub

'//

Private Sub mnuContextPreview_Click(Index As Integer)

  Dim lColor As Long

    Select Case Index
            
        Case 0 '-- Background color...
            
            lColor = mDialogColor.SelectColor(Me.hWnd, ucPlayer.BackColor, Extended:=True)
            If (lColor <> -1) Then
                ucPlayer.BackColor = lColor
                uAPP_SETTINGS.PreviewBackColor = lColor
            End If
            
        Case 2 '-- Pause/Resume
            
            If (ucPlayer.ImageFrames >= 2) Then
                If (ucPlayer.IsPlaying) Then
                    Call ucPlayer.PauseAnimation
                  Else
                    Call ucPlayer.ResumeAnimation
                End If
            End If
        
        Case 4 '-- Rotate +90º
            
            If (ucPlayer.ImageFrames <= 1) Then
                Call VBA.DoEvents
                Call ucPlayer.Rotate90CW
            End If
        
        Case 5 '-- Rotate -90º
            
            If (ucPlayer.ImageFrames <= 1) Then
                Call VBA.DoEvents
                Call ucPlayer.Rotate90ACW
            End If
            
        Case 6 '-- Copy image
            
            If (ucPlayer.HasImage) Then
                Call VBA.DoEvents
                Call ucPlayer.CopyImage
            End If
    End Select
End Sub

Private Sub mnuContextThumbnail_Click(Index As Integer)

  Dim lItm As Long
  Dim uSEI As SHELLEXECUTEINFO
  Dim lRet As Long
    
    Select Case Index
    
        Case 0 To 4 '-- Shell (needs fix for W9x)
        
            With uSEI
                
                .cbSize = Len(uSEI)
                .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
                .hWnd = Me.hWnd
                .lpParameters = vbNullChar
                .lpDirectory = vbNullChar
                
                Select Case Index
            
                    Case 0 '-- Properties
                        .lpVerb = "properties"
                        .lpFile = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
                        .nShow = 0
                    
                    Case 2 '-- Shell open...
                        .lpVerb = "open"
                        .lpFile = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
                        .nShow = 0
                    
                    Case 3 '-- Shell edit...
                        .lpVerb = "edit"
                        .lpFile = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
                        .nShow = 0
                    
                    Case 4 '-- Explore folder...
                        .lpVerb = "open"
                        .lpFile = ucFolderView.Path & vbNullChar
                        .nShow = SW_NORMAL
                End Select
            End With
            
            Call VBA.DoEvents
            lRet = ShellExecuteEx(uSEI)
        
        Case 6 To 7 '-- Database
        
            Call VBA.DoEvents
            Screen.MousePointer = vbArrowHourglass
            
            Select Case Index
                
                Case 6 '-- Update item
                    lItm = ucThumbnailView.ItemFindState(, [tvFocused])
                    Call mThumbnail.UpdateItem(ucFolderView.Path, lItm)
                    Call ucThumbnailView_ItemClick(lItm)
        
                Case 7 '-- Update folder
                    Call ucPlayer.Clear
                    Call ucThumbnailView.Clear
                    Call mThumbnail.DeleteFolderThumbnails(ucFolderView.Path)
                    Call mnuView_Click(0)
            End Select
            Screen.MousePointer = vbDefault
    End Select
End Sub



'========================================================================================
' Toolbar
'========================================================================================

Private Sub ucToolbar_ButtonClick(ByVal Button As Long)
    
    Select Case Button
    
        Case 1  '-- Back
            Call mnuGo_Click(0)
      
        Case 2  '-- Forward
            Call mnuGo_Click(1)
      
        Case 3  '-- Up
            Call mnuGo_Click(2)
      
        Case 5  '-- Refresh
            Call mnuView_Click(0)
       
        Case 7  '-- View
            Select Case ucThumbnailView.ViewMode
                Case [tvThumbnail]
                    Call mnuView_Click(3)
                Case [tvDetails]
                    Call mnuView_Click(2)
            End Select
      
        Case 8  '-- Full screen
            Call ucPlayer_DblClick
      
        Case 10 '-- Database
            Call mnuDatabase_Click(0)
    End Select
End Sub

Private Sub ucToolbar_ButtonDropDown(ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    
    '-- Drop-down menu (view mode)
    Call PopupMenu(mnuViewModeTop, , x, y)
End Sub



'========================================================================================
' Changing path
'========================================================================================

Private Sub ucFolderView_ChangeBefore(ByVal NewPath As String, Cancel As Boolean)

    If (Not m_bEnding And Not ucFolderView.PathIsValid(NewPath)) Then
            
        '-- Invalid path
        Call MsgBox("The specified path is invalid or does not exist.")
        Call SendMessage(cbPath.hWnd, CB_SETCURSEL, 0, ByVal 0)
        Cancel = True
        
      Else
        '-- Stop thumbnailing / Clear
        Call mThumbnail.Cancel
        Call ucPlayer.Clear
        Call ucThumbnailView.Clear
    End If
End Sub

Private Sub ucFolderView_ChangeAfter(ByVal OldPath As String)
    tmrExploreFolder.Enabled = False
    tmrExploreFolder.Enabled = True
End Sub

Private Sub tmrExploreFolder_Timer()

    tmrExploreFolder.Enabled = False
    
    If (Not m_bEnding) Then
        
        ucProgress.Visible = True
        Screen.MousePointer = vbArrowHourglass
        
        '-- Add to recent paths
        Call pvAddPath(ucFolderView.Path): m_bSkipPath = False

        '-- Add items from path
        Call mThumbnail.UpdateFolder(ucFolderView.Path)
        
        '-- Items ?
        If (ucThumbnailView.Count) Then
            
            '-- Select first by default
            If (ucThumbnailView.ItemFindState(, [tvSelected]) = -1) Then
                ucThumbnailView.ItemSelected(0) = True
            End If
            
          Else
            ucStatusbar.PanelText(1) = vbNullString
            ucStatusbar.PanelText(2) = vbNullString
            ucStatusbar.PanelText(3) = vbNullString
        End If
        
        '-- Show # of items found
        ucStatusbar.PanelText(3) = Format$(ucThumbnailView.Count, "#,#0 image/s found")
        
        ucProgress.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cbPath_GotFocus()
    m_bComboHasFocus = True
End Sub
Private Sub cbPath_LostFocus()
    m_bComboHasFocus = False
End Sub

Private Sub cbPath_Click()
    
    '-- Path selected
    If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0) Then
        
        With ucFolderView
            If (.Path <> cbPath.Text) Then
                .Path = cbPath.Text
            End If
        End With
    End If
End Sub

Private Sub cbPath_KeyDown(KeyCode As Integer, Shift As Integer)
    
  Dim lIdx As Long
  
    Select Case KeyCode
    
        '-- New path typed
        Case vbKeyReturn
            
            '-- Check combo's list state (visible)
            If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) <> 0) Then
                '-- Get current list box selected (hot) item
                lIdx = SendMessage(cbPath.hWnd, CB_GETCURSEL, 0, ByVal 0)
                If (lIdx <> CB_ERR) Then
                    Call SendMessage(cbPath.hWnd, CB_SETCURSEL, lIdx, ByVal 0)
                End If
            End If
            
            '-- Hide combo's list and force combo click
            Call SendMessage(cbPath.hWnd, CB_SHOWDROPDOWN, 0, ByVal 0)
            Call cbPath_Click
      
        '-- Avoids navigation when list hidden (also avoids mouse-wheel navigation).
        Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
            
            '-- Preserve manual drop-down
            If (Shift <> vbAltMask) Then
                If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0) Then
                    KeyCode = 0
                End If
            End If
    End Select
End Sub



'========================================================================================
' Displaying image / 'full screen' mode
'========================================================================================

Private Sub ucThumbnailView_ItemClick(ByVal Item As Long)
    
    Screen.MousePointer = vbArrowHourglass
    
    With ucPlayer
        
        '-- Try loading image
        If (.ImportImage(ucFolderView.Path & ucThumbnailView.ItemText(Item, [tvFileName]))) Then
                                    
            '-- Success: show info
            ucStatusbar.PanelText(1) = ucThumbnailView.ItemText(Item, [tvFileName])
            ucStatusbar.PanelText(2) = .ImageWidth & "x" & .ImageHeight & IIf(.ImageTimeString <> vbNullString, " - " & .ImageTimeString, vbNullString)
            ucToolbar.ButtonEnabled(8) = True
        
          Else
            '-- Destroy image
            Call .DestroyImage
            Call .Refresh
            
            '-- Show info
            ucStatusbar.PanelText(1) = "Error!"
            ucStatusbar.PanelText(2) = vbNullString
            ucToolbar.ButtonEnabled(8) = False
        End If
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub ucThumbnailView_ItemDblClick(ByVal Item As Long)

    '-- Force call (full screen)
    Call ucPlayer_DblClick
End Sub

Private Sub ucPlayer_DblClick()

    If (ucPlayer.HasImage) Then
    
        '-- Toggle 'full screen'
        If (Not fFullScreen.Loaded) Then
            Call fFullScreen.Show(vbModal, Me)
          Else
            Call Unload(fFullScreen)
        End If
    End If
End Sub



'========================================================================================
' Context menus
'========================================================================================

Private Sub ucThumbnailView_ItemRightClick(ByVal Item As Long)
    
    '-- Thumbnail context menu
    Call Me.PopupMenu(mnuContextThumbnailTop, , , , mnuContextThumbnail(0))
End Sub

Private Sub ucPlayer_RightClick()
        
    '-- Check available
    mnuContextPreview(2).Enabled = (ucPlayer.HasImage And ucPlayer.ImageFrames >= 2)
    mnuContextPreview(4).Enabled = (ucPlayer.HasImage And ucPlayer.ImageFrames <= 1)
    mnuContextPreview(5).Enabled = (ucPlayer.HasImage And ucPlayer.ImageFrames <= 1)
    mnuContextPreview(6).Enabled = (ucPlayer.HasImage And ucPlayer.HasImage)
    
    '-- Preview context menu
    Call Me.PopupMenu(mnuContextPreviewTop)
End Sub



'========================================================================================
' Navigating
'========================================================================================

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Const SCROLL_FACTOR As Long = 5
  Dim lFocused        As Long
  Dim bResize         As Boolean
    
    Select Case Shift
    
        Case vbAltMask
    
            If (Not m_bComboHasFocus) Then

                If (Not fFullScreen.Loaded) Then
                
                    Select Case KeyCode
                                    
                        Case vbKeyLeft  '-- Back
                            Call mnuGo_Click(0)
                        
                        Case vbKeyRight '-- Forward
                            Call mnuGo_Click(1)
                        
                        Case vbKeyUp    '-- Up
                            Call mnuGo_Click(2)
                    End Select
                    KeyCode = 0
                End If
            End If
      
        Case vbCtrlMask
       
            Select Case KeyCode
                
                Case vbKeyP        '-- Pause/Resume
                    Call mnuContextPreview_Click(2)
                
                Case vbKeyAdd      '-- Pause/Resume
                    Call mnuContextPreview_Click(4)
                
                Case vbKeySubtract '-- Pause/Resume
                    Call mnuContextPreview_Click(5)
                    
                Case vbKeyC        '-- Copy image
                    Call mnuContextPreview_Click(6)
            End Select
            KeyCode = 0
               
        Case Else
            
            Select Case KeyCode
                    
                '-- Navigating thumbnails (full-screen)
                Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
                        
                    If (Not m_bComboHasFocus) Then
                        
                        If (fFullScreen.Loaded) Then
        
                            With ucThumbnailView
                                
                                '-- Currently selected
                                lFocused = .ItemFindState(, [tvFocused])
                                
                                Select Case KeyCode
                            
                                    Case vbKeyPageUp   '-- Previous
                                        .ItemSelected(lFocused + 1 * (lFocused > 0)) = True
                            
                                    Case vbKeyPageDown '-- Next
                                        .ItemSelected(lFocused - 1 * (lFocused < .Count - 1)) = True
                            
                                    Case vbKeyHome     '-- First
                                        .ItemSelected(0) = True
                            
                                    Case vbKeyEnd      '-- Last
                                        .ItemSelected(.Count - 1) = True
                                End Select
                                
                                Call .ItemEnsureVisible(.ItemFindState(, [tvFocused]))
                            End With
                            KeyCode = 0
                        End If
                    End If
                       
                '-- Best fit mode / zoom
                Case vbKeySpace, vbKeyAdd, vbKeySubtract
                        
                    If (Not m_bComboHasFocus) Then
                        
                        With ucPlayer
                        
                            Select Case KeyCode
                                
                                Case vbKeySpace    '-- Best fit mode on/off
                                    .BestFitMode = Not .BestFitMode: bResize = True
                                    
                                Case vbKeyAdd      '-- Zoom +
                                    If (.Zoom < 15 And Not .BestFitMode) Then
                                        .Zoom = .Zoom + 1: bResize = True
                                    End If
                                    
                                Case vbKeySubtract '-- Zoom -
                                    If (.Zoom > 1 And Not .BestFitMode) Then
                                        .Zoom = .Zoom - 1: bResize = True
                                    End If
                            End Select
                            
                            If (bResize) Then
                                
                                Call .Refresh
                                
                                If (fFullScreen.Loaded) Then
                                    uAPP_SETTINGS.FullScreenBestFit = .BestFitMode
                                    uAPP_SETTINGS.FullScreenZoom = .Zoom
                                  Else
                                    uAPP_SETTINGS.PreviewBestFit = .BestFitMode
                                    uAPP_SETTINGS.PreviewZoom = .Zoom
                                End If
                            End If
                        End With
                        KeyCode = 0
                    End If
                    
                '-- Scrolling preview
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                        
                    If (Not m_bComboHasFocus) Then
                        
                        With ucPlayer
                            
                            '-- Scroll <SCROLL_FACTOR> pixels
                            Select Case KeyCode
                                
                                Case vbKeyUp
                                    Call .Scroll(0, SCROLL_FACTOR * .Zoom)
                                    
                                Case vbKeyDown
                                    Call .Scroll(0, -SCROLL_FACTOR * .Zoom)
                                    
                                Case vbKeyLeft
                                    Call .Scroll(-SCROLL_FACTOR * .Zoom, 0)
                                
                                Case vbKeyRight
                                    Call .Scroll(SCROLL_FACTOR * .Zoom, 0)
                            End Select
                        End With
                        KeyCode = 0
                    End If
                         
                '-- Toggle 'full screen'
                Case vbKeyReturn
                    If (Not m_bComboHasFocus) Then
                        Call ucPlayer_DblClick
                        KeyCode = 0
                    End If
                    
                '-- Restore combo edit text
                Case vbKeyEscape
                    Call SendMessage(cbPath.hWnd, CB_SETCURSEL, 0, ByVal 0)
                    KeyCode = 0
                    
                '-- Avoid combo drop-down
                Case vbKeyF4
                    KeyCode = 0
            End Select
    End Select
End Sub

'========================================================================================
' Misc
'========================================================================================

Private Sub ucThumbnailView_ColumnResize(ByVal ColumnID As tvColumnIDConstants)
    
    With uAPP_SETTINGS
        .ViewColumnWidth(ColumnID) = ucThumbnailView.ColumnWidth(ColumnID)
    End With
End Sub



'========================================================================================
' Private
'========================================================================================

Private Sub pvUndoPath()

    If (m_PathsPos > 1) Then
        m_PathsPos = m_PathsPos - 1
        
        '-- Update path
        m_bSkipPath = True
        ucFolderView.Path = m_Paths(m_PathsPos)
        
        '-- Update buttons
        Call pvCheckNavigationButtons
    End If
End Sub

Private Sub pvRedoPath()
  
    If (m_PathsPos < m_PathsMax) Then
        m_PathsPos = m_PathsPos + 1
        
        '-- Update path
        m_bSkipPath = True
        ucFolderView.Path = m_Paths(m_PathsPos)
        
        '-- Update buttons
        Call pvCheckNavigationButtons
    End If
End Sub

Private Sub pvAddPath(ByVal sPath As String)
  
 Dim lc   As Long
 Dim lPtr As Long
    
    With uAPP_SETTINGS
           
        '-- Add to recent paths list
        For lc = 0 To cbPath.ListCount - 1
            If (sPath = cbPath.List(lc)) Then
                Call cbPath.RemoveItem(lc)
                Exit For
            End If
        Next lc
        If (cbPath.ListCount = 25) Then
            Call cbPath.RemoveItem(cbPath.ListCount - 1)
        End If
        Call cbPath.AddItem(sPath, 0): cbPath.ListIndex = 0
        
        If (m_bSkipPath = False) Then
            
            If (m_PathsPos = m_PathLevels) Then
                '-- Move down items
                lPtr = StrPtr(m_Paths(1))
                Call CopyMemory(ByVal VarPtr(m_Paths(1)), ByVal VarPtr(m_Paths(2)), (m_PathLevels - 1) * 4)
                Call CopyMemory(ByVal VarPtr(m_Paths(m_PathLevels)), lPtr, 4)
              Else
                '-- One position up
                m_PathsPos = m_PathsPos + 1
                m_PathsMax = m_PathsPos
            End If
            
            '-- Store path
            m_Paths(m_PathsPos) = sPath
        End If
    End With
    
    '-- Update buttons
    Call pvCheckNavigationButtons
End Sub

Private Sub pvCheckNavigationButtons()
    
    '-- Menu buttons
    mnuGo(0).Enabled = (m_PathsPos > 1)
    mnuGo(1).Enabled = (m_PathsPos < m_PathsMax)
    mnuGo(2).Enabled = Not ucFolderView.PathParentIsRoot And Not ucFolderView.PathIsRoot
    
    '-- Toolbar buttons
    ucToolbar.ButtonEnabled(1) = mnuGo(0).Enabled
    ucToolbar.ButtonEnabled(2) = mnuGo(1).Enabled
    ucToolbar.ButtonEnabled(3) = mnuGo(2).Enabled
    ucToolbar.ButtonEnabled(8) = ucPlayer.HasImage
End Sub

Private Sub pvChangeDropDownListHeight(oCombo As ComboBox, ByVal lHeight As Long)
    
    With oCombo
        '-- Drop down list height
        Call MoveWindow(.hWnd, .Left \ Screen.TwipsPerPixelX, .Top \ Screen.TwipsPerPixelY, .Width \ Screen.TwipsPerPixelX, lHeight, 0)
    End With
End Sub

'//

Private Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE())
   InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function
