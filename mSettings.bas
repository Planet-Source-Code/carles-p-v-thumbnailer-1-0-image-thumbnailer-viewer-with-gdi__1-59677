Attribute VB_Name = "mSettings"
Option Explicit
Option Compare Text

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Type APP_SETTINGS
    ThumbnailWidth     As Long
    ThumbnailHeight    As Long
    ViewMode           As tvViewModeConstants
    ViewColumnWidth(3) As Long
    PreviewBackColor   As Long
    PreviewBestFit     As Boolean
    PreviewZoom        As Long
    FullScreenBestFit  As Boolean
    FullScreenZoom     As Long
End Type

'//

Public uAPP_SETTINGS As APP_SETTINGS

Public Sub LoadSettings()
    
  Dim sINI As String
  Dim sVal As String
  Dim lc   As Long
  
    sINI = pvAppPath & "Thumbnailer.ini"
    
    With fMain
    
        '-- Main window
        .Width = pvGetINI(sINI, "Form", "MainWidth", .Width)
        .Height = pvGetINI(sINI, "Form", "MainHeight", .Height)
        .Top = pvGetINI(sINI, "Form", "MainTop", (Screen.Height - .Height) \ 2)
        .Left = pvGetINI(sINI, "Form", "MainLeft", (Screen.Width - .Width) \ 2)
        .WindowState = pvGetINI(sINI, "Form", "MainWindowState", .WindowState)
        
        '-- Splitters
        .ucSplitterH.Left = pvGetINI(sINI, "Form", "SplitterHPos", 250)
        .ucSplitterV.Top = pvGetINI(sINI, "Form", "SplitterVPos", 325)
    End With
        
    '-- Options
    With uAPP_SETTINGS
    
        '-- Various
        .ThumbnailWidth = pvGetINI(sINI, "Options", "ThumbnailWidth", 80)
        .ThumbnailHeight = pvGetINI(sINI, "Options", "ThumbnailHeight", 80)
        .ViewMode = pvGetINI(sINI, "Options", "ViewMode", [tvThumbnail])
        .ViewColumnWidth(0) = pvGetINI(sINI, "Options", "ViewColumnWidth0", 100)
        .ViewColumnWidth(1) = pvGetINI(sINI, "Options", "ViewColumnWidth1", 100)
        .ViewColumnWidth(2) = pvGetINI(sINI, "Options", "ViewColumnWidth2", 100)
        .ViewColumnWidth(3) = pvGetINI(sINI, "Options", "ViewColumnWidth3", 100)
        .PreviewBackColor = pvGetINI(sINI, "Options", "PreviewBackColor", vbBlack)
        .PreviewBestFit = pvGetINI(sINI, "Options", "PreviewBestFit", True)
        .PreviewZoom = pvGetINI(sINI, "Options", "PreviewZoom", 1)
        .FullScreenBestFit = pvGetINI(sINI, "Options", "FullScreenBestFit", False)
        .FullScreenZoom = pvGetINI(sINI, "Options", "FullScreenZoom", 1)
        
        '-- Recent paths
        For lc = 0 To 24
            sVal = pvGetINI(sINI, "Options", "RecentPaths" & Format$(lc, "00"), vbNullString)
            If (sVal <> vbNullString) Then
                Call fMain.cbPath.AddItem(sVal)
              Else
                Exit For
            End If
        Next lc
    End With
End Sub

Public Sub SaveSettings()
    
  Dim sINI As String
  Dim sVal As String
  Dim lc   As Long
    
    sINI = pvAppPath & "Thumbnailer.ini"
    
    With fMain
        
        '-- Main window
        If (.WindowState = vbNormal) Then
            Call pvPutINI(sINI, "Form", "MainWidth", .Width)
            Call pvPutINI(sINI, "Form", "MainHeight", .Height)
            Call pvPutINI(sINI, "Form", "MainTop", .Top)
            Call pvPutINI(sINI, "Form", "MainLeft", .Left)
        End If
        Call pvPutINI(sINI, "Form", "MainWindowState", .WindowState)
        
        '-- Splitters
        Call pvPutINI(sINI, "Form", "SplitterHPos", .ucSplitterH.Left)
        Call pvPutINI(sINI, "Form", "SplitterVPos", .ucSplitterV.Top)
    End With
        
    '-- Options
    With uAPP_SETTINGS
        
        '-- Various
        Call pvPutINI(sINI, "Options", "ThumbnailWidth", .ThumbnailWidth)
        Call pvPutINI(sINI, "Options", "ThumbnailHeight", .ThumbnailHeight)
        Call pvPutINI(sINI, "Options", "ViewMode", .ViewMode)
        Call pvPutINI(sINI, "Options", "ViewColumnWidth0", .ViewColumnWidth(0))
        Call pvPutINI(sINI, "Options", "ViewColumnWidth1", .ViewColumnWidth(1))
        Call pvPutINI(sINI, "Options", "ViewColumnWidth2", .ViewColumnWidth(2))
        Call pvPutINI(sINI, "Options", "ViewColumnWidth3", .ViewColumnWidth(3))
        Call pvPutINI(sINI, "Options", "PreviewBackColor", .PreviewBackColor)
        Call pvPutINI(sINI, "Options", "PreviewBestFit", .PreviewBestFit)
        Call pvPutINI(sINI, "Options", "PreviewZoom", .PreviewZoom)
        Call pvPutINI(sINI, "Options", "FullScreenBestFit", .FullScreenBestFit)
        Call pvPutINI(sINI, "Options", "FullScreenZoom", .FullScreenZoom)
        
        '-- Recent paths
        For lc = 0 To 24
            sVal = fMain.cbPath.List(lc)
            If (sVal <> vbNullString) Then
                Call pvPutINI(sINI, "Options", "RecentPaths" & Format$(lc, "00"), sVal)
              Else
                Exit For
            End If
        Next lc
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvPutINI(ByVal INIFile As String, ByVal INIHead As String, ByVal INIKey As String, ByVal INIVal As String)
  
    Call WritePrivateProfileString(INIHead, INIKey, INIVal, INIFile)
End Sub

Private Function pvGetINI(ByVal INIFile As String, ByVal INIHead As String, ByVal INIKey As String, ByVal INIDefault As String) As String

  Dim sTemp As String * 260
    
    Call GetPrivateProfileString(INIHead, INIKey, INIDefault, sTemp, Len(sTemp), INIFile)
    pvGetINI = pvStripNulls(sTemp)
End Function

Private Function pvAppPath() As String
    pvAppPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString)
End Function

Private Function pvStripNulls(ByVal sString As String) As String
    
  Dim lPos As Long
    
    pvStripNulls = sString
    
    lPos = InStr(sString, vbNullChar)
    If (lPos > 1) Then
        pvStripNulls = Left$(pvStripNulls, lPos - 1)
    ElseIf (lPos = 1) Then
        pvStripNulls = vbNullString
    End If
End Function
