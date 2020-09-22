Attribute VB_Name = "mThumbnail"
Option Explicit
Option Compare Text

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

'//

Private Const LOCALE_USER_DEFAULT     As Long = &H400
Private Const LOCALE_NOUSEROVERRIDE   As Long = &H80000000
Private Const DATE_SHORTDATE          As Long = &H1

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long

'//

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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

Public Const IMAGETYPES_MASK As String = "|BMP|DIB|RLE|GIF|JPG|JPEG|JPE|JFIF|PNG|TIF|TIFF|EMF|WMF|"

Public Type DATABASE_INFO
    Size    As Long
    Entries As Long
End Type

'//

Private Type FILE_INFO
    Filename As String
    FileDate As String
    FileSize As Long
End Type

Private m_sDatabasePath As String
Private m_oDatabase     As Database
Private m_oRecordset    As Recordset
Private m_bThumbnailing As Boolean
Private m_oTile         As cTile
Private m_sFolder       As String



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeModule()
    
    '-- Database full path
    m_sDatabasePath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "Thumbs.mdb"
    '-- Open database
    Set m_oDatabase = Workspaces(0).OpenDatabase(m_sDatabasePath)
    
    '-- Initialize pattern brush
    Set m_oTile = New cTile
    Call m_oTile.CreatePatternFromStdPicture(LoadResPicture("PATTERN_4X4", vbResBitmap))
End Sub

Public Sub TerminateModule()
    
    '-- Close all
    If (Not m_oRecordset Is Nothing) Then
        Call m_oRecordset.Close
        Set m_oRecordset = Nothing
    End If
    If (Not m_oDatabase Is Nothing) Then
        Call m_oDatabase.Close
        Set m_oDatabase = Nothing
    End If
    Set m_oTile = Nothing
End Sub

Public Sub Cancel()

    '-- Cancel thumbnailing
    m_bThumbnailing = False
End Sub

Public Sub UpdateFolder(ByVal sFolder As String)
  
  Dim lItem   As Long
  Dim uFile() As FILE_INFO
  
  Dim aTBIH() As Byte, uTBIH As BITMAPINFOHEADER
  Dim aData() As Byte
    
    '-- Folder path
    Let m_sFolder = sFolder
    
    '-- Thumbnailing...
    m_bThumbnailing = True
    
    On Local Error GoTo errH
    
    With fMain.ucThumbnailView
        
        '-- Get files list
        If (pvGetFiles(IMAGETYPES_MASK, uFile())) Then
        
            '-- Disable redraw and set items count
            Call .SetRedraw(bRedraw:=False)
            Call .ItemSetCount(UBound(uFile()) + 1)
                
            '-- Add items
            For lItem = 0 To UBound(uFile())
                Call .ItemAdd(lItem, uFile(lItem).Filename, uFile(lItem).FileSize, uFile(lItem).FileDate)
            Next lItem
            
            '-- Enable redraw and ensure visible first item
            Call .ItemEnsureVisible(0)
            Call .SetRedraw(bRedraw:=True)
           
            '-- Open recordset (table -> Seek search)
            Set m_oRecordset = m_oDatabase.OpenRecordset("tblThumbnail", dbOpenTable)
            Let m_oRecordset.Index = "IDPath"
            
            '-- Add/Get thumbnails
            fMain.ucProgress.Max = UBound(uFile()) + 1
            
            For lItem = 0 To .Count - 1
                
                '-- Current progress
                fMain.ucProgress.Value = lItem + 1
                
                '-- Find item...
                Call m_oRecordset.Seek("=", m_sFolder & uFile(lItem).Filename)
                
                If (m_oRecordset.NoMatch) Then
                    
                    '-- Not found: add
                    Call m_oRecordset.AddNew
                    Call pvSetThumbnail(uFile(lItem), lItem)
                    Call m_oRecordset.Update
                    
                  Else
                    
                    '-- Found: check file date/time
                    If (m_oRecordset("Date") <> uFile(lItem).FileDate) Then
                        
                        '-- Date/time has changed: update
                        Call m_oRecordset.Edit
                        Call pvSetThumbnail(uFile(lItem), lItem)
                        Call m_oRecordset.Update
                      
                      Else
                      
                        '-- No changes: get from database
                        aTBIH() = m_oRecordset("Thumbnail").GetChunk(0, Len(uTBIH))
                        Call CopyMemory(uTBIH, aTBIH(0), Len(uTBIH))
                        aData() = m_oRecordset("Thumbnail").GetChunk(Len(uTBIH), uTBIH.biSizeImage)
                        Call .ThumbnailInfo_SetTBIH(lItem, aTBIH())
                        Call .ThumbnailInfo_SetData(lItem, aData())
                     End If
                End If
                
                '-- Refresh added item
                Call .RefreshItems(lItem, lItem)
                
                Call VBA.DoEvents
                If (Not m_bThumbnailing) Then Exit For
            Next lItem
        End If
    End With

errH:
    '-- Close recordset
    If (Not m_oRecordset Is Nothing) Then
        Call m_oRecordset.Close
        Set m_oRecordset = Nothing
    End If

    m_bThumbnailing = False
    fMain.ucProgress.Value = 0
End Sub

Public Sub UpdateItem(ByVal sPath As String, ByVal lItem As Long)
  
  Dim uFile   As FILE_INFO
  Dim uWFD    As WIN32_FIND_DATA
  Dim hSearch As Long
  
    m_sFolder = sPath
    hSearch = FindFirstFile(m_sFolder & fMain.ucThumbnailView.ItemText(lItem, [tvFileName]) & vbNullChar, uWFD)
    
    If (hSearch <> INVALID_HANDLE_VALUE) Then
        
        '-- Get file name, date and size
        With uFile
            .Filename = pvStripNulls(uWFD.cFileName)
            .FileDate = pvGetFileDateTimeStr(uWFD.ftLastWriteTime)
            .FileSize = uWFD.nFileSizeHigh * &HFFFF0000 + uWFD.nFileSizeLow
        End With
        
        '-- Search database item
        Set m_oRecordset = m_oDatabase.OpenRecordset("tblThumbnail", dbOpenTable)
        Let m_oRecordset.Index = "IDPath"
        Call m_oRecordset.Seek("=", m_sFolder & uFile.Filename)
        
        '-- Update
        If (Not m_oRecordset.NoMatch) Then
            Call m_oRecordset.Edit
            Call pvSetThumbnail(uFile, lItem)
            Call m_oRecordset.Update
            '-- Refresh item
            Call fMain.ucThumbnailView.RefreshItems(lItem, lItem)
        End If
    End If
    Call FindClose(hSearch)
    
    If (Not m_oRecordset Is Nothing) Then
        Call m_oRecordset.Close
        Set m_oRecordset = Nothing
    End If
End Sub

Public Function GetDatabaseInfo() As DATABASE_INFO
    
  On Error GoTo errH
  
    Set m_oRecordset = m_oDatabase.OpenRecordset("tblThumbnail", dbOpenTable)
    
    With GetDatabaseInfo
        .Size = FileLen(m_sDatabasePath)
        .Entries = m_oRecordset.RecordCount
    End With
    
    If (Not m_oRecordset Is Nothing) Then
        Call m_oRecordset.Close
        Set m_oRecordset = Nothing
    End If

errH:
End Function

Public Sub CheckAllThumbnails(Optional oProgress As ucProgress, Optional lblInfo As Label)
    
  Dim uWFD    As WIN32_FIND_DATA
  Dim hSearch As Long
  
  On Error GoTo errH
    
    Set m_oRecordset = m_oDatabase.OpenRecordset("tblThumbnail", dbOpenTable)
    Call m_oRecordset.MoveFirst
    
    If (Not oProgress Is Nothing) Then
        oProgress.Max = m_oRecordset.RecordCount
    End If
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Checking all thumbnails..."
        lblInfo.Refresh
    End If
    
    Do While Not m_oRecordset.EOF
        
        If (Not oProgress Is Nothing) Then
            oProgress.Value = oProgress.Value + 1
        End If
        
        hSearch = FindFirstFile(m_oRecordset("Path") & vbNullChar, uWFD)
        
        If (hSearch = INVALID_HANDLE_VALUE) Then
            Call m_oRecordset.Delete
          Else
            If (m_oRecordset("Date") <> pvGetFileDateTimeStr(uWFD.ftLastWriteTime)) Then
                Call m_oRecordset.Delete
            End If
        End If
        Call m_oRecordset.MoveNext
        
        Call FindClose(hSearch)
    Loop

    If (Not m_oRecordset Is Nothing) Then
        Call m_oRecordset.Close
        Set m_oRecordset = Nothing
    End If
    
errH:
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Done"
    End If
    If (Not oProgress Is Nothing) Then
        oProgress.Value = 0
    End If
End Sub

Public Sub DeleteAllThumbnails(Optional lblInfo As Label)

  On Error GoTo errH
    
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Deleting all thumbnails..."
        lblInfo.Refresh
    End If
    
    Call m_oDatabase.Execute("DELETE * FROM [tblThumbnail]")
    
errH:
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Done"
    End If
End Sub

Public Sub DeleteFolderThumbnails(ByVal Folder As String, Optional lblInfo As Label)

  On Error GoTo errH
    
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Deleting folder thumbnails..."
        lblInfo.Refresh
    End If

    Call m_oDatabase.Execute("DELETE * FROM [tblThumbnail] where [Path] like '" & Folder & "*'")
    
errH:
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Done"
    End If
End Sub

Public Sub CompactDatabase(Optional lblInfo As Label)
    
  Dim sOldName As String
  Dim sNewName As String
   
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Compacting database..."
        lblInfo.Refresh
    End If
    
    '-- Close and free
    Call m_oDatabase.Close
    Set m_oDatabase = Nothing
    
    '-- Temporary database name
    sOldName = m_sDatabasePath
    sNewName = Left$(m_sDatabasePath, Len(sOldName) - 1) & Chr$(126)
    
    '-- Delete old re-named (if any)
    On Error Resume Next
    Call VBA.Kill(sNewName)
    On Error GoTo 0
    
    '-- Compact...
    On Error GoTo errH
    Call DAO.DBEngine.CompactDatabase(sOldName, sNewName, dbLangGeneral, dbVersion30)
    
    '-- Delete old (if any)
    On Error Resume Next
    Call VBA.Kill(sOldName)
    On Error GoTo 0
    
    '-- Rename
    Name sNewName As sOldName
        
errH:
    If (Not lblInfo Is Nothing) Then
        lblInfo.Caption = "Done"
    End If
    Set m_oDatabase = Workspaces(0).OpenDatabase(m_sDatabasePath)
End Sub



'========================================================================================
' Private
'========================================================================================

Private Sub pvSetThumbnail( _
            ByRef uFile As FILE_INFO, _
            ByVal lItem As Long)
            
  Dim sExt      As String
  Dim oDIBThumb As cDIB
  Dim hImage    As Long
  Dim hGraphics As Long
  
  Dim bfx As Long, bfW As Long, W As Long
  Dim bfy As Long, bfH As Long, H As Long

  Dim aTBIH(39) As Byte, uTBIH As BITMAPINFOHEADER
  Dim aData()   As Byte
  
    '-- Type
    sExt = Mid$(m_sFolder & uFile.Filename, InStrRev(m_sFolder & uFile.Filename, ".") + 1)
    
    '-- Generate thumbnail...
    If (mGDIplus.GdipLoadImageFromFile(StrConv(m_sFolder & uFile.Filename, vbUnicode), hImage) = [Ok]) Then
        
        '-- Initialize DIB
        Set oDIBThumb = New cDIB

        '-- Image size
        Call mGDIplus.GdipGetImageWidth(hImage, W)
        Call mGDIplus.GdipGetImageHeight(hImage, H)
        
        '-- Best fit to current thumbnail max. size
        Call oDIBThumb.GetBestFitInfo(W, H, fMain.ucThumbnailView.ThumbnailWidth, fMain.ucThumbnailView.ThumbnailHeight, bfx, bfy, bfW, bfH)
        Call oDIBThumb.Create(bfW, bfH, [16_bpp])

        '-- Prepare target surface
        Call mGDIplus.GdipCreateFromHDC(oDIBThumb.hDC, hGraphics)
        
        '-- Tile 'transparent' layer and render thumbnail
        Call m_oTile.Tile(oDIBThumb.hDC, 0, 0, bfW, bfH)
        Call mGDIplus.GdipDrawImageRectI(hGraphics, hImage, 0, 0, bfW, bfH)
        
        '-- Clean up
        Call mGDIplus.GdipDeleteGraphics(hGraphics)
        Call mGDIplus.GdipDisposeImage(hImage)
        
        '-- Prepare bitmap header (thumbnail)
        With uTBIH
            .biSize = Len(uTBIH)
            .biBitCount = oDIBThumb.BPP
            .biWidth = oDIBThumb.Width
            .biHeight = oDIBThumb.Height
            .biSizeImage = oDIBThumb.Size
            .biPlanes = 1
        End With
        
        '-- Prepare data
        ReDim aData(oDIBThumb.Size - 1)
        Call CopyMemory(aTBIH(0), uTBIH, Len(uTBIH))
        Call CopyMemory(aData(0), ByVal oDIBThumb.lpBits, oDIBThumb.Size)
        
        '-- Transfer to database
        Call m_oRecordset("Thumbnail").AppendChunk(aTBIH())
        Call m_oRecordset("Thumbnail").AppendChunk(aData())
      
      Else
        '-- *Null* transfer to database
        ReDim aData(0)
        Call ZeroMemory(aTBIH(0), Len(uTBIH))
        Call m_oRecordset("Thumbnail").AppendChunk(aTBIH())
    End If
    
    '-- Add file path, date, image dimensions and description
    m_oRecordset("Path") = m_sFolder & uFile.Filename
    m_oRecordset("Date") = uFile.FileDate

    '-- Transfer to thumbnail viewer
    Call fMain.ucThumbnailView.ThumbnailInfo_SetTBIH(lItem, aTBIH())
    Call fMain.ucThumbnailView.ThumbnailInfo_SetData(lItem, aData())
End Sub

Private Function pvGetFiles(ByVal sMask As String, uFile() As FILE_INFO) As Boolean
  
  Dim uFileTmp()  As FILE_INFO
  Dim sExt        As String
  Dim lExtSep     As Long
  Dim lCount      As Long
  Dim lc          As Long
  
  Dim uWFD        As WIN32_FIND_DATA
  Dim hSearch     As Long
  Dim hNext       As Long
    
    '-- Initial storage
    ReDim uFileTmp(100)

    '-- Start searching files (all)
    hNext = 1
    hSearch = FindFirstFile(m_sFolder & "*.*" & vbNullChar, uWFD)
    
    If (hSearch <> INVALID_HANDLE_VALUE) Then
        
        Do While hNext
        
            If (uWFD.dwFileAttributes <> FILE_ATTRIBUTE_DIRECTORY) Then
                
                '-- Get file name, date and size
                With uFileTmp(lCount)
                    .Filename = pvStripNulls(uWFD.cFileName)
                    .FileDate = pvGetFileDateTimeStr(uWFD.ftLastWriteTime)
                    .FileSize = uWFD.nFileSizeHigh * &HFFFF0000 + uWFD.nFileSizeLow
                End With
                lCount = lCount + 1
                
                '-- Resize array [?]
                If ((lCount Mod 100) = 0) Then
                    ReDim Preserve uFileTmp(UBound(uFileTmp()) + 100)
                End If
            End If
            hNext = FindNextFile(hSearch, uWFD)
        Loop
        hNext = FindClose(hSearch)
    End If
    ReDim Preserve uFileTmp(lCount - -(lCount > 0))
    
    '-- Filter files
    If (lCount > 0) Then
        lCount = 0
        ReDim uFile(100)
        
        '-- Check all files
        For lc = 0 To UBound(uFileTmp())
        
            '-- Extension ?
            lExtSep = InStrRev(uFileTmp(lc).Filename, ".")
            If (lExtSep) Then
                
                '-- Get extension
                sExt = "|" & Mid$(uFileTmp(lc).Filename, lExtSep + 1) & "|"
                
                '-- Supported file
                If (InStr(1, sMask, sExt)) Then
                    
                    '-- Get this file
                    uFile(lCount) = uFileTmp(lc)
                    lCount = lCount + 1
                    
                    '-- Resize array [?]
                    If ((lCount Mod 100) = 0) Then
                        ReDim Preserve uFile(UBound(uFile()) + 100)
                    End If
                End If
            End If
        Next lc
        ReDim Preserve uFile(lCount - -(lCount > 0))
    End If
    
    '-- Success
    pvGetFiles = (lCount > 0)
End Function

Private Static Function pvGetFileDateTimeStr(uFileTime As FILETIME) As String
  
  Dim uFT As FILETIME
  Dim uST As SYSTEMTIME

    Call FileTimeToLocalFileTime(uFileTime, uFT)
    Call FileTimeToSystemTime(uFT, uST)
  
    pvGetFileDateTimeStr = pvGetFileDateStr(uST) & " " & pvGetFileTimeStr(uST)
End Function

Private Static Function pvGetFileDateStr(uSystemTime As SYSTEMTIME) As String
  
  Dim sDate As String * 32
  Dim lLen  As Long
  
    lLen = GetDateFormat(LOCALE_USER_DEFAULT, LOCALE_NOUSEROVERRIDE Or DATE_SHORTDATE, uSystemTime, vbNullString, sDate, 64)
    If (lLen) Then
        pvGetFileDateStr = Left$(sDate, lLen - 1)
    End If
End Function

Private Static Function pvGetFileTimeStr(uSystemTime As SYSTEMTIME) As String
  
  Dim sTime As String * 32
  Dim lLen  As Long
  
    lLen = GetTimeFormat(LOCALE_USER_DEFAULT, LOCALE_NOUSEROVERRIDE, uSystemTime, vbNullString, sTime, 64)
    If (lLen) Then
        pvGetFileTimeStr = Left$(sTime, lLen - 1)
    End If
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
