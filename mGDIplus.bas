Attribute VB_Name = "mGDIplus"
' From great stuff:
'
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'   by Avery
'
'   Platform SDK Redistributable: GDI+ RTM
'   http://www.microsoft.com/downloads/release.asp?releaseid=32738

Option Explicit

Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Public Enum GpStatus
    [Ok] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

Public Enum GpUnit
    [UnitWorld]
    [UnitDisplay]
    [UnitPixel]
    [UnitPoint]
    [UnitInch]
    [UnitDocument]
    [UnitMillimeter]
End Enum

Public Enum InterpolationMode
    [InterpolationModeInvalid] = -1
    [InterpolationModeDefault]
    [InterpolationModeLowQuality]
    [InterpolationModeHighQuality]
    [InterpolationModeBilinear]
    [InterpolationModeBicubic]
    [InterpolationModeNearestNeighbor]
    [InterpolationModeHighQualityBilinear]
    [InterpolationModeHighQualityBicubic]
End Enum

Public Enum PixelOffsetMode
    [PixelOffsetModeInvalid] = -1
    [PixelOffsetModeDefault]
    [PixelOffsetModeHighSpeed]
    [PixelOffsetModeHighQuality]
    [PixelOffsetModeNone]
    [PixelOffsetModeHalf]
End Enum

Public Enum QualityMode
    [QualityModeInvalid] = -1
    [QualityModeDefault]
    [QualityModeLow]
    [QualityModeHigh]
End Enum

Public Enum ImageLockMode
    [ImageLockModeRead] = &H1
    [ImageLockModeWrite] = &H2
    [ImageLockModeUserInputBuf] = &H4
End Enum

Public Enum RotateFlipType
    [RotateNoneFlipNone] = 0
    [Rotate90FlipNone] = 1
    [Rotate180FlipNone] = 2
    [Rotate270FlipNone] = 3
    [RotateNoneFlipX] = 4
    [Rotate90FlipX] = 5
    [Rotate180FlipX] = 6
    [Rotate270FlipX] = 7
    [RotateNoneFlipY] = Rotate180FlipX
    [Rotate90FlipY] = Rotate270FlipX
    [Rotate180FlipY] = RotateNoneFlipX
    [Rotate270FlipY] = Rotate90FlipX
    [RotateNoneFlipXY] = Rotate180FlipNone
    [Rotate90FlipXY] = Rotate270FlipNone
    [Rotate180FlipXY] = RotateNoneFlipNone
    [Rotate270FlipXY] = Rotate90FlipNone
End Enum

'//

Public Const PixelFormat24bppRGB      As Long = &H21808

Public Const PropertyTagTypeByte      As Long = 1
Public Const PropertyTagTypeASCII     As Long = 2
Public Const PropertyTagTypeShort     As Long = 3
Public Const PropertyTagTypeLong      As Long = 4
Public Const PropertyTagTypeRational  As Long = 5
Public Const PropertyTagTypeUndefined As Long = 7
Public Const PropertyTagTypeSLONG     As Long = 9
Public Const PropertyTagTypeSRational As Long = 10

Public Const PropertyTagFrameDelay    As Long = &H5100
Public Const PropertyTagLoopCount     As Long = &H5101

Public Const FrameDimensionTime       As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
Public Const FrameDimensionResolution As String = "{84236F7B-3BD3-428F-8DAB-4EA1439CA315}"
Public Const FrameDimensionPage       As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"

'//

Public Type BITMAPDATA
    Width       As Long
    Height      As Long
    Stride      As Long
    PixelFormat As Long
    Scan0       As Long
    Reserved    As Long
End Type

Public Type RECTL
    x As Long
    y As Long
    W As Long
    H As Long
End Type

Public Type CLSID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Public Type PropertyItem
    propId As Long
    Length As Long
    Type   As Integer
    Value  As Long
End Type
    
Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus

Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As String, hImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, Height As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, hBitmap As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus

Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As InterpolationMode) As GpStatus
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As PixelOffsetMode) As GpStatus

Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, RECT As RECTL, ByVal Flags As Long, ByVal PixelFormat As Long, LockedBitmapData As BITMAPDATA) As GpStatus
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, LockedBitmapData As BITMAPDATA) As GpStatus

Public Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal hImage As Long, ByVal rfType As RotateFlipType) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus

Public Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, dimensionID As CLSID, Count As Long) As GpStatus
Public Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, dimensionID As CLSID, ByVal frameIndex As Long) As GpStatus
Public Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, Size As Long) As GpStatus
Public Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, ByVal propSize As Long, buffer As PropertyItem) As GpStatus

'//

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

Private Type ARGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type


Public Function ColorARGB(ByVal Color As Long, ByVal Alpha As Byte) As Long
  
  Dim uARGB As ARGBQUAD
  Dim aSwap As Byte

   Call CopyMemory(uARGB, Color, 4)
   With uARGB
        .A = Alpha: aSwap = .R: .R = .B: .B = aSwap
   End With
   Call CopyMemory(ColorARGB, uARGB, 4)
End Function



'========================================================================================
' Helpers
'========================================================================================

Public Function GetPropertyValue(Item As PropertyItem) As Variant
   
    If (Item.Value = 0 Or Item.Length = 0) Then Call Err.Raise(5, "GetPropertyValue")

    '-- We'll make Undefined types a Btye array as it seems the safest choice...
    Select Case Item.Type
        
        Case PropertyTagTypeByte, PropertyTagTypeUndefined
         
            ReDim buffByte(1 To Item.Length) As Byte
            Call CopyMemory(buffByte(1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffByte()
            Erase buffByte()

        Case PropertyTagTypeASCII
         
            GetPropertyValue = PtrToStrA(Item.Value)
         
        Case PropertyTagTypeShort
         
            ReDim buffShort(1 To (Item.Length / 2)) As Integer
            Call CopyMemory(buffShort(1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffShort()
            Erase buffShort()
         
        Case PropertyTagTypeLong, PropertyTagTypeSLONG
         
            ReDim buffLong(1 To (Item.Length / 4)) As Long
            Call CopyMemory(buffLong(1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffLong()
            Erase buffLong()
         
        Case PropertyTagTypeRational, PropertyTagTypeSRational
         
            ReDim buffLongPair(1 To (Item.Length / 8), 1 To 2) As Long
            Call CopyMemory(buffLongPair(1, 1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffLongPair()
            Erase buffLongPair()

        Case Else
            
            Call Err.Raise(461, "GetPropertyValue")
    End Select
End Function

Public Sub DEFINE_GUID(ByVal sGuid As String, uCLSID As CLSID)
    
    Call CLSIDFromString(StrPtr(sGuid), uCLSID)
End Sub

Public Function StretchDIB24Ex( _
                oDIB24 As cDIB, _
                ByVal hDC As Long, _
                ByVal x As Long, ByVal y As Long, _
                ByVal nWidth As Long, ByVal nHeight As Long, _
                Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, _
                Optional ByVal Interpolate As Boolean = False _
                ) As Long

  Dim gplRet As Long
  
  Dim hGraphics As Long
  Dim hBitmap   As Long
  Dim bmpRect   As RECTL
  Dim bmpData   As BITMAPDATA
  
    If (oDIB24.BPP = 24) Then
        
        If (nSrcWidth = 0) Then nSrcWidth = oDIB24.Width
        If (nSrcHeight = 0) Then nSrcHeight = oDIB24.Height
      
        '-- Prepare image info
        With bmpRect
            .W = oDIB24.Width
            .H = oDIB24.Height
        End With
        With bmpData
            .Width = oDIB24.Width
            .Height = oDIB24.Height
            .Stride = -oDIB24.BytesPerScanline
            .PixelFormat = [PixelFormat24bppRGB]
            .Scan0 = oDIB24.lpBits - .Stride * (oDIB24.Height - 1)
        End With
        
        '-- Initialize Graphics object
        gplRet = GdipCreateFromHDC(hDC, hGraphics)
        
        '-- Initialize blank Bitmap and assign DIB data
        gplRet = GdipCreateBitmapFromScan0(oDIB24.Width, oDIB24.Height, 0, [PixelFormat24bppRGB], ByVal 0, hBitmap)
        gplRet = GdipBitmapLockBits(hBitmap, bmpRect, [ImageLockModeWrite] Or [ImageLockModeUserInputBuf], [PixelFormat24bppRGB], bmpData)
        gplRet = GdipBitmapUnlockBits(hBitmap, bmpData)

        '-- Render
        gplRet = GdipSetInterpolationMode(hGraphics, [InterpolationModeNearestNeighbor] + -(2 * Interpolate))
        gplRet = GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHighQuality])
        gplRet = GdipDrawImageRectRectI(hGraphics, hBitmap, x, y, nWidth, nHeight, xSrc, ySrc, nSrcWidth, nSrcHeight, [UnitPixel], 0)
        
        '-- Clean up
        gplRet = GdipDeleteGraphics(hGraphics)
        gplRet = GdipDisposeImage(hBitmap)
        
        '-- Success
        StretchDIB24Ex = (gplRet = [Ok])
    End If
End Function

'//

Private Function PtrToStrW(ByVal lpsz As Long) As String
  
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        PtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function

Private Function PtrToStrA(ByVal lpsz As Long) As String
  
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenA(lpsz)

    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        PtrToStrA = sOut
    End If
End Function
