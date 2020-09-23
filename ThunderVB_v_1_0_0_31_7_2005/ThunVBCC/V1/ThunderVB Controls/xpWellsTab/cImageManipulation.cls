VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageManipulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Thanks to Microsoft.


Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(1) As RGBQUAD
End Type

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'DrawIconEx Flags
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

'DIB Section constants
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs


'Raster Operation Codes
Private Const DSna = &H220326 '0x00220326

'VB Errors
Private Const giINVALID_PICTURE As Integer = 481

Private m_hpalHalftone As Long



'-------------------------------------------------------------------------
'Purpose:   Draws a transparent bitmap to a DC.  The pixels of the passed
'           bitmap that match the passed mask color will not be painted
'           to the destination DC
'In:
'   [hdcDest]
'           Device context to paint the picture on
'   [xDest]
'           X coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [yDest]
'           Y coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [Width]
'           Width of picture area to paint in pixels.  Note: If this value
'           is outrageous (i.e.: you passed a forms ScaleWidth in twips
'           instead of the pictures' width in pixels), this procedure will
'           attempt to create bitmaps that require outrageous
'           amounts of memory.
'   [Height]
'           Height of picture area to paint in pixels.  Note: If this
'           value is outrageous (i.e.: you passed a forms ScaleHeight in
'           twips instead of the pictures' height in pixels), this
'           procedure will attempt to create bitmaps that require
'           outrageous amounts of memory.
'   [hdcSrc]
'           Device context that contains the source picture
'   [xSrc]
'           X coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'   [ySrc]
'           Y coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'   [clrMask]
'           Color of pixels to be masked out
'   [hPal]
'           Handle of palette to select into the memory DC's used to create
'           the painting effect.
'           If not provided, a HalfTone palette is used.
'-------------------------------------------------------------------------
Public Sub PaintTransparentDC(ByVal hDCDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal clrMask As OLE_COLOR = 16711935, _
                                    Optional ByVal hPal As Long = 0)
    '<EhHeader>
    On Error GoTo PaintTransparentDC_Err
    '</EhHeader>
    Dim hdcMask As Long        'hDC of the created mask image
    Dim hdcColor As Long       'hDC of the created color image
    Dim hbmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hDCScreen As Long
    Dim hdcScnBuffer As Long         'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    
    hDCScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    
    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hDCScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hDCScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hDCDest, xDest, yDest, vbSrcCopy
    
    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hDCScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hDCScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTheTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hDCScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTheTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hDC.  All black (0)
    'bits are transformed to the foreground color.
    SetTheTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hDCDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer
    
    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hDCScreen
    '<EhFooter>
    Exit Sub

PaintTransparentDC_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cImageManipulation.PaintTransparentDC " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

''-------------------------------------------------------------------------
''Purpose:   Draws a transparent bitmap to a DC.  The pixels of the passed
''           bitmap that match the passed mask color will not be painted
''           to the destination DC
''In:
''   [hdcDest]
''           Device context to paint the picture on
''   [xDest]
''           X coordinate of the upper left corner of the area that the
''           picture is to be painted on. (in pixels)
''   [yDest]
''           Y coordinate of the upper left corner of the area that the
''           picture is to be painted on. (in pixels)
''   [Width]
''           Width of picture area to paint in pixels.  Note: If this value
''           is outrageous (i.e.: you passed a forms ScaleWidth in twips
''           instead of the pictures' width in pixels), this procedure will
''           attempt to create bitmaps that require outrageous
''           amounts of memory.
''   [Height]
''           Height of picture area to paint in pixels.  Note: If this
''           value is outrageous (i.e.: you passed a forms ScaleHeight in
''           twips instead of the pictures' height in pixels), this
''           procedure will attempt to create bitmaps that require
''           outrageous amounts of memory.
''   [picSource]
''           Standard Picture object to be used as the image source
''   [xSrc]
''           X coordinate of the upper left corner of the area in the picture
''           to use as the source. (in pixels)
''           Ignored if picSource is an Icon.
''   [ySrc]
''           Y coordinate of the upper left corner of the area in the picture
''           to use as the source. (in pixels)
''           Ignored if picSource is an Icon.
''   [clrMask]
''           Color of pixels to be masked out
''   [hPal]
''           Handle of palette to select into the memory DC's used to create
''           the painting effect.
''           If not provided, a HalfTone palette is used.
''-------------------------------------------------------------------------
Public Sub PaintTransparentPicture(ByVal hDCDest As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal clrMask As OLE_COLOR = 16711935, _
                                    Optional ByVal hPal As Long = 0)
    '<EhHeader>
    On Error GoTo PaintTransparentPicture_Err
    '</EhHeader>
    Dim hdcSrc As Long         'hDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hDCScreen As Long
    Dim hPalOld As Long
    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintTransparentStdPic_InvalidParam

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hDCScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            'Select passed picture into an hDC
            hdcSrc = CreateCompatibleDC(hDCScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal

            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hDCScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hDCScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hDCScreen)
            
            
            'hbmMemSrc = CreateCompatibleBitmap(hDCScreen, Width, Height)
            'hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            'hPalOld = SelectPalette(hdcSrc, hPal, True)
            'RealizePalette hdcSrc
            'Draw Icon onto DC
            'udtRect.Bottom = Height
            'udtRect.Right = Width
            'OleTranslateColor clrMask, 0&, lMaskColor
            'hbrMask = CreateSolidBrush(lMaskColor)
            'FillRect hdcSrc, udtRect, hbrMask
            'DeleteObject hbrMask
            'DrawIcon hdcSrc, 0, 0, picSource.handle
            'Draw Transparent image
            'BitBlt hDCDest, xDest, yDest, Width, Height, hdcSrc, Width, Height, vbSrcCopy
            'PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, lMaskColor, hPal
            'Clean up
            'DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            'SelectPalette hdcSrc, hPalOld, True
            'RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hDCScreen
        Case Else
            
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub
PaintTransparentStdPic_InvalidParam:
'    Err.Raise giINVALID_PICTURE
    Exit Sub
    '<EhFooter>
    Exit Sub

PaintTransparentPicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cImageManipulation.PaintTransparentPicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'CSEH: ErrAskLog
Private Sub Class_Initialize()
    '<EhHeader>
    On Error GoTo Class_Initialize_Err
    '</EhHeader>
    Dim hDCScreen As Long
    'Create halftone palette
    hDCScreen = GetDC(0&)
    m_hpalHalftone = CreateHalftonePalette(hDCScreen)
    ReleaseDC 0&, hDCScreen
    '<EhFooter>
    Exit Sub

Class_Initialize_Err:
        MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cImageManipulation.Class_Terminate " & _
           "at line " & Erl
    '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error GoTo Class_Terminate_Err
    '</EhHeader>
    DeleteObject m_hpalHalftone
    '<EhFooter>
    Exit Sub

Class_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cImageManipulation.Class_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub




