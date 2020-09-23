VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====================================================================
' Filename: cTile.Cls
' Author:   Steve McMahon
' Date:     15 June 1999
'
' Tiles a picture over the specified area in a DC
'
' --------------------------------------------------------------------
' vbAccelerator - Advanced, Free Source Code:
' http://vbaccelerator.com/
' ====================================================================

Private m_lHdc As Long
Private m_lHBmp As Long
Private m_lHBmpOld As Long
Private m_lhPalOld As Long
Private m_pic As StdPicture
Private m_sFileName As String
Private m_lXOriginOffset As Long
Private m_lYOriginOffset As Long
Private m_lBitmapW As Long
Private m_lBitmapH As Long

Private Const cTileErrorBase = 5600

Public Property Get PicturehDC() As Long
    '<EhHeader>
    On Error GoTo PicturehDC_Err
    '</EhHeader>
    PicturehDC = m_lHdc
    '<EhFooter>
    Exit Property

PicturehDC_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.PicturehDC " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get XOriginOffset() As Long
    '<EhHeader>
    On Error GoTo XOriginOffset_Err
    '</EhHeader>
    XOriginOffset = m_lXOriginOffset
    '<EhFooter>
    Exit Property

XOriginOffset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.XOriginOffset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let XOriginOffset(ByVal lPixels As Long)
    '<EhHeader>
    On Error GoTo XOriginOffset_Err
    '</EhHeader>
    m_lXOriginOffset = lPixels
    '<EhFooter>
    Exit Property

XOriginOffset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.XOriginOffset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get YOriginOffset() As Long
    '<EhHeader>
    On Error GoTo YOriginOffset_Err
    '</EhHeader>
    YOriginOffset = m_lYOriginOffset
    '<EhFooter>
    Exit Property

YOriginOffset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.YOriginOffset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let YOriginOffset(ByVal lPiYels As Long)
    '<EhHeader>
    On Error GoTo YOriginOffset_Err
    '</EhHeader>
    m_lYOriginOffset = lPiYels
    '<EhFooter>
    Exit Property

YOriginOffset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.YOriginOffset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get BitmapWidth() As Long
    '<EhHeader>
    On Error GoTo BitmapWidth_Err
    '</EhHeader>
    BitmapWidth = m_lBitmapW
    '<EhFooter>
    Exit Property

BitmapWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.BitmapWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get BitmapHeight() As Long
    '<EhHeader>
    On Error GoTo BitmapHeight_Err
    '</EhHeader>
    BitmapHeight = m_lBitmapH
    '<EhFooter>
    Exit Property

BitmapHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.BitmapHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Private Sub pErr(lNumber As Long, smsg As String)
    '<EhHeader>
    On Error GoTo pErr_Err
    '</EhHeader>
    MsgBox "Error: " & smsg & ", " & lNumber, vbExclamation
    '<EhFooter>
    Exit Sub

pErr_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.pErr " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Let Filename( _
        ByVal sFileName As String _
    )
    ' Load a picture from a file:
    '<EhHeader>
    On Error GoTo Filename_Err
    '</EhHeader>
    If (m_sFileName <> sFileName) Then
        pClearUp
        If (pbLoadPicture(sFileName)) Then
            m_sFileName = sFileName
        End If
    End If
    '<EhFooter>
    Exit Property

Filename_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.Filename " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Filename() As String
    '<EhHeader>
    On Error GoTo Filename_Err
    '</EhHeader>
    Filename = m_sFileName
    '<EhFooter>
    Exit Property

Filename_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.Filename " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Picture() As StdPicture
    '<EhHeader>
    On Error GoTo Picture_Err
    '</EhHeader>
    Set Picture = m_pic
    '<EhFooter>
    Exit Property

Picture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.Picture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Picture(oPic As StdPicture)
   ' Load a picture from a StdPicture object:
    '<EhHeader>
    On Error GoTo Picture_Err
    '</EhHeader>
   pClearUp
   If Not oPic Is Nothing Then
      If (pbEnsurePicture()) Then
          Set m_pic = oPic
          If (Err.Number = 0) Then
              pbGetBitmapIntoDC
          End If
      End If
   End If
    '<EhFooter>
    Exit Property

Picture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.Picture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Private Function pbEnsurePicture() As Boolean
    '<EhHeader>
    On Error GoTo pbEnsurePicture_Err
    '</EhHeader>
On Error Resume Next
    pbEnsurePicture = True
    If (m_pic Is Nothing) Then
        Set m_pic = New StdPicture
        If (Err.Number <> 0) Then
            pErr 3, "Unable to allocate memory for picture object."
            pbEnsurePicture = False
        Else
        End If
    End If
On Error GoTo 0
    Exit Function
    '<EhFooter>
    Exit Function

pbEnsurePicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.pbEnsurePicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Function pbLoadPictureFromFile(sFile As String) As Boolean
    '<EhHeader>
    On Error GoTo pbLoadPictureFromFile_Err
    '</EhHeader>
On Error Resume Next
   Set m_pic = LoadPicture(sFile)
   If (Err.Number <> 0) Then
       pErr 0, "Load Picture Failed: " & Err.Description
   Else
       pbLoadPictureFromFile = True
   End If
On Error GoTo 0
    Exit Function
    '<EhFooter>
    Exit Function

pbLoadPictureFromFile_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.pbLoadPictureFromFile " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Function pbLoadPicture(sFile As String) As Boolean
    '<EhHeader>
    On Error GoTo pbLoadPicture_Err
    '</EhHeader>

    If (pbEnsurePicture()) Then
        If (pbLoadPictureFromFile(sFile)) Then
            pbLoadPicture = pbGetBitmapIntoDC()
        End If
    End If
    
    '<EhFooter>
    Exit Function

pbLoadPicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.pbLoadPicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Function pbGetBitmapIntoDC() As Boolean
    '<EhHeader>
    On Error GoTo pbGetBitmapIntoDC_Err
    '</EhHeader>
Dim tB As BITMAP
Dim lHDC As Long, lHwnd As Long
Dim lHDCTemp As Long, lHBmpTempOld As Long

    ' Make a DC to hold the picture bitmap which we can blt from:
    lHwnd = GetDesktopWindow()
    lHDC = GetDC(lHwnd)
    m_lHdc = CreateCompatibleDC(lHDC)
    lHDCTemp = CreateCompatibleDC(lHDC)
    If (m_lHdc <> 0) Then
        ' Get size of bitmap:
        A_GetObject m_pic.handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth
        m_lBitmapH = tB.bmHeight
        lHBmpTempOld = SelectObject(lHDCTemp, m_pic.handle)
        
        m_lHBmp = CreateCompatibleBitmap(lHDC, m_lBitmapW, m_lBitmapH)
        m_lHBmpOld = SelectObject(m_lHdc, m_lHBmp)
        
        BitBlt m_lHdc, 0, 0, m_lBitmapW, m_lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
        
        SelectObject lHDCTemp, lHBmpTempOld
        DeleteDC lHDCTemp
        
        If (m_lHBmpOld <> 0) Then
            pbGetBitmapIntoDC = True
            If (m_sFileName = "") Then
               m_sFileName = "PICTURE"
            End If
        Else
            pClearUp
            pErr 2, "Unable to select bitmap into DC"
        End If
    Else
        pErr 1, "Unable to create compatible DC"
    End If
    ReleaseDC lHwnd, lHDC
    
    '<EhFooter>
    Exit Function

pbGetBitmapIntoDC_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.pbGetBitmapIntoDC " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Public Property Get Palette() As StdPicture
    '<EhHeader>
    On Error GoTo Palette_Err
    '</EhHeader>
    Set Palette = m_pic
    '<EhFooter>
    Exit Property

Palette_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.Palette " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Private Sub pClearUp()
    ' Clear reference to the filename:
    '<EhHeader>
    On Error GoTo pClearUp_Err
    '</EhHeader>
    m_sFileName = ""
    ' If we have a DC, then clear up:
    If (m_lHdc <> 0) Then
        ' Select the bitmap out of DC:
        If (m_lHBmpOld <> 0) Then
            SelectObject m_lHdc, m_lHBmpOld
            ' The original bitmap does not have to deleted because it is owned by m_pic
        End If
        If (m_lHBmp <> 0) Then
            DeleteObject m_lHBmp
        End If
        ' Remove the DC:
        DeleteDC m_lHdc
    End If
    '<EhFooter>
    Exit Sub

pClearUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.pClearUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub TileArea( _
        ByRef hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long _
    )
    '<EhHeader>
    On Error GoTo TileArea_Err
    '</EhHeader>
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((X + m_lXOriginOffset) Mod m_lBitmapW)
    lSrcStartY = ((Y + m_lYOriginOffset) Mod m_lBitmapH)
    lSrcStartWidth = (m_lBitmapW - lSrcStartX)
    lSrcStartHeight = (m_lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = Y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, m_lHdc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = m_lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = m_lBitmapH
    Loop
    '<EhFooter>
    Exit Sub

TileArea_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.TileArea " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub Class_Terminate()
    ' Ensure all GDI objects are freed:
    '<EhHeader>
    On Error GoTo Class_Terminate_Err
    '</EhHeader>
    pClearUp
    ' Clear up the picture:
    Set m_pic = Nothing
    '<EhFooter>
    Exit Sub

Class_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTile.Class_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

