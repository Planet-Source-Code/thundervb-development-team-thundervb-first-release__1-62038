VERSION 5.00
Begin VB.UserControl HzxYOption 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ClipControls    =   0   'False
   FillColor       =   &H8000000F&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000F&
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "HzxYOption.ctx":0000
End
Attribute VB_Name = "HzxYOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long



Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

Private Enum OptionState
    FalseNormal = 0
    TrueNormal = 1
    FalseDisabled = 2
    TrueDisabled = 3
    FalseOver = 4
    TrueOver = 5
    FalseDown = 6
    TrueDown = 7
End Enum

Private Enum picScaleMe
    vbUser = 0
    vbTwips = 1
    vbPoints = 2
    vbPixels = 3
    vbCharacters = 4
    vbInches = 5
    vbMillimeters = 6
    vbCentimeters = 7
    vbHimetric = 8
    vbContainerPosition = 9
    vbContainerSize = 10
End Enum

Private Enum DT
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_CHARSTREAM = 4
    DT_DISPFILE = 6
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_METAFILE = 5
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_PLOTTER = 0
    DT_RASCAMERA = 3
    DT_RASDISPLAY = 1
    DT_RASPRINTER = 2
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
    DT_WORD_ELLIPSIS = &H40000
    DT_END_ELLIPSIS = 32768
    DT_PATH_ELLIPSIS = &H4000
    DT_EDITCONTROL = &H2000
    '===================
    DT_INCENTER = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
End Enum

Private Enum CPR
    ALTERNATE = 1
    BDR_SUNKENINNER = &H8
    BDR_RAISEDOUTER = &H1
    BDR_RAISEDINNER = &H4
    BDR_SUNKENOUTER = &H2
    BF_LEFT = &H1
    BF_RIGHT = &H4
    BF_TOP = &H2
    BF_BOTTOM = &H8
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
End Enum

Private Enum IconStates
    Icon_Normal = 0
    Icon_Grey = 1
    Icon_Disabled = 2
End Enum

Private Enum IconDrawMe
    DI_MASK = &H1
    DI_IMAGE = &H2
    DI_NORMAL = DI_MASK Or DI_IMAGE
End Enum

Private m_Value As Boolean
Private m_Caption As String
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_State As OptionState
Private optImage(7) As StdPicture
Private CorX_Pic As Long
Private CorY_Pic As Long
Private CorX_Cap As Long
Private CorY_Cap As Long
Private CaptionHeight As Long
Private lngFormat As Long
Private CaptionRect As RECT

Private Const m_def_State = FalseNormal

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_Initialize()
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>
    UserControl.ScaleMode = vbPixels
    UserControl.PaletteMode = vbPaletteModeContainer
    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_Initialize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_InitProperties()
    '<EhHeader>
    On Error GoTo UserControl_InitProperties_Err
    '</EhHeader>
    m_Caption = Ambient.DisplayName
    Enabled = True
    m_Value = False
    Set UserControl.Font = Ambient.Font
    m_BackColor = parent.BackColor
    m_ForeColor = parent.ForeColor
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    m_State = m_def_State
    Dim loop1 As Integer
    For loop1 = LBound(optImage) To UBound(optImage)
        Set optImage(loop1) = Nothing
    Next
    '<EhFooter>
    Exit Sub

UserControl_InitProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_InitProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
    With PropBag
        m_Value = .ReadProperty("Value", False)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_BackColor = .ReadProperty("BackColor", parent.BackColor)
        m_ForeColor = .ReadProperty("ForeColor", parent.ForeColor)
        UserControl.BackColor = m_BackColor
        UserControl.ForeColor = m_ForeColor
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        Set optImage(0) = .ReadProperty("Pic_FalseNormal", Nothing)
        Set optImage(1) = .ReadProperty("Pic_TrueNormal", Nothing)
        Set optImage(2) = .ReadProperty("Pic_FalseDisabled", Nothing)
        Set optImage(3) = .ReadProperty("Pic_TrueDisabled", Nothing)
        Set optImage(4) = .ReadProperty("Pic_FalseOver", Nothing)
        Set optImage(5) = .ReadProperty("Pic_TrueOver", Nothing)
        Set optImage(6) = .ReadProperty("Pic_FalseDown", Nothing)
        Set optImage(7) = .ReadProperty("Pic_TrueDown", Nothing)
    End With
    m_State = IIf(m_Value, TrueNormal, FalseNormal)
    If Enabled = False Then m_State = m_State + FalseDisabled
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>
    With PropBag
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("Value", m_Value, False)
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("BackColor", m_BackColor, parent.BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, parent.ForeColor)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("Pic_FalseNormal", optImage(0), Nothing)
        Call .WriteProperty("Pic_TrueNormal", optImage(1), Nothing)
        Call .WriteProperty("Pic_FalseDisabled", optImage(2), Nothing)
        Call .WriteProperty("Pic_TrueDisabled", optImage(3), Nothing)
        Call .WriteProperty("Pic_FalseOver", optImage(4), Nothing)
        Call .WriteProperty("Pic_TrueOver", optImage(5), Nothing)
        Call .WriteProperty("Pic_FalseDown", optImage(6), Nothing)
        Call .WriteProperty("Pic_TrueDown", optImage(7), Nothing)
    End With
    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_WriteProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Get BackColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
    BackColor = m_BackColor
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
    DrawPicture m_State
    DrawCaption
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Caption() As String
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    Caption = m_Caption
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Caption(ByVal New_Caption As String)
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Refresh
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Enabled() As Boolean
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
    Enabled = UserControl.Enabled
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Enabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
    If New_Enabled <> UserControl.Enabled Then
        UserControl.Enabled() = New_Enabled
        m_State = IIf(New_Enabled, m_State Mod 2, 2 + (m_State Mod 2))
        DrawCaption
        DrawPicture m_State
    End If
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Enabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ForeColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo ForeColor_Err
    '</EhHeader>
    ForeColor = m_ForeColor
    '<EhFooter>
    Exit Property

ForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ForeColor_Err
    '</EhHeader>
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    DrawCaption
    '<EhFooter>
    Exit Property

ForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_TrueNormal() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_TrueNormal_Err
    '</EhHeader>
    Set Pic_TrueNormal = optImage(1)
    '<EhFooter>
    Exit Property

Pic_TrueNormal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueNormal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_TrueNormal(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_TrueNormal_Err
    '</EhHeader>
    Set optImage(1) = newPic
    PropertyChanged "Pic_TrueNormal"
    If m_State = TrueNormal Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_TrueNormal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueNormal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_TrueDisabled() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_TrueDisabled_Err
    '</EhHeader>
    Set Pic_TrueDisabled = optImage(3)
    '<EhFooter>
    Exit Property

Pic_TrueDisabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueDisabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_TrueDisabled(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_TrueDisabled_Err
    '</EhHeader>
    Set optImage(3) = newPic
    PropertyChanged "Pic_TrueDisabled"
    If m_State = TrueDisabled Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_TrueDisabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueDisabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_TrueDown() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_TrueDown_Err
    '</EhHeader>
    Set Pic_TrueDown = optImage(7)
    '<EhFooter>
    Exit Property

Pic_TrueDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_TrueDown(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_TrueDown_Err
    '</EhHeader>
    Set optImage(7) = newPic
    PropertyChanged "Pic_TrueDown"
    If m_State = TrueDown Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_TrueDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_TrueOver() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_TrueOver_Err
    '</EhHeader>
    Set Pic_TrueOver = optImage(5)
    '<EhFooter>
    Exit Property

Pic_TrueOver_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueOver " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_TrueOver(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_TrueOver_Err
    '</EhHeader>
    Set optImage(5) = newPic
    PropertyChanged "Pic_TrueOver"
    If m_State = TrueOver Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_TrueOver_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_TrueOver " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_FalseNormal() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_FalseNormal_Err
    '</EhHeader>
    Set Pic_FalseNormal = optImage(0)
    '<EhFooter>
    Exit Property

Pic_FalseNormal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseNormal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_FalseNormal(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_FalseNormal_Err
    '</EhHeader>
    Set optImage(0) = newPic
    PropertyChanged "Pic_FalseNormal"
    If m_State = FalseNormal Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_FalseNormal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseNormal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_FalseDisabled() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_FalseDisabled_Err
    '</EhHeader>
    Set Pic_FalseDisabled = optImage(2)
    '<EhFooter>
    Exit Property

Pic_FalseDisabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseDisabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_FalseDisabled(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_FalseDisabled_Err
    '</EhHeader>
    Set optImage(2) = newPic
    PropertyChanged "Pic_FalseDisabled"
    If m_State = FalseDisabled Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_FalseDisabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseDisabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_FalseDown() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_FalseDown_Err
    '</EhHeader>
    Set Pic_FalseDown = optImage(6)
    '<EhFooter>
    Exit Property

Pic_FalseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_FalseDown(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_FalseDown_Err
    '</EhHeader>
    Set optImage(6) = newPic
    PropertyChanged "Pic_FalseDown"
    If m_State = FalseDown Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_FalseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Pic_FalseOver() As StdPicture
    '<EhHeader>
    On Error GoTo Pic_FalseOver_Err
    '</EhHeader>
    Set Pic_FalseOver = optImage(4)
    '<EhFooter>
    Exit Property

Pic_FalseOver_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseOver " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Pic_FalseOver(ByVal newPic As StdPicture)
    '<EhHeader>
    On Error GoTo Pic_FalseOver_Err
    '</EhHeader>
    Set optImage(4) = newPic
    PropertyChanged "Pic_FalseOver"
    If m_State = FalseOver Then DrawPicture m_State
    '<EhFooter>
    Exit Property

Pic_FalseOver_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Pic_FalseOver " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Value() As Boolean
    '<EhHeader>
    On Error GoTo Value_Err
    '</EhHeader>
    Value = m_Value
    '<EhFooter>
    Exit Property

Value_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Value " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Value(ByVal vNewValue As Boolean)
    '<EhHeader>
    On Error GoTo Value_Err
    '</EhHeader>
    If m_Value <> vNewValue Then
        m_Value = vNewValue
        PropertyChanged "Value"
        If m_Value Then
            m_State = 2 * Int(m_State / 2) + 1
        Else
            m_State = 2 * Int(m_State / 2)
        End If
        DrawPicture m_State
        If m_Value Then ContainerCheck
    End If
    '<EhFooter>
    Exit Property

Value_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Value " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Font() As Font
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set Font = UserControl.Font
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Font(ByVal New_Font As Font)
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set UserControl.Font = New_Font
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontUnderline() As Boolean
    '<EhHeader>
    On Error GoTo FontUnderline_Err
    '</EhHeader>
    FontUnderline = UserControl.FontUnderline
    '<EhFooter>
    Exit Property

FontUnderline_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontUnderline " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    '<EhHeader>
    On Error GoTo FontUnderline_Err
    '</EhHeader>
    UserControl.FontUnderline() = New_FontUnderline
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

FontUnderline_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontUnderline " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontStrikethru() As Boolean
    '<EhHeader>
    On Error GoTo FontStrikethru_Err
    '</EhHeader>
    FontStrikethru = UserControl.FontStrikethru
    '<EhFooter>
    Exit Property

FontStrikethru_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontStrikethru " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    '<EhHeader>
    On Error GoTo FontStrikethru_Err
    '</EhHeader>
    UserControl.FontStrikethru() = New_FontStrikethru
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

FontStrikethru_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontStrikethru " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontSize() As Single
    '<EhHeader>
    On Error GoTo FontSize_Err
    '</EhHeader>
    FontSize = UserControl.FontSize
    '<EhFooter>
    Exit Property

FontSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontSize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    '<EhHeader>
    On Error GoTo FontSize_Err
    '</EhHeader>
    UserControl.FontSize() = New_FontSize
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

FontSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontSize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontName() As String
    '<EhHeader>
    On Error GoTo FontName_Err
    '</EhHeader>
    FontName = UserControl.FontName
    '<EhFooter>
    Exit Property

FontName_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontName " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontName(ByVal New_FontName As String)
    '<EhHeader>
    On Error GoTo FontName_Err
    '</EhHeader>
    UserControl.FontName() = New_FontName
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

FontName_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontName " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontItalic() As Boolean
    '<EhHeader>
    On Error GoTo FontItalic_Err
    '</EhHeader>
    FontItalic = UserControl.FontItalic
    '<EhFooter>
    Exit Property

FontItalic_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontItalic " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    '<EhHeader>
    On Error GoTo FontItalic_Err
    '</EhHeader>
    UserControl.FontItalic() = New_FontItalic
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

FontItalic_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontItalic " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontBold() As Boolean
    '<EhHeader>
    On Error GoTo FontBold_Err
    '</EhHeader>
    FontBold = UserControl.FontBold
    '<EhFooter>
    Exit Property

FontBold_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontBold " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    '<EhHeader>
    On Error GoTo FontBold_Err
    '</EhHeader>
    UserControl.FontBold() = New_FontBold
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

FontBold_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.FontBold " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get hdc() As Long
    '<EhHeader>
    On Error GoTo hdc_Err
    '</EhHeader>
    hdc = UserControl.hdc
    '<EhFooter>
    Exit Property

hdc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.hdc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get hwnd() As Long
    '<EhHeader>
    On Error GoTo hwnd_Err
    '</EhHeader>
    hwnd = UserControl.hwnd
    '<EhFooter>
    Exit Property

hwnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.hwnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get MouseIcon() As StdPicture
    '<EhHeader>
    On Error GoTo MouseIcon_Err
    '</EhHeader>
    Set MouseIcon = UserControl.MouseIcon
    '<EhFooter>
    Exit Property

MouseIcon_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.MouseIcon " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    '<EhHeader>
    On Error GoTo MouseIcon_Err
    '</EhHeader>
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
    '<EhFooter>
    Exit Property

MouseIcon_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.MouseIcon " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get MousePointer() As MousePointerConstants
    '<EhHeader>
    On Error GoTo MousePointer_Err
    '</EhHeader>
    MousePointer = UserControl.MousePointer
    '<EhFooter>
    Exit Property

MousePointer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.MousePointer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    '<EhHeader>
    On Error GoTo MousePointer_Err
    '</EhHeader>
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    '<EhFooter>
    Exit Property

MousePointer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.MousePointer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Sub UserControl_Click()
    '<EhHeader>
    On Error GoTo UserControl_Click_Err
    '</EhHeader>
    If Not Value Then Value = True
    RaiseEvent Click
    '<EhFooter>
    Exit Sub

UserControl_Click_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_Click " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_DblClick()
    '<EhHeader>
    On Error GoTo UserControl_DblClick_Err
    '</EhHeader>
    m_State = IIf(Value, TrueDown, FalseDown)
    DrawPicture m_State
    '<EhFooter>
    Exit Sub

UserControl_DblClick_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_DblClick " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyDown_Err
    '</EhHeader>
    RaiseEvent KeyDown(KeyCode, Shift)
    '<EhFooter>
    Exit Sub

UserControl_KeyDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_KeyDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyPress_Err
    '</EhHeader>
    RaiseEvent KeyPress(KeyAscii)
    '<EhFooter>
    Exit Sub

UserControl_KeyPress_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_KeyPress " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyUp_Err
    '</EhHeader>
    RaiseEvent KeyUp(KeyCode, Shift)
    '<EhFooter>
    Exit Sub

UserControl_KeyUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_KeyUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseDown_Err
    '</EhHeader>
    If Enabled = True Then
        If Value Then
            m_State = TrueDown
        Else
            m_State = FalseDown
        End If
    End If
    DrawPicture m_State
    RaiseEvent MouseDown(Button, Shift, X, Y)
    '<EhFooter>
    Exit Sub

UserControl_MouseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_MouseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
    
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseMove_Err
    '</EhHeader>
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hwnd
    If PointInControl(X, Y) Then
        If m_State < FalseOver Then
            If Button = vbLeftButton Then
                If Value Then
                    m_State = TrueDown
                Else
                    m_State = FalseDown
                End If
            Else
                If Value Then
                    m_State = TrueOver
                Else
                    m_State = FalseOver
                End If
            End If
            DrawPicture m_State
        End If
    Else
        If Value Then
            m_State = TrueNormal
        Else
            m_State = FalseNormal
        End If
        DrawPicture m_State
        RaiseEvent MouseOut
        ReleaseCapture
    End If
    '<EhFooter>
    Exit Sub

UserControl_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Function PointInControl(X As Single, Y As Single) As Boolean
    '<EhHeader>
    On Error GoTo PointInControl_Err
    '</EhHeader>
  If X >= 0 And X <= UserControl.ScaleWidth And _
    Y >= 0 And Y <= UserControl.ScaleHeight Then
    PointInControl = True
  End If
    '<EhFooter>
    Exit Function

PointInControl_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.PointInControl " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseUp_Err
    '</EhHeader>
    If Button = vbLeftButton Then
        m_State = m_State - 2
        DrawPicture m_State
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
    '<EhFooter>
    Exit Sub

UserControl_MouseUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_MouseUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Paint()
    '<EhHeader>
    On Error GoTo UserControl_Paint_Err
    '</EhHeader>
    Refresh
    '<EhFooter>
    Exit Sub

UserControl_Paint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_Paint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
    Refresh
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error GoTo UserControl_Terminate_Err
    '</EhHeader>
    Dim loop1 As Integer
    For loop1 = LBound(optImage) To UBound(optImage)
        Set optImage(loop1) = Nothing
    Next
    '<EhFooter>
    Exit Sub

UserControl_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.UserControl_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub Refresh()
    '<EhHeader>
    On Error GoTo Refresh_Err
    '</EhHeader>
    UserControl.Cls
    CalPosition
    DrawCaption
    DrawPicture m_State
    '<EhFooter>
    Exit Sub

Refresh_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Refresh " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub CalPosition()
    '<EhHeader>
    On Error GoTo CalPosition_Err
    '</EhHeader>
        
    Dim tmpRect As RECT
    Dim TextSize As SIZEL
    
    UserControl.ScaleMode = vbPixels
    CorX_Pic = 0

    With UserControl
        W_GetTextExtentPoint32 .hdc, m_Caption, Len(m_Caption), TextSize
        .Width = (TextSize.cx + 17) * 15
        Call SetRect(tmpRect, 17, 0, .ScaleWidth, .ScaleHeight)
    End With
    lngFormat = DT_WORDBREAK Or DT_LEFT
    CaptionHeight = W_DrawText(UserControl.hdc, m_Caption, -1, tmpRect, lngFormat Or DT_CALCRECT)
    If CaptionHeight > 1 Then
        With UserControl
            .Height = IIf(CaptionHeight >= 13, CaptionHeight * 15, 195)
            Call SetRect(CaptionRect, 17, Int((.ScaleHeight - CaptionHeight) / 2), .ScaleWidth, Int((.ScaleHeight + CaptionHeight) / 2))
            CorY_Pic = .ScaleHeight \ 2 - 6
        End With
    End If
    '<EhFooter>
    Exit Sub

CalPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.CalPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawCaption()
    '<EhHeader>
    On Error GoTo DrawCaption_Err
    '</EhHeader>
    Dim TmpRGBColor1 As Long
    
    If UserControl.Enabled Then
        If Trim(m_Caption) > 0 Then
            TmpRGBColor1 = BreakApart(m_ForeColor)
            SetTextColor UserControl.hdc, TmpRGBColor1
            W_DrawText UserControl.hdc, m_Caption, -1, CaptionRect, lngFormat
        End If
    Else
        If Trim(m_Caption) > 0 Then
            TmpRGBColor1 = BreakApart(&H80000011)
            SetTextColor UserControl.hdc, TmpRGBColor1
            W_DrawText UserControl.hdc, m_Caption, -1, CaptionRect, lngFormat
        End If
    End If
    '<EhFooter>
    Exit Sub

DrawCaption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.DrawCaption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawPicture(CurState As OptionState)
    '<EhHeader>
    On Error GoTo DrawPicture_Err
    '</EhHeader>
    
    Dim Str As String
    Dim tempPic As StdPicture
    
    Select Case CurState
    Case FalseNormal
        Str = "optFalseNormal"
    Case TrueNormal
        Str = "optTrueNormal"
    Case FalseDisabled
        Str = "optFalseDisabled"
    Case TrueDisabled
        Str = "optTrueDisabled"
    Case FalseOver
        Str = "optFalseOver"
    Case TrueOver
        Str = "optTrueOver"
    Case FalseDown
        Str = "optFalseDown"
    Case TrueDown
        Str = "optTrueDown"
    End Select

    Select Case CurState
    Case FalseNormal, TrueNormal
        If Not optImage(CurState) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState)
        Else
            Set tempPic = LoadResPicture(Str, vbResIcon)
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, tempPic
            Set tempPic = Nothing
        End If
    Case FalseDisabled, TrueDisabled
        If Not optImage(CurState) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState)
        ElseIf Not optImage(CurState Mod 2) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState Mod 2), Icon_Grey
        Else
            Set tempPic = LoadResPicture(Str, vbResIcon)
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, tempPic
            Set tempPic = Nothing
        End If
    Case FalseOver, TrueOver, FalseDown, TrueDown
        If Not optImage(CurState) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState)
        ElseIf Not optImage(CurState Mod 2) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState Mod 2)
        Else
            Set tempPic = LoadResPicture(Str, vbResIcon)
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, tempPic
            Set tempPic = Nothing
        End If
    End Select
    '<EhFooter>
    Exit Sub

DrawPicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.DrawPicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal IconState As IconStates = Icon_Normal, Optional ByVal ShadowColor As Long = -1)
    '<EhHeader>
    On Error GoTo TransBlt_Err
    '</EhHeader>
    
    If DstW = 0 Or DstH = 0 Then Exit Sub
    
    Dim OriW As Long, OriH As Long
    Dim SrcDC As Long, SrcRect As RECT, SrcBmp As Long, SrcObj As Long

    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
    Dim ToBeChange As Boolean
    Dim loopx As Long, loopy As Long
    Dim i As Long, iTop As Long, iLeft As Long
    Dim DisabledRGB As RGBTRIPLE, HighLightRGB As RGBTRIPLE, ShadowRGB As RGBTRIPLE
    Dim HaveChanged As Boolean

    OriW = UserControl.ScaleX(SrcPic.Width, vbHimetric, vbPixels)
    OriH = UserControl.ScaleY(SrcPic.Height, vbHimetric, vbPixels)
    
    Select Case IconState
    Case Icon_Normal
        Select Case SrcPic.Type
        Case vbPicTypeBitmap
            SrcDC = CreateCompatibleDC(hdc)
            SrcBmp = CreateCompatibleBitmap(hdc, DstW, DstH)
            SrcObj = SelectObject(SrcDC, SrcPic)
            
            StretchBlt DstDC, DstX, DstY, DstW, DstH, SrcDC, 0, 0, OriW, OriH, vbSrcCopy
            
'            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteObject SrcBmp
            DeleteDC SrcDC
        Case vbPicTypeIcon
            DrawIconEx DstDC, DstX, DstY, SrcPic.handle, DstW, DstH, 0, 0, DI_NORMAL
        End Select
    
    Case Icon_Disabled
        
        Const cShadow = &H808080
        Const cHighLight = &HFFFFFF
        
        Select Case SrcPic.Type
        Case vbPicTypeBitmap
            DrawRectangle DstDC, DstX, DstY, DstW, DstH, cShadow
            Dim tmpRect As RECT
            tmpRect.Left = DstX
            tmpRect.Right = DstX + DstW
            tmpRect.Top = DstY
            tmpRect.Bottom = DstY + DstH
            DrawEdge DstDC, tmpRect, BDR_SUNKENINNER, BF_RECT
        Case vbPicTypeIcon
            SrcDC = CreateCompatibleDC(DstDC)
            SrcBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
            SrcObj = SelectObject(SrcDC, SrcBmp)
            BitBlt SrcDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
        
            TmpDC = CreateCompatibleDC(SrcDC)
            TmpBmp = CreateCompatibleBitmap(SrcDC, DstW, DstH)
            TmpObj = SelectObject(TmpDC, TmpBmp)
            BitBlt SrcDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
            DrawIconEx TmpDC, 0, 0, SrcPic.handle, DstW, DstH, 0, 0, DI_NORMAL
            
            ReDim Data1(DstW * DstH * 3 - 1)
            ReDim Data2(UBound(Data1))
            With Info.bmiHeader
                .biSize = Len(Info.bmiHeader)
                .biWidth = DstW
                .biHeight = DstH
                .biPlanes = 1
                .biBitCount = 24
            End With
    
            GetDIBits SrcDC, SrcBmp, 0, DstH, Data1(0), Info, 0
            GetDIBits TmpDC, TmpBmp, 0, DstH, Data2(0), Info, 0
            
            With DisabledRGB
                .rgbBlue = (cShadow \ &H10000) Mod &H100
                .rgbGreen = (cShadow \ &H100) Mod &H100
                .rgbRed = cShadow And &HFF
            End With
            
            With HighLightRGB
                .rgbBlue = (cHighLight \ &H10000) Mod &H100
                .rgbGreen = (cHighLight \ &H100) Mod &H100
                .rgbRed = cHighLight And &HFF
            End With
    
            For loopy = 0 To DstH - 1
                For loopx = DstW - 1 To 0 Step -1
                    i = loopy * DstW + loopx
                    If Data2(i).rgbRed = Data1(i).rgbRed And Data2(i).rgbGreen = Data1(i).rgbGreen And Data2(i).rgbBlue = Data1(i).rgbBlue Then '±³¾°É«
                        HaveChanged = False
                        If loopy < DstH - 1 Then
                            iTop = (loopy + 1) * DstW + loopx
                            If Data2(iTop).rgbRed <> Data1(iTop).rgbRed Or Data2(iTop).rgbGreen <> Data1(iTop).rgbGreen Or Data2(iTop).rgbBlue <> Data1(iTop).rgbBlue Then
                                HaveChanged = True
                                Data2(i) = HighLightRGB
                            End If
                        End If
                        If loopx > 0 And (Not HaveChanged) Then
                            iLeft = i - 1
                            If Data2(iLeft).rgbRed <> Data1(iLeft).rgbRed Or Data2(iLeft).rgbGreen <> Data1(iLeft).rgbGreen Or Data2(iLeft).rgbBlue <> Data1(iLeft).rgbBlue Then
                                Data2(i) = HighLightRGB
                            End If
                        End If
                    Else
                        Data2(i) = DisabledRGB
                    End If
                Next
            Next

            SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0), Info, 0

            Erase Data1, Data2
            DeleteObject SelectObject(TmpDC, TmpObj)
            DeleteObject TmpBmp
            DeleteDC TmpDC
            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteObject SrcBmp
            DeleteDC SrcDC
        
        End Select
        
    Case Icon_Grey
        
        If ShadowColor <> -1 Then
            With ShadowRGB
                .rgbBlue = (cShadow \ &H10000) Mod &H100
                .rgbGreen = (cShadow \ &H100) Mod &H100
                .rgbRed = cShadow And &HFF
            End With
        End If
        
        Select Case SrcPic.Type
        Case vbPicTypeBitmap
            SrcDC = CreateCompatibleDC(DstDC)
            SrcObj = SelectObject(SrcDC, SrcPic)
            
            TmpDC = CreateCompatibleDC(SrcDC)
            TmpBmp = CreateCompatibleBitmap(SrcDC, DstW, DstH)
            TmpObj = SelectObject(TmpDC, TmpBmp)
            StretchBlt TmpDC, 0, 0, DstW, DstH, SrcDC, 0, 0, OriW, OriH, vbSrcCopy
        
            ReDim Data2(DstW * DstH * 3 - 1)
            With Info.bmiHeader
                .biSize = Len(Info.bmiHeader)
                .biWidth = DstW
                .biHeight = DstH
                .biPlanes = 1
                .biBitCount = 24
            End With
            
            GetDIBits TmpDC, TmpBmp, 0, DstH, Data2(0), Info, 0
        
            For loopy = 0 To DstH - 1
                For loopx = DstW - 1 To 0 Step -1
                    i = loopy * DstW + loopx
                    If ShadowColor <> -1 Then
                        Data2(i) = ShadowRGB
                    Else
                        With Data2(i)
                            gCol = CLng(.rgbRed * 0.3) + .rgbGreen * 0.59 + .rgbBlue * 0.11
                            .rgbRed = gCol
                            .rgbGreen = gCol
                            .rgbBlue = gCol
                        End With
                    End If
                Next
            Next
        
            SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0), Info, 0
        
            Erase Data2
            DeleteObject SelectObject(TmpDC, TmpObj)
            DeleteObject TmpBmp
            DeleteDC TmpDC
'            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteDC SrcDC
        Case vbPicTypeIcon
            SrcDC = CreateCompatibleDC(DstDC)
            SrcBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
            SrcObj = SelectObject(SrcDC, SrcBmp)
            BitBlt SrcDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
        
            TmpDC = CreateCompatibleDC(SrcDC)
            TmpBmp = CreateCompatibleBitmap(SrcDC, DstW, DstH)
            TmpObj = SelectObject(TmpDC, TmpBmp)
            BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
            DrawIconEx TmpDC, 0, 0, SrcPic.handle, DstW, DstH, 0, 0, DI_NORMAL
            
            ReDim Data1(DstW * DstH * 3 - 1)
            ReDim Data2(UBound(Data1))
            With Info.bmiHeader
                .biSize = Len(Info.bmiHeader)
                .biWidth = DstW
                .biHeight = DstH
                .biPlanes = 1
                .biBitCount = 24
            End With
    
            GetDIBits SrcDC, SrcBmp, 0, DstH, Data1(0), Info, 0
            GetDIBits TmpDC, TmpBmp, 0, DstH, Data2(0), Info, 0
            
            For loopy = 0 To DstH - 1
                For loopx = DstW - 1 To 0 Step -1
                    i = loopy * DstW + loopx
                    If Data2(i).rgbRed <> Data1(i).rgbRed Or Data2(i).rgbGreen <> Data1(i).rgbGreen Or Data2(i).rgbBlue <> Data1(i).rgbBlue Then
                        If ShadowColor <> -1 Then
                            Data2(i) = ShadowRGB
                        Else
                            With Data2(i)
                                gCol = CLng(.rgbRed * 0.3) + .rgbGreen * 0.59 + .rgbBlue * 0.11
                                .rgbRed = gCol
                                .rgbGreen = gCol
                                .rgbBlue = gCol
                            End With
                        End If
                    End If
                Next
            Next
        
            SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0), Info, 0
        
            Erase Data1, Data2
            DeleteObject SelectObject(TmpDC, TmpObj)
            DeleteObject TmpBmp
            DeleteDC TmpDC
            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteObject SrcBmp
            DeleteDC SrcDC
        End Select
    
    End Select

    '<EhFooter>
    Exit Sub

TransBlt_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.TransBlt " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawRectangle(DstDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
    '<EhHeader>
    On Error GoTo DrawRectangle_Err
    '</EhHeader>

    Dim bRECT As RECT
    Dim hBrush As Long

    bRECT.Left = X
    bRECT.Top = Y
    bRECT.Right = X + Width
    bRECT.Bottom = Y + Height

    hBrush = CreateSolidBrush(Color)

    If OnlyBorder Then
        FrameRect DstDC, bRECT, hBrush
    Else
        FillRect DstDC, bRECT, hBrush
    End If

    DeleteObject hBrush
    '<EhFooter>
    Exit Sub

DrawRectangle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.DrawRectangle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Function BreakApart(ByVal Color As Long) As Long
    '<EhHeader>
    On Error GoTo BreakApart_Err
    '</EhHeader>
    Dim R As Integer, G As Integer, B As Integer
    R = getRedVal(Color)
    G = getGreenVal(Color)
    B = getBlueVal(Color)
    BreakApart = RGB(R, G, B)
    '<EhFooter>
    Exit Function

BreakApart_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.BreakApart " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Function getBlueVal(ByVal RGBCol As Long) As Integer
    '<EhHeader>
    On Error GoTo getBlueVal_Err
    '</EhHeader>
    RGBCol = Sys2RGB(RGBCol)
    If RGBCol < 0 Then RGBCol = 0
    getBlueVal = (RGBCol And &HFF0000) / &H10000
    '<EhFooter>
    Exit Function

getBlueVal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.getBlueVal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Function getGreenVal(ByVal RGBCol As Long) As Integer
    '<EhHeader>
    On Error GoTo getGreenVal_Err
    '</EhHeader>
    RGBCol = Sys2RGB(RGBCol)
    If RGBCol < 0 Then RGBCol = 0
    getGreenVal = ((RGBCol And &H100FF00) / &H100)
    '<EhFooter>
    Exit Function

getGreenVal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.getGreenVal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Function getRedVal(ByVal RGBCol As Long) As Integer
    '<EhHeader>
    On Error GoTo getRedVal_Err
    '</EhHeader>
    RGBCol = Sys2RGB(RGBCol)
    If RGBCol < 0 Then RGBCol = 0
    getRedVal = RGBCol And &HFF
    '<EhFooter>
    Exit Function

getRedVal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.getRedVal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Function Sys2RGB(RGBCol As Long) As Long
    '<EhHeader>
    On Error GoTo Sys2RGB_Err
    '</EhHeader>
    If RGBCol < 0 Then
        OleTranslateColor RGBCol, 0&, Sys2RGB
    Else
        Sys2RGB = RGBCol
    End If
    '<EhFooter>
    Exit Function

Sys2RGB_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.Sys2RGB " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Sub ContainerCheck()
    '<EhHeader>
    On Error GoTo ContainerCheck_Err
    '</EhHeader>
    Dim Control As Object
    For Each Control In UserControl.parent.Controls
        If TypeOf Control Is HzxYOption Then
            If Control.Container.hwnd = UserControl.ContainerHwnd Then
                If Control.hdc <> UserControl.hdc Then
                    If Control.Value = True Then Control.Value = False
                End If
            End If
        End If
    Next
    '<EhFooter>
    Exit Sub

ContainerCheck_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYOption.ContainerCheck " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


