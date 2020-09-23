VERSION 5.00
Begin VB.UserControl HzxYFrame 
   Appearance      =   0  'Flat
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   ControlContainer=   -1  'True
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ToolboxBitmap   =   "HzxYFrame.ctx":0000
End
Attribute VB_Name = "HzxYFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
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

Enum fraBorderStyles
    fraNone = 0
    fraFixed_Single = 1
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

Private Enum CP
    PS_SOLID = 0
    PS_DASH = 1
    PS_DOT = 2
    PS_DASHDOT = 3
    PS_DASHDOTDOT = 4
    PS_NULL = 5
    PS_INSIDEFRAME = 6
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

Private Enum OperaRGN
    RGN_AND = 1
    RGN_OR = 2
    RGN_XOR = 3
    RGN_DIFF = 4
    RGN_COPY = 5
    RGN_MAX = RGN_COPY
    RGN_MIN = RGN_AND
End Enum

Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_BorderStyle As fraBorderStyles
Private m_BorderColor As OLE_COLOR
Private m_Caption As String
Private m_Image As StdPicture
Private m_ImageWidth As Long
Private m_ImageHeight As Long
Private CorX_Pic As Long
Private CorY_Pic As Long
Private CorXLeft_Cap As Long
Private CorXRight_Cap As Long
Private CorY_Cap As Long
Private CorY_TopLine As Long
Private CaptionHeight As Long
Private lngFormat As Long
Private CaptionRect As RECT
Private m_ControlContainedControls As Boolean

Private Const m_def_ForeColor = &HD54600
Private Const m_def_BorderColor = &HA09C98
Private Const m_def_ImageWidth = 16
Private Const m_def_ImageHeight = 16
Private Const m_def_BaseLeft = 6
'Events
Event Click()
Event DblClick()
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
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_Initialize " & _
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
    Set UserControl.Font = parent.Font
    m_BackColor = parent.BackColor
    m_ForeColor = m_def_ForeColor
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    m_ControlContainedControls = True
    m_BorderStyle = fraFixed_Single
    m_BorderColor = m_def_BorderColor
    Set m_Image = Nothing
    m_ImageWidth = m_def_ImageWidth
    m_ImageHeight = m_def_ImageHeight
    '<EhFooter>
    Exit Sub

UserControl_InitProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_InitProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
    With PropBag
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", parent.Font)
        m_BackColor = .ReadProperty("BackColor", parent.BackColor)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        UserControl.BackColor = m_BackColor
        UserControl.ForeColor = m_ForeColor
        m_ControlContainedControls = .ReadProperty("ControlContainedControls", True)
        m_BorderStyle = .ReadProperty("BorderStyle", fraFixed_Single)
        m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        Set m_Image = .ReadProperty("Image", Nothing)
        m_ImageWidth = .ReadProperty("ImageWidth", m_def_ImageWidth)
        m_ImageHeight = .ReadProperty("ImageHeight", m_def_ImageHeight)
    End With
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>
    Dim loop1 As Integer
    With PropBag
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("BackColor", m_BackColor, parent.BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
        Call .WriteProperty("ControlContainedControls", m_ControlContainedControls, True)
        Call .WriteProperty("BorderStyle", m_BorderStyle, fraFixed_Single)
        Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("Image", m_Image, Nothing)
        Call .WriteProperty("ImageWidth", m_ImageWidth, m_def_ImageWidth)
        Call .WriteProperty("ImageHeight", m_ImageHeight, m_def_ImageHeight)
    End With
    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_WriteProperties " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
    If m_BackColor <> New_BackColor Then
        m_BackColor = New_BackColor
        PropertyChanged "BackColor"
        Refresh
    End If
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get BorderStyle() As fraBorderStyles
    '<EhHeader>
    On Error GoTo BorderStyle_Err
    '</EhHeader>
    BorderStyle = m_BorderStyle
    '<EhFooter>
    Exit Property

BorderStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.BorderStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As fraBorderStyles)
    '<EhHeader>
    On Error GoTo BorderStyle_Err
    '</EhHeader>
    If m_BorderStyle <> New_BorderStyle Then
        m_BorderStyle = New_BorderStyle
        PropertyChanged "BorderStyle"
        Refresh
    End If
    '<EhFooter>
    Exit Property

BorderStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.BorderStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get BorderColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo BorderColor_Err
    '</EhHeader>
    BorderColor = m_BorderColor
    '<EhFooter>
    Exit Property

BorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.BorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo BorderColor_Err
    '</EhHeader>
    If m_BorderColor <> New_BorderColor Then
        m_BorderColor = New_BorderColor
        PropertyChanged "BorderColor"
        DrawBorder
    End If
    '<EhFooter>
    Exit Property

BorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.BorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ControlContainedControls() As Boolean
    '<EhHeader>
    On Error GoTo ControlContainedControls_Err
    '</EhHeader>
    ControlContainedControls = m_ControlContainedControls
    '<EhFooter>
    Exit Property

ControlContainedControls_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ControlContainedControls " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ControlContainedControls(ByVal New_ControlContainedControls As Boolean)
    '<EhHeader>
    On Error GoTo ControlContainedControls_Err
    '</EhHeader>
    m_ControlContainedControls = New_ControlContainedControls
    PropertyChanged "ControlContainedControls"
    '<EhFooter>
    Exit Property

ControlContainedControls_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ControlContainedControls " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Caption(NewCaption As String)
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    m_Caption = NewCaption
    PropertyChanged "Caption"
'    DrawRectangle UserControl.hdc, CorXLeft_Cap, 0, CorXRight_Cap, CorY_TopLine + 1, BreakApart(m_BackColor)
    Refresh
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
    Enabled = UserControl.Enabled
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Enabled " & _
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
        PropertyChanged "Enabled"
        DrawCaption
        DrawPicture
        DrawBorder
        Dim Control As Object
        If m_ControlContainedControls Then
            For Each Control In UserControl.ContainedControls
                Control.Enabled = New_Enabled
            Next
        Else
            For Each Control In UserControl.ContainedControls
                Control.Refresh
            Next
        End If
    End If
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Enabled " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ForeColor_Err
    '</EhHeader>
    If m_ForeColor <> New_ForeColor Then
        m_ForeColor = New_ForeColor
        PropertyChanged "ForeColor"
        DrawCaption
    End If
    '<EhFooter>
    Exit Property

ForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set Font = UserControl.Font
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Font(ByVal New_Font As Font)
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Refresh
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Font " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.FontUnderline " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    '<EhHeader>
    On Error GoTo FontUnderline_Err
    '</EhHeader>
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "Font"
    Refresh
    '<EhFooter>
    Exit Property

FontUnderline_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.FontUnderline " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.FontStrikethru " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    '<EhHeader>
    On Error GoTo FontStrikethru_Err
    '</EhHeader>
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "Font"
    Refresh
    '<EhFooter>
    Exit Property

FontStrikethru_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.FontStrikethru " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.FontSize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    '<EhHeader>
    On Error GoTo FontSize_Err
    '</EhHeader>
    UserControl.FontSize() = New_FontSize
    PropertyChanged "Font"
    Refresh
    '<EhFooter>
    Exit Property

FontSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.FontSize " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.FontName " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontName(ByVal New_FontName As String)
    '<EhHeader>
    On Error GoTo FontName_Err
    '</EhHeader>
    UserControl.FontName() = New_FontName
    PropertyChanged "Font"
    Refresh
    '<EhFooter>
    Exit Property

FontName_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.FontName " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.FontItalic " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    '<EhHeader>
    On Error GoTo FontItalic_Err
    '</EhHeader>
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "Font"
    Refresh
    '<EhFooter>
    Exit Property

FontItalic_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.FontItalic " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.FontBold " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    '<EhHeader>
    On Error GoTo FontBold_Err
    '</EhHeader>
    UserControl.FontBold() = New_FontBold
    PropertyChanged "Font"
    Refresh
    '<EhFooter>
    Exit Property

FontBold_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.FontBold " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.hdc " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.hwnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Image() As StdPicture
    '<EhHeader>
    On Error GoTo Image_Err
    '</EhHeader>
    Set Image = m_Image
    '<EhFooter>
    Exit Property

Image_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Image " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Image(ByVal NewImage As StdPicture)
    '<EhHeader>
    On Error GoTo Image_Err
    '</EhHeader>
    Set m_Image = NewImage
    PropertyChanged "Image"
    Refresh
    '<EhFooter>
    Exit Property

Image_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Image " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ImageHeight() As Long
    '<EhHeader>
    On Error GoTo ImageHeight_Err
    '</EhHeader>
    ImageHeight = m_ImageHeight
    '<EhFooter>
    Exit Property

ImageHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ImageHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ImageHeight(ByVal NewImageHeight As Long)
    '<EhHeader>
    On Error GoTo ImageHeight_Err
    '</EhHeader>
    If m_ImageHeight <> NewImageHeight Then
        m_ImageHeight = NewImageHeight
        PropertyChanged "ImageHeight"
        If Not m_Image Is Nothing Then Refresh
    End If
    '<EhFooter>
    Exit Property

ImageHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ImageHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ImageWidth() As Long
    '<EhHeader>
    On Error GoTo ImageWidth_Err
    '</EhHeader>
    ImageWidth = m_ImageWidth
    '<EhFooter>
    Exit Property

ImageWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ImageWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ImageWidth(ByVal NewImageWidth As Long)
    '<EhHeader>
    On Error GoTo ImageWidth_Err
    '</EhHeader>
    If m_ImageWidth <> NewImageWidth Then
        m_ImageWidth = NewImageWidth
        PropertyChanged "ImageWidth"
        If Not m_Image Is Nothing Then Refresh
    End If
    '<EhFooter>
    Exit Property

ImageWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ImageWidth " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.MouseIcon " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.MouseIcon " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.MousePointer " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.MousePointer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Sub UserControl_Click()
    '<EhHeader>
    On Error GoTo UserControl_Click_Err
    '</EhHeader>
    RaiseEvent Click
    '<EhFooter>
    Exit Sub

UserControl_Click_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_Click " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_DblClick()
    '<EhHeader>
    On Error GoTo UserControl_DblClick_Err
    '</EhHeader>
    RaiseEvent DblClick
    '<EhFooter>
    Exit Sub

UserControl_DblClick_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_DblClick " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseDown_Err
    '</EhHeader>
    RaiseEvent MouseDown(Button, Shift, X, Y)
    '<EhFooter>
    Exit Sub

UserControl_MouseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_MouseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseMove_Err
    '</EhHeader>
    RaiseEvent MouseMove(Button, Shift, X, Y)
    '<EhFooter>
    Exit Sub

UserControl_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseUp_Err
    '</EhHeader>
    RaiseEvent MouseUp(Button, Shift, X, Y)
    '<EhFooter>
    Exit Sub

UserControl_MouseUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_MouseUp " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_Paint " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub Refresh()
    '<EhHeader>
    On Error GoTo Refresh_Err
    '</EhHeader>
    DrawRectangle UserControl.hdc, m_def_BaseLeft + 1, 0, CorXRight_Cap, CorY_TopLine + 1, BreakApart(m_BackColor)
    CalPosition
    DrawBlock
    If Trim(m_Caption) <> "" Then DrawCaption
    If Not m_Image Is Nothing Then DrawPicture
    If m_BorderStyle = fraFixed_Single Then DrawBorder
    RoundCorners
    '<EhFooter>
    Exit Sub

Refresh_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.Refresh " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub CalPosition()
    '<EhHeader>
    On Error GoTo CalPosition_Err
    '</EhHeader>
        
    Dim tmpRect As RECT
    Dim TextSize As SIZEL
    Dim BaseLeft As Long
    
    UserControl.ScaleMode = vbPixels
    BaseLeft = m_def_BaseLeft
    
    CorY_TopLine = 0
    CorX_Pic = 0
    CorY_Pic = 0
    CorXLeft_Cap = 0
    CorXRight_Cap = 0
    CorY_Cap = 0
    
    If Not m_Image Is Nothing Then
        CorY_TopLine = m_ImageHeight \ 2
        CorX_Pic = BaseLeft + 2
        BaseLeft = CorX_Pic + m_ImageWidth
        CorXRight_Cap = BaseLeft + 2
    End If
    
    If Trim(m_Caption) <> "" Then
        CorXLeft_Cap = BaseLeft + 2
        W_GetTextExtentPoint32 UserControl.hdc, m_Caption, Len(m_Caption), TextSize
        CorXRight_Cap = CorXLeft_Cap + TextSize.cx + 2
        Call SetRect(tmpRect, CorXLeft_Cap + 1, 0, CorXRight_Cap - 1, UserControl.ScaleHeight)
        lngFormat = DT_WORDBREAK Or DT_LEFT
        CaptionHeight = W_DrawText(UserControl.hdc, m_Caption, -1, tmpRect, lngFormat Or DT_CALCRECT)
        If CaptionHeight > 1 Then
            If CaptionHeight \ 2 >= CorY_TopLine Then
                CorY_TopLine = CaptionHeight \ 2
                CorY_Pic = (CaptionHeight - m_ImageHeight) \ 2
                Call SetRect(CaptionRect, CorXLeft_Cap + 1, 0, CorXRight_Cap - 1, CaptionHeight)
            Else
                CorY_Cap = CorY_TopLine - CaptionHeight \ 2
                Call SetRect(CaptionRect, CorXLeft_Cap + 1, CorY_Cap, CorXRight_Cap - 1, CorY_Cap + CaptionHeight)
            End If
        End If
    End If

    '<EhFooter>
    Exit Sub

CalPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.CalPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawBlock()
    '<EhHeader>
    On Error GoTo DrawBlock_Err
    '</EhHeader>
    Dim Wi As Long, He As Long
    
    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
        DrawRectangle .hdc, 0, CorY_TopLine, Wi, He, BreakApart(m_BackColor)
        DrawRectangle .hdc, m_def_BaseLeft + 1, 0, CorXRight_Cap, CorY_TopLine + 1, BreakApart(m_BackColor)
    End With
    '<EhFooter>
    Exit Sub

DrawBlock_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.DrawBlock " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.DrawCaption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawPicture()
    '<EhHeader>
    On Error GoTo DrawPicture_Err
    '</EhHeader>
    
    If Not m_Image Is Nothing Then
        If UserControl.Enabled Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, m_Image
        Else
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, BreakApart(m_BackColor)
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, m_Image, Icon_Grey
        End If
    End If
    '<EhFooter>
    Exit Sub

DrawPicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.DrawPicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawBorder()
    '<EhHeader>
    On Error GoTo DrawBorder_Err
    '</EhHeader>

    Dim Color As Long
    Dim loop1 As Integer
    Dim Wi As Long, He As Long
    Dim TabLeftPos As Long
    Dim oldPen As Long, hPen As Long

    If m_BorderStyle = fraNone Then Exit Sub
    
    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
    End With
    
    Color = IIf(UserControl.Enabled, m_BorderColor, ShiftColor(&HFFFFFF, -&H3C, True))
    
    DrawLine 0, CorY_TopLine, 0, He - 1, Color
    DrawLine Wi - 1, CorY_TopLine, Wi - 1, He - 1, Color
    DrawLine 0, He - 1, Wi - 1, He - 1, Color
    
    If m_Image Is Nothing And Trim(m_Caption) = "" Then
        DrawLine 0, CorY_TopLine, Wi - 1, CorY_TopLine, Color
    ElseIf m_Image Is Nothing Then
        DrawLine 0, CorY_TopLine, CorXLeft_Cap - 2, CorY_TopLine, Color
        DrawLine CorXRight_Cap, CorY_TopLine, Wi - 1, CorY_TopLine, Color
    Else
        DrawLine 0, CorY_TopLine, m_def_BaseLeft, CorY_TopLine, Color
        DrawLine CorXRight_Cap, CorY_TopLine, Wi - 1, CorY_TopLine, Color
    End If
    
    With UserControl
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hdc, hPen)
        Arc .hdc, 0, CorY_TopLine, 8, CorY_TopLine + 8, 4, CorY_TopLine, 0, CorY_TopLine + 4
        Arc .hdc, Wi - 8, CorY_TopLine, Wi, CorY_TopLine + 8, Wi, CorY_TopLine + 4, Wi - 4, CorY_TopLine
        Arc .hdc, 0, He - 8, 8, He, 0, He - 4, 4, He
        Arc .hdc, Wi - 8, He - 8, Wi, He, Wi - 4, He, Wi, He - 4
        SelectObject .hdc, oldPen
        DeleteObject hPen
    End With

    '<EhFooter>
    Exit Sub

DrawBorder_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.DrawBorder " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub RoundCorners()
    '<EhHeader>
    On Error GoTo RoundCorners_Err
    '</EhHeader>
    Dim TempRect As Long, TempRect1 As Long, TempRect2 As Long, TempRect3 As Long
    Dim He As Long, Wi As Long
    Dim loop1 As Integer
    Dim re As Long
    Dim TabLeftPos As Long
    
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    
    TempRect = CreateRectRgn(0, 0, Wi, He)
    TempRect1 = CreateRoundRectRgn(0, CorY_TopLine - 1, Wi + 1, He + 1, 8, 8)
    TempRect2 = CreateRectRgn(0, CorY_TopLine, Wi + 1, He + 1)
    CombineRgn TempRect, TempRect2, TempRect1, RGN_AND
    DeleteObject TempRect2
    DeleteObject TempRect1
        
    If CorXRight_Cap > 0 Then
        TempRect1 = CreateRectRgn(m_def_BaseLeft, 0, CorXRight_Cap, CorY_TopLine + 1)
        CombineRgn TempRect, TempRect, TempRect1, RGN_OR
        DeleteObject TempRect1
    End If
    
    SetWindowRgn UserControl.hwnd, TempRect, True
    DeleteObject TempRect
    '<EhFooter>
    Exit Sub

RoundCorners_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.RoundCorners " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.TransBlt " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.DrawRectangle " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.BreakApart " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.getBlueVal " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.getGreenVal " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.getRedVal " & _
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
           "in ThunVBCC_v1_0.HzxYFrame.Sys2RGB " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
    '<EhHeader>
    On Error GoTo ShiftColor_Err
    '</EhHeader>
    Dim Red As Long, Blue As Long, Green As Long
    
    If Not isXP Then 'for XP button i use a work-aroud that works fine
        Value = Value \ 2 'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
        Blue = ((Color \ &H10000) Mod &H100) + Value
    Else
        Blue = ((Color \ &H10000) Mod &H100)
        Blue = Blue + ((Blue * Value) \ &HC0)
    End If
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value
    
    If Value > 0 Then
        If Red > 255 Then Red = 255
        If Green > 255 Then Green = 255
        If Blue > 255 Then Blue = 255
    ElseIf Value < 0 Then
        If Red < 0 Then Red = 0
        If Green < 0 Then Green = 0
        If Blue < 0 Then Blue = 0
    End If
    
    ShiftColor = Red + 256& * Green + 65536 * Blue
    '<EhFooter>
    Exit Function

ShiftColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.ShiftColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'Private Sub DrawLine(DstDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'¿ìËÙ»­Ïß
    '<EhHeader>
    On Error GoTo DrawLine_Err
    '</EhHeader>
    Dim pt As POINTAPI
    Dim oldPen As Long, hPen As Long

    With UserControl
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hdc, hPen)
    
        MoveToEx .hdc, X1, Y1, pt
        LineTo .hdc, X2, Y2
    
        SelectObject .hdc, oldPen
        DeleteObject hPen
    End With

    '<EhFooter>
    Exit Sub

DrawLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.HzxYFrame.DrawLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
