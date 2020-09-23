VERSION 5.00
Begin VB.UserControl ISCombo 
   BackColor       =   &H80000005&
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   ToolboxBitmap   =   "ISCombo.ctx":0000
   Begin VB.Timer timUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   660
   End
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1860
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1980
      Width           =   435
   End
   Begin ThunVBCC_v1.vbaRichEdit txtText 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Version         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ViewMode        =   0
      Border          =   0   'False
      ControlRightMargin=   100000000
      AutoURLDetect   =   0   'False
      TextOnly        =   -1  'True
      SingleLine      =   -1  'True
      ScrollBars      =   0
   End
   Begin VB.Image ImgItem 
      Height          =   195
      Left            =   300
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   375
   End
End
Attribute VB_Name = "ISCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      ControlName:    ISCombo.
''      Version:        2.10
''      Author:         Alfredo Córdova Pérez ( fred_cpp )
''      e-mail:         fred_cpp@hotmail.com
''                      fred_cpp@yahoo.com.mx
''
''      Description:
''
''      This is the second Release of the ISCombo Control.
''      This is a Custom ImageCombo, that supports some aditional Features Like:
''      *Style:
''          Version 2.10 Support Multiple Styles:
''          * Normal:       Classic ImageCombo Style
''          * MSO2000:      Flat Office 2000 Style
''          * MSOXP:        Flat Office XP Style
''          * WINXP:        Windows XP style
''
''      * KeyBoard Navigation Supported
''          All Keyboard functionality is now supported
''          (Including AltKey + DownKey to show the Scroll Window)
''
''      * Keep Bookmark.
''          Automatically The previous Selection Is Displayed when user activates Scroll Window
''
''      * Text Align:
''          Supports This Property like TextBoxes
''
''      * DefaultIcon
''          Is the Icon that is Shown When there isn't a Selected Item.
''
''      * Backcolor
''          You can select the Backcolor for all styles
''
''      * HoverColor
''          You can select also the Hovercolor in all styles
''
''      * MSOXPColor
''          The Color for the Button in MSOXPStyle
''
''      * MSOXPHoverColor
''          The HoverColor for the button in the MSOXPStyle
''
''      * WINXPColor
''          The color for the Button in WINXPStyle
''
''      * WINXPHoverColor
''          The HoverColor for the Button in WINXPStyle
''
''      * WINXPBorderColor
''          The Color in the Border in WINXPStyle
''
''      * DropDownListBackColor
''          The color for the DropDownList in all styles
''
''      * DropDownListHoverColor
''          The color for the selected items in DropDownList in all styles
''
''      * DropDownListBorderColor
''          The color for the Border in DropDownList in all styles
''
''      * FontColor
''          The color for the text in all styles
''
''      * FontHighLightColor
''          The HighLight Color fot the text in all styles.
''
''      * RestoreOriginalColors
''          Set all Colors to the default value.
''
''      * Autocomplete
''          Now Complete, The problem is that The control doesn't show the dropdown list
''          when the autocomplte is on process
''
''      Special thanks:
''      Charles P.V.    :Lot of tips
''      Lucifer         :Almost all autocomplete routines.
''      Chad            :
''      you know, you can use this freely, just give me credit.
''      Votes and suggestions are wellcome.
''

Option Explicit

'************************************************************'
'*                                                          *'
'*  Declaratios (Structures, Constants and API)             *'
'*                                                          *'
'************************************************************'

' Type Declarations

Public Enum ISAlign
    AlignLeft
    AlignRight
    AlignCenter
End Enum

Private Enum State
    Normal
    Hover
    pushed
    disabled
End Enum

Public Enum iscStyle
    ISNormal
    ISMSO2000
    ISMSOXP
    ISWINXP
End Enum

''Color Constants

Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_ADJ_MAX = 100
Private Const COLOR_ADJ_MIN = -100
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8

' 3D border styles
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

' Border flags
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10

'' Private Variables
Private InOut As Boolean
Private iState As State
Private OnClicking As Boolean
Private OnFocus As Boolean
Private bPressed As Boolean
Private bPreserve As Boolean
Private bListIsVisible As Boolean
Private gScaleX As Single
Private gScaleY As Single
Private bRead As Byte
Private bInput As Boolean
Private m_Focused As Boolean
Private m_ImageSize As Integer
Private m_Items As New Collection
Private m_Images As New Collection
Private m_ItemsCount As Integer
Private m_Autocomplete As Boolean
Private m_Editable As Boolean
Private m_SelectedItem As Integer

Private WithEvents cDown As wndDown
Attribute cDown.VB_VarHelpID = -1

'Default Property Values:
Const m_def_Enabled = True
Const m_def_Autocomplete = True
Const m_def_IconAlign = 0
Const m_def_IconSize = 0
Const m_def_TextAlign = 4
Const m_def_BackColor = &HFFFFFF
Const m_def_HoverColor = &HFFFFFF
'Const m_def_Default = False
'Const m_def_Editable = False
Const m_def_Style = 1
Const m_def_FontColor = 0
Const m_def_FontHighlightColor = &H80000
Const m_def_MSOXPColor = &HC08080
Const m_def_MSOXPHoverColor = &H800000
Const m_def_WINXPColor = &HFF8D6F
Const m_def_WINXPHoverColor = &HFF9D7F
Const m_def_WINXPBorderColor = &HB99D7F
Const m_def_DropDownListBackColor = &H80000005
Const m_def_DropDownListHoverColor = &H8000000D
Const m_def_DropDownListBorderColor = 0
Const m_def_DropDownListIconsBackColor = &H80000005

'Property Variables:
Dim m_Enabled As Boolean
Dim m_FontColor As OLE_COLOR
Dim m_FontHighlightColor As OLE_COLOR
Dim m_TextAlign As ISAlign
Dim m_Icon As Picture
Dim m_BackColor As OLE_COLOR
Dim m_HoverColor As OLE_COLOR
Dim m_Style As iscStyle
Dim m_MSOXPColor As OLE_COLOR
Dim m_MSOXPHoverColor As OLE_COLOR
Dim m_WINXPColor As OLE_COLOR
Dim m_WINXPHoverColor As OLE_COLOR
Dim m_WINXPBorderColor As OLE_COLOR
Dim m_DropDownListBackColor As OLE_COLOR
Dim m_DropDownListHoverColor As OLE_COLOR
Dim m_DropDownListBorderColor As OLE_COLOR
Dim m_DropDownListIconsBackColor As OLE_COLOR

'Event Declarations:
Event ItemClick(iItem As Integer)
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut()
Event MouseHover()
Event KeyPress(KeyAscii As Integer)
Event ButtonClick()
Event Change()
Const pBorderColor = &HC08080

' API Declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage2W Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Const CB_SHOWDROPDOWN = 335

'************************************************************'
'*                                                          *'
'*  Autocomplete routines                                   *'
'*                                                          *'
'***********************************************************

Function Search(Text As String) As Integer
    '<EhHeader>
    On Error GoTo Search_Err
    '</EhHeader>
    Dim a, i
    a = m_Items.Count

    For i = 1 To a
        If Text = m_Items.Item(i) Then
            'Item Exists
            Search = 1
            Exit For
        End If
    Next i
    '<EhFooter>
    Exit Function

Search_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Search " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Function Complete() As Integer
    '<EhHeader>
    On Error GoTo Complete_Err
    '</EhHeader>
    Dim a, ni
    a = m_Items.Count

    If txtText.Text <> "" And bInput = False Then
        bRead = Len(txtText.Text)
        For ni = 1 To a
            If LCase(txtText.Text) = LCase(m_Items.Item(ni)) Then
                Exit Function
            ElseIf LCase(txtText.Text) = LCase(Left(m_Items.Item(ni), bRead)) Then
                bInput = True
                txtText.SetFocus
                txtText.Text = m_Items.Item(ni)
                txtText.SetSelection bRead, Len(txtText.Text) - bRead - 1
                'txtText.SelStart = bRead
                'txtText.SelLength = Len(txtText.Text) - bRead
                m_SelectedItem = ni - 1
                'This should open the Dropdown List, but this combo is diferent, so I need Code all manually :(
                'SendMessage2 txtText.hwnd, CB_SHOWDROPDOWN, 1, txtText.text
                'I'll edit Dropdown Functionality
                If Not cDown Is Nothing Then cDown.m_bPreserve = False
                Exit For
            End If
        Next ni
        m_SelectedItem = ni - 1
    Else
        bInput = False
    End If
    '<EhFooter>
    Exit Function

Complete_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Complete " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'************************************************************'
'*                                                          *'
'*  Drawing Routines                                        *'
'*                                                          *'
'***********************************************************

Private Sub APIFillRect(hdc As Long, rc As RECT, Color As Long)
    '<EhHeader>
    On Error GoTo APIFillRect_Err
    '</EhHeader>
  Dim OldBrush As Long
  Dim NewBrush As Long
  
  NewBrush& = CreateSolidBrush(Color&)
  Call FillRect(hdc&, rc, NewBrush&)
  Call DeleteObject(NewBrush&)
    '<EhFooter>
    Exit Sub

APIFillRect_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.APIFillRect " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Function APILine(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional Color As OLE_COLOR = -1) As Long
    '<EhHeader>
    On Error GoTo APILine_Err
    '</EhHeader>
    Dim hPen As Long, hPenOld As Long
    Dim R
    Dim pt As POINTAPI
    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X1, Y1, pt
    LineTo hdc, X2, Y2
    SelectObject hdc, hPenOld
    DeleteObject hPen
    '<EhFooter>
    Exit Function

APILine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.APILine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Function APIRectangle(ByVal hdc As Long, rtRect As RECT, Optional Color As OLE_COLOR = -1) As Long
    '<EhHeader>
    On Error GoTo APIRectangle_Err
    '</EhHeader>
    Dim hPen As Long, hPenOld As Long
    Dim R
    Dim pt As POINTAPI
    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, rtRect.Left, rtRect.Top, pt
    LineTo hdc, rtRect.Right, rtRect.Top
    LineTo hdc, rtRect.Right, rtRect.Bottom
    LineTo hdc, rtRect.Left, rtRect.Bottom
    LineTo hdc, rtRect.Left, rtRect.Top
    SelectObject hdc, hPenOld
    DeleteObject hPen
    '<EhFooter>
    Exit Function

APIRectangle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.APIRectangle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'' Make Soft a color
Function SoftColor(lColor As OLE_COLOR) As OLE_COLOR
    '<EhHeader>
    On Error GoTo SoftColor_Err
    '</EhHeader>
    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lR As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
    lR = (lColor And &HFF)
    lg = ((lColor And 65280) \ 256)
    lb = ((lColor) And 16711680) \ 65536
    lRed = (76 - Int(((lColor And &HFF) + 32) \ 64) * 19)
    lGreen = (76 - Int((((lColor And 65280) \ 256) + 32) \ 64) * 19)
    lBlue = (76 - Int((((lColor And &HFF0000) \ &H10000) + 32) / 64) * 19)
    SoftColor = RGB(lR + lRed, lg + lGreen, lb + lBlue)
    '<EhFooter>
    Exit Function

SoftColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.SoftColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

''  Offset a color
Function OffsetColor(lColor As OLE_COLOR, lOffset As Long) As OLE_COLOR
    '<EhHeader>
    On Error GoTo OffsetColor_Err
    '</EhHeader>
    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lR As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
    lR = (lColor And &HFF)
    lg = ((lColor And 65280) \ 256)
    lb = ((lColor) And 16711680) \ 65536
    lRed = (lOffset + lR)
    lGreen = (lOffset + lg)
    lBlue = (lOffset + lb)
    If lRed > 255 Then lRed = 255
    If lRed < 0 Then lRed = 0
    If lGreen > 255 Then lGreen = 255
    If lGreen < 0 Then lGreen = 0
    If lBlue > 255 Then lBlue = 255
    If lBlue < 0 Then lBlue = 0
    OffsetColor = RGB(lRed, lGreen, lBlue)
    '<EhFooter>
    Exit Function

OffsetColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.OffsetColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'' Detect If the Mouse is over a Window Rect
Private Function InBox(ObjectHWnd As Long) As Boolean
    '<EhHeader>
    On Error GoTo InBox_Err
    '</EhHeader>
    Dim mpos As POINTAPI
    Dim oRect As RECT
    GetCursorPos mpos
    GetWindowRect ObjectHWnd, oRect
    If mpos.X >= oRect.Left And mpos.X <= oRect.Right And _
        mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
        InBox = True
    Else
        InBox = False
   End If
    '<EhFooter>
    Exit Function

InBox_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.InBox " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'************************************************************'
'*                                                          *'
'*  Event Procedures                                        *'
'*                                                          *'
'***********************************************************

Private Sub cDown_ItemClick(iItem As Integer, sText As String)
    '<EhHeader>
    On Error GoTo cDown_ItemClick_Err
    '</EhHeader>
    UserControl.ImgItem.Picture = m_Images(iItem + 1)
    txtText.Text = sText
    txtText.SetSelection 0, Len(sText)
    'txtText.SelStart = 0
    'txtText.SelLength = Len(sText)
    m_SelectedItem = iItem
    UserControl.SetFocus
    iState = Hover
    bListIsVisible = False
    Unload cDown
    Set cDown = Nothing
    RaiseEvent ItemClick(iItem)
    '<EhFooter>
    Exit Sub

cDown_ItemClick_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.cDown_ItemClick " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cDown_Hide()
    '<EhHeader>
    On Error GoTo cDown_Hide_Err
    '</EhHeader>
    Unload cDown
    Set cDown = Nothing
    bListIsVisible = False
    '<EhFooter>
    Exit Sub

cDown_Hide_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.cDown_Hide " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ImgItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo ImgItem_MouseMove_Err
    '</EhHeader>
    UserControl_MouseMove Button, Shift, X, Y
    '<EhFooter>
    Exit Sub

ImgItem_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.ImgItem_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo picButton_MouseDown_Err
    '</EhHeader>
    OnClicking = True
    If Button = vbLeftButton Then
        bPressed = True
        iState = pushed
        DrawFace
    End If
    '<EhFooter>
    Exit Sub

picButton_MouseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.picButton_MouseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo picButton_MouseMove_Err
    '</EhHeader>
        UserControl_MouseMove Button, Shift, X, Y
    '<EhFooter>
    Exit Sub

picButton_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.picButton_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo picButton_MouseUp_Err
    '</EhHeader>
    OnClicking = False
    bPressed = False
    DrawFace
    '<EhFooter>
    Exit Sub

picButton_MouseUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.picButton_MouseUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub picButton_Paint()
    'DrawFace
    '<EhHeader>
    On Error GoTo picButton_Paint_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

picButton_Paint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.picButton_Paint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub timUpdate_Timer()
    '<EhHeader>
    On Error GoTo timUpdate_Timer_Err
    '</EhHeader>
    If InBox(UserControl.hwnd) Then
        If InOut = False Then
            iState = Hover
            UserControl_Paint
            RaiseEvent MouseHover
        Else
            iState = Normal
        End If
        InOut = True
    Else
        If InOut Then
            timUpdate.Enabled = False
            If OnFocus Then
                iState = Hover
            Else
                iState = Normal
            End If
            UserControl_Paint
            RaiseEvent MouseOut
        End If
        InOut = False
    End If
    '<EhFooter>
    Exit Sub

timUpdate_Timer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.timUpdate_Timer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtText_Change()
    '<EhHeader>
    On Error GoTo txtText_Change_Err
    '</EhHeader>
    If m_Autocomplete Then
        Dim a
        a = Complete()   'Autocomplete
    End If
    RaiseEvent Change
    '<EhFooter>
    Exit Sub

txtText_Change_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.txtText_Change " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtText_GotFocus()
    '<EhHeader>
    On Error GoTo txtText_GotFocus_Err
    '</EhHeader>
    txtText.SetSelection 0, Len(txtText.Text)
    'txtText.SelStart = 0
    'txtText.SelLength = Len(txtText.Text)
    '<EhFooter>
    Exit Sub

txtText_GotFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.txtText_GotFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo txtText_KeyDown_Err
    '</EhHeader>
   If Not m_Autocomplete Then Exit Sub
   If KeyCode = 8 Then

        If bRead > 0 And txtText.SelLength > 0 Then
            txtText.SelStart = bRead - 1
            txtText.SelLength = Len(txtText.Text) - bRead + 1
        End If
    ElseIf KeyCode = 46 Then
        If txtText.SelLength <> 0 Then
            txtText.Text = Left(txtText.Text, bRead)
            bInput = True
        End If
    End If
    '<EhFooter>
    Exit Sub

txtText_KeyDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.txtText_KeyDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub txtText_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo txtText_KeyPress_Err
    '</EhHeader>
    On Error Resume Next
    Dim a
    Select Case KeyAscii
    Case 13:
        a = Search(txtText.Text)
        ' si a = 0 The especified Text doesn't exist in the list, so then Add the item
        ' si a = 1 Exist, Don't add the item.
        If a = 0 And txtText.Text <> "" Then
            Me.AddItem txtText.Text, , m_Icon
        End If
        ImgItem.Picture = m_Images(m_SelectedItem + 1)
        txtText.SelStart = 0
        txtText.SelLength = Len(txtText.Text)
        If bListIsVisible Then cDown.Hide
        bListIsVisible = False
    Case 27
        If bListIsVisible Then cDown.Hide
        txtText.Text = Left(txtText.Text, txtText.SelStart)
        bListIsVisible = False
    End Select
    RaiseEvent KeyPress(0 + KeyAscii)
    RaiseEvent ItemClick(m_SelectedItem)
    '<EhFooter>
    Exit Sub

txtText_KeyPress_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.txtText_KeyPress " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
    '' Process the KeyBoard Events to Move Or Edit Text
    '<EhHeader>
    On Error GoTo txtText_KeyUp_Err
    '</EhHeader>
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If (Shift And 7) = 4 Then
            picButton_Click
        Else
            If KeyCode = vbKeyUp Then
                m_SelectedItem = m_SelectedItem - 1
                If m_SelectedItem < 0 Then m_SelectedItem = 0
            ElseIf KeyCode = vbKeyDown Then
                m_SelectedItem = m_SelectedItem + 1
                If m_SelectedItem > m_Items.Count - 1 Then m_SelectedItem = m_Items.Count - 1
            End If
            txtText.Text = CStr(m_Items(m_SelectedItem + 1))
            ImgItem.Picture = m_Images(m_SelectedItem + 1)
            txtText.SelStart = 0
            txtText.SelLength = Len(txtText.Text)
        End If
    End If
    '<EhFooter>
    Exit Sub

txtText_KeyUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.txtText_KeyUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo txtText_MouseMove_Err
    '</EhHeader>
    UserControl_MouseMove Button, Shift, X, Y
    '<EhFooter>
    Exit Sub

txtText_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.txtText_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_EnterFocus()
    '<EhHeader>
    On Error GoTo UserControl_EnterFocus_Err
    '</EhHeader>
    bPreserve = True
    If Not cDown Is Nothing Then cDown.m_bPreserve = True
    OnFocus = True
    iState = Hover
    If m_Enabled Then
        DrawFace
    End If
    DoEvents
    bPreserve = False
    If Not cDown Is Nothing Then cDown.m_bPreserve = False
    '<EhFooter>
    Exit Sub

UserControl_EnterFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_EnterFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_ExitFocus()
    '<EhHeader>
    On Error GoTo UserControl_ExitFocus_Err
    '</EhHeader>
    OnFocus = False
    iState = Normal
    If Not cDown Is Nothing Then If cDown.m_ShowingList And Not bPreserve Then cDown.Reset
    If m_Enabled Then
        DrawFace
    End If
    '<EhFooter>
    Exit Sub

UserControl_ExitFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_ExitFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Initialize()
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>
    InOut = False
    OnClicking = False
    
    gScaleX = Screen.TwipsPerPixelX
    gScaleY = Screen.TwipsPerPixelY
    m_ImageSize = 16
    UserControl_Resize
    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_Initialize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyDown_Err
    '</EhHeader>
    If m_Enabled Then DrawFace
    '<EhFooter>
    Exit Sub

UserControl_KeyDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_KeyDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyPress_Err
    '</EhHeader>
    If m_Enabled Then
        RaiseEvent Click
        RaiseEvent KeyPress(KeyAscii)
    End If
    '<EhFooter>
    Exit Sub

UserControl_KeyPress_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_KeyPress " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseDown_Err
    '</EhHeader>
    If m_Enabled Then
        iState = pushed
        UserControl_Paint
        timUpdate.Enabled = False
        OnClicking = True
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
    '<EhFooter>
    Exit Sub

UserControl_MouseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_MouseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseMove_Err
    '</EhHeader>
    If m_Enabled Then
        If Button = 0 Then timUpdate.Enabled = True
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
    '<EhFooter>
    Exit Sub

UserControl_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseUp_Err
    '</EhHeader>
    If m_Enabled Then
        iState = Hover
        UserControl_Paint
        timUpdate.Enabled = True
        RaiseEvent MouseUp(Button, Shift, X, Y)
        If InBox(UserControl.hwnd) Then
            RaiseEvent Click
        End If
        OnClicking = False
    End If
    '<EhFooter>
    Exit Sub

UserControl_MouseUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_MouseUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '   Text Position
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
    UserControl.ScaleMode = 3
    If ScaleWidth < 56 Then Width = 56 * gScaleX
    If ScaleHeight < 21 Then Height = 21 * gScaleY
    ImgItem.Move 4, (UserControl.ScaleHeight - m_ImageSize) / 2, m_ImageSize, m_ImageSize
    txtText.Move 7 + m_ImageSize, (UserControl.ScaleHeight - txtText.Height) / 2, ScaleWidth - m_ImageSize - picButton.Width - 10
    picButton.Move ScaleWidth - 18, 2, 16, ScaleHeight - 4
    Select Case m_TextAlign
        Case 0  '   Left
            'txtText.TextAlign = fmTextAlignLeft
        Case 1  '   Right
            'txtText.TextAlign = fmTextAlignRight
        Case 2  '   Top
            'txtText.TextAlign = fmTextAlignCenter
    End Select
    'Locate Button
    UserControl_Paint
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_DblClick()
    '<EhHeader>
    On Error GoTo UserControl_DblClick_Err
    '</EhHeader>
    If m_Enabled Then RaiseEvent DblClick
    '<EhFooter>
    Exit Sub

UserControl_DblClick_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_DblClick " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Show()
    '<EhHeader>
    On Error GoTo UserControl_Show_Err
    '</EhHeader>
    UserControl_Resize
    UserControl_Paint
    '<EhFooter>
    Exit Sub

UserControl_Show_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_Show " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub picButton_Click()
    ' Show de Auxiliar Window
    '<EhHeader>
    On Error GoTo picButton_Click_Err
    '</EhHeader>
    Dim ni As Integer

    If m_Enabled Then
        Set cDown = New wndDown
        For ni = 1 To m_Items.Count
            cDown.m_Items.Add m_Items(ni)
            cDown.m_Images.Add m_Images(ni)
        Next ni
        RaiseEvent ButtonClick
        Dim rT As RECT
        GetWindowRect UserControl.hwnd, rT
        cDown.m_BackColor = m_DropDownListBackColor
        cDown.m_BorderColor = m_DropDownListBorderColor
        cDown.m_HoverColor = m_DropDownListHoverColor
        cDown.m_IconsBackColor = m_DropDownListIconsBackColor
        cDown.SetParentHeight UserControl.ScaleHeight
        bPreserve = True
        cDown.PopUp rT.Left * gScaleX, rT.Bottom * gScaleY, UserControl.Width, UserControl.Extender.parent, m_SelectedItem
        bListIsVisible = True
        DoEvents
        bPreserve = False
    End If
NoItemsToShow:
    '<EhFooter>
    Exit Sub

picButton_Click_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.picButton_Click " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error GoTo UserControl_Terminate_Err
    '</EhHeader>
    On Error Resume Next
    If Not cDown Is Nothing Then
        Unload cDown
        Set cDown = Nothing
    End If
    '<EhFooter>
    Exit Sub

UserControl_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'************************************************************'
'*                                                          *'
'*  Properties                                              *'
'*                                                          *'
'***********************************************************

Public Property Get Style() As iscStyle
    '<EhHeader>
    On Error GoTo Style_Err
    '</EhHeader>
    Style = m_Style
    '<EhFooter>
    Exit Property

Style_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Style " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Style(ByVal New_Style As iscStyle)
    '<EhHeader>
    On Error GoTo Style_Err
    '</EhHeader>
    m_Style = New_Style
    InOut = False
    timUpdate.Enabled = True
    DoEvents
    UserControl_Paint
    PropertyChanged "Style"
    '<EhFooter>
    Exit Property

Style_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Style " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Caption() As String
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    Caption = txtText.Text
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Caption(ByVal New_Caption As String)
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    txtText.Text = New_Caption
    PropertyChanged "Caption"
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Font() As Font
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set Font = txtText.Font
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set Font(ByVal New_Font As Font)
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set txtText.Font = New_Font
    txtText.Height = 1 'TextBox Will Automacally adjust to Minimum Height
    UserControl_Resize
    PropertyChanged "Font"
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ToolTipText() As String
    '<EhHeader>
    On Error GoTo ToolTipText_Err
    '</EhHeader>
    ToolTipText = txtText.ToolTipText
    '<EhFooter>
    Exit Property

ToolTipText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.ToolTipText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    '<EhHeader>
    On Error GoTo ToolTipText_Err
    '</EhHeader>
    txtText.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
    '<EhFooter>
    Exit Property

ToolTipText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.ToolTipText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Icon() As Picture
    '<EhHeader>
    On Error GoTo Icon_Err
    '</EhHeader>
    Set Icon = m_Icon
    '<EhFooter>
    Exit Property

Icon_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Icon " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    '<EhHeader>
    On Error GoTo Icon_Err
    '</EhHeader>
    Set m_Icon = New_Icon
    Set ImgItem.Picture = New_Icon
    PropertyChanged "Icon"
    '<EhFooter>
    Exit Property

Icon_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Icon " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TextAlign() As ISAlign
    '<EhHeader>
    On Error GoTo TextAlign_Err
    '</EhHeader>
    TextAlign = m_TextAlign
    '<EhFooter>
    Exit Property

TextAlign_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.TextAlign " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TextAlign(ByVal New_TextAlign As ISAlign)
    '<EhHeader>
    On Error GoTo TextAlign_Err
    '</EhHeader>
    m_TextAlign = New_TextAlign
    UserControl_Resize
    PropertyChanged "TextAlign"
    '<EhFooter>
    Exit Property

TextAlign_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.TextAlign " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get BackColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
    BackColor = m_BackColor
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let BackColor(newBackColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
    m_BackColor = newBackColor
    UserControl_Paint
    PropertyChanged "Backcolor"
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get HoverColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo HoverColor_Err
    '</EhHeader>
    HoverColor = m_HoverColor
    '<EhFooter>
    Exit Property

HoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.HoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let HoverColor(newHoverColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo HoverColor_Err
    '</EhHeader>
    m_HoverColor = newHoverColor
    UserControl_Paint
    PropertyChanged "Hovercolor"
    '<EhFooter>
    Exit Property

HoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.HoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' Office XP Color
Public Property Get MSOXPColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo MSOXPColor_Err
    '</EhHeader>
    MSOXPColor = m_MSOXPColor
    '<EhFooter>
    Exit Property

MSOXPColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.MSOXPColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let MSOXPColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo MSOXPColor_Err
    '</EhHeader>
    m_MSOXPColor = NewColor
    UserControl_Paint
    PropertyChanged "MSOXPColor"
    '<EhFooter>
    Exit Property

MSOXPColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.MSOXPColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' Office XP Color (Pushed)
Public Property Get MSOXPHoverColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo MSOXPHoverColor_Err
    '</EhHeader>
    MSOXPHoverColor = m_MSOXPHoverColor
    '<EhFooter>
    Exit Property

MSOXPHoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.MSOXPHoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let MSOXPHoverColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo MSOXPHoverColor_Err
    '</EhHeader>
    m_MSOXPHoverColor = NewColor
    UserControl_Paint
    PropertyChanged "MSOXPHoverColor"
    '<EhFooter>
    Exit Property

MSOXPHoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.MSOXPHoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' Windows XP Color
Public Property Get WINXPColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo WINXPColor_Err
    '</EhHeader>
    WINXPColor = m_WINXPColor
    '<EhFooter>
    Exit Property

WINXPColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.WINXPColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let WINXPColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo WINXPColor_Err
    '</EhHeader>
    m_WINXPColor = NewColor
    UserControl_Paint
    PropertyChanged "WINXPColor"
    '<EhFooter>
    Exit Property

WINXPColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.WINXPColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' Windows XP Hover Color
Public Property Get WINXPHoverColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo WINXPHoverColor_Err
    '</EhHeader>
    WINXPHoverColor = m_WINXPHoverColor
    '<EhFooter>
    Exit Property

WINXPHoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.WINXPHoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let WINXPHoverColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo WINXPHoverColor_Err
    '</EhHeader>
    m_WINXPHoverColor = NewColor
    UserControl_Paint
    PropertyChanged "WINXPHoverColor"
    '<EhFooter>
    Exit Property

WINXPHoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.WINXPHoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' Windows XP Border Color
Public Property Get WINXPBorderColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo WINXPBorderColor_Err
    '</EhHeader>
    WINXPBorderColor = m_WINXPBorderColor
    '<EhFooter>
    Exit Property

WINXPBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.WINXPBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let WINXPBorderColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo WINXPBorderColor_Err
    '</EhHeader>
    m_WINXPBorderColor = NewColor
    UserControl_Paint
    PropertyChanged "WINXPBorderColor"
    '<EhFooter>
    Exit Property

WINXPBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.WINXPBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' DropDownList BackColor
Public Property Get DropDownListBackColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo DropDownListBackColor_Err
    '</EhHeader>
    DropDownListBackColor = m_DropDownListBackColor
    '<EhFooter>
    Exit Property

DropDownListBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let DropDownListBackColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo DropDownListBackColor_Err
    '</EhHeader>
    m_DropDownListBackColor = NewColor
    UserControl_Paint
    PropertyChanged "DropDownListBackColor"
    '<EhFooter>
    Exit Property

DropDownListBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' DropDownList HoverColor
Public Property Get DropDownListHoverColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo DropDownListHoverColor_Err
    '</EhHeader>
    DropDownListHoverColor = m_DropDownListHoverColor
    '<EhFooter>
    Exit Property

DropDownListHoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListHoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let DropDownListHoverColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo DropDownListHoverColor_Err
    '</EhHeader>
    m_DropDownListHoverColor = NewColor
    PropertyChanged "DropDownListHoverColor"
    '<EhFooter>
    Exit Property

DropDownListHoverColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListHoverColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' DropDownList BorderColor
Public Property Get DropDownListBorderColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo DropDownListBorderColor_Err
    '</EhHeader>
    DropDownListBorderColor = m_DropDownListBorderColor
    '<EhFooter>
    Exit Property

DropDownListBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let DropDownListBorderColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo DropDownListBorderColor_Err
    '</EhHeader>
    m_DropDownListBorderColor = NewColor
    PropertyChanged "DropDownListBorderColor"
    '<EhFooter>
    Exit Property

DropDownListBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' DropDownList Icons Backcolor
Public Property Get DropDownListIconsBackColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo DropDownListIconsBackColor_Err
    '</EhHeader>
    DropDownListIconsBackColor = m_DropDownListIconsBackColor
    '<EhFooter>
    Exit Property

DropDownListIconsBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListIconsBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let DropDownListIconsBackColor(NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo DropDownListIconsBackColor_Err
    '</EhHeader>
    m_DropDownListIconsBackColor = NewColor
    PropertyChanged "DropDownListIconsBackColor"
    '<EhFooter>
    Exit Property

DropDownListIconsBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DropDownListIconsBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

''' Font Colors
Public Property Get FontColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo FontColor_Err
    '</EhHeader>
    FontColor = m_FontColor
    '<EhFooter>
    Exit Property

FontColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.FontColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo FontColor_Err
    '</EhHeader>
    m_FontColor = New_FontColor
    txtText.ForeColor = New_FontColor
    PropertyChanged "FontColor"
    '<EhFooter>
    Exit Property

FontColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.FontColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get FontHighlightColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo FontHighlightColor_Err
    '</EhHeader>
    FontHighlightColor = m_FontHighlightColor
    '<EhFooter>
    Exit Property

FontHighlightColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.FontHighlightColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let FontHighlightColor(ByVal New_FontHighlightColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo FontHighlightColor_Err
    '</EhHeader>
    m_FontHighlightColor = New_FontHighlightColor
    PropertyChanged "FontHighlightColor"
    '<EhFooter>
    Exit Property

FontHighlightColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.FontHighlightColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Autocomplete() As Boolean
    '<EhHeader>
    On Error GoTo Autocomplete_Err
    '</EhHeader>
    Autocomplete = m_Autocomplete
    '<EhFooter>
    Exit Property

Autocomplete_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Autocomplete " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Autocomplete(ByVal New_Autocomplete As Boolean)
    '<EhHeader>
    On Error GoTo Autocomplete_Err
    '</EhHeader>
    m_Autocomplete = New_Autocomplete
    PropertyChanged "Autocomplete"
    '<EhFooter>
    Exit Property

Autocomplete_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Autocomplete " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Enabled() As Boolean
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
    Enabled = m_Enabled
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Enabled " & _
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
           "in ThunVBCC_v1_0.ISCombo.hwnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
    m_Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    picButton.Enabled = New_Enabled
    ImgItem.Enabled = New_Enabled
    If New_Enabled Then
        txtText.BackColor = vbWindowBackground
        UserControl.BackColor = vbWindowBackground
        txtText.ForeColor = vbButtonText
        txtText.ReadOnly = False
        txtText.Enabled = True
        iState = Normal
    Else
        txtText.BackColor = vb3DFace
        UserControl.BackColor = vb3DFace
        txtText.ForeColor = vbGrayText
        txtText.ReadOnly = True
        txtText.Enabled = False
        iState = disabled
    End If
    UserControl_Paint
    PropertyChanged "Enabled"
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Enabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Sub UserControl_InitProperties()
    '<EhHeader>
    On Error GoTo UserControl_InitProperties_Err
    '</EhHeader>
    Set m_Icon = LoadPicture("")
    m_TextAlign = m_def_TextAlign
    m_FontColor = m_def_FontColor
    m_FontHighlightColor = m_def_FontHighlightColor
    
    m_HoverColor = m_def_HoverColor
    m_BackColor = m_def_BackColor
    m_MSOXPColor = m_def_MSOXPColor
    m_MSOXPHoverColor = m_def_MSOXPHoverColor
    m_WINXPColor = m_def_WINXPColor
    m_WINXPHoverColor = m_def_WINXPHoverColor
    m_WINXPBorderColor = m_def_WINXPBorderColor
    
    m_DropDownListBackColor = m_def_DropDownListBackColor
    m_DropDownListHoverColor = m_def_DropDownListHoverColor
    m_DropDownListBorderColor = m_def_DropDownListBorderColor
    m_DropDownListIconsBackColor = m_def_DropDownListIconsBackColor
    
    m_Autocomplete = m_def_Autocomplete
    txtText.Text = UserControl.Extender.Name
    m_Enabled = m_def_Enabled
    m_Style = m_def_Style
    '<EhFooter>
    Exit Sub

UserControl_InitProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_InitProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
    Dim picNormal As Picture
    With PropBag
        Set picNormal = PropBag.ReadProperty("Icon", Nothing)
        If Not (picNormal Is Nothing) Then Set Icon = picNormal
    End With

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_HoverColor = PropBag.ReadProperty("HoverColor", m_def_HoverColor)
    m_MSOXPColor = PropBag.ReadProperty("MSOXPColor", m_def_MSOXPColor)
    m_MSOXPHoverColor = PropBag.ReadProperty("MSOXPHoverColor", m_def_MSOXPHoverColor)
    m_WINXPColor = PropBag.ReadProperty("WINXPColor", m_def_WINXPColor)
    m_WINXPHoverColor = PropBag.ReadProperty("WINXPHoverColor", m_def_WINXPHoverColor)
    m_WINXPBorderColor = PropBag.ReadProperty("WINXPBorderColor", m_def_WINXPBorderColor)
    
    m_DropDownListBackColor = PropBag.ReadProperty("DropDownListBackColor", m_def_DropDownListBackColor)
    m_DropDownListHoverColor = PropBag.ReadProperty("DropDownListHoverColor", m_def_DropDownListHoverColor)
    m_DropDownListBorderColor = PropBag.ReadProperty("DropDownListBorderColor", m_def_DropDownListBorderColor)
    m_DropDownListIconsBackColor = PropBag.ReadProperty("DropDownListIconsBackColor", m_def_DropDownListIconsBackColor)
    
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_FontHighlightColor = PropBag.ReadProperty("FontHighlightColor", m_def_FontHighlightColor)
    
    txtText.Text = PropBag.ReadProperty("Caption", "Caption")
    Set txtText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtText.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    
    m_Autocomplete = PropBag.ReadProperty("Autocomplete", m_def_Autocomplete)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    txtText.Enabled = m_Enabled
    picButton.Enabled = m_Enabled
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>

    Call PropBag.WriteProperty("HoverColor", m_HoverColor, m_def_HoverColor)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("MSOXPColor", m_MSOXPColor, m_def_MSOXPColor)
    Call PropBag.WriteProperty("MSOXPHoverColor", m_MSOXPHoverColor, m_def_MSOXPHoverColor)
    Call PropBag.WriteProperty("WINXPColor", m_WINXPColor, m_def_WINXPColor)
    Call PropBag.WriteProperty("WINXPHoverColor", m_WINXPHoverColor, m_def_WINXPHoverColor)
    Call PropBag.WriteProperty("WINXPBorderColor", m_WINXPBorderColor, m_def_WINXPBorderColor)
    
    Call PropBag.WriteProperty("DropDownListBackColor", m_DropDownListBackColor, m_def_DropDownListBackColor)
    Call PropBag.WriteProperty("DropDownListHoverColor", m_DropDownListHoverColor, m_def_DropDownListHoverColor)
    Call PropBag.WriteProperty("DropDownListBorderColor", m_DropDownListBorderColor, m_def_DropDownListBorderColor)
    Call PropBag.WriteProperty("DropDownListIconsBackColor", m_DropDownListIconsBackColor, m_def_DropDownListIconsBackColor)
    
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("FontHighlightColor", m_FontHighlightColor, m_def_FontHighlightColor)
    
    Call PropBag.WriteProperty("Caption", txtText.Text, "Caption")
    Call PropBag.WriteProperty("Font", txtText.Font, Ambient.Font)
    Call PropBag.WriteProperty("ToolTipText", txtText.ToolTipText, "")
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    
    Call PropBag.WriteProperty("Autocomplete", m_Autocomplete, m_def_Autocomplete)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_WriteProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'************************************************************'
'*                                                          *'
'*  Section: Miscelaneus Functions                          *'
'*                                                          *'
'***********************************************************

Private Sub UserControl_Paint()
    'Call All Drawing code
    '<EhHeader>
    On Error GoTo UserControl_Paint_Err
    '</EhHeader>
    Call DrawFace
    '<EhFooter>
    Exit Sub

UserControl_Paint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.UserControl_Paint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawArrow(Optional bBlack As Boolean = True)
    '' This sub Draws the Small Down Arrow in the Classic Combo Box Button
    '<EhHeader>
    On Error GoTo DrawArrow_Err
    '</EhHeader>
    Dim lHDC As Long
    Dim lcw As Long, lch As Long
    Dim hPen As Long, hPenOld As Long
    Dim pt As POINTAPI
    lcw = picButton.Width / 2
    lch = picButton.Height / 2
    lHDC = picButton.hdc
    hPen = CreatePen(0, 1, IIf(bBlack, 0, vbWhite))
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx lHDC, lcw - 4, lch - 1, pt
    LineTo lHDC, lcw + 3, lch - 1
    MoveToEx lHDC, lcw - 3, lch, pt
    LineTo lHDC, lcw + 2, lch
    MoveToEx lHDC, lcw - 2, lch + 1, pt
    LineTo lHDC, lcw + 1, lch + 1
    MoveToEx lHDC, lcw - 1, lch + 2, pt
    LineTo lHDC, lcw, lch + 2
    SelectObject lHDC, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld
    '<EhFooter>
    Exit Sub

DrawArrow_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DrawArrow " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawWinXPButton(Mode As State)
    '' This Sub Draws the XPStyle Button
    '<EhHeader>
    On Error GoTo DrawWinXPButton_Err
    '</EhHeader>
    Dim lHDC As Long
    Dim tempColor As Long
    Dim lH As Long, lW As Long
    Dim lcw As Long, lch As Long
    Dim lStep As Single
    Dim ni As Single
    lW = picButton.Width
    lH = picButton.Height
    lHDC = picButton.hdc
    lcw = lW / 2
    lch = lH / 2
    UserControl.picButton.Cls
    lStep = 25 / lH
    Select Case Mode
    Case 0, 1:
        tempColor = IIf((Mode = Hover), OffsetColor(m_WINXPHoverColor, &H30), OffsetColor(m_WINXPColor, &H30))
        For ni = 0 To lH
            APILine lHDC, 0, lH - ni, lW, lH - ni, OffsetColor(tempColor, ni * lStep)
        Next ni
        APILine lHDC, 0, lH - 1, lW - 0, lH - 1, OffsetColor(tempColor, -64)
        APILine lHDC, 1, lH - 2, lW - 1, lH - 2, OffsetColor(tempColor, -32)
        APILine lHDC, lW - 1, 0, lW - 1, lH - 0, OffsetColor(tempColor, -32)
        APILine lHDC, lW - 2, 1, lW - 2, lH - 1, OffsetColor(tempColor, -12)
        APILine lHDC, 1, 0, lW - 1, 0, OffsetColor(tempColor, 19)
        APILine lHDC, 0, 1, lW - 1, 1, OffsetColor(tempColor, 32)
        APILine lHDC, 1, 2, 1, lH - 2, OffsetColor(tempColor, 4)
        picButton.PSet (1, 1), OffsetColor(tempColor, 48)
        picButton.PSet (1, lH - 2), OffsetColor(tempColor, -12)
        picButton.PSet (lW - 2, 1), OffsetColor(tempColor, 40)
        picButton.PSet (lW - 2, lH - 2), OffsetColor(tempColor, -64)
        picButton.PSet (0, 0), m_BackColor 'OffsetColor(tempColor, 92) ' &HFCEEE6
        picButton.PSet (0, lH - 1), m_BackColor 'OffsetColor(tempColor, 80) ' &HF9E6DC
        picButton.PSet (lW - 1, 0), m_BackColor 'OffsetColor(tempColor, 80) '&HF6E3D9
        picButton.PSet (lW - 1, lH - 1), m_BackColor 'OffsetColor(tempColor, 64) '&HF8E3D8
        
    Case 2:
        tempColor = OffsetColor(m_WINXPHoverColor, &H30) 'OffsetColor(RGB(127, 157, 255), &H30)
        For ni = 0 To lH
            APILine lHDC, 0, lH - ni, lW, lH - ni, OffsetColor(tempColor, -ni * lStep)
        Next ni
        APILine lHDC, 0, lH - 1, lW - 0, lH - 1, OffsetColor(tempColor, -16)
        APILine lHDC, 1, lH - 2, lW - 1, lH - 2, OffsetColor(tempColor, -24)
        APILine lHDC, lW - 1, 0, lW - 1, lH - 0, OffsetColor(tempColor, -32)
        APILine lHDC, lW - 2, 1, lW - 2, lH - 1, OffsetColor(tempColor, -24)
        APILine lHDC, 1, 0, lW - 1, 0, OffsetColor(tempColor, -64)
        APILine lHDC, 0, 1, lW - 1, 1, OffsetColor(tempColor, -32)
        APILine lHDC, 0, 1, 0, lH - 1, OffsetColor(tempColor, -64)
        APILine lHDC, 1, 2, 1, lH - 2, OffsetColor(tempColor, -32)
        picButton.PSet (1, 1), &HBF8D6F
        picButton.PSet (1, lH - 2), &HDFAD8F
        picButton.PSet (lW - 2, 1), &HDFAD8F
        picButton.PSet (lW - 2, lH - 2), &HFBC9AB
        
        picButton.PSet (0, 0), m_HoverColor '&HFCEEE6
        picButton.PSet (0, lH - 1), m_HoverColor '&HF9E6DC
        picButton.PSet (lW - 1, 0), m_HoverColor '&HF6E3D9
        picButton.PSet (lW - 1, lH - 1), m_HoverColor '&HF8E3D8
        lch = lch + 1
        lcw = lcw + 1
    Case 3:
        tempColor = GetSysColor(COLOR_BTNFACE) 'OffsetColor(GetSysColor(COLOR_BTNFACE), &H30)
        For ni = 0 To lH
            APILine lHDC, 0, lH - ni, lW, lH - ni, OffsetColor(tempColor, ni * lStep)
        Next ni
        APILine lHDC, 0, lH - 1, lW - 0, lH - 1, OffsetColor(tempColor, -64)
        APILine lHDC, 1, lH - 2, lW - 1, lH - 2, OffsetColor(tempColor, -32)
        APILine lHDC, lW - 1, 0, lW - 1, lH - 0, OffsetColor(tempColor, -32)
        APILine lHDC, lW - 2, 1, lW - 2, lH - 1, OffsetColor(tempColor, -12)
        APILine lHDC, 1, 0, lW - 1, 0, OffsetColor(tempColor, 19)
        APILine lHDC, 0, 1, lW - 1, 1, OffsetColor(tempColor, 32)
        'APILine lhdc, 0, 1, 0, lh - 1, vbRed 'OffsetColor(tempColor, -4)
        APILine lHDC, 1, 2, 1, lH - 2, OffsetColor(tempColor, 4)
        picButton.PSet (1, 1), &HFFEFD2
        picButton.PSet (1, lH - 2), &HFBC9AB
        picButton.PSet (lW - 2, 1), &HFFE3C5
        picButton.PSet (lW - 2, lH - 2), &HBF8D6F
        picButton.PSet (0, 0), &HFCEEE6
        picButton.PSet (0, lH - 1), &HF9E6DC
        picButton.PSet (lW - 1, 0), &HF6E3D9
        picButton.PSet (lW - 1, lH - 1), &HF8E3D8
    End Select
    '' Draw The XP Style Arrow
    APILine lHDC, lcw - 5, lch - 2, lcw, lch + 3, 0
    APILine lHDC, lcw - 4, lch - 2, lcw, lch + 2, 0
    APILine lHDC, lcw - 4, lch - 3, lcw, lch + 1, 0
    
    APILine lHDC, lcw + 3, lch - 2, lcw - 2, lch + 3, 0
    APILine lHDC, lcw + 2, lch - 2, lcw - 2, lch + 2, 0
    APILine lHDC, lcw + 2, lch - 3, lcw - 2, lch + 1, 0
        
    '<EhFooter>
    Exit Sub

DrawWinXPButton_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DrawWinXPButton " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub DrawFace()
    ''Code for Button Drawing
    '<EhHeader>
    On Error GoTo DrawFace_Err
    '</EhHeader>
    UserControl.ScaleMode = 3
    Dim rT As RECT, rtout As RECT
    Dim tmpState As State
    
    rT.Top = 0
    rT.Left = 0
    rT.Bottom = picButton.ScaleHeight
    rT.Right = picButton.ScaleWidth
    rtout.Top = 1
    rtout.Left = picButton.Left - 1
    rtout.Right = UserControl.ScaleWidth - 2
    rtout.Bottom = UserControl.ScaleHeight - 2
    
    tmpState = iState
    If OnFocus Then tmpState = Hover
    If bPressed Then tmpState = 2
    ''Control drawing
    
    Dim lW As Long, lH As Long
    Dim lUCHDC As Long, lPBHDC As Long
    Dim colorShadow As Long, colorLight As Long, colorBack As Long, colorFace As Long
    Dim ucRT        As RECT, rtIn As RECT
    
    lW = ScaleWidth
    lH = ScaleHeight
    lUCHDC = UserControl.hdc
    lPBHDC = picButton.hdc
    ucRT.Top = 0
    ucRT.Left = 0
    ucRT.Bottom = UserControl.ScaleHeight
    ucRT.Right = UserControl.ScaleWidth
    rtIn.Top = 1
    rtIn.Left = 1
    rtIn.Right = UserControl.ScaleWidth - 2
    rtIn.Bottom = UserControl.ScaleHeight - 2
    UserControl.BackColor = IIf(iState = disabled, GetSysColor(COLOR_BTNFACE), IIf((iState <> Normal), m_HoverColor, m_BackColor))
    txtText.BackColor = IIf(iState = disabled, GetSysColor(COLOR_BTNFACE), IIf(iState = Normal, m_BackColor, m_HoverColor))
    Select Case m_Style
    Case 0: 'Normal
        UserControl.Cls
        Call DrawEdge(UserControl.hdc, ucRT, EDGE_SUNKEN, BF_RECT)
        picButton.BackColor = GetSysColor(COLOR_BTNFACE)
        Select Case tmpState
        Case 0  'Normal
            DrawEdge picButton.hdc, rT, EDGE_RAISED, BF_RECT
            txtText.ForeColor = m_FontColor
        Case 1  'Hover
            DrawEdge picButton.hdc, rT, EDGE_RAISED, BF_RECT
            txtText.ForeColor = m_FontHighlightColor
        Case 2  'Pushed
            txtText.ForeColor = m_FontHighlightColor
            DrawEdge picButton.hdc, rT, EDGE_SUNKEN, BF_RECT
        Case 3  'Disabled
            DrawEdge picButton.hdc, rT, EDGE_RAISED, BF_RECT
        End Select
        DrawArrow True
    Case 1: 'MSO2000
        Select Case tmpState
        Case 0
            UserControl.BackColor = m_BackColor
            UserControl.Cls
            txtText.ForeColor = m_FontColor
            ucRT.Bottom = UserControl.ScaleHeight - 1
            ucRT.Right = UserControl.ScaleWidth - 1
            APIRectangle UserControl.hdc, ucRT, GetSysColor(COLOR_BTNFACE)
            APIRectangle UserControl.hdc, rtIn, GetSysColor(COLOR_BTNFACE)
            picButton.Cls
            rT.Bottom = picButton.ScaleHeight - 1
            rT.Right = picButton.ScaleWidth - 1
            APIRectangle picButton.hdc, rT, m_BackColor 'GetSysColor(COLOR_WINDOW)
        Case 1
            UserControl.BackColor = m_HoverColor
            UserControl.Cls
            txtText.ForeColor = m_FontHighlightColor
            APIRectangle UserControl.hdc, rtIn, GetSysColor(COLOR_BTNFACE)
            Call DrawEdge(UserControl.hdc, ucRT, BDR_SUNKENOUTER, BF_RECT)
            picButton.Cls
            APILine UserControl.hdc, picButton.Left - 1, 2, picButton.Left - 1, UserControl.ScaleHeight - 2, GetSysColor(COLOR_BTNFACE)
            DrawEdge picButton.hdc, rT, BDR_RAISEDINNER, BF_RECT
        Case 2
            UserControl.BackColor = m_HoverColor
            UserControl.Cls
            txtText.ForeColor = m_FontHighlightColor
            APIRectangle UserControl.hdc, rtIn, GetSysColor(COLOR_BTNFACE)
            Call DrawEdge(UserControl.hdc, ucRT, BDR_SUNKENOUTER, BF_RECT)
            picButton.Cls
            APILine UserControl.hdc, picButton.Left - 1, 2, picButton.Left - 1, UserControl.ScaleHeight - 2, GetSysColor(COLOR_BTNFACE)
            DrawEdge picButton.hdc, rT, BDR_SUNKENINNER, BF_RECT
        Case 3
            UserControl.Cls
            ucRT.Bottom = UserControl.ScaleHeight - 1
            ucRT.Right = UserControl.ScaleWidth - 1
            APIRectangle UserControl.hdc, rtIn, GetSysColor(COLOR_WINDOW)
            APIFillRect picButton.hdc, rT, GetSysColor(COLOR_BTNFACE)
        End Select
        DrawArrow True
    Case 2: 'MSOXP
        ucRT.Bottom = UserControl.ScaleHeight - 1
        ucRT.Right = UserControl.ScaleWidth - 1
        Select Case tmpState
        Case 0
            UserControl.Cls
            txtText.ForeColor = m_FontColor
            ucRT.Bottom = UserControl.ScaleHeight - 1
            ucRT.Right = UserControl.ScaleWidth - 1
            APIRectangle UserControl.hdc, ucRT, GetSysColor(COLOR_BTNFACE)
            APIRectangle UserControl.hdc, rtIn, GetSysColor(COLOR_BTNFACE)
            picButton.Cls
            rT.Bottom = picButton.ScaleHeight - 1
            rT.Right = picButton.ScaleWidth - 1
            APIRectangle picButton.hdc, rT, m_BackColor
            DrawArrow True
        Case 1
            UserControl.Cls
            txtText.ForeColor = m_FontHighlightColor
            colorShadow = RGB(0, 0, 0)
            APIRectangle UserControl.hdc, rtIn, m_HoverColor
            APIRectangle UserControl.hdc, ucRT, m_MSOXPColor
            picButton.Cls
            APIFillRect picButton.hdc, rT, m_MSOXPColor '&HC08080
            APILine UserControl.hdc, picButton.Left - 1, 2, picButton.Left - 1, UserControl.ScaleHeight - 2, vbRed 'GetSysColor(COLOR_BTNFACE)
            APIRectangle UserControl.hdc, rtout, m_MSOXPColor '&HC08080
            DrawArrow False
        Case 2
            UserControl.Cls
            colorShadow = RGB(0, 0, 0)
            txtText.ForeColor = m_FontHighlightColor
            APIRectangle UserControl.hdc, rtIn, m_HoverColor
            APIRectangle UserControl.hdc, ucRT, m_MSOXPHoverColor
            picButton.Cls
            APIFillRect picButton.hdc, rT, m_MSOXPHoverColor
            APIRectangle UserControl.hdc, rtout, m_MSOXPHoverColor
            DrawArrow False
        Case 3
            UserControl.Cls
            ucRT.Bottom = UserControl.ScaleHeight - 1
            ucRT.Right = UserControl.ScaleWidth - 1
            APIRectangle UserControl.hdc, rtIn, GetSysColor(COLOR_WINDOW)
            APIFillRect picButton.hdc, rT, GetSysColor(COLOR_BTNFACE)
            DrawArrow False
        End Select
    Case 3: 'WINXP
        ucRT.Bottom = UserControl.ScaleHeight - 1
        ucRT.Right = UserControl.ScaleWidth - 1
        If iState = disabled Then
            picButton.Cls
            APIRectangle UserControl.hdc, ucRT, GetSysColor(COLOR_BTNSHADOW)
            APIRectangle UserControl.hdc, rtIn, GetSysColor(COLOR_WINDOW)
        Else
            UserControl.Cls
            If iState = Normal Then
                txtText.ForeColor = m_FontColor
                APIRectangle UserControl.hdc, rtIn, m_BackColor
            Else
                txtText.ForeColor = m_FontHighlightColor
                APIRectangle UserControl.hdc, rtIn, m_HoverColor
            End If
            APIRectangle UserControl.hdc, ucRT, m_WINXPBorderColor 'RGB(127, 157, 185)
        End If
        '' This is too Complex, so I decided Put the Drawing Code in another sub.
        DrawWinXPButton tmpState
    End Select
    '<EhFooter>
    Exit Sub

DrawFace_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.DrawFace " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Restore Original Colors

Public Sub RestoreOriginalColors()
    '<EhHeader>
    On Error GoTo RestoreOriginalColors_Err
    '</EhHeader>
    Me.BackColor = m_def_BackColor
    Me.HoverColor = m_def_HoverColor
    Me.MSOXPColor = m_def_MSOXPColor
    Me.MSOXPHoverColor = m_def_MSOXPHoverColor
    Me.WINXPColor = m_def_WINXPColor
    Me.WINXPHoverColor = m_def_WINXPHoverColor
    Me.FontColor = m_def_FontColor
    Me.FontHighlightColor = m_def_FontHighlightColor
    Me.WINXPBorderColor = m_def_WINXPBorderColor
    '<EhFooter>
    Exit Sub

RestoreOriginalColors_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.RestoreOriginalColors " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Remove all Items In the List
Public Sub Clear()
    '<EhHeader>
    On Error GoTo Clear_Err
    '</EhHeader>
    Dim Item
    For Each Item In m_Items
        m_Items.Remove (Item)
        m_Images.Remove (Item)
    Next Item
    m_SelectedItem = 0
    '<EhFooter>
    Exit Sub

Clear_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Clear " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub Remove(Item)
    '<EhHeader>
    On Error GoTo Remove_Err
    '</EhHeader>
    m_Images.Remove (Item)
    m_Items.Remove (Item)
    '<EhFooter>
    Exit Sub

Remove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.Remove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Add a new Item to the Combo List
Public Sub AddItem(Text As String, Optional Index As Integer, Optional iImage As Picture)
    '<EhHeader>
    On Error GoTo AddItem_Err
    '</EhHeader>
    Dim ImageTemp As Picture
    If IsMissing(iImage) Then
        Set ImageTemp = LoadPicture()
    Else
        Set ImageTemp = iImage
    End If
    m_Items.Add Text, Text
    m_Images.Add ImageTemp, Text
    '<EhFooter>
    Exit Sub

AddItem_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ISCombo.AddItem " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


