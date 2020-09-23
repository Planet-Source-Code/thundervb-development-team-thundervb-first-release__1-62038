VERSION 5.00
Begin VB.Form wndDown 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pScroller 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   0
      Top             =   120
      Width           =   2715
      Begin VB.VScrollBar vsb 
         Height          =   2295
         Left            =   2220
         Max             =   115
         SmallChange     =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   60
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   1
         Top             =   120
         Width           =   2055
         Begin VB.Timer timUpdate 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   1260
            Top             =   600
         End
         Begin ThunVBCC_v1.UniLabel lblCaption 
            Height          =   375
            Index           =   0
            Left            =   480
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            CaptionB        =   "wndDown.frx":0000
            CaptionLen      =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image ImgItem 
            Height          =   240
            Index           =   0
            Left            =   60
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "wndDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
''      This is a Custom ImageCombo, that supports some aditional Features
''      See ISCombo.ctl For Detailed Info.:
''      you know, you can use this freely, just give me credit.
''      Votes and suggestions are wellcome.
''

Option Explicit



Dim temiPos As Integer
Private Const WM_SIZE = &H5
Private Const WM_MOVE = &H3
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_KILLFOCUS = &H8
' = (-4)

Private Const AW_HOR_POSITIVE = &H1
Private Const AW_HOR_NEGATIVE = &H2
Private Const AW_VER_POSITIVE = &H4
Private Const AW_VER_NEGATIVE = &H8
Private Const AW_CENTER = &H10
Private Const AW_HIDE = &H10000
Private Const AW_ACTIVATE = &H20000
Private Const AW_SLIDE = &H40000
Private Const AW_BLEND = &H80000

Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Dim iPos As Integer
Dim iItems As Integer
Dim IsInside As Boolean
Dim iPrevPos As Integer
Dim iFirstVisible As Integer
Dim bAnimateWindow As Boolean
Dim bMoveBykeyBoard As Boolean
Dim m_lParentHeight As Long

Public m_bShowByAutocomplete As Boolean
Public m_bPreserve As Boolean
Public m_Items As New Collection
Public m_Images As New Collection
Public m_ShowingList As Boolean
Public ItemClick As Integer
Public m_BackColor As OLE_COLOR
Public m_HoverColor As OLE_COLOR
Public m_BorderColor As OLE_COLOR
Public m_IconsBackColor As OLE_COLOR

Event ItemClick(iItem As Integer, sText As String)
Event Hide()
Private nValue As Long
Private OriginalWndProc As Long


'' Detect if the Mouse cursor is inside a Window
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
           "in ThunVBCC_v1_0.wndDown.InBox " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Sub Form_Paint()
    '' Draw The Border of the Window
    '<EhHeader>
    On Error GoTo Form_Paint_Err
    '</EhHeader>
    Line (0, 0)-(ScaleWidth, 0), m_BorderColor
    Line (0, 0)-(0, ScaleHeight), m_BorderColor
    Line (ScaleWidth - 1, 0)-(ScaleWidth - 1, ScaleHeight - 1), m_BorderColor
    Line (0, ScaleHeight - 1)-(ScaleWidth - 1, ScaleHeight - 1), m_BorderColor
    '<EhFooter>
    Exit Sub

Form_Paint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.Form_Paint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Draw All items
Private Sub DrawAll(ActiveItem As Integer)
    ''Customizable Colors :)
    ''  thanks to Lucifer for this Suggestion
    '<EhHeader>
    On Error GoTo DrawAll_Err
    '</EhHeader>
    lblCaption(iPrevPos).BackColor = m_BackColor ' vbWindowBackground
    lblCaption(iPrevPos).ForeColor = vbButtonText
    lblCaption(ActiveItem).BackColor = m_HoverColor ' vbHighlight
    lblCaption(ActiveItem).ForeColor = vbHighlightText
    If ActiveItem <= 0 Then
        iPrevPos = 0
    Else
        iPrevPos = ActiveItem
    End If
    '<EhFooter>
    Exit Sub

DrawAll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.DrawAll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub SetParentHeight(lParentHeight As Long)
    '<EhHeader>
    On Error GoTo SetParentHeight_Err
    '</EhHeader>
    m_lParentHeight = lParentHeight
    '<EhFooter>
    Exit Sub

SetParentHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.SetParentHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub imgItem_Click(Index As Integer)
    '<EhHeader>
    On Error GoTo imgItem_Click_Err
    '</EhHeader>
    lblCaption_Click Index
    '<EhFooter>
    Exit Sub

imgItem_Click_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.imgItem_Click " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ImgItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo ImgItem_MouseMove_Err
    '</EhHeader>
    lblCaption_MouseMove Index, Button, Shift, X, Y
    '<EhFooter>
    Exit Sub

ImgItem_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.ImgItem_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Raise the ItemClick event
Private Sub lblCaption_Click(Index As Integer)
    '<EhHeader>
    On Error GoTo lblCaption_Click_Err
    '</EhHeader>
    Reset
    RaiseEvent ItemClick(Index, lblCaption(Index).Caption)
    '<EhFooter>
    Exit Sub

lblCaption_Click_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.lblCaption_Click " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Detect the mouse movement
Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '' If the user moves the mouse over the text area, It's selected.
    '' But not if the move is a result of KeyBoard action.
    '<EhHeader>
    On Error GoTo lblCaption_MouseMove_Err
    '</EhHeader>
    If Button = 0 And Not bMoveBykeyBoard Then
        timUpdate.Enabled = True
        iPos = Index
    End If
    '<EhFooter>
    Exit Sub

lblCaption_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.lblCaption_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

''  Hide and unload if the window lost the focus
Private Sub picGroup_LostFocus()
    '<EhHeader>
    On Error GoTo picGroup_LostFocus_Err
    '</EhHeader>
    If Not m_bPreserve Then
        Reset
        RaiseEvent Hide
    End If
    '<EhFooter>
    Exit Sub

picGroup_LostFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.picGroup_LostFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

''  Hide and unload if the window lost the focus
Private Sub Form_LostFocus()
    'If Not m_bPreserve Then
    '<EhHeader>
    On Error GoTo Form_LostFocus_Err
    '</EhHeader>
        Reset
        RaiseEvent Hide
    'End If
    '<EhFooter>
    Exit Sub

Form_LostFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.Form_LostFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Activate the TipUpdate Timer
Private Sub picGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo picGroup_MouseMove_Err
    '</EhHeader>
    If Button = 0 Then
        timUpdate.Enabled = True
    End If
    '<EhFooter>
    Exit Sub

picGroup_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.picGroup_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub pScroller_KeyUp(KeyCode As Integer, Shift As Integer)
    '' Process the KeyBoard Events
    '<EhHeader>
    On Error GoTo pScroller_KeyUp_Err
    '</EhHeader>
    bMoveBykeyBoard = True
    Select Case KeyCode
        Case vbKeyUp, vbKeyLeft
            'Select previous Item
            If iPos >= 1 Then
                iPos = iPos - 1
                '' If the combo has a ScrollBar, then
                If vsb.visible Then
                    '' check that the selected item is visible
                    If iPos <= vsb.Value - 1 Then
                        vsb.Value = iPos
                        DoEvents
                    End If
                End If
                DrawAll iPos
            End If
        Case vbKeyDown, vbKeyRight
            'Select Next Item
            If iPos <= m_Items.Count - 2 Then
                iPos = iPos + 1
                '' If the combo has a ScrollBar, then
                If vsb.visible Then
                    '' check that the selected item is visible
                    If iPos >= vsb.Value + 8 Then
                        vsb.Value = iPos - 7
                        DoEvents
                    End If
                End If
                DrawAll iPos
            End If
        Case vbKeyEnd
            iPos = m_Items.Count - 1
            If vsb.visible Then
                vsb.Value = iPos - 7
                DoEvents
            End If
        Case vbKeyHome
            iPos = 0
            If vsb.visible Then
                vsb.Value = 0
                DoEvents
            End If
        Case vbKeyReturn
            'Click on the Selected Item.
            lblCaption_Click iPos
        Case vbKeyEscape
            'Cancel and Reset
            Reset
            RaiseEvent Hide
        Case vbKeyTab
            Reset
            RaiseEvent Hide
    End Select
    bMoveBykeyBoard = False
    '<EhFooter>
    Exit Sub

pScroller_KeyUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.pScroller_KeyUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

''  Hide and unload if the window lost the focus
Private Sub pScroller_LostFocus()
    'If Not m_bPreserve Then
    '<EhHeader>
    On Error GoTo pScroller_LostFocus_Err
    '</EhHeader>
        Reset
        RaiseEvent Hide
    'End If
    '<EhFooter>
    Exit Sub

pScroller_LostFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.pScroller_LostFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Detect the position of the cursor
''  (Only if the cursor is in the Window)
Private Sub timUpdate_Timer()
    '<EhHeader>
    On Error GoTo timUpdate_Timer_Err
    '</EhHeader>

    If InBox(picGroup.hwnd) Then
        If IsInside Then
            If temiPos <> iPos Then
                DrawAll iPos
            End If
        Else
            IsInside = True
        End If
    Else
        timUpdate.Enabled = False
        DrawAll 0
        IsInside = False
    End If
    temiPos = iPos
    '<EhFooter>
    Exit Sub

timUpdate_Timer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.timUpdate_Timer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' Change the position of the items when the ScrollBar changes
Private Sub vsb_Change()
    '<EhHeader>
    On Error GoTo vsb_Change_Err
    '</EhHeader>
    On Error Resume Next
    picGroup.Move 0, 1 - 17 * vsb.Value
    'Me.SetFocus
    '<EhFooter>
    Exit Sub

vsb_Change_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.vsb_Change " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub SetSelectedItem(iSelectedItem As Integer)
    'Select A Specified Item
    '<EhHeader>
    On Error GoTo SetSelectedItem_Err
    '</EhHeader>
    iPos = iSelectedItem
    If iPos > 8 Then vsb.Value = iPos - 7
    DrawAll iPos
    '<EhFooter>
    Exit Sub

SetSelectedItem_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.SetSelectedItem " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'' Hide Window and Save state in Variable
Public Sub Reset()
    '<EhHeader>
    On Error GoTo Reset_Err
    '</EhHeader>
    Hide
    m_ShowingList = False
    '<EhFooter>
    Exit Sub

Reset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.Reset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'' This function Show the cDown Window, And adds the items
Public Function PopUp(X As Long, Y As Long, lWidth As Single, parent As Object, iSelectedItem As Integer) As Boolean
    '<EhHeader>
    On Error GoTo PopUp_Err
    '</EhHeader>
    Dim ni As Integer
    Dim ht As Single
    Dim lHeight As Single
    m_ShowingList = True
    ht = (17 * (m_Items.Count) + 2) * Screen.TwipsPerPixelY
    picGroup.BackColor = m_IconsBackColor
    For ni = 1 To m_Items.Count + 2
        Load lblCaption(ni)
        Load ImgItem(ni)
    Next ni
    If m_Items.Count <= 8 Then
        lHeight = ht
        vsb.visible = False
    Else
        lHeight = (8 * 17 + 2) * Screen.TwipsPerPixelY
        vsb.visible = True
        vsb.Min = 0
        vsb.Max = m_Items.Count - 8
        vsb.SmallChange = 1
        vsb.LargeChange = m_Items.Count - 8
    End If
    On Error GoTo LimitOfItems
    For ni = 1 To m_Items.Count
        lblCaption(ni - 1).BackColor = m_BackColor
        lblCaption(ni - 1).visible = True
        lblCaption(ni - 1).Caption = m_Items.Item(ni)
        lblCaption(ni - 1).Move 24, 17 * (ni - 1), lWidth - 28
        ImgItem(ni - 1).visible = True
        Set ImgItem(ni - 1).Picture = m_Images(ni)
        ImgItem(ni - 1).Move 2, 17 * (ni - 1)
    Next ni
LimitOfItems:
    ''Check If dropdown list exceeds screen area then dropup
    ''  If Is OK, show. . .
    '' This is a suggestion made by Charles P. V.
    If Y + lHeight <= Screen.Height Then
        Me.Move X, Y, lWidth, lHeight
    Else
        Me.Move X, Y - lHeight - m_lParentHeight * Screen.TwipsPerPixelY, lWidth, lHeight '- parent.ScaleHeight
    End If
    
    'Show The DropDown List
    If bAnimateWindow Then
        AnimateWindow Me.hwnd, 250, AW_VER_POSITIVE + AW_SLIDE + AW_ACTIVATE
    Else
        show
    End If
    
    picGroup.Move 0, 0, ScaleWidth - 4, ht - 4
    vsb.Move ScaleWidth - vsb.Width - 2, 0, vsb.Width, ScaleHeight - 2
    pScroller.Move 1, 1, ScaleWidth - 2, ScaleHeight - 2
    iPrevPos = 0
    iPos = iSelectedItem
    On Error Resume Next
    If iPos > 8 Then vsb.Value = iPos - 7
    
    'SetWindowPos hWnd, -1, 0, 0, 0, 0, 1 Or 2
    Me.SetFocus
    Form_Paint
    DrawAll iPos
    iPrevPos = iPos
    temiPos = iPos
    
    '<EhFooter>
    Exit Function

PopUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.PopUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Sub vsb_Scroll()
    '<EhHeader>
    On Error GoTo vsb_Scroll_Err
    '</EhHeader>
    vsb_Change
    '<EhFooter>
    Exit Sub

vsb_Scroll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.wndDown.vsb_Scroll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
