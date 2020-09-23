VERSION 5.00
Begin VB.UserControl xpWellsTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   DrawWidth       =   56
   MaskColor       =   &H00974D37&
   PropertyPages   =   "xpWellsTab.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "xpWellsTab.ctx":002A
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   720
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer tmr_addctrl 
      Interval        =   1000
      Left            =   2520
      Top             =   1200
   End
End
Attribute VB_Name = "xpWellsTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CSEH: ErrMsgBox
Option Explicit
'This code is not mine
' I (drkIIRaziel) did just some bugfixes /changes
' namely :
' I fixed a bug that caused the last tab to do not be "unhoted" if the mouse was
' on it's right ..
' I changed the behavior of the control so that the yellow line over the tabs
' does not apear at teh selected ones..

'Original Code header

'Acknowledgements:

'Ariad Software.
'For letting me look through there ToolBar code
'to see how they use Property Pages

'Manjula Dharmawardhana at www.manjulapra.com.
'For his simple Common Dialog without the .OCX sample

'Special Thanks:
'Steve McMahon ( The Man ) at www.vbaccelerator.com
'for showing us mere mortals how to make quality ActiveX controls.
'Without his generosity and skills, this control would not have happened.

'Planet Source Code, and the people who submit there code:
'For providing the #1 source code site for VB`ers on the net.

    Dim cPic As cImageManipulation
    Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

'Tab Alignment
'Not yet implimented..
    Public Enum eTabAlignment
        tTop = 0
        tLeft = 1
        tRight = 2
        tBottom = 3
    End Enum
    Private eTab As eTabAlignment
'//
Private m_MouseOver                 As Boolean
Private iSelectedTab                As Long
Private iHotTab                     As Long
Private MouseInBody                 As Boolean
Private MouseInTab                  As Boolean
Private hasFocus                    As Boolean
Private sAccessKeys                 As String
Private iPrevTab                    As Integer
'Property Variables
    Private iTabHeight              As Long
    Private oBackColor              As OLE_COLOR
    Private oForeColor              As OLE_COLOR
    Private oActiveForeColor        As OLE_COLOR
    Private oForeColorHot           As OLE_COLOR
    Private oFrameColor             As OLE_COLOR
    Private oMaskColor              As OLE_COLOR
    Private iNumberOfTabs           As Long
    Private Tabs()                  As New cTabs
    Dim rcTabs()                    As RECT
    Dim rcBody                      As RECT
'Events
Event TabPressed(PreviousTab As Integer)
Event MouseIn(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim WithEvents sco As SubclassEventImpl
Attribute sco.VB_VarHelpID = -1

Private Sub sco_BefWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, CallWProc As Boolean, CallAftProc As Boolean, ByVal pOldProc As Long)
  
  If uMsg = WM_LBUTTONDOWN Then
      Call UserControl_MouseDown(vbLeftButton, 0, lParam And &HFFFF&, lParam \ &H10000 And &HFFFF&)
  End If
  
End Sub


Private Sub tmr_addctrl_Timer()
    AddCotnrolsToTab
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
Dim i As Long
    i = GetTabAccessKey(KeyAscii)
    SelectedTab = i
End Sub

Private Sub UserControl_Initialize()

    If App.LogMode <> 0 Then
        Set sco = New SubclassEventImpl
        sco.SubClass UserControl.hwnd, wproc_notify
    End If
    
End Sub

Private Sub UserControl_InitProperties()
    AddTab
    iTabHeight = 22
    oBackColor = UserControl.parent.BackColor
    UserControl.BackColor = oBackColor
    oForeColor = vbButtonText
    UserControl.ForeColor = oForeColor
    oActiveForeColor = RGB(56, 80, 152)
    oForeColorHot = RGB(0, 0, 255)
    oFrameColor = RGB(152, 160, 160)
    oMaskColor = RGB(255, 0, 255)
    Set UserControl.Font = Ambient.Font
    DrawTab
End Sub

Private Sub UserControl_Terminate()
    Erase Tabs
    Erase rcTabs
    If sco.IsSubclassed = False Then sco.UnSubClass
End Sub

Private Sub UserControl_GotFocus()
    'hasFocus = True
    DrawTab
End Sub

Private Sub UserControl_LostFocus()
    'hasFocus = False
    DrawTab
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hRgn        As Long
Dim i           As Long

    For i = 1 To iNumberOfTabs
    hRgn = CreateRectRgnIndirect(rcTabs(i))
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            iSelectedTab = i
            SelectedTab = i
            DeleteObjectReference hRgn
        Else
            DeleteObjectReference hRgn
            RaiseEvent MouseDown(Button, Shift, X, Y)
        End If
    Next i
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hRgn            As Long
Dim i               As Long
Dim iLocalHotTab    As Long
Dim DoRedraw        As Boolean
    If MouseOver(UserControl.hwnd) = True Then
        RaiseEvent MouseIn(Button, Shift, X, Y)
        hRgn = CreateRectRgnIndirect(rcBody)
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            iHotTab = 0
            If MouseInBody = False Then
                MouseInTab = False
                DrawTab
                MouseInBody = True
                DeleteObjectReference hRgn
                Exit Sub
            End If
        End If
        'iHotTab = -1
        For i = 1 To iNumberOfTabs
        hRgn = CreateRectRgnIndirect(rcTabs(i))
            If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
                iLocalHotTab = i
                If iLocalHotTab <> iHotTab Then
                    DoRedraw = True
                    MouseInTab = True
                    MouseInBody = False
                    iHotTab = i
                End If
                DeleteObjectReference hRgn
                Exit For
            Else
                If iHotTab <> 0 Then
                    DoRedraw = True
                    MouseInTab = False
                    MouseInBody = False
                    iHotTab = 0
                End If
                DeleteObjectReference hRgn
            End If
        Next i
        If DoRedraw = True Then
            DrawTab
        End If
    Else
        RaiseEvent MouseOut(Button, Shift, X, Y)
        iHotTab = 0
        DrawTab
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft, vbKeyUp
            If iSelectedTab = 1 Then
                SelectedTab = 1
            Else
                SelectedTab = SelectedTab - 1
            End If
        Case vbKeyRight, vbKeyDown
            If iSelectedTab = iNumberOfTabs Then
                SelectedTab = iNumberOfTabs
            Else
                SelectedTab = SelectedTab + 1
            End If
    End Select
End Sub

Private Sub UserControl_Resize()
    DrawTab
End Sub

'CSEH: ErrMsgBox
Public Sub DrawVisibleControls()
'Dim t As Object
'On Error Resume Next
    'For Each t In ContainedControls
    '
    '    If t.WhatsThisHelpID = 500 + iSelectedTab * 200 Then
    '        'SetParent t.hwnd, UserControl.hwnd
    '        t.visible = True
    '    Else
    '        'SetParent t.hwnd, Picture1.hwnd
    '        t.visible = False
    '    End If
    '
    'Next t
    
End Sub

Public Sub AddCotnrolsToTab()
Dim t As Object

    'For Each t In ContainedControls
    '    If t.WhatsThisHelpID < 500 Then
    '        t.WhatsThisHelpID = 500 + iSelectedTab * 200
    '    End If
    'Next t
    
    DrawVisibleControls
    
End Sub
Public Property Get TabHeight() As Long
    TabHeight = iTabHeight
End Property

Public Property Let TabHeight(ByVal NewTabHeight As Long)
    iTabHeight = NewTabHeight
    PropertyChanged "TabHeight"
    DrawTab
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = oBackColor
End Property

Public Property Let BackColor(ByVal newBackColor As OLE_COLOR)
    oBackColor = newBackColor
    UserControl.BackColor = oBackColor
    PropertyChanged "BackColor"
    DrawTab
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = oForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    oForeColor = NewForeColor
    UserControl.ForeColor = oForeColor
    PropertyChanged "ForeColor"
    DrawTab
End Property

Public Property Get ForeColorActive() As OLE_COLOR
    ForeColorActive = oActiveForeColor
End Property

Public Property Let ForeColorActive(ByVal NewForeColorActive As OLE_COLOR)
    oActiveForeColor = NewForeColorActive
    PropertyChanged "ForeColorActive"
    DrawTab
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = oFrameColor
End Property

Public Property Let FrameColor(ByVal NewFrameColor As OLE_COLOR)
    oFrameColor = NewFrameColor
    PropertyChanged "FrameColor"
    DrawTab
End Property

Public Property Get TabWidth(ByVal Index As Long) As Long
    TabWidth = Tabs(Index).TabWidth
End Property

Public Property Let TabWidth(ByVal Index As Long, ByVal NewTabWidth As Long)
    Tabs(Index).TabWidth = NewTabWidth
    DrawTab
    PropertyChanged "TabWidth"
End Property

Public Property Get TabCaption(ByVal Index As Long) As String
    TabCaption = Tabs(Index).TabCaption
End Property

Public Property Let TabCaption(ByVal Index As Long, ByVal NewTabCaption As String)
    Tabs(Index).TabCaption = NewTabCaption
    SetTabAccessKeys
    DrawTab
    PropertyChanged "TabCaption"
End Property

Public Property Get TabPicture(ByVal Index As Long) As StdPicture
    Set TabPicture = Tabs(Index).TabIcon
End Property

Public Property Set TabPicture(ByVal Index As Long, ByVal NewTabPicture As StdPicture)
    Set Tabs(Index).TabPicture = NewTabPicture
    DrawTab
    PropertyChanged "TabPicture"
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set UserControl.Font = NewFont
    PropertyChanged "Font"
    DrawTab
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = oMaskColor
End Property

Public Property Let MaskColor(ByVal NewMaskColor As OLE_COLOR)
    oMaskColor = NewMaskColor
    PropertyChanged "MaskColor"
    DrawTab
End Property

Public Property Get NumberOfTabs() As Long
    NumberOfTabs = iNumberOfTabs
End Property

Public Property Get Alignment() As eTabAlignment
    Alignment = eTab
End Property

Public Property Let Alignment(ByVal NewAlignment As eTabAlignment)
    eTab = NewAlignment
    PropertyChanged "Alignment"
    DrawTab
End Property

Public Property Get ForeColorHot() As OLE_COLOR
    ForeColorHot = oForeColorHot
End Property

Public Property Let ForeColorHot(ByVal NewForeColorHot As OLE_COLOR)
    oForeColorHot = NewForeColorHot
    PropertyChanged "ForeColorHot"
End Property

Public Property Get SelectedTab() As Long
    SelectedTab = iSelectedTab
End Property

Public Property Let SelectedTab(ByVal NewSelectedTab As Long)
    
    AddCotnrolsToTab
    iSelectedTab = NewSelectedTab
    RaiseEvent TabPressed(iPrevTab)
    iPrevTab = iSelectedTab
    PropertyChanged "SelectedTab"
    DrawTab
    DrawVisibleControls
    
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Long
    On Error Resume Next
    With PropBag
        Alignment = .ReadProperty("Alignment", eTab)
        TabHeight = .ReadProperty("TabHeight", 25)
        BackColor = .ReadProperty("BackColor", RGB(240, 240, 224))
        ForeColor = .ReadProperty("ForeColor", vbButtonText)
        ForeColorActive = .ReadProperty("ForeColorActive", RGB(56, 80, 152))
        ForeColorHot = .ReadProperty("ForeColorHot", RGB(0, 0, 255))
        FrameColor = .ReadProperty("FrameColor", RGB(152, 160, 160))
        MaskColor = .ReadProperty("MaskColor", RGB(255, 0, 255))
        SelectedTab = .ReadProperty("SelectedTab", 1)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        iNumberOfTabs = .ReadProperty("NumberOfTabs", 0)
        ReDim Tabs(iNumberOfTabs) As New cTabs
    End With
    For i = 1 To iNumberOfTabs
        With Tabs(i)
            .TabWidth = PropBag.ReadProperty("TabWidth" & i)
            .TabCaption = PropBag.ReadProperty("TabText" & i)
            Set .TabIcon = PropBag.ReadProperty("TabPicture" & i)
        End With
    Next i
    SetTabAccessKeys
    DrawTab
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long
    With PropBag
        .WriteProperty "Alignment", eTab
        .WriteProperty "TabHeight", iTabHeight
        .WriteProperty "BackColor", oBackColor
        .WriteProperty "ForeColor", oForeColor
        .WriteProperty "ForeColorActive", oActiveForeColor
        .WriteProperty "ForeColorHot", oForeColorHot
        .WriteProperty "FrameColor", oFrameColor
        .WriteProperty "MaskColor", oMaskColor
        .WriteProperty "SelectedTab", iSelectedTab
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "NumberOfTabs", iNumberOfTabs
    End With
    For i = 1 To iNumberOfTabs
        With Tabs(i)
            PropBag.WriteProperty "TabWidth" & i, .TabWidth
            PropBag.WriteProperty "TabText" & i, .TabCaption
            PropBag.WriteProperty "TabPicture" & i, .TabIcon, Nothing
        End With
    Next i
End Sub

Public Sub DrawTab()
    Select Case eTab
        Case 0
            DrawTabTop
        Case 1
            DrawTabTop
        Case 2
        
        Case 3
    End Select
End Sub

Public Sub DrawTabTop()
Dim rc As RECT
Dim i As Long
Dim rcTemp              As RECT
Dim X                   As Single
Dim Y                   As Single
Dim X1                  As Single
Dim Y1                  As Single
Dim iOffSet             As Long
Dim iBodyHeight         As Long
Dim pX                  As Single
Dim pY                  As Single

    Set cPic = New cImageManipulation
    
    rc = GetRect(UserControl.hwnd)
    rc.Top = rc.Top + iTabHeight
    Cls
    DrawASquare UserControl.hdc, rc, oFrameColor
    Y = 2
    X = 2
    'Height of the Tab
    Y1 = iTabHeight
    '//
    
    'Loop through the Tabs
    For i = 1 To iNumberOfTabs
        With Tabs(i)
        'Position the Tabs.
            .TabLeft = X
            .TabTop = Y
            X1 = .TabWidth
            .TabHeight = Y1
        '//
        'Create a RECT area using the above dimentions to draw into.
            With rc
                .Left = X
                .Top = Y
                .Right = .Left + X1
                .Bottom = Y1
                ReDim Preserve rcTabs(i)
                'Save the rect
                rcTabs(i) = rc
                'Left
                DrawALine UserControl.hdc, .Left, .Top + 2, .Left, .Bottom, oFrameColor
                'Top
                DrawALine UserControl.hdc, .Left + 2, .Top, .Right - 1, .Top, oFrameColor
                'Right
                DrawALine UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, oFrameColor
                'Left Corner
                DrawADot UserControl.hdc, .Left + 1, .Top + 1, oFrameColor
                'Right Corner
                DrawADot UserControl.hdc, .Right - 1, .Top + 1, oFrameColor
            End With
            X = (X + 2) + .TabWidth
        End With
    Next i
'Draw the gradients
    For i = 1 To iNumberOfTabs
        If i <> iSelectedTab Then
            ClearRect rcTemp
            CopyTheRect rcTemp, rcTabs(i)
            rcTemp.Left = rcTemp.Left + 1
            DrawGradient UserControl.hdc, TranslateColorToRGB(oBackColor, 0, 0, 0, 5), TranslateColorToRGB(oBackColor, 0, 0, 0, -15), rcTemp, iTabHeight - 5
        End If
    Next i
    
'Draw the selected and hot tab ( if required ).
    ClearRect rc

    If iNumberOfTabs > 0 Then
        If iSelectedTab = 0 Then
            iSelectedTab = 1
        End If
        With Tabs(iSelectedTab)
            rcTabs(iSelectedTab) = ResizeRect(rcTabs(iSelectedTab), 2, 2)
            DrawASquare UserControl.hdc, rcTabs(iSelectedTab), oBackColor, True
            
            DrawALine UserControl.hdc, rcTabs(iSelectedTab).Left, rcTabs(iSelectedTab).Top + 2, rcTabs(iSelectedTab).Left, rcTabs(iSelectedTab).Bottom - 1, oFrameColor
            DrawALine UserControl.hdc, rcTabs(iSelectedTab).Right, rcTabs(iSelectedTab).Top + 2, rcTabs(iSelectedTab).Right, rcTabs(iSelectedTab).Bottom - 1, oFrameColor
            DrawALine UserControl.hdc, rcTabs(iSelectedTab).Left + 2, rcTabs(iSelectedTab).Top, rcTabs(iSelectedTab).Right - 1, rcTabs(iSelectedTab).Top, oFrameColor
            DrawADot UserControl.hdc, 0, iTabHeight + 1, oFrameColor
            
            DrawADot UserControl.hdc, rcTabs(iSelectedTab).Left + 1, rcTabs(iSelectedTab).Top + 1, oFrameColor
            DrawADot UserControl.hdc, rcTabs(iSelectedTab).Right - 1, rcTabs(iSelectedTab).Top + 1, oFrameColor
            
            'HighLights
            If hasFocus = True Then
                DrawALine UserControl.hdc, rcTabs(iSelectedTab).Left + 2, rcTabs(iSelectedTab).Top, rcTabs(iSelectedTab).Right - 1, rcTabs(iSelectedTab).Top, RGB(232, 144, 40)
                DrawALine UserControl.hdc, rcTabs(iSelectedTab).Left + 2, rcTabs(iSelectedTab).Top + 1, rcTabs(iSelectedTab).Right - 1, rcTabs(iSelectedTab).Top + 1, RGB(255, 208, 56)
                DrawALine UserControl.hdc, rcTabs(iSelectedTab).Left + 1, rcTabs(iSelectedTab).Top + 2, rcTabs(iSelectedTab).Right, rcTabs(iSelectedTab).Top + 2, RGB(255, 200, 56)
            End If
        End With
        
        'Hot tab
        If iHotTab <> 0 And iHotTab <> iSelectedTab Then
            With Tabs(iHotTab)
                DrawALine UserControl.hdc, rcTabs(iHotTab).Left + 2, rcTabs(iHotTab).Top, rcTabs(iHotTab).Right - 1, rcTabs(iHotTab).Top, RGB(232, 144, 40)
                DrawALine UserControl.hdc, rcTabs(iHotTab).Left + 2, rcTabs(iHotTab).Top + 1, rcTabs(iHotTab).Right - 1, rcTabs(iHotTab).Top + 1, RGB(255, 208, 56)
                DrawALine UserControl.hdc, rcTabs(iHotTab).Left + 1, rcTabs(iHotTab).Top + 2, rcTabs(iHotTab).Right, rcTabs(iHotTab).Top + 2, RGB(255, 200, 56)
            End With
        End If
    End If
'//


    rcBody = GetRect(UserControl.hwnd)
    rcBody.Top = rcBody.Top + iTabHeight
    ClearRect rcTemp
    CopyTheRect rcTemp, rcBody
    ResizeRect rcTemp, -1, -1
    PositionRect rcTemp, 0, 0
    iBodyHeight = rcTemp.Bottom - rcTemp.Top - 2
'//

'Draw the caption and pictures
    For i = 1 To iNumberOfTabs
        ClearRect rcTemp
        CopyTheRect rcTemp, rcTabs(i)
        GetTextRect UserControl.hdc, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp
        GetPictureSize Tabs(i).TabIcon, pX, pY
        PositionRect rcTemp, 8 + pX, ((iTabHeight + rcTemp.Top) - rcTemp.Bottom) / 2
        cPic.PaintTransparentPicture UserControl.hdc, Tabs(i).TabIcon, rcTabs(i).Left + 4, ((iTabHeight + rcTemp.Top) - rcTemp.Bottom) / 2, pX, pX, , , oMaskColor
        If i = iSelectedTab Then
            SetTheTextColor UserControl.hdc, oActiveForeColor
            DrawTheText UserControl.hdc, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp, DrawTextFlags.Center
        Else
        If i = iHotTab Then
            SetTheTextColor UserControl.hdc, oForeColorHot
            DrawTheText UserControl.hdc, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp, DrawTextFlags.Center
        Else
            SetTheTextColor UserControl.hdc, oForeColor
            DrawTheText UserControl.hdc, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp, DrawTextFlags.Center
        End If
        End If
    Next i
    Set cPic = Nothing
'//
Refresh
End Sub

Public Function AddTab(Optional iTabWidth As Long = 60, _
                        Optional sTabText As String = "", _
                        Optional pTabPicture As StdPicture = Nothing) As Long
    Dim i As Long
    
        iNumberOfTabs = iNumberOfTabs + 1
        ReDim Preserve Tabs(iNumberOfTabs) As New cTabs
                
        With Tabs(iNumberOfTabs)
            .TabWidth = iTabWidth
            If sTabText = "" Then
                sTabText = "Tab" & iNumberOfTabs
                .TabCaption = sTabText
            Else
                .TabCaption = sTabText
                SetTabAccessKeys
            End If
            Set .TabPicture = pTabPicture
        End With
        PropertyChanged "NumberOfTabs"
        AddTab = iNumberOfTabs
        DrawTab
End Function

Public Sub DeleteTab()
    If iNumberOfTabs > 1 Then
        iNumberOfTabs = iNumberOfTabs - 1
    End If
    PropertyChanged "NumberOfTabs"
    DrawTab
End Sub

Public Sub GetPictureSize(pPicture As StdPicture, picX As Single, picY As Single)
    picX = 0
    picY = 0
    If pPicture Is Nothing Then Exit Sub
    picX = ScaleX(pPicture.Width, 8, 3)
    picY = ScaleY(pPicture.Height, 8, 3)
End Sub



Public Sub SetTabAccessKeys()
Dim i As Long
    sAccessKeys = ""
    For i = 1 To iNumberOfTabs
        sAccessKeys = sAccessKeys & SetAccessKey(Tabs(i).TabCaption)
    Next i
    UserControl.AccessKeys = sAccessKeys
End Sub

Public Function GetTabAccessKey(iKey As Integer) As Long
Dim i As Long
Dim sChr As String
    For i = 1 To Len(sAccessKeys)
        sChr = Mid$(sAccessKeys, i, 1)
        If Asc(sChr) = iKey Then
            GetTabAccessKey = i
            Exit For
        End If
    Next i
End Function

