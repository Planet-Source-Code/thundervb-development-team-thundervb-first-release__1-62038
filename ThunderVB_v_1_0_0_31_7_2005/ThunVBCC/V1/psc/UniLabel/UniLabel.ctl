VERSION 5.00
Begin VB.UserControl UniLabel 
   CanGetFocus     =   0   'False
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   DefaultCancel   =   -1  'True
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   ToolboxBitmap   =   "UniLabel.ctx":0000
End
Attribute VB_Name = "UniLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' UniLabel version 1.1 - 9th of September 2004
' by Merri aka Vesa Piittinen <vesa@merri.net>
'
' Made this label so you don't need to have extra hassle
' on getting normal labels to show Unicode characters.
' NOTE! You must install Far-East language support to show
' most of the special Unicode characters. You can do this at
' Regional and Language Settings in the Control Panel.
'
' ----------------------------------------------------------------------
'
' Change history:
'
' Version 1.1 (09/09/2004)
' - added support for Windows 95/98/98SE/ME (no Unicode though)
' - added events (Change, Click, DblClick, MouseDown, MouseMove, MouseUp
'
' Version 1.0 (09/01/2004)
' - initial release
'
' ----------------------------------------------------------------------
'
' Future:
'
' - make it possible to have a command button mode
'   - label
'   - flat button
'   - button
'   - officexp button
'   - officexp bold (the borders 2 pixels wide instead of 1)
'   - set customizable the amount text moves when mouse down
' - I don't know if I have enough interest to do this, but atleast
'   I do have ideas! :D

Option Explicit

'constants for API calls
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_NOCLIP As Long = &H100
Private Const DT_RIGHT As Long = &H2
Private Const DT_WORDBREAK As Long = &H10

'custom types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'API declarations
Private Declare Function DrawTextANSI Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextUnicode Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpArrPtr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'public enums
'Public Enum uniLabelMode
'    ulbLabel
'    ulbFlatButton
'    ulbButton
'End Enum

'events
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'default public properties
Private Const m_def_Alignment As Byte = vbLeftJustify
Private Const m_def_AutoRedraw As Boolean = False
Private Const m_def_AutoSize As Boolean = False
Private Const m_def_WordWrap As Boolean = False

'public properties
Dim m_Alignment As AlignmentConstants
Dim m_AutoRedraw As Boolean
Dim m_AutoSize As Boolean
Dim m_BackColor As Long
Dim m_CaptionB() As Byte
Dim m_CaptionLen As Long
Dim m_ForeColor As Long
Dim m_WordWrap As Boolean

'helper variables
Dim m_Caption As String 'only used under Windows 95/98/98SE/ME
Dim m_DTMODE As Long
Dim m_RECT As RECT
Dim m_WINNT As Boolean
Public Property Get Alignment() As AlignmentConstants
    '<EhHeader>
    On Error GoTo Alignment_Err
    '</EhHeader>
    Alignment = m_Alignment
    '<EhFooter>
    Exit Property

Alignment_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.Alignment " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Alignment(ByVal NewAlignment As AlignmentConstants)
    '<EhHeader>
    On Error GoTo Alignment_Err
    '</EhHeader>
    m_Alignment = NewAlignment
    'repaint
    UpdateRect
    '<EhFooter>
    Exit Property

Alignment_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.Alignment " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get AutoRedraw() As Boolean
    '<EhHeader>
    On Error GoTo AutoRedraw_Err
    '</EhHeader>
    AutoRedraw = m_AutoRedraw
    '<EhFooter>
    Exit Property

AutoRedraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.AutoRedraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let AutoRedraw(ByVal NewMode As Boolean)
    '<EhHeader>
    On Error GoTo AutoRedraw_Err
    '</EhHeader>
    Dim OldMode As Boolean, EmptyImage As IPictureDisp
    OldMode = m_AutoRedraw
    m_AutoRedraw = NewMode
    UserControl.AutoRedraw = NewMode
    'when autoredraw mode changes, old content is set as picture: clear with an empty image
    If NewMode = False And OldMode = True Then UserControl.Picture = EmptyImage
    UpdateRect
    '<EhFooter>
    Exit Property

AutoRedraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.AutoRedraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get AutoSize() As Boolean
    '<EhHeader>
    On Error GoTo AutoSize_Err
    '</EhHeader>
    AutoSize = m_AutoSize
    '<EhFooter>
    Exit Property

AutoSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.AutoSize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let AutoSize(ByVal NewMode As Boolean)
    '<EhHeader>
    On Error GoTo AutoSize_Err
    '</EhHeader>
    m_AutoSize = NewMode
    'repaint
    UpdateRect
    '<EhFooter>
    Exit Property

AutoSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.AutoSize " & _
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
           "in ThunVBCC_v1_0.UniLabel.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
    m_BackColor = NewColor
    'change backcolor
    UserControl.BackColor = m_BackColor
    'repaint
    UserControl_Paint
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Caption() As String
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    Caption = m_CaptionB
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Caption(ByVal NewCaption As String)
    'check for null string
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    If LenB(NewCaption) > 0 Then
        'non-null string
        m_CaptionB = NewCaption
        m_CaptionLen = (UBound(m_CaptionB) + 1) \ 2
    Else
        'null string
        Erase m_CaptionB
        m_CaptionLen = 0
    End If
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
'only called under Windows 95/98/98SE/ME
Private Sub CaptionChange()
    'check if the array is empty or not
    '<EhHeader>
    On Error GoTo CaptionChange_Err
    '</EhHeader>
    If (Not m_CaptionB) <> True Then
        m_Caption = m_CaptionB
    Else
        m_Caption = vbNullString
    End If
    '<EhFooter>
    Exit Sub

CaptionChange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.CaptionChange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Get CaptionLen() As Long
    'return the length of the text
    '<EhHeader>
    On Error GoTo CaptionLen_Err
    '</EhHeader>
    CaptionLen = m_CaptionLen
    '<EhFooter>
    Exit Property

CaptionLen_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.CaptionLen " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let CaptionLen(ByVal NewLength As Long)
    '<EhHeader>
    On Error GoTo CaptionLen_Err
    '</EhHeader>
    Dim NewUBound As Long
    'check for invalid length
    If NewLength < 0 Then NewLength = 0
    'if somebody is really doing something this silly...
    If NewLength > &H3FFFFFFF Then NewLength = &H3FFFFFFF
    'and of course a check for this, no need to update!
    If m_CaptionLen = NewLength Then Exit Property
    'change it
    m_CaptionLen = NewLength
    If NewLength = 0 Then
        'null string
        Erase m_CaptionB
        'clear control
        UserControl.BackColor = m_BackColor
    Else
        'new array size
        NewUBound = NewLength + NewLength - 1
        'change byte array / string size
        ReDim Preserve m_CaptionB(NewUBound)
    End If
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
    '<EhFooter>
    Exit Property

CaptionLen_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.CaptionLen " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get CaptionAscB(ByVal Index As Long) As Byte
    'no out of bounds checking...
    '<EhHeader>
    On Error GoTo CaptionAscB_Err
    '</EhHeader>
    CaptionAscB = m_CaptionB(Index + Index)
    '<EhFooter>
    Exit Property

CaptionAscB_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.CaptionAscB " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let CaptionAscB(ByVal Index As Long, ByVal NewCode As Byte)
    'we have no out of bounds checking here...
    '<EhHeader>
    On Error GoTo CaptionAscB_Err
    '</EhHeader>
    m_CaptionB(Index + Index) = NewCode
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
    '<EhFooter>
    Exit Property

CaptionAscB_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.CaptionAscB " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get CaptionAscW(ByVal Index As Long) As Integer
    '<EhHeader>
    On Error GoTo CaptionAscW_Err
    '</EhHeader>
    Dim CurIndex As Long
    'check we are not out of bounds
    If Index < 0 Or Index > m_CaptionLen - 1 Then Exit Property
    'very minor speed optimization
    CurIndex = Index + Index
    'is the highest bit active?
    If (m_CaptionB(CurIndex + 1) And &H80) = 0 Then
        'not active
        'convert two bytes into an integer
        CaptionAscW = m_CaptionB(CurIndex) Or (CInt(m_CaptionB(CurIndex + 1)) * &H100)
    Else
        'active
        'convert two bytes into an integer and mark highest bit active
        CaptionAscW = m_CaptionB(CurIndex) Or (CInt(m_CaptionB(CurIndex + 1) And &H7F) * &H100) Or &H8000
    End If
    '<EhFooter>
    Exit Property

CaptionAscW_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.CaptionAscW " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let CaptionAscW(ByVal Index As Long, ByVal NewCode As Integer)
    '<EhHeader>
    On Error GoTo CaptionAscW_Err
    '</EhHeader>
    Dim Byte1 As Byte, Byte2 As Byte, CurIndex As Long
    'check we are not out of bounds
    If Index < 0 Or Index > m_CaptionLen - 1 Then Exit Property
    'rip lower byte
    Byte1 = CByte(NewCode And &HFF)
    'rip higher byte: check if the highest bit is active
    If NewCode < 0 Then
        'highest bit active
        Byte2 = ((NewCode And &H7F00) \ &H100) Or &H80
    Else
        'highest bit not active
        Byte2 = (NewCode And &H7F00) \ &H100
    End If
    'very minor speed optimization
    CurIndex = Index + Index
    'update data in array
    m_CaptionB(CurIndex) = Byte1
    m_CaptionB(CurIndex + 1) = Byte2
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
    '<EhFooter>
    Exit Property

CaptionAscW_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.CaptionAscW " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Font() As IFontDisp
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set Font = UserControl.Font
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Font(ByVal NewFont As IFontDisp)
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
    Set UserControl.Font = NewFont
    'repaint
    UpdateRect
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.Font " & _
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
           "in ThunVBCC_v1_0.UniLabel.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ForeColor_Err
    '</EhHeader>
    m_ForeColor = NewColor
    UserControl.ForeColor = m_ForeColor
    'repaint
    UserControl_Paint
    '<EhFooter>
    Exit Property

ForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Function GetCaptionB() As Byte()
    '<EhHeader>
    On Error GoTo GetCaptionB_Err
    '</EhHeader>
    GetCaptionB = m_CaptionB
    '<EhFooter>
    Exit Function

GetCaptionB_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.GetCaptionB " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Public Sub SetCaptionB(ByRef NewCaption() As Byte)
    'check if the array is empty
    '<EhHeader>
    On Error GoTo SetCaptionB_Err
    '</EhHeader>
    If (Not NewCaption) <> True Then
        'array with data
        m_CaptionB = NewCaption
        m_CaptionLen = (UBound(m_CaptionB) + 1) \ 2
    Else
        'empty array
        Erase m_CaptionB
        m_CaptionLen = 0
    End If
    If Not m_WINNT Then CaptionChange
    'repaint
    UpdateRect
    'event
    RaiseEvent Change
    '<EhFooter>
    Exit Sub

SetCaptionB_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.SetCaptionB " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub UpdateRect()
    '<EhHeader>
    On Error GoTo UpdateRect_Err
    '</EhHeader>
On Error GoTo 0
    Static HereAlready As Boolean
    'check if this sub is running already
    If HereAlready Then Exit Sub
    'mark we are running this sub
    HereAlready = True
    'alignment for painting
    Select Case Alignment
        Case vbLeftJustify
            'paint left justified
            m_DTMODE = DT_LEFT
        Case vbCenter
            'paint centered
            m_DTMODE = DT_CENTER
        Case vbRightJustify
            'paint right justified
            m_DTMODE = DT_RIGHT
    End Select
    'autosize mode?
    If Not m_AutoSize Then
        'no autosize, use control width and height as the painting area
        With m_RECT
            .Top = 0
            .Left = 0
            .Bottom = UserControl.ScaleHeight
            .Right = UserControl.ScaleWidth
        End With
        'set wordwrapping settings
        If m_WordWrap Then
            'paint wordwrapped
            m_DTMODE = m_DTMODE Or DT_WORDBREAK
        End If
    Else
        With m_RECT
            'reset all of these
            .Top = 0
            .Left = 0
            .Bottom = 0
            .Right = UserControl.ScaleWidth
            If m_WINNT Then 'UNICODE
                'set wordwrapping settings
                If m_WordWrap Then
                    'paint wordwrapped
                    m_DTMODE = m_DTMODE Or DT_WORDBREAK
                    'get paint area height
                    DrawTextUnicode UserControl.hdc, VarPtr(m_CaptionB(0)), m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control height
                    UserControl.Height = .Bottom
                    'set paint area width (correct height is returned in drawtext
                    .Right = UserControl.ScaleWidth
                Else
                    'no wordwrapping
                    m_DTMODE = m_DTMODE Or DT_NOCLIP
                    'get paint area width and height
                    DrawTextUnicode UserControl.hdc, VarPtr(m_CaptionB(0)), m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control width to the same as painting area width
                    UserControl.Width = .Right
                    'set control height to the same as painting area height
                    UserControl.Height = .Bottom
                End If
            Else 'ANSI
                'set wordwrapping settings
                If m_WordWrap Then
                    'paint wordwrapped
                    m_DTMODE = m_DTMODE Or DT_WORDBREAK
                    'get paint area height
                    DrawTextANSI UserControl.hdc, m_Caption, m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control height
                    UserControl.Height = .Bottom
                    'set paint area width (correct height is returned in drawtext
                    .Right = UserControl.ScaleWidth
                Else
                    'no wordwrapping
                    m_DTMODE = m_DTMODE Or DT_NOCLIP
                    'get paint area width and height
                    DrawTextANSI UserControl.hdc, m_CaptionB, m_CaptionLen, m_RECT, DT_CALCRECT Or m_DTMODE
                    'set control width to the same as painting area width
                    UserControl.Width = .Right
                    'set control height to the same as painting area height
                    UserControl.Height = .Bottom
                End If
            End If
        End With
    End If
    'mark we are done
    HereAlready = False
    'repaint
    UserControl_Paint
    '<EhFooter>
    Exit Sub

UpdateRect_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.UpdateRect " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Get WordWrap() As Boolean
    '<EhHeader>
    On Error GoTo WordWrap_Err
    '</EhHeader>
    WordWrap = m_WordWrap
    '<EhFooter>
    Exit Property

WordWrap_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.WordWrap " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let WordWrap(ByVal NewMode As Boolean)
    '<EhHeader>
    On Error GoTo WordWrap_Err
    '</EhHeader>
    m_WordWrap = NewMode
    UpdateRect
    '<EhFooter>
    Exit Property

WordWrap_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.WordWrap " & _
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
           "in ThunVBCC_v1_0.UniLabel.UserControl_Click " & _
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
           "in ThunVBCC_v1_0.UniLabel.UserControl_DblClick " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub UserControl_Initialize()
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>
    m_WINNT = (Environ$("OS") = "Windows_NT")
    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.UserControl_Initialize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub UserControl_InitProperties()
    'get default settings
    '<EhHeader>
    On Error GoTo UserControl_InitProperties_Err
    '</EhHeader>
    m_Alignment = m_def_Alignment
    m_AutoRedraw = m_def_AutoRedraw
    m_AutoSize = m_def_AutoSize
    m_BackColor = UserControl.parent.BackColor
    m_CaptionB = UserControl.Name
    m_CaptionLen = Len(UserControl.Name)
    m_ForeColor = UserControl.parent.ForeColor
    m_WordWrap = m_def_WordWrap
    With m_RECT
        .Top = 0
        .Left = 0
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    '<EhFooter>
    Exit Sub

UserControl_InitProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.UserControl_InitProperties " & _
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
           "in ThunVBCC_v1_0.UniLabel.UserControl_MouseDown " & _
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
           "in ThunVBCC_v1_0.UniLabel.UserControl_MouseMove " & _
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
           "in ThunVBCC_v1_0.UniLabel.UserControl_MouseUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub UserControl_Paint()
    '<EhHeader>
    On Error GoTo UserControl_Paint_Err
    '</EhHeader>
    Static HereAlready As Boolean
    'if we wanted to prevent blinking we'd set AutoRedraw = True
    'and do nothing here and do painting in UpdateRect
    
    If HereAlready Then Exit Sub
    'check we are not running this sub already to prevent running this forever
    HereAlready = True
    'clear before redraw
    UserControl.Cls
    'check if not a null string and also that the colors differ
    If m_CaptionLen > 0 And m_BackColor <> m_ForeColor Then
        'check OS
        If m_WINNT Then
            'Windows NT/2000/XP
            DrawTextUnicode UserControl.hdc, VarPtr(m_CaptionB(0)), m_CaptionLen, m_RECT, m_DTMODE
        Else
            'Windows 95/98/98SE/ME (no Unicode support)
            DrawTextANSI UserControl.hdc, m_Caption, m_CaptionLen, m_RECT, m_DTMODE
        End If
    End If
    'make change visible (this is the main reason we use HereAlready)
    If m_AutoRedraw Then UserControl.Refresh
    'mark we are done
    HereAlready = False
    '<EhFooter>
    Exit Sub

UserControl_Paint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.UserControl_Paint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'get all saved properties
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_BackColor = PropBag.ReadProperty("BackColor", UserControl.parent.BackColor)
    m_CaptionB = PropBag.ReadProperty("CaptionB", UserControl.Name)
    m_CaptionLen = PropBag.ReadProperty("CaptionLen", (UBound(m_CaptionB) + 1) \ 2)
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.parent.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", UserControl.parent.ForeColor)
    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
    'use the settings
    With UserControl
        .AutoRedraw = m_AutoRedraw
        .BackColor = m_BackColor
        .ForeColor = m_ForeColor
    End With
    'initial draw
    UpdateRect
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub UserControl_Resize()
    'refresh
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
    UpdateRect
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save all properties
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>
    PropBag.WriteProperty "Alignment", m_Alignment, m_def_Alignment
    PropBag.WriteProperty "AutoRedraw", m_AutoRedraw, m_def_AutoRedraw
    PropBag.WriteProperty "AutoSize", m_AutoSize, m_def_AutoSize
    PropBag.WriteProperty "BackColor", m_BackColor, UserControl.parent.BackColor
    PropBag.WriteProperty "CaptionB", m_CaptionB, UserControl.Name
    PropBag.WriteProperty "CaptionLen", m_CaptionLen, Len(UserControl.Name)
    PropBag.WriteProperty "Font", UserControl.Font, UserControl.parent.Font
    PropBag.WriteProperty "ForeColor", m_ForeColor, UserControl.parent.ForeColor
    PropBag.WriteProperty "WordWrap", m_WordWrap, m_def_WordWrap
    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.UniLabel.UserControl_WriteProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
