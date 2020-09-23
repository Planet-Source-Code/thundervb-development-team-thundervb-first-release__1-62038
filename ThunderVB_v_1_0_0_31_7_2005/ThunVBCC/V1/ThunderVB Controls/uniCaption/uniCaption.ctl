VERSION 5.00
Begin VB.UserControl uniCaption 
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   4545
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "uniCaption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1515
   End
End
Attribute VB_Name = "uniCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'    ------------------
'    --- uniCaption --- version 0.7.1
'    ------------------
'
'    made by Libor for ThunderVB
'    note : this UserControl needs "SubClassing Engine" made by drkIIRaziel
'
' Maybe you know that VB forms are ANSI windows. We need to display unicode string in caption (in titlebar).
' This UserControl does the job by drawing the unicode string on the titlebar.
'
' public properties
'  - Caption        - unicode caption
'  - CaptionColor   - caption color
'  - CaptionAlign   - caption align
'  - ShowVbCaption  - show VB caption
'
' public methods
'  - Redraw         - refresh unicode caption
'  - Caption_Start  - show unicode caption
'  - Caption_Stop   - hide unicode caption
'

Implements ISubclass_Callbacks

Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function GetTextExtentPoint Lib "gdi32.dll" Alias "GetTextExtentPointW" (ByVal hdc As Long, ByVal pszString As Long, ByVal cbString As Long, lpSize As Size) As Long

Private Type Size
    cx As Long
    cy As Long
End Type

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hdc As Long, ByVal pStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_SINGLELINE As Long = &H20

Public Enum TEXT_ALIGN
    Align_Left = &H0
    Align_Right = &H2
    Align_Top = &H0
    Align_Bottom = &H8
    Align_vcenter = &H4
    Align_Center = &H1
    Align_RtLreading = &H20000
End Enum

Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Const TRANSPARENT As Long = 1

Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Const DEFAULT_CHARSET As Long = 1
Private Const LF_FACESIZE  As Long = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type


Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPI_GETNONCLIENTMETRICS As Long = 41

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Const SM_CXBORDER As Long = 5
Private Const SM_CXSIZE As Long = 30
Private Const SM_CXFRAME As Long = 32
Private Const SM_CXSIZEFRAME As Long = SM_CXFRAME
Private Const SM_CYFRAME As Long = 33
Private Const SM_CYSIZEFRAME As Long = SM_CYFRAME
Private Const SM_CYBORDER As Long = 6
Private Const SM_CXDLGFRAME As Long = 7
Private Const SM_CXFIXEDFRAME As Long = SM_CXDLGFRAME
Private Const SM_CYDLGFRAME As Long = 8
Private Const SM_CYFIXEDFRAME As Long = SM_CYDLGFRAME
Private Const SM_CYCAPTION As Long = 4
Private Const SM_CXSMICON As Long = 49

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE

Private Const WM_NCACTIVATE As Long = &H86
Private Const WM_NCPAINT As Long = &H85

Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16

Private Const WS_EX_APPWINDOW As Long = &H40000

Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_SIZEBOX As Long = WS_THICKFRAME
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_EX_TOOLWINDOW As Long = &H80&

Private Const ARIAL_UNICODE As String = "Arial Unicode MS"

Private hwnd As Long    'hWnd of window on which we will draw unicode caption
Private hFont As Long   'handle of unicode font

'--- properties ---
Private sCaption As String             'unicode text (caption)
Private lCaptionColor As Long          'color of caption
Private lCaptionAlign As Long          'align of caption (use TEXT_ALIGN enum)
Private bShowVbCaption As Boolean

Private xLeft As Long, xRight As Long  'position of caption in window
Private yTop As Long, yBottom

'we use after callback for drawing caption
Private Sub ISubclass_Callbacks_AftWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal CalledWProc As Boolean, ByVal pOriginalPRoc As Long)
    '<EhHeader>
    On Error GoTo ISubclass_Callbacks_AftWndProc_Err
    '</EhHeader>
    
    'draw caption
    If uMsg = WM_NCPAINT Or uMsg = WM_NCACTIVATE Then
        Call DrawCaption
    End If

    '<EhFooter>
    Exit Sub

ISubclass_Callbacks_AftWndProc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.ISubclass_Callbacks_AftWndProc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ISubclass_Callbacks_BefWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, CallWProc As Boolean, CallAftProc As Boolean, ByVal pOriginalPRoc As Long)
'
    '<EhHeader>
    On Error GoTo ISubclass_Callbacks_BefWndProc_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

ISubclass_Callbacks_BefWndProc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.ISubclass_Callbacks_BefWndProc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ISubclass_Callbacks_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal pOriginalPRoc As Long)
'
    '<EhHeader>
    On Error GoTo ISubclass_Callbacks_WindowProc_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

ISubclass_Callbacks_WindowProc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.ISubclass_Callbacks_WindowProc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'draw caption on titlebar
Private Sub DrawCaption()
    '<EhHeader>
    On Error GoTo DrawCaption_Err
    '</EhHeader>
Dim hdc As Long, tWinRect As RECT, tCapRect As RECT
Dim tSize As Size, sTemp As String, sCurCaption As String

'get DC
hdc = GetWindowDC(hwnd)

'get window rect :-)
GetWindowRect hwnd, tWinRect
'get rect of caption
SetRect tCapRect, xLeft, yTop, tWinRect.lRight - tWinRect.lLeft - xRight, yBottom

'create new DC and bitmap
Dim hBMPsource As Long, hNewDC As Long
hBMPsource = CreateCompatibleBitmap(hdc, tCapRect.lRight - tCapRect.lLeft, tCapRect.lBottom - tCapRect.lTop)
hNewDC = CreateCompatibleDC(hdc)
SelectObject hNewDC, hBMPsource

'get window text
sCurCaption = GetWinText

'we do not want VB caption so clear it
If bShowVbCaption = False Then SendMessage hwnd, WM_SETTEXT, 0, ByVal ""
'copy titlebar background
BitBlt hNewDC, 0, 0, tCapRect.lRight - tCapRect.lLeft, tCapRect.lBottom - tCapRect.lTop, hdc, tCapRect.lLeft, tCapRect.lTop, vbSrcCopy
'restore VB caption
If bShowVbCaption = False Then SendMessage hwnd, WM_SETTEXT, 0, ByVal sCurCaption

'set new font and text color
SelectObject hNewDC, hFont
SetBkMode hNewDC, TRANSPARENT
SetTextColor hNewDC, lCaptionColor

sTemp = sCaption
recalc:
'get size of unicode caption
GetTextExtentPoint hNewDC, StrPtr(sTemp), Len(sTemp), tSize

'is unicode text smaller then caption?
If tCapRect.lRight - tCapRect.lLeft >= tSize.cx Then
Dim tt As RECT
    'new rect
    SetRect tt, 0, 0, tCapRect.lRight - tCapRect.lLeft, tCapRect.lBottom - tCapRect.lTop
    'draw text on out bitmap
    DrawText hNewDC, StrPtr(sTemp), Len(sTemp), tt, lCaptionAlign Or DT_SINGLELINE
Else
    'text is bigger then caption so truncate it (replace last char with "...")
    If Mid(sTemp, Len(sTemp) - 2, 3) = "..." Then
        sTemp = Mid(sTemp, 1, Len(sTemp) - 4) & "..."
    Else
        sTemp = Mid(sTemp, 1, Len(sTemp) - 3) & "..."
    End If
    'go back
    GoTo recalc
End If

'copy bitmap to titlebar
BitBlt hdc, tCapRect.lLeft, tCapRect.lTop, tCapRect.lRight - tCapRect.lLeft, tCapRect.lBottom - tCapRect.lTop, hNewDC, 0, 0, vbSrcCopy

'clean up
DeleteObject hNewDC
DeleteObject hBMPsource

ReleaseDC hwnd, hdc
    
    '<EhFooter>
    Exit Sub

DrawCaption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.DrawCaption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'start subclassing (hWindow - handle of window on our unicode caption will be drawn)
Public Function Caption_Start(ByVal hWindow As Long, ByVal sFontName As String) As Boolean
    '<EhHeader>
    On Error GoTo Caption_Start_Err
    '</EhHeader>
Dim tNonClient As NONCLIENTMETRICS, tFont As LOGFONT, i As Long
Dim lStyle As Long, lExStyle As Long

    Caption_Start = False
    hwnd = hWindow

    'get style of window
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)

    'if window has not caption or it's a tool window then exit
    If (lStyle And WS_CAPTION) <> WS_CAPTION Then Exit Function
    If (lExStyle And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW Then Exit Function

    'we need info about font in titlebar
    tNonClient.cbSize = Len(tNonClient)
    If SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(tNonClient), ByVal VarPtr(tNonClient), 0) = 0 Then Exit Function

    'replace font name with our unicode font
    tFont = tNonClient.lfCaptionFont
    For i = 1 To LF_FACESIZE
        If i <= Len(sFontName) Then
            tFont.lfFaceName(i) = Asc(Mid(sFontName, i, 1))
        Else
            tFont.lfFaceName(i) = 0
        End If
    Next i
    tFont.lfCharSet = DEFAULT_CHARSET

    'create font
    hFont = CreateFontIndirect(tFont)
    If hFont = 0 Then
        Exit Function
    End If

    'init
    xLeft = 0: xRight = 0: yTop = 0: yBottom = 0

    'get offset of caption
    xRight = GetSystemMetrics(SM_CXSIZE)
    If (lStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then xRight = xRight + GetSystemMetrics(SM_CXSIZE)
    If (lStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Then xRight = xRight + GetSystemMetrics(SM_CXSIZE)
    If (lStyle And WS_SYSMENU) = WS_SYSMENU Then xLeft = GetSystemMetrics(SM_CXSMICON) '+ 4

    If (lStyle And WS_SIZEBOX) = WS_SIZEBOX Then
    
        'sizable window

        xLeft = xLeft + GetSystemMetrics(SM_CXBORDER) + GetSystemMetrics(SM_CXSIZEFRAME)
        xRight = xRight + GetSystemMetrics(SM_CXSIZEFRAME) + GetSystemMetrics(SM_CXBORDER) '- 4

        yTop = GetSystemMetrics(SM_CYSIZEFRAME) + GetSystemMetrics(SM_CYBORDER)
        yBottom = GetSystemMetrics(SM_CYCAPTION)

    Else

        'dialog

        xLeft = xLeft + GetSystemMetrics(SM_CXFIXEDFRAME) + GetSystemMetrics(SM_CXBORDER)
        xRight = xRight + GetSystemMetrics(SM_CXFIXEDFRAME) + GetSystemMetrics(SM_CXBORDER)

        yTop = GetSystemMetrics(SM_CYFIXEDFRAME) + GetSystemMetrics(SM_CYBORDER)
        yBottom = GetSystemMetrics(SM_CYCAPTION)

    End If

    'subclass window
    SubClasshWnd hwnd, Me, wproc_notify
    Redraw
    
    Caption_Start = True

    '<EhFooter>
    Exit Function

Caption_Start_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.Caption_Start " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Sub Caption_stop()
    '<EhHeader>
    On Error GoTo Caption_stop_Err
    '</EhHeader>
    
    If hwnd = 0 Then Exit Sub
    
    'delete font and stop subclassing
    DeleteObject hFont
    UnSubClasshWnd hwnd
    Redraw
    hwnd = 0
    
    '<EhFooter>
    Exit Sub

Caption_stop_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.Caption_stop " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'refresh window
Public Sub Redraw()
    '<EhHeader>
    On Error GoTo Redraw_Err
    '</EhHeader>
    If hwnd = 0 Then Exit Sub
    SendMessage hwnd, WM_NCPAINT, 0, ByVal 0
    '<EhFooter>
    Exit Sub

Redraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.Redraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'get window caption
Private Function GetWinText() As String
    '<EhHeader>
    On Error GoTo GetWinText_Err
    '</EhHeader>
Dim l As Long, sText As String
    
    If hwnd = 0 Then Exit Function
    
    'get caption length
    l = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
    If l <> 0 Then
        'create buffer
        sText = Space(l + 1)
        'get text
        l = SendMessage(hwnd, WM_GETTEXT, Len(sText), ByVal sText)
        'trim null
        If l <> 0 Then GetWinText = Mid(sText, 1, l)
    End If

    '<EhFooter>
    Exit Function

GetWinText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.GetWinText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'show in taskbar (bShow - if true then show me in taskbar otherwise not)
Public Sub TaskBar(ByVal bShow As Boolean)
    '<EhHeader>
    On Error GoTo TaskBar_Err
    '</EhHeader>
Dim lStyle As Long
    
    If hwnd = 0 Then Exit Sub
    
    'get window style
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If bShow = True Then
        'show
        If (lStyle Or WS_EX_APPWINDOW) = lStyle Then Exit Sub
        lStyle = lStyle Or WS_EX_APPWINDOW
    Else
        'hide
        If (lStyle Or WS_EX_APPWINDOW) <> lStyle Then Exit Sub
        lStyle = lStyle Or WS_EX_APPWINDOW
        lStyle = lStyle And Not WS_EX_APPWINDOW
    End If
    
    'we have to refresh window, so we hide it
    LockWindowUpdate hwnd
    ShowWindow hwnd, SW_HIDE
    
    'change style
    SetWindowLong hwnd, GWL_EXSTYLE, lStyle
    
    'and show it
    ShowWindow hwnd, SW_SHOW
    LockWindowUpdate 0
    
    '<EhFooter>
    Exit Sub

TaskBar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.TaskBar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Property Get Caption() As String
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    Caption = sCaption
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Caption(ByVal sNewValue As String)
    '<EhHeader>
    On Error GoTo Caption_Err
    '</EhHeader>
    sCaption = sNewValue
    Call Redraw
    '<EhFooter>
    Exit Property

Caption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.Caption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get CaptionColor() As Long
    '<EhHeader>
    On Error GoTo CaptionColor_Err
    '</EhHeader>
    CaptionColor = lCaptionColor
    '<EhFooter>
    Exit Property

CaptionColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.CaptionColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let CaptionColor(ByVal lNewValue As Long)
    '<EhHeader>
    On Error GoTo CaptionColor_Err
    '</EhHeader>
    lCaptionColor = lNewValue
    Call Redraw
    '<EhFooter>
    Exit Property

CaptionColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.CaptionColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get CaptionAlign() As Long
    '<EhHeader>
    On Error GoTo CaptionAlign_Err
    '</EhHeader>
    CaptionAlign = lCaptionAlign
    '<EhFooter>
    Exit Property

CaptionAlign_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.CaptionAlign " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let CaptionAlign(ByVal lNewValue As TEXT_ALIGN)
    '<EhHeader>
    On Error GoTo CaptionAlign_Err
    '</EhHeader>
    lCaptionAlign = lNewValue
    Call Redraw
    '<EhFooter>
    Exit Property

CaptionAlign_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.CaptionAlign " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ShowVbCaption() As Boolean
    '<EhHeader>
    On Error GoTo ShowVbCaption_Err
    '</EhHeader>
    ShowVbCaption = bShowVbCaption
    '<EhFooter>
    Exit Property

ShowVbCaption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.ShowVbCaption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ShowVbCaption(ByVal bNewValue As Boolean)
    '<EhHeader>
    On Error GoTo ShowVbCaption_Err
    '</EhHeader>
    bShowVbCaption = bNewValue
    Call Redraw
    '<EhFooter>
    Exit Property

ShowVbCaption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.ShowVbCaption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ARIAL_UNICODE_MS() As String
    '<EhHeader>
    On Error GoTo ARIAL_UNICODE_MS_Err
    '</EhHeader>
    ARIAL_UNICODE_MS = ARIAL_UNICODE
    '<EhFooter>
    Exit Property

ARIAL_UNICODE_MS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.ARIAL_UNICODE_MS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
    lblCaption.Move 0, 0
    UserControl.Height = lblCaption.Height
    UserControl.Width = lblCaption.Width
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniCaption.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
