VERSION 5.00
Begin VB.UserControl vbaRichEdit 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ScaleHeight     =   3600
   ScaleWidth      =   5055
   ToolboxBitmap   =   "VBRichEdit.ctx":0000
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "vbAccelerator Rich Edit Control"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4875
   End
End
Attribute VB_Name = "vbaRichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ======================================================================
' Visit vbAccelerator at http://vbaccelerator.com/
' - the VB Programmer's Resource
' ======================================================================

' ======================================================================
' vbalRichEdit
' Copyright © 1998 Steve McMahon (steve@vbaccelerator.com)
' 14 June 1998
'
' A lightweight RichEdit control all in VB with lots of great features
' Requires:
'  mRichEdit.Bas
'  mWinGeneral.Bas
'  SSUBTMR.DLL
' ======================================================================

' ======================================================================
' Enums:
' ======================================================================
Public Enum ERECControlVersion
    eRICHED32
    eRICHED20
End Enum
Public Enum ERECFileTypes
    SF_TEXT = &H1
    SF_RTF = &H2
End Enum

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long

Public Enum ERECSetFormatRange
   ercSetFormatAll = SCF_ALL
   ercSetFormatSelection = SCF_SELECTION
   ercSetFormatWord = SCF_WORD Or SCF_SELECTION
End Enum
Public Enum ERECTextTypes
   ercTextNormal
   ercTextSuperscript
   ercTextSubscript
End Enum
Public Enum ERECViewModes
   ercDefault = 0
   ercWordWrap = 1
   ercWYSIWYG = 2
End Enum
' /*  UndoName info */
Public Enum ERECUndoTypeConstants
    ercUID_UNKNOWN = 0
    ercUID_TYPING = 1
    ercUID_DELETE = 2
    ercUID_DRAGDROP = 3
    ercUID_CUT = 4
    ercUID_PASTE = 5
End Enum
Public Enum ERECSelectionTypeConstants
   SEL_EMPTY = &H0
   SEL_TEXT = &H1
   SEL_OBJECT = &H2
   SEL_MULTICHAR = &H4
   SEL_MULTIOBJECT = &H8
End Enum
Public Enum ERECFindTypeOptions
   FR_DEFAULT = &H0
   FR_DOWN = &H1
   FR_WHOLEWORD = &H2
   FR_MATCHCASE = &H4&
End Enum
Public Enum ERECOptionTypeConstants
' /* Edit control options */
   ECO_AUTOWORDSELECTION = &H1&
   ECO_AUTOVSCROLL = &H40&
   ECO_AUTOHSCROLL = &H80&
   ECO_NOHIDESEL = &H100&
   ECO_READONLY = &H800&
   ECO_WANTRETURN = &H1000&
   ECO_SAVESEL = &H8000&
   ECO_SELECTIONBAR = &H1000000
   ECO_VERTICAL = &H400000                  ' /* FE specific */
End Enum

Public Enum ERECInbuiltShortcutConstants
   [_First] = 1
   ' Inbuilt methods
   ercCut_CtrlX = 1
   ercCopy_CtrlC = 2
   ercPaste_CtrlV = 3
   ercUndo_CtrlZ = 4
   ercSelectAll_CtrlA = 5
   
   ' Supplied methods:
   ercBold_CtrlB = 6
   ercItalic_CtrlI = 7
   ercUnderline_CtrlU = 8
   ercPrint_CtrlP = 9
   ercRedo_CtrlY = 10
   
   ercSuperscript_CtrlPlus = 11
   ercSubscript_CtrlMinus = 12
   
   ercNew_CtrlN = 13
   [_Last] = 13
End Enum

Public Enum ERECProgressTypeConstants
   ercNone = 0
   ercLoad = 1
   ercSave = 2
   ercPrint = 3
End Enum

Public Enum ERECParagraphNumberingConstants
   ercParaNone = 0
   ercParaBullet = PFN_BULLET
   ercParaArabicNumbers_NS = 2
   ercParaLowerCaseLetters_NS = 3
   ercParaUpperCaseLetters_NS = 4
   ercParaLowerCaseRoman_NS = 5
   ercParaUpperCaseRoman_NS = 6
   ercParaCustomNumber_NS = 7
End Enum

Public Enum ERECParagraphAlignmentConstants
   ercParaLeft = PFA_LEFT
   ercParaCentre = PFA_CENTER
   ercParaRight = PFA_RIGHT
   ercParaJustify = PFA_JUSTIFY
End Enum

Public Enum ERECTabAlignmentConstants
   ercTabOrdinary = 0
   ercTabCentre_NS = 1
   ercTabRight_NS = 2
   ercTabDecimal_NS = 3
   ercTabWordBarTab_NS = 4
End Enum

Public Enum ERECTabLeaderConstants
   ercTabNoLeader = 0
   ercTabDottedLeader_NS = 1
   ercTabDashedLeader_NS = 2
   ercTabUnderlinedLeader_NS = 3
   ercTabThickLineLeader_NS = 4
   ercTabDoubleLineLeader_NS = 5
End Enum

Public Enum ERECParagraphLineSpacingConstants
   ercLineSpacingSingle = 0
   ercLineSpacingOneAndAHalf = 1
   ercLineSpacingDouble = 2
   ercLineSpacingTwips = 3
   ercLineSpacingTwipsAnyMinimum = 4
   ercLineSpacingTwentiethLine = 5
End Enum

Public Enum ERECLinkEventTypeCOnstants
   ercLButtonDblClick = WM_LBUTTONDBLCLK
   ercLButtonDown = WM_LBUTTONDOWN
   ercLButtonUp = WM_LBUTTONUP
   ercMouseMove = WM_MOUSEMOVE
   ercRButtonDblClick = WM_RBUTTONDBLCLK
   ercRButtonDown = WM_RBUTTONDOWN
   ercRBUttonUp = WM_RBUTTONUP
   ercSetCursor = WM_SETCURSOR
End Enum

Public Enum ERECScrollBarConstants
   ercScrollBarsNone = 1
   ercScrollBarsHorizontal
   ercScrollBarsVertical
   ercScrollBarsBoth
End Enum

' ======================================================================
' Internal Control Variables:
' ======================================================================
Private m_hWnd As Long
Private m_hWndParent As Long
Private m_hWndForm  As Long
Private m_bRunTime As Boolean
Private m_bSubClassing As Boolean
Private m_hLib As Long
Private m_eVersion As ERECControlVersion
Private m_eViewMode As ERECViewModes
Private m_bRedraw As Boolean
Private m_sText As String
Private m_bAllowMethod(ERECInbuiltShortcutConstants.[_First] To ERECInbuiltShortcutConstants.[_Last]) As Boolean
Private m_sFileName As String
Private m_eProgressType As ERECProgressTypeConstants
Private m_sLastFindText As String
Private m_eLastFindMode As ERECFindTypeOptions
Private m_bLastFindNext As Boolean
Private m_eCharFormatRange As ERECSetFormatRange
Private m_bBorder As Boolean
Private m_lLeftMargin As Long
Private m_lRightMargin As Long
Private m_lTopMargin As Long
Private m_lBottomMargin As Long
Private m_lLeftMarginPixels As Long
Private m_lRightMarginPixels As Long
Private m_lLimit As Long
Private m_bTrapTab As Boolean
Private m_bAutoURLDetect As Boolean
Private m_bReadOnly As Boolean
Private m_bTextOnly As Boolean
Private m_bTransparent As Boolean
Private m_bSingleLine As Boolean
Private m_bDisableNoScroll As Boolean
Private m_bPassword As Boolean
Private m_sPasswordChar As String
Private m_eScrollBars As ERECScrollBarConstants
Private m_bHideSelection As Boolean
Private m_bEnabled As Boolean

' Over-riding VB UserControl's default IOLEInPlaceActivate:
Private m_IPAOHookStruct As IPAOHookStructRE
' Tiling images
Private m_cTile As cTile

' ======================================================================
' Events:
' ======================================================================
Public Event SelectionChange(ByVal lMin As Long, ByVal lMax As Long, ByVal eSelType As ERECSelectionTypeConstants)
Attribute SelectionChange.VB_Description = "Raised when the current selection changes."
Public Event LinkOver(ByVal iType As ERECLinkEventTypeCOnstants, ByVal lMin As Long, ByVal lMax As Long)
Attribute LinkOver.VB_Description = "Raised when the user moves the mouse over a hyperlink."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Raised when the user depresses a key on the control."
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Raised when the user depresses a character key on the control and the key has been converted into an Ascii code."
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Raised when the user releases a key on the control."
Public Event DblClick(X As Single, Y As Single)
Attribute DblClick.VB_Description = "Raised when the control is double clicked."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ProgressStatus(ByVal lAmount As Long, ByVal lTotal As Long)
Public Event ModifyProtected(ByRef bDoIt As Boolean, ByVal lMin As Long, ByVal lMax As Long)
Attribute ModifyProtected.VB_Description = "Raised when the user attempts to modify text marked as protected.  Set bDoIt to True to accept the modification."
Public Event VScroll()
Attribute VScroll.VB_Description = "Raised when the control is scrolled vertically."
Public Event HScroll()
Attribute HScroll.VB_Description = "Raised when the control is scrolled horizontally."
Public Event Change()

' ======================================================================
' Subclassing:
' ======================================================================
Implements ISubclass
Private m_emr As EMsgResponse


Public Property Get ScrollBars() As ERECScrollBarConstants
    '<EhHeader>
    On Error GoTo ScrollBars_Err
    '</EhHeader>
   ScrollBars = m_eScrollBars
    '<EhFooter>
    Exit Property

ScrollBars_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ScrollBars " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ScrollBars(ByVal eBars As ERECScrollBarConstants)
    '<EhHeader>
    On Error GoTo ScrollBars_Err
    '</EhHeader>
   If m_hWnd <> 0 Then
      Select Case eBars
      Case ercScrollBarsNone
         If pSetStyle(WS_HSCROLL Or WS_VSCROLL, False) Then
            m_eScrollBars = eBars
         End If
      Case ercScrollBarsVertical
         If pSetStyle(WS_HSCROLL, False) Then
            If pSetStyle(WS_VSCROLL, True) Then
               m_eScrollBars = eBars
            End If
         End If
      Case ercScrollBarsHorizontal
         If pSetStyle(WS_HSCROLL, True) Then
            If pSetStyle(WS_VSCROLL, False) Then
               m_eScrollBars = eBars
            End If
         End If
      Case ercScrollBarsBoth
         If pSetStyle(WS_HSCROLL Or WS_VSCROLL, True) Then
            m_eScrollBars = eBars
         End If
      End Select
   Else
      m_eScrollBars = eBars
   End If
    '<EhFooter>
    Exit Property

ScrollBars_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ScrollBars " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisableNoScroll() As Boolean
    '<EhHeader>
    On Error GoTo DisableNoScroll_Err
    '</EhHeader>
   DisableNoScroll = m_bDisableNoScroll
    '<EhFooter>
    Exit Property

DisableNoScroll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.DisableNoScroll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let DisableNoScroll(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo DisableNoScroll_Err
    '</EhHeader>
   If m_hWnd <> 0 Then
      If pSetStyle(ES_DISABLENOSCROLL, m_bDisableNoScroll) Then
         m_bDisableNoScroll = bState
      End If
   Else
      m_bDisableNoScroll = bState
   End If
   PropertyChanged "DisableNoScroll"
    '<EhFooter>
    Exit Property

DisableNoScroll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.DisableNoScroll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get HideSelection() As Boolean
    '<EhHeader>
    On Error GoTo HideSelection_Err
    '</EhHeader>
   HideSelection = m_bHideSelection
    '<EhFooter>
    Exit Property

HideSelection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.HideSelection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let HideSelection(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo HideSelection_Err
    '</EhHeader>
   If m_hWnd <> 0 Then
      If pSetStyle(ES_NOHIDESEL, Not (bState)) Then
         m_bHideSelection = bState
      End If
   Else
      m_bHideSelection = bState
   End If
   PropertyChanged "HideSelection"
    '<EhFooter>
    Exit Property

HideSelection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.HideSelection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get SingleLine() As Boolean
    '<EhHeader>
    On Error GoTo SingleLine_Err
    '</EhHeader>
   SingleLine = m_bSingleLine
    '<EhFooter>
    Exit Property

SingleLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SingleLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let SingleLine(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo SingleLine_Err
    '</EhHeader>
   If m_hWnd <> 0 Then
      If pSetStyle(ES_MULTILINE, Not (bState)) Then
         m_bSingleLine = bState
      End If
   Else
      m_bSingleLine = bState
   End If
   PropertyChanged "SingleLine"
    '<EhFooter>
    Exit Property

SingleLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SingleLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get PasswordChar() As String
    '<EhHeader>
    On Error GoTo PasswordChar_Err
    '</EhHeader>
   PasswordChar = m_sPasswordChar
    '<EhFooter>
    Exit Property

PasswordChar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.PasswordChar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let PasswordChar(ByVal sChar As String)
   ' Validate
    '<EhHeader>
    On Error GoTo PasswordChar_Err
    '</EhHeader>
   If Len(sChar) > 1 Then sChar = Left$(sChar, 1)
   ' set it:
   If Len(sChar) > 0 Then
      If m_hWnd <> 0 Then
         If pSetStyle(ES_PASSWORD, True) Then
            W_SendMessage ByVal m_hWnd, ByVal EM_SETPASSWORDCHAR, ByVal AscW(sChar), ByVal 0
            m_bPassword = True
            m_sPasswordChar = sChar
         End If
      Else
         m_bPassword = True
         m_sPasswordChar = sChar
      End If
   Else
      If m_hWnd <> 0 Then
         m_bPassword = False
         m_sPasswordChar = ""
         W_SendMessage ByVal m_hWnd, ByVal EM_SETPASSWORDCHAR, ByVal 0, ByVal 0
      Else
         m_bPassword = False
         m_sPasswordChar = ""
      End If
   End If
   PropertyChanged "PasswordChar"
    '<EhFooter>
    Exit Property

PasswordChar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.PasswordChar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Private Function pSetStyle(ByVal lStyle As Long, ByVal bState As Boolean) As Boolean
    '<EhHeader>
    On Error GoTo pSetStyle_Err
    '</EhHeader>
Dim lS As Long
   
   ' Get current style:
   lS = GetWindowLong(m_hWnd, GWL_STYLE)
   ' Apply the flag:
   If bState Then
      lS = lS Or lStyle
   Else
      lS = lS And Not lStyle
   End If
   ' Set the style:
   SetWindowLong m_hWnd, GWL_STYLE, lS
   ' Force window to notice style change:
   pStyleChanged
   
   ' Success?
   pSetStyle = (GetWindowLong(m_hWnd, GWL_STYLE) = lS)
   
    '<EhFooter>
    Exit Function

pSetStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pSetStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Property Get Enabled() As Boolean
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
   Enabled = m_bEnabled
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Enabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Enabled(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
   m_bEnabled = bState
   UserControl.Enabled = bState
   If Not m_hWnd = 0 Then
      EnableWindow m_hWnd, Abs(bState)
   End If
   PropertyChanged "Enabled"
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Enabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set Picture(ByRef sPic As IPicture)
Attribute Picture.VB_Description = "Gets/sets the background picture tiled behind the control when Transparent is set to True."
    '<EhHeader>
    On Error GoTo Picture_Err
    '</EhHeader>
   If m_hWnd = 0 Then
      Set UserControl.Picture = sPic
   Else
      m_cTile.Picture = sPic
   End If
   PropertyChanged "Picture"
    '<EhFooter>
    Exit Property

Picture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Picture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Picture(ByRef sPic As IPicture)
    '<EhHeader>
    On Error GoTo Picture_Err
    '</EhHeader>
   If m_hWnd = 0 Then
      Set UserControl.Picture = sPic
   Else
      m_cTile.Picture = sPic
   End If
   PropertyChanged "Picture"
    '<EhFooter>
    Exit Property

Picture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Picture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Picture() As IPicture
    '<EhHeader>
    On Error GoTo Picture_Err
    '</EhHeader>
   Set Picture = UserControl.Picture
    '<EhFooter>
    Exit Property

Picture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Picture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TRANSPARENT() As Boolean
Attribute TRANSPARENT.VB_Description = "Gets/sets whether the control is transparent and displays the Picture or not."
    '<EhHeader>
    On Error GoTo TRANSPARENT_Err
    '</EhHeader>
   TRANSPARENT = m_bTransparent
    '<EhFooter>
    Exit Property

TRANSPARENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TRANSPARENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let TRANSPARENT(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo TRANSPARENT_Err
    '</EhHeader>
Dim lS As Long
   m_bTransparent = bState
   If m_hWnd <> 0 Then
      lS = GetWindowLong(m_hWnd, GWL_EXSTYLE)
      If bState Then
         lS = lS Or WS_EX_TRANSPARENT
      Else
         lS = lS And Not WS_EX_TRANSPARENT
      End If
      SetWindowLong m_hWnd, GWL_EXSTYLE, lS
      pStyleChanged
   End If
   PropertyChanged "Transparent"
    '<EhFooter>
    Exit Property

TRANSPARENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TRANSPARENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Private Sub pStyleChanged(Optional ByVal hwnd As Long = 0)
    '<EhHeader>
    On Error GoTo pStyleChanged_Err
    '</EhHeader>
   If hwnd = 0 Then hwnd = m_hWnd
   SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
    '<EhFooter>
    Exit Sub

pStyleChanged_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pStyleChanged " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Friend Function TranslateAccelerator(lpMsg As msg) As Long
    '<EhHeader>
    On Error GoTo TranslateAccelerator_Err
    '</EhHeader>
    
   TranslateAccelerator = S_FALSE
   If m_hWnd <> 0 Then
      ' Here you can modify the response to the key down
      ' accelerator command using the values in lpMsg.  This
      ' can be used to capture Tabs, Returns, Arrows etc.
      ' Just process the message as required and return S_OK.
      If lpMsg.message = WM_KEYDOWN Or lpMsg.message = WM_CHAR Or lpMsg.message = WM_KEYUP Then
         Select Case lpMsg.wParam And &HFFFF&
         Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyReturn
            W_SendMessage ByVal m_hWnd, ByVal lpMsg.message, ByVal lpMsg.wParam, ByVal lpMsg.lParam
            TranslateAccelerator = S_OK
         Case vbKeyTab
            If Not ReadOnly Then
               If m_bTrapTab Then
                  ' Allow shift-tab to move out of control:
                  If GetAsyncKeyState(vbKeyShift) = 0 Then
                     ' Default handling of tab:
                     If lpMsg.message = WM_KEYDOWN Then
                        W_SendMessage ByVal m_hWnd, ByVal WM_CHAR, ByVal lpMsg.wParam, ByVal lpMsg.lParam
                     End If
                     TranslateAccelerator = S_OK
                  End If
               End If
            End If
         End Select
      End If
   End If
   
    '<EhFooter>
    Exit Function

TranslateAccelerator_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TranslateAccelerator " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Property Get TrapTab() As Boolean
Attribute TrapTab.VB_Description = "Gets/sets whether the control traps the tab key or not."
    '<EhHeader>
    On Error GoTo TrapTab_Err
    '</EhHeader>
   TrapTab = m_bTrapTab
    '<EhFooter>
    Exit Property

TrapTab_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TrapTab " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let TrapTab(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo TrapTab_Err
    '</EhHeader>
   m_bTrapTab = bState
   PropertyChanged "TrapTab"
    '<EhFooter>
    Exit Property

TrapTab_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TrapTab " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TextLimit() As Long
Attribute TextLimit.VB_Description = "Same as MaxLength (!)"
    '<EhHeader>
    On Error GoTo TextLimit_Err
    '</EhHeader>
   TextLimit = m_lLimit
    '<EhFooter>
    Exit Property

TextLimit_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TextLimit " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let TextLimit(ByVal lLimit As Long)
    '<EhHeader>
    On Error GoTo TextLimit_Err
    '</EhHeader>
Dim lR As Long
   m_lLimit = lLimit
   If m_hWnd <> 0 Then
      lR = W_SendMessage(ByVal m_hWnd, ByVal EM_EXLIMITTEXT, ByVal 0, ByVal lLimit)
   End If
   PropertyChanged "TextLimit"
    '<EhFooter>
    Exit Property

TextLimit_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TextLimit " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Gets/sets the maximum length of text or RTF loaded into the control."
    '<EhHeader>
    On Error GoTo MaxLength_Err
    '</EhHeader>
   If (m_hWnd <> 0) Then
      MaxLength = W_SendMessage(m_hWnd, EM_GETLIMITTEXT, 0, 0)
   End If
    '<EhFooter>
    Exit Property

MaxLength_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.MaxLength " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let MaxLength(ByVal lMax As Long)
    '<EhHeader>
    On Error GoTo MaxLength_Err
    '</EhHeader>
   If (m_hWnd <> 0) Then
      W_SendMessage m_hWnd, EM_EXLIMITTEXT, 0, lMax
   End If
    '<EhFooter>
    Exit Property

MaxLength_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.MaxLength " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Border() As Boolean
Attribute Border.VB_Description = "Gets/sets whether the control has a 3D border."
    '<EhHeader>
    On Error GoTo Border_Err
    '</EhHeader>
   Border = m_bBorder
    '<EhFooter>
    Exit Property

Border_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Border " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Border(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo Border_Err
    '</EhHeader>
Dim dwStyle As Long
Dim dwExStyle As Long

   m_bBorder = bState
   If m_hWnd <> 0 Then
      ' Make sure that the RichEdit never has a border:
      dwStyle = GetWindowLong(m_hWnd, GWL_STYLE)
      dwExStyle = GetWindowLong(m_hWnd, GWL_EXSTYLE)
      dwStyle = dwStyle And Not ES_SUNKEN
      dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
      SetWindowLong m_hWnd, GWL_STYLE, dwStyle
      SetWindowLong m_hWnd, GWL_EXSTYLE, dwExStyle
      pStyleChanged
   End If
   UserControl.BorderStyle() = Abs(bState)
   
    '<EhFooter>
    Exit Property

Border_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Border " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ProgressType() As ERECProgressTypeConstants
    '<EhHeader>
    On Error GoTo ProgressType_Err
    '</EhHeader>
   ProgressType = m_eProgressType
    '<EhFooter>
    Exit Property

ProgressType_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ProgressType " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Function FindText( _
      ByVal sText As String, _
      Optional ByVal eOptions As ERECFindTypeOptions = FR_DEFAULT, _
      Optional ByVal bFindNext As Boolean = True, _
      Optional ByVal bFIndInSelection As Boolean = False, _
      Optional ByRef lMin As Long, _
      Optional ByRef lMax As Long _
   ) As Long
Attribute FindText.VB_Description = "Finds the specified text in the control."
    '<EhHeader>
    On Error GoTo FindText_Err
    '</EhHeader>
Dim tEx1 As FINDTEXTEX_A
'Dim tEx2 As FINDTEXTEX_W
Dim tCR As CHARRANGE
Dim lR As Long
Dim lJunk As Long
Dim B() As Byte

   m_sLastFindText = sText
   m_eLastFindMode = eOptions
   m_bLastFindNext = bFindNext
   
   lMin = -1: lMax = -1
   If (bFIndInSelection) Then
      GetSelection tCR.cpMax, tCR.cpMax
   Else
      If (bFindNext) Then
         GetSelection tCR.cpMin, lJunk
         If (lJunk >= tCR.cpMin) Then
            tCR.cpMin = lJunk + 1
         End If
         tCR.cpMax = -1
      Else
         tCR.cpMin = 0
         tCR.cpMax = -1
      End If
   End If
   
   B = StrConv(sText, vbFromUnicode)
   ' VB won't do the terminating null for you!
   ReDim Preserve B(0 To UBound(B) + 1) As Byte
   B(UBound(B)) = 0
   tEx1.lpstrText = VarPtr(B(0))
   LSet tEx1.chrg = tCR
   
   lR = W_SendMessageAnyRef(m_hWnd, EM_FINDTEXTEX, eOptions, tEx1)
   
   LSet tCR = tEx1.chrgText
   If (lR <> -1) Then
      lMax = tCR.cpMax
      lMin = lMax - Len(sText)
   End If
   FindText = lR
   
    '<EhFooter>
    Exit Function

FindText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FindText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Public Property Get LastFindText() As String
    '<EhHeader>
    On Error GoTo LastFindText_Err
    '</EhHeader>
   LastFindText = m_sLastFindText
    '<EhFooter>
    Exit Property

LastFindText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.LastFindText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get LastFindMode() As ERECFindTypeOptions
    '<EhHeader>
    On Error GoTo LastFindMode_Err
    '</EhHeader>
   LastFindMode = m_eLastFindMode
    '<EhFooter>
    Exit Property

LastFindMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.LastFindMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get LastFindNext() As Boolean
    '<EhHeader>
    On Error GoTo LastFindNext_Err
    '</EhHeader>
   LastFindNext = m_bLastFindNext
    '<EhFooter>
    Exit Property

LastFindNext_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.LastFindNext " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Gets/sets the font of the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
   If (m_eCharFormatRange = ercSetFormatAll) Or (m_hWnd = 0) Then
      Set Font = UserControl.Font
   Else
      Dim sFnt As New StdFont
      Set Font = GetFont(True)
   End If
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Set Font(ByRef sFnt As StdFont)
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
   With UserControl.Font
      .Name = sFnt.Name
      .Size = sFnt.Size
      .bold = sFnt.bold
      .italic = sFnt.italic
      .underline = sFnt.underline
      .Strikethrough = sFnt.Strikethrough
      .charSet = sFnt.charSet
   End With
   If (m_hWnd <> 0) Then
      SetFont sFnt, , , , m_eCharFormatRange
   End If
   PropertyChanged "Font"
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the background colour of the control."
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
   BackColor = UserControl.BackColor
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo BackColor_Err
    '</EhHeader>
   UserControl.BackColor = oColor
   lblText.BackColor = oColor
   If (m_hWnd <> 0) Then
      W_SendMessage m_hWnd, EM_SETBKGNDCOLOR, 0, TranslateColor(oColor)
   End If
   PropertyChanged "BackColor"
    '<EhFooter>
    Exit Property

BackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.BackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gets/sets the forecolour of the control."
    '<EhHeader>
    On Error GoTo ForeColor_Err
    '</EhHeader>
   ForeColor = UserControl.ForeColor
    '<EhFooter>
    Exit Property

ForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ForeColor_Err
    '</EhHeader>
   UserControl.ForeColor = oColor
   If (m_hWnd <> 0) Then
      SetFont UserControl.Font, TranslateColor(oColor), , , ercSetFormatAll
   End If
   PropertyChanged "ForeColor"
    '<EhFooter>
    Exit Property

ForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Gets the text contained in the control."
    '<EhHeader>
    On Error GoTo Text_Err
    '</EhHeader>
   If (m_hWnd = 0) Then
      Text = m_sText
      If (m_sText = "") Then
         'blText.Caption = "vbAccelerator Rich Edit Control"
      Else
         lblText.Caption = m_sText
      End If
   Else
      Text = Contents(SF_TEXT)
   End If
    '<EhFooter>
    Exit Property

Text_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Text " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Text(ByRef sText As String)
    '<EhHeader>
    On Error GoTo Text_Err
    '</EhHeader>
   If (m_hWnd = 0) Then
      m_sText = sText
   Else
      Contents(SF_TEXT) = sText
   End If
    '<EhFooter>
    Exit Property

Text_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Text " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Modified() As Boolean
Attribute Modified.VB_Description = "Gets/sets whether the contents of the control have been modified."
    '<EhHeader>
    On Error GoTo Modified_Err
    '</EhHeader>
   If (m_hWnd <> 0) Then
      Modified = (W_SendMessage(m_hWnd, EM_GETMODIFY, 0, 0) <> 0)
   End If
    '<EhFooter>
    Exit Property

Modified_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Modified " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Modified(ByVal bModified As Boolean)
    '<EhHeader>
    On Error GoTo Modified_Err
    '</EhHeader>
   If (m_hWnd <> 0) Then
      W_SendMessage m_hWnd, EM_SETMODIFY, Abs(bModified), 0
   End If
    '<EhFooter>
    Exit Property

Modified_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Modified " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TextOnly() As Boolean
Attribute TextOnly.VB_Description = "Gets/sets whether the control acts as a text-only control or not."
    '<EhHeader>
    On Error GoTo TextOnly_Err
    '</EhHeader>
   If m_eVersion = eRICHED20 Then
      TextOnly = m_bTextOnly
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

TextOnly_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TextOnly " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let TextOnly(ByVal bTextOnly As Boolean)
    '<EhHeader>
    On Error GoTo TextOnly_Err
    '</EhHeader>
Dim lStyle As Long
   If m_eVersion = eRICHED20 Then
      m_bTextOnly = bTextOnly
      If m_hWnd <> 0 Then
         If m_bTextOnly Then
            lStyle = TM_PLAINTEXT Or TM_MULTILEVELUNDO Or TM_MULTICODEPAGE
         Else
            lStyle = TM_RICHTEXT Or TM_MULTILEVELUNDO Or TM_MULTICODEPAGE
         End If
         W_SendMessage m_hWnd, EM_SETTEXTMODE, lStyle, 0
      End If
      PropertyChanged "TextOnly"
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

TextOnly_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TextOnly " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get RichEditOption( _
      ByVal eOption As ERECOptionTypeConstants _
   ) As Boolean
Attribute RichEditOption.VB_Description = "Gets/sets various options affecting the operation of the RichEdit control."
    '<EhHeader>
    On Error GoTo RichEditOption_Err
    '</EhHeader>
Dim lR As Long
   lR = W_SendMessage(m_hWnd, EM_GETOPTIONS, 0, 0)
   RichEditOption = ((lR And eOption) = eOption)
    '<EhFooter>
    Exit Property

RichEditOption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.RichEditOption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let RichEditOption( _
      ByVal eOption As ERECOptionTypeConstants, _
      ByVal bState As Boolean _
   )
    '<EhHeader>
    On Error GoTo RichEditOption_Err
    '</EhHeader>
Dim lOptions As Long
Dim lR As Long
   lOptions = W_SendMessage(m_hWnd, EM_GETOPTIONS, 0, 0)
   If (bState) Then
      lOptions = lOptions Or eOption
   Else
      lOptions = lOptions And Not eOption
   End If
   lR = W_SendMessage(m_hWnd, EM_SETOPTIONS, 0, lOptions)
    '<EhFooter>
    Exit Property

RichEditOption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.RichEditOption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get LineForCharacterIndex(ByVal lIndex As Long) As Long
Attribute LineForCharacterIndex.VB_Description = "Same as LineForCharacter (!)"
    '<EhHeader>
    On Error GoTo LineForCharacterIndex_Err
    '</EhHeader>
   LineForCharacterIndex = (W_SendMessage(m_hWnd, EM_EXLINEFROMCHAR, 0, lIndex))
    '<EhFooter>
    Exit Property

LineForCharacterIndex_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.LineForCharacterIndex " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Function Unsupported(Optional ByVal iType As Integer = 0)
    '<EhHeader>
    On Error GoTo Unsupported_Err
    '</EhHeader>
   If (iType = 0) Then
      'Debug.Assert "Function not supported in eRICHED32 mode, use RICHED20" = ""
   ElseIf (iType = 1) Then
      Debug.Assert "Property is read-only at run-time" = ""
   End If
    '<EhFooter>
    Exit Function

Unsupported_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Unsupported " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Property Get SelectedText() As String
Attribute SelectedText.VB_Description = "Gets the selected text from the control."
    '<EhHeader>
    On Error GoTo SelectedText_Err
    '</EhHeader>
Dim sBuff As String
Dim lStart As Long
Dim lEnd As Long
Dim lR As Long

   GetSelection lStart, lEnd
   If (lEnd > lStart) Then
      sBuff = String$(lEnd - lStart + 1, 0)
      lR = W_SendMessageStr(m_hWnd, EM_GETSELTEXT, 0, sBuff)
      If (lR > 0) Then
         SelectedText = Left$(sBuff, lR)
      End If
   End If
    '<EhFooter>
    Exit Property

SelectedText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SelectedText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get SelectedContents(ByVal eType As ERECFileTypes) As String
Attribute SelectedContents.VB_Description = "Gets the text or RichText in the current selection in the control."
    '<EhHeader>
    On Error GoTo SelectedContents_Err
    '</EhHeader>
Dim tStream As EDITSTREAM
        
   m_eProgressType = ercSave
        
   tStream.dwCookie = m_hWnd
   tStream.pfnCallback = plAddressOf(AddressOf SaveCallBack)
   tStream.dwError = 0
   ' The text will be streamed out though the SaveCallback function:
   ClearStreamText
   RichEdit = Me
   W_SendMessageAnyRef m_hWnd, EM_STREAMOUT, eType Or SF_UNICODE, tStream
   ClearRichEdit
   
   SelectedContents = StreamText()
    
   m_eProgressType = ercNone
    
    '<EhFooter>
    Exit Property

SelectedContents_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SelectedContents " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TextInRange(ByVal lStart As Long, ByVal lEnd As Long)
Attribute TextInRange.VB_Description = "Gets the text in a specified range without changing the selection."
    '<EhHeader>
    On Error GoTo TextInRange_Err
    '</EhHeader>
Dim tR As TEXTRANGE
Dim lR As Long
Dim sText As String
Dim B() As Byte
      
   tR.chrg.cpMin = lStart
   tR.chrg.cpMax = lEnd
   
   sText = String$(lEnd - lStart + 1, 0)
   B = StrConv(sText, vbFromUnicode)
   ' VB won't do the terminating null for you!
   ReDim Preserve B(0 To UBound(B) + 1) As Byte
   B(UBound(B)) = 0
   tR.lpstrText = VarPtr(B(0))

   lR = W_SendMessageAnyRef(m_hWnd, EM_GETTEXTRANGE, 0, tR)
   If (lR > 0) Then
      sText = StrConv(B, vbUnicode)
      TextInRange = Left$(sText, lR)
   End If
    '<EhFooter>
    Exit Property

TextInRange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.TextInRange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let AutoURLDetect(ByVal bState As Boolean)
Attribute AutoURLDetect.VB_Description = "Gets/sets whether the control will automatically detect hyperlinks prefixed by certain URL identifiers (e.g. http:)"
    '<EhHeader>
    On Error GoTo AutoURLDetect_Err
    '</EhHeader>
Dim lR As Long
   If (m_eVersion = eRICHED20) Then
      m_bAutoURLDetect = bState
      If m_hWnd <> 0 Then
         lR = W_SendMessage(m_hWnd, EM_AUTOURLDETECT, Abs(bState), 0)
         Debug.Assert (lR = 0)
      End If
      PropertyChanged m_bAutoURLDetect
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

AutoURLDetect_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.AutoURLDetect " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get AutoURLDetect() As Boolean
    '<EhHeader>
    On Error GoTo AutoURLDetect_Err
    '</EhHeader>
   AutoURLDetect = m_bAutoURLDetect
    '<EhFooter>
    Exit Property

AutoURLDetect_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.AutoURLDetect " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ReadOnly(ByVal bState As Boolean)
Attribute ReadOnly.VB_Description = "Gets/sets whether the control is read-only."
    '<EhHeader>
    On Error GoTo ReadOnly_Err
    '</EhHeader>
   m_bReadOnly = bState
   If m_hWnd <> 0 Then
      W_SendMessage m_hWnd, EM_SETREADONLY, Abs(bState), 0
   End If
   PropertyChanged "ReadOnly"
    '<EhFooter>
    Exit Property

ReadOnly_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ReadOnly " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ReadOnly() As Boolean
    '<EhHeader>
    On Error GoTo ReadOnly_Err
    '</EhHeader>
Dim lStyle As Long
   If (m_hWnd <> 0) Then
      lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
      If (lStyle And ES_READONLY) = ES_READONLY Then
         ReadOnly = True
      End If
   End If
    '<EhFooter>
    Exit Property

ReadOnly_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ReadOnly " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get LineCount() As Long
Attribute LineCount.VB_Description = "Returns the number of lines in the control."
    '<EhHeader>
    On Error GoTo LineCount_Err
    '</EhHeader>
   LineCount = W_SendMessage(m_hWnd, EM_GETLINECOUNT, 0, 0)
    '<EhFooter>
    Exit Property

LineCount_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.LineCount " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get FirstVisibleLine() As Long
Attribute FirstVisibleLine.VB_Description = "Gets the 0 based index of the first visible line within the control."
    '<EhHeader>
    On Error GoTo FirstVisibleLine_Err
    '</EhHeader>
   FirstVisibleLine = W_SendMessage(m_hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
    '<EhFooter>
    Exit Property

FirstVisibleLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FirstVisibleLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get CurrentLine() As Long
Attribute CurrentLine.VB_Description = "Gets the 0 based index of the line containing the cursor."
    '<EhHeader>
    On Error GoTo CurrentLine_Err
    '</EhHeader>
Dim lStart As Long, lEnd As Long
   GetSelection lStart, lEnd
   ' Use EX to ensure we can cope with > 32k text
   CurrentLine = W_SendMessage(m_hWnd, EM_EXLINEFROMCHAR, 0, lStart)
    '<EhFooter>
    Exit Property

CurrentLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CurrentLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get LineForCharacter(ByVal lCharacter As Long)
Attribute LineForCharacter.VB_Description = "Gets the line containing the specified 0 based character index."
   ' Use EX to ensure we can cope with > 32k text
    '<EhHeader>
    On Error GoTo LineForCharacter_Err
    '</EhHeader>
   LineForCharacter = W_SendMessage(m_hWnd, EM_EXLINEFROMCHAR, 0, lCharacter)
    '<EhFooter>
    Exit Property

LineForCharacter_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.LineForCharacter " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get CharFromPos(ByVal xPixels As Long, ByVal yPixels As Long)
Attribute CharFromPos.VB_Description = "Gets the 0 based index of the character at the specified position in pixels."
    '<EhHeader>
    On Error GoTo CharFromPos_Err
    '</EhHeader>
Dim tP As POINTAPI
   tP.X = xPixels
   tP.Y = yPixels
   CharFromPos = W_SendMessageAnyRef(m_hWnd, EM_CHARFROMPOS, 0, tP)
    '<EhFooter>
    Exit Property

CharFromPos_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CharFromPos " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Sub GetPosFromChar(ByVal lIndex As Long, ByRef xPixels As Long, ByRef yPixels As Long)
Attribute GetPosFromChar.VB_Description = "Returns the position in pixels for a given 0 based character index."
    '<EhHeader>
    On Error GoTo GetPosFromChar_Err
    '</EhHeader>
Dim lXY As Long
   lXY = W_SendMessage(m_hWnd, EM_POSFROMCHAR, lIndex, 0)
   xPixels = (lXY And &HFFFF&)
   yPixels = (lXY \ &H10000) And &HFFFF&
    '<EhFooter>
    Exit Sub

GetPosFromChar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetPosFromChar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub GetSelection(ByRef lStart As Long, ByRef lEnd As Long)
Attribute GetSelection.VB_Description = "Gets the start and end of the current position."
    '<EhHeader>
    On Error GoTo GetSelection_Err
    '</EhHeader>
Dim tCR As CHARRANGE
   W_SendMessageAnyRef m_hWnd, EM_EXGETSEL, 0, tCR
   lStart = tCR.cpMin
   lEnd = tCR.cpMax
    '<EhFooter>
    Exit Sub

GetSelection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetSelection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Property Get SelLength() As Long

    Dim t1 As Long, t2 As Long
    GetSelection t1, t2
    SelLength = t2 - t1
    
End Property

Public Property Let SelLength(value As Long)
Dim t1 As Long
    
    GetSelection t1, 0
    SetSelection t1, t1 + value
    
End Property


Public Property Get SelStart() As Long

    Dim t1 As Long
    GetSelection t1, 0
    SelStart = t1
    
End Property

Public Property Let SelStart(value As Long)
Dim t2 As Long
    
    GetSelection 0, t2
    SetSelection value, t2
    
End Property

Public Property Get SelEnd() As Long

    Dim t1 As Long
    GetSelection t1, 0
    SelStart = t1
    
End Property

Public Property Let SelEnd(value As Long)
Dim t1 As Long
    
    GetSelection t1, 0
    SetSelection t1, value
    
End Property


Public Sub SetSelection(ByVal lStart As Long, ByVal lEnd As Long)
Attribute SetSelection.VB_Description = "Sets the current selection."
    '<EhHeader>
    On Error GoTo SetSelection_Err
    '</EhHeader>
Dim tCR As CHARRANGE
   tCR.cpMin = lStart
   tCR.cpMax = lEnd
   W_SendMessageAnyRef m_hWnd, EM_EXSETSEL, 0, tCR
    '<EhFooter>
    Exit Sub

SetSelection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetSelection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub SelectAll()
Attribute SelectAll.VB_Description = "Selects the contents of the control."
    '<EhHeader>
    On Error GoTo SelectAll_Err
    '</EhHeader>
   SetSelection 0, -1
    '<EhFooter>
    Exit Sub

SelectAll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SelectAll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub SelectNone()
Attribute SelectNone.VB_Description = "Clears any selection in the control."
    '<EhHeader>
    On Error GoTo SelectNone_Err
    '</EhHeader>
Dim tc As CHARRANGE
    tc.cpMax = 0
    tc.cpMin = 0
    W_SendMessageAnyRef m_hWnd, EM_EXSETSEL, 0, tc
    '<EhFooter>
    Exit Sub

SelectNone_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SelectNone " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Property Get CanPaste() As Boolean
Attribute CanPaste.VB_Description = "Returns whether Pasting is possible."
    '<EhHeader>
    On Error GoTo CanPaste_Err
    '</EhHeader>
   CanPaste = W_SendMessage(m_hWnd, EM_CANPASTE, 0, 0)
    '<EhFooter>
    Exit Property

CanPaste_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CanPaste " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get CanCopy() As Boolean
Attribute CanCopy.VB_Description = "Returns whether the copying is possible."
    '<EhHeader>
    On Error GoTo CanCopy_Err
    '</EhHeader>
Dim lStart As Long, lEnd As Long
   GetSelection lStart, lEnd
   If (lEnd > lStart) Then
      CanCopy = True
   End If
    '<EhFooter>
    Exit Property

CanCopy_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CanCopy " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get CanUndo() As Boolean
Attribute CanUndo.VB_Description = "Returns whether an Undo operation is possible."
    '<EhHeader>
    On Error GoTo CanUndo_Err
    '</EhHeader>
   CanUndo = W_SendMessage(m_hWnd, EM_CANUNDO, 0, 0)
    '<EhFooter>
    Exit Property

CanUndo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CanUndo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get CanRedo() As Boolean
Attribute CanRedo.VB_Description = "Returns whether a Redo operation is possible."
    '<EhHeader>
    On Error GoTo CanRedo_Err
    '</EhHeader>
   If m_eVersion = eRICHED20 Then
      CanRedo = W_SendMessage(m_hWnd, EM_CANREDO, 0, 0)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

CanRedo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CanRedo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get UndoType() As ERECUndoTypeConstants
Attribute UndoType.VB_Description = "Gets the type of action which will be undone."
    '<EhHeader>
    On Error GoTo UndoType_Err
    '</EhHeader>
   If m_eVersion = eRICHED20 Then
      UndoType = W_SendMessage(m_hWnd, EM_GETUNDONAME, 0, 0)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

UndoType_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UndoType " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get RedoType() As ERECUndoTypeConstants
Attribute RedoType.VB_Description = "Gets the type of action which will be redone."
    '<EhHeader>
    On Error GoTo RedoType_Err
    '</EhHeader>
   If m_eVersion = eRICHED20 Then
      RedoType = W_SendMessage(m_hWnd, EM_GETREDONAME, 0, 0)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

RedoType_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.RedoType " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Sub Cut()
Attribute Cut.VB_Description = "Performs the control's copy operation.  Check CanCut to see if it is possible to cut."
    '<EhHeader>
    On Error GoTo Cut_Err
    '</EhHeader>
   W_SendMessage m_hWnd, WM_CUT, 0, 0
    '<EhFooter>
    Exit Sub

Cut_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Cut " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub Copy()
Attribute Copy.VB_Description = "Performs the control's copy operation.  Check CanCopy to see if it is possible to copy."
    '<EhHeader>
    On Error GoTo Copy_Err
    '</EhHeader>
   W_SendMessage m_hWnd, WM_COPY, 0, 0
    '<EhFooter>
    Exit Sub

Copy_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Copy " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub Paste()
Attribute Paste.VB_Description = "Performs the control's Paste action.  Use CanPaste to determine if the Paste action can be performed."
    '<EhHeader>
    On Error GoTo Paste_Err
    '</EhHeader>
   W_SendMessage m_hWnd, WM_PASTE, 0, 0
    '<EhFooter>
    Exit Sub

Paste_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Paste " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub PasteSpecial()
Attribute PasteSpecial.VB_Description = "<For future development>"
    '<EhHeader>
    On Error GoTo PasteSpecial_Err
    '</EhHeader>
   W_SendMessage m_hWnd, EM_PASTESPECIAL, 0, 0
    '<EhFooter>
    Exit Sub

PasteSpecial_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.PasteSpecial " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub Undo()
Attribute Undo.VB_Description = "Performs the control's Undo action.  Check the CanUndo property to see if the Undo action can be performed."
    '<EhHeader>
    On Error GoTo Undo_Err
    '</EhHeader>
   W_SendMessage m_hWnd, EM_UNDO, 0, 0
    '<EhFooter>
    Exit Sub

Undo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Undo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub Redo()
Attribute Redo.VB_Description = "Performs the control's Redo action.  Check the CanRedo property to see if this action is available."
    '<EhHeader>
    On Error GoTo Redo_Err
    '</EhHeader>
   If (m_eVersion = eRICHED20) Then
      W_SendMessage m_hWnd, EM_REDO, 0, 0
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Sub

Redo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Redo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub Delete()
Attribute Delete.VB_Description = "Performs the control's delete operation.  Check CanCut to see if it is possible to delete."
   ' TODO
    '<EhHeader>
    On Error GoTo Delete_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

Delete_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Delete " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub InsertContents(ByVal eType As ERECFileTypes, ByRef sText As String)
Attribute InsertContents.VB_Description = "Inserts a text or RTF string into the control."
    '<EhHeader>
    On Error GoTo InsertContents_Err
    '</EhHeader>
Dim tStream As EDITSTREAM
Dim lR As Long
   ' Don't redraw:
   Redraw = False
   ' Insert the text:
   tStream.dwCookie = m_hWnd
   tStream.pfnCallback = plAddressOf(AddressOf LoadCallBack)
   tStream.dwError = 0
   StreamText = sText
   ' The text will be streamed in though the LoadCallback function:
   lR = W_SendMessageAnyRef(m_hWnd, EM_STREAMIN, eType Or SFF_SELECTION Or SF_UNICODE, tStream)
   ' Redraw again:
   Redraw = True
   
    '<EhFooter>
    Exit Sub

InsertContents_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.InsertContents " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Property Get ViewMode() As ERECViewModes
Attribute ViewMode.VB_Description = "Gets/sets who the control lays out the text on screen."
    '<EhHeader>
    On Error GoTo ViewMode_Err
    '</EhHeader>
   ViewMode = m_eViewMode
    '<EhFooter>
    Exit Property

ViewMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ViewMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ViewMode(ByVal eViewMode As ERECViewModes)
    '<EhHeader>
    On Error GoTo ViewMode_Err
    '</EhHeader>
   If (eViewMode <> m_eViewMode) Then
      m_eViewMode = eViewMode
      pSetViewMode eViewMode
   End If
    '<EhFooter>
    Exit Property

ViewMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ViewMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Private Sub pSetViewMode(ByVal eViewMode As ERECViewModes)
    '<EhHeader>
    On Error GoTo pSetViewMode_Err
    '</EhHeader>
   Select Case m_eViewMode
   Case ercWYSIWYG
      On Error Resume Next
      W_SendMessage m_hWnd, EM_SETTARGETDEVICE, Printer.hdc, Printer.Width
   Case ercWordWrap
      W_SendMessage m_hWnd, EM_SETTARGETDEVICE, 0, 0
   Case ercDefault
      W_SendMessage m_hWnd, EM_SETTARGETDEVICE, 0, 1
   End Select
    '<EhFooter>
    Exit Sub

pSetViewMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pSetViewMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Get CharFormatRange() As ERECSetFormatRange
Attribute CharFormatRange.VB_Description = "Gets/sets the range to which font formatting will apply."
    '<EhHeader>
    On Error GoTo CharFormatRange_Err
    '</EhHeader>
   CharFormatRange = m_eCharFormatRange
    '<EhFooter>
    Exit Property

CharFormatRange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CharFormatRange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let CharFormatRange(ByVal eRange As ERECSetFormatRange)
    '<EhHeader>
    On Error GoTo CharFormatRange_Err
    '</EhHeader>
   m_eCharFormatRange = eRange
    '<EhFooter>
    Exit Property

CharFormatRange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CharFormatRange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get CharacterCount() As Long
Attribute CharacterCount.VB_Description = "Returns the number of characters of text in the control."
    '<EhHeader>
    On Error GoTo CharacterCount_Err
    '</EhHeader>
   If m_eVersion = eRICHED20 Then
      CharacterCount = W_SendMessage(m_hWnd, WM_GETTEXTLENGTH, 0, 0)
   Else
      CharacterCount = W_SendMessage(m_hWnd, EM_GETTEXTLENGTHEX, 0, 0)
   End If
    '<EhFooter>
    Exit Property

CharacterCount_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.CharacterCount " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Gets/sets whether the font is bold for the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontBold_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_BOLD
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
   FontBold = ((tCF.dwEffects And CFE_BOLD) = CFE_BOLD)
    '<EhFooter>
    Exit Property

FontBold_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontBold " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontBold(ByVal bBold As Boolean)
    '<EhHeader>
    On Error GoTo FontBold_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_BOLD
   If (bBold) Then
      tCF.dwEffects = CFE_BOLD
   End If
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    '<EhFooter>
    Exit Property

FontBold_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontBold " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Gets/sets whether the font is italic for the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontItalic_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_ITALIC
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
   FontItalic = ((tCF.dwEffects And CFE_ITALIC) = CFE_ITALIC)
    '<EhFooter>
    Exit Property

FontItalic_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontItalic " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontItalic(ByVal bItalic As Boolean)
    '<EhHeader>
    On Error GoTo FontItalic_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_ITALIC
   If (bItalic) Then
      tCF.dwEffects = CFE_ITALIC
   End If
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    '<EhFooter>
    Exit Property

FontItalic_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontItalic " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Gets/sets whether the font is underlined for the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontUnderline_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_UNDERLINE
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
   FontUnderline = ((tCF.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
    '<EhFooter>
    Exit Property

FontUnderline_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontUnderline " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontUnderline(ByVal bUnderline As Boolean)
    '<EhHeader>
    On Error GoTo FontUnderline_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_UNDERLINE
   If (bUnderline) Then
      tCF.dwEffects = CFE_UNDERLINE
   End If
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    '<EhFooter>
    Exit Property

FontUnderline_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontUnderline " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontStrikeOut() As Boolean
Attribute FontStrikeOut.VB_Description = "Gets/sets whether the font is struck out for the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontStrikeOut_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_STRIKEOUT
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
   FontStrikeOut = ((tCF.dwEffects And CFE_STRIKEOUT) = CFE_STRIKEOUT)
    '<EhFooter>
    Exit Property

FontStrikeOut_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontStrikeOut " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontStrikeOut(ByVal bStrikeOut As Boolean)
    '<EhHeader>
    On Error GoTo FontStrikeOut_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
   tCF.dwMask = CFM_STRIKEOUT
   If (bStrikeOut) Then
      tCF.dwEffects = CFE_STRIKEOUT
   End If
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    '<EhFooter>
    Exit Property

FontStrikeOut_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontStrikeOut " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get FontColour() As OLE_COLOR
Attribute FontColour.VB_Description = "Gets/sets the colour of the font for the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontColour_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
Dim lColour As Long
   tCF.dwMask = CFM_COLOR
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
   FontColour = tCF.crTextColor
    '<EhFooter>
    Exit Property

FontColour_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontColour " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontColour(ByVal oColour As OLE_COLOR)
    '<EhHeader>
    On Error GoTo FontColour_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long
Dim lColour As Long
   If oColour = -1 Then
      tCF.dwMask = CFM_COLOR
      tCF.dwEffects = CFE_AUTOCOLOR
      tCF.crTextColor = -1
   Else
      tCF.crTextColor = TranslateColor(oColour)
      tCF.dwMask = CFM_COLOR
   End If
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    '<EhFooter>
    Exit Property

FontColour_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontColour " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontBackColour() As OLE_COLOR
Attribute FontBackColour.VB_Description = "Gets/sets the background colour of the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontBackColour_Err
    '</EhHeader>
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED20) Then
      tCF2.dwMask = CFM_BACKCOLOR
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
      FontBackColour = tCF2.crBackColor
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

FontBackColour_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontBackColour " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontBackColour(ByVal oColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo FontBackColour_Err
    '</EhHeader>
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED20) Then
      If oColor = -1 Then
         tCF2.dwMask = CFM_BACKCOLOR
         tCF2.dwEffects = CFE_AUTOBACKCOLOR
         tCF2.crBackColor = -1
      Else
         tCF2.dwMask = CFM_BACKCOLOR
         tCF2.crBackColor = TranslateColor(oColor)
      End If
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

FontBackColour_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontBackColour " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontLink() As Boolean
Attribute FontLink.VB_Description = "Gets/sets whether the selection acts as a hyperlink.  Set CharFormatRange to selection."
    '<EhHeader>
    On Error GoTo FontLink_Err
    '</EhHeader>
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED20) Then
      tCF2.dwMask = CFM_LINK
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
      FontLink = ((tCF2.dwEffects And CFE_LINK) = CFE_LINK)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

FontLink_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontLink " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontLink(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo FontLink_Err
    '</EhHeader>
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED20) Then
      tCF2.dwMask = CFM_LINK
      If (bState) Then
         tCF2.dwEffects = CFE_LINK
      End If
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

FontLink_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontLink " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontProtected() As Boolean
Attribute FontProtected.VB_Description = "Gets/sets whether the selection is protected (raises the ModifyRequest event).  Set CharFormatRange to selection."
    '<EhHeader>
    On Error GoTo FontProtected_Err
    '</EhHeader>
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED20) Then
      tCF2.dwMask = CFM_PROTECTED
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
      FontProtected = ((tCF2.dwEffects And CFE_PROTECTED) = CFE_PROTECTED)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

FontProtected_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontProtected " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontProtected(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo FontProtected_Err
    '</EhHeader>
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED20) Then
      tCF2.dwMask = CFM_PROTECTED
      If (bState) Then
         tCF2.dwEffects = CFE_PROTECTED
      End If
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
   Else
      Unsupported
   End If
    '<EhFooter>
    Exit Property

FontProtected_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontProtected " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get FontSuperScript() As Boolean
Attribute FontSuperScript.VB_Description = "Gets/sets whether the font is superscripted for the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontSuperScript_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED32) Then
      tCF.dwMask = CFM_OFFSET
      tCF.cbSize = Len(tCF)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
      FontSuperScript = (tCF.yOffset > 0)
   Else
      tCF2.dwMask = CFM_SUPERSCRIPT
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
      FontSuperScript = ((tCF2.dwEffects And CFE_SUPERSCRIPT) = CFE_SUPERSCRIPT)
   End If
    '<EhFooter>
    Exit Property

FontSuperScript_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontSuperScript " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get FontSubScript() As Boolean
Attribute FontSubScript.VB_Description = "Gets/sets whether the font is subscripted for the control or selection, depending on the setting of CharFormatRange."
    '<EhHeader>
    On Error GoTo FontSubScript_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim tCF2 As CHARFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED32) Then
      tCF.dwMask = CFM_OFFSET
      tCF.cbSize = Len(tCF)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
   Else
      tCF2.dwMask = CFM_SUBSCRIPT
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
      FontSuperScript = ((tCF2.dwEffects And CFE_SUBSCRIPT) = CFE_SUBSCRIPT)
   End If
    '<EhFooter>
    Exit Property

FontSubScript_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontSubScript " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontSuperScript(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo FontSuperScript_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim tCF2 As CHARFORMAT2
Dim lR As Long
Dim Y As Long

   If (m_eVersion = eRICHED32) Then
      ' Get the current font size in twips:
      tCF.dwMask = CFM_SIZE
      tCF.cbSize = Len(tCF)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, ercSetFormatSelection, tCF)
      Y = tCF.yHeight \ 2
      
      ' Set the offset:
      tCF.dwMask = CFM_OFFSET
      tCF.cbSize = Len(tCF)
      If (bState) Then
         tCF.yOffset = Y
      Else
         tCF.yOffset = 0
      End If
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
   Else
      tCF2.dwMask = CFM_SUPERSCRIPT
      If (bState) Then
         tCF2.dwEffects = CFE_SUPERSCRIPT
      End If
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
   End If
    '<EhFooter>
    Exit Property

FontSuperScript_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontSuperScript " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let FontSubScript(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo FontSubScript_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim tCF2 As CHARFORMAT2
Dim lR As Long
Dim Y As Long

   If (m_eVersion = eRICHED32) Then
      ' Get the current font size in twips:
      tCF.dwMask = CFM_SIZE
      tCF.cbSize = Len(tCF)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, ercSetFormatSelection, tCF)
      Y = tCF.yHeight \ -2
      
      ' Set the offset:
      tCF.dwMask = CFM_OFFSET
      tCF.cbSize = Len(tCF)
      If (bState) Then
         tCF.yOffset = Y
      Else
         tCF.yOffset = 0
      End If
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
   Else
      tCF2.dwMask = CFM_SUBSCRIPT
      If (bState) Then
         tCF2.dwEffects = CFE_SUBSCRIPT
      End If
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
   End If
    '<EhFooter>
    Exit Property

FontSubScript_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.FontSubScript " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Sub SetFont( _
      ByRef fntThis As StdFont, _
      Optional ByVal oColor As OLE_COLOR = vbWindowText, _
      Optional ByVal eType As ERECTextTypes = ercTextNormal, _
      Optional ByVal bHyperLink As Boolean = False, _
      Optional ByVal eRange As ERECSetFormatRange = ercSetFormatSelection _
   )
    '<EhHeader>
    On Error GoTo SetFont_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim tCF2 As CHARFORMAT2
Dim dwEffects As Long
Dim dwMask As Long
Dim i As Long
   
   tCF.cbSize = Len(tCF)
   tCF.crTextColor = TranslateColor(oColor)
   dwMask = CFM_COLOR
   If fntThis.bold Then
      dwEffects = dwEffects Or CFE_BOLD
   End If
   dwMask = dwMask Or CFM_BOLD
   If fntThis.italic Then
      dwEffects = dwEffects Or CFE_ITALIC
   End If
   dwMask = dwMask Or CFM_ITALIC
   If fntThis.Strikethrough Then
      dwEffects = dwEffects Or CFE_STRIKEOUT
   End If
   dwMask = dwMask Or CFM_STRIKEOUT
   If fntThis.underline Then
      dwEffects = dwEffects Or CFE_UNDERLINE
   End If
   dwMask = dwMask Or CFM_UNDERLINE

   If bHyperLink Then
      dwEffects = dwEffects Or CFE_LINK
   End If
   dwMask = dwMask Or CFM_LINK
   
   tCF.dwEffects = dwEffects
   tCF.dwMask = dwMask Or CFM_FACE Or CFM_SIZE
   
   For i = 1 To Len(fntThis.Name)
      tCF.szFaceName(i - 1) = Asc(Mid$(fntThis.Name, i, 1))
   Next i
   tCF.yHeight = (fntThis.Size * 20)
   If (eType = ercTextSubscript) Then
      tCF.yOffset = -tCF.yHeight \ 2
   End If
   If (eType = ercTextSuperscript) Then
      tCF.yOffset = tCF.yHeight \ 2
   End If
   
   If (m_eVersion = eRICHED32) Then
      W_SendMessageAnyRef m_hWnd, EM_SETCHARFORMAT, eRange, tCF
   Else
      CopyMemory tCF2, tCF, Len(tCF)
      tCF2.cbSize = Len(tCF2)
      tCF.yOffset = 0
      If (eType = ercTextSubscript) Then
         tCF.dwEffects = tCF.dwEffects Or CFE_SUBSCRIPT
         tCF.dwMask = tCF.dwMask Or CFM_SUBSCRIPT
      End If
      If (eType = ercTextSuperscript) Then
         tCF.dwEffects = tCF.dwEffects Or CFE_SUPERSCRIPT
         tCF.dwMask = tCF.dwMask Or CFM_SUPERSCRIPT
      End If
      W_SendMessageAnyRef m_hWnd, EM_SETCHARFORMAT, eRange, tCF2
   End If
   
    '<EhFooter>
    Exit Sub

SetFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Function GetFont( _
      Optional ByVal bForSelection As Boolean = False, _
      Optional ByRef oColor As OLE_COLOR, _
      Optional ByRef bHyperLink As Boolean, _
      Optional ByVal eType As ERECTextTypes = ercTextNormal _
   ) As StdFont
    '<EhHeader>
    On Error GoTo GetFont_Err
    '</EhHeader>
Dim sFnt As New StdFont
Dim tCF As CHARFORMAT
Dim tCF2 As CHARFORMAT2
Dim dwEffects As Long
Dim dwMask As Long
Dim i As Long
Dim sName As String
   
   tCF.cbSize = Len(tCF)
   dwMask = dwMask Or CFM_COLOR

   dwMask = dwMask Or CFM_BOLD
   dwMask = dwMask Or CFM_ITALIC
   dwMask = dwMask Or CFM_STRIKEOUT
   dwMask = dwMask Or CFM_UNDERLINE
   dwMask = dwMask Or CFM_LINK
   If (m_eVersion = eRICHED32) Then
      tCF.dwEffects = dwEffects
      tCF.dwMask = dwMask Or CFM_FACE Or CFM_SIZE
      W_SendMessageAnyRef m_hWnd, EM_GETCHARFORMAT, Abs(bForSelection), tCF
   Else
      CopyMemory tCF2, tCF, Len(tCF)
      tCF2.cbSize = Len(tCF2)
      W_SendMessageAnyRef m_hWnd, EM_GETCHARFORMAT, Abs(bForSelection), tCF2
   End If
      
   If (m_eVersion = eRICHED32) Then
      'tCF.crTextColor = TranslateColor(oColor)
      oColor = tCF.crTextColor
      For i = 1 To LF_FACESIZE
         sName = sName & Chr$(tCF.szFaceName(i - 1))
      Next i
      sFnt.Name = sName
      sFnt.Size = tCF.yHeight \ 20
      sFnt.bold = ((tCF.dwEffects And CFE_BOLD) = CFE_BOLD)
      sFnt.italic = ((tCF.dwEffects And CFE_ITALIC) = CFE_ITALIC)
      sFnt.underline = ((tCF.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
      sFnt.Strikethrough = ((tCF.dwEffects And CFE_STRIKEOUT) = CFE_STRIKEOUT)
      bHyperLink = ((tCF.dwEffects And CFE_LINK) = CFE_LINK)
      If (tCF.yOffset = 0) Then
         eType = ercTextNormal
      ElseIf (tCF.yOffset < 0) Then
         eType = ercTextSubscript
      Else
         eType = ercTextSuperscript
      End If
   Else
      oColor = tCF2.crTextColor
      For i = 1 To LF_FACESIZE
         sName = sName & Chr$(tCF2.szFaceName(i - 1))
      Next i
      sFnt.Size = tCF2.yHeight \ 20
      sFnt.bold = ((tCF2.dwEffects And CFE_BOLD) = CFE_BOLD)
      sFnt.italic = ((tCF2.dwEffects And CFE_ITALIC) = CFE_ITALIC)
      sFnt.underline = ((tCF2.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
      sFnt.Strikethrough = ((tCF2.dwEffects And CFE_STRIKEOUT) = CFE_STRIKEOUT)
      bHyperLink = ((tCF2.dwEffects And CFE_LINK) = CFE_LINK)
      eType = ercTextNormal
      If ((tCF2.dwEffects And CFE_SUPERSCRIPT) = CFE_SUPERSCRIPT) Then
         eType = ercTextSuperscript
      End If
      If ((tCF2.dwEffects And CFE_SUBSCRIPT) = CFE_SUBSCRIPT) Then
         eType = ercTextSubscript
      End If
      sFnt.Name = sName
   End If
   Set GetFont = sFnt
    '<EhFooter>
    Exit Function

GetFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Property Get ParagraphNumbering() As ERECParagraphNumberingConstants
Attribute ParagraphNumbering.VB_Description = "Gets/sets whether the selected paragraph has bullets or not."
    '<EhHeader>
    On Error GoTo ParagraphNumbering_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long

   If (m_eVersion = eRICHED32) Then
      tP.dwMask = PFM_NUMBERING
      tP.cbSize = Len(tP)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP)
      ParagraphNumbering = tP.wNumbering
   Else
      tP2.dwMask = PFM_NUMBERING
      tP2.cbSize = Len(tP2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP2)
      ParagraphNumbering = tP2.wNumbering
   End If
    '<EhFooter>
    Exit Property

ParagraphNumbering_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ParagraphNumbering " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ParagraphNumbering(ByVal eStyle As ERECParagraphNumberingConstants)
    '<EhHeader>
    On Error GoTo ParagraphNumbering_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long

   If (m_eVersion = eRICHED32) Then
      tP.dwMask = PFM_NUMBERING
      tP.cbSize = Len(tP)
      tP.wNumbering = eStyle
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP)
   Else
      tP2.dwMask = PFM_NUMBERING
      tP2.wNumbering = eStyle
      tP2.cbSize = Len(tP2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP2)
   End If
    '<EhFooter>
    Exit Property

ParagraphNumbering_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ParagraphNumbering " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Sub GetParagraphOffsets( _
      ByRef lStartIndent As Long, _
      ByRef lLeftOffset As Long, _
      ByRef lRightOffset As Long _
   )
Attribute GetParagraphOffsets.VB_Description = "Gets the paragraph offsets (left, right and initial line)."
    '<EhHeader>
    On Error GoTo GetParagraphOffsets_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long

   If (m_eVersion = eRICHED32) Then
      tP.dwMask = PFM_STARTINDENT Or PFM_RIGHTINDENT Or PFM_OFFSET
      tP.cbSize = Len(tP)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP)
      lStartIndent = tP.dxStartIndent
      lLeftOffset = tP.dxOffset
      lRightOffset = tP.dxRightIndent
   Else
      tP2.dwMask = PFM_STARTINDENT Or PFM_RIGHTINDENT Or PFM_OFFSET
      tP2.cbSize = Len(tP2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP2)
      lStartIndent = tP2.dxStartIndent
      lLeftOffset = tP2.dxOffset
      lRightOffset = tP2.dxRightIndent
   End If
    '<EhFooter>
    Exit Sub

GetParagraphOffsets_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetParagraphOffsets " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub SetParagraphOffsets( _
      ByVal lStartIndent As Long, _
      ByVal lLeftOffset As Long, _
      ByVal lRightOffset As Long _
   )
Attribute SetParagraphOffsets.VB_Description = "Sets the offsets (left, right and initial line) for the current paragraph."
    '<EhHeader>
    On Error GoTo SetParagraphOffsets_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long

   If (m_eVersion = eRICHED32) Then
      tP.dwMask = PFM_STARTINDENT Or PFM_RIGHTINDENT Or PFM_OFFSET
      tP.dxStartIndent = lStartIndent
      tP.dxOffset = lLeftOffset
      tP.dxRightIndent = lRightOffset
      tP.cbSize = Len(tP)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP)
   Else
      tP2.dwMask = PFM_STARTINDENT Or PFM_RIGHTINDENT Or PFM_OFFSET
      tP2.dxStartIndent = lStartIndent
      tP2.dxOffset = lLeftOffset
      tP2.dxRightIndent = lRightOffset
      tP2.cbSize = Len(tP2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP2)
   End If
      
    '<EhFooter>
    Exit Sub

SetParagraphOffsets_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetParagraphOffsets " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Get ParagraphAlignment() As ERECParagraphAlignmentConstants
Attribute ParagraphAlignment.VB_Description = "Gets/Sets the alignment of the selected paragraph."
    '<EhHeader>
    On Error GoTo ParagraphAlignment_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long

   If (m_eVersion = eRICHED32) Then
      tP.dwMask = PFM_ALIGNMENT
      tP.cbSize = Len(tP)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP)
      ParagraphAlignment = tP.wAlignment
   Else
      tP2.dwMask = PFM_ALIGNMENT
      tP2.cbSize = Len(tP2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP2)
      ParagraphAlignment = tP2.wAlignment
   End If

    '<EhFooter>
    Exit Property

ParagraphAlignment_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ParagraphAlignment " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ParagraphAlignment(ByVal eAlign As ERECParagraphAlignmentConstants)
    '<EhHeader>
    On Error GoTo ParagraphAlignment_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long

   If (m_eVersion = eRICHED32) Then
      If (eAlign = ercParaJustify) Then
         Unsupported
      Else
         tP.dwMask = PFM_ALIGNMENT
         tP.cbSize = Len(tP)
         tP.wAlignment = eAlign
         lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP)
      End If
   Else
      tP2.dwMask = PFM_ALIGNMENT
      tP2.cbSize = Len(tP2)
      tP2.wAlignment = eAlign
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP2)
   End If

    '<EhFooter>
    Exit Property

ParagraphAlignment_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ParagraphAlignment " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Sub GetParagraphTabs( _
      ByRef iCount As Integer, _
      ByRef lTabSize() As Long, _
      Optional ByRef eTabAlignment As Variant, _
      Optional ByRef eTabLeader As Variant _
   )
Attribute GetParagraphTabs.VB_Description = "Gets the tab stops for the current paragraph."
    '<EhHeader>
    On Error GoTo GetParagraphTabs_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long
Dim lNumTabs As Long
Dim lPtrTabs As Long
Dim lTabs() As Long
Dim i As Long
Dim lAlign() As Long
Dim lLeader() As Long


   Erase lTabSize
   eTabAlignment = 0
   eTabLeader = 0
   iCount = 0
   
   If (m_eVersion = eRICHED32) Then
      tP.dwMask = PFM_TABSTOPS
      tP.cbSize = Len(tP)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP)
      lNumTabs = tP.cTabCount
      If (lNumTabs > 0) Then
         iCount = tP.cTabCount
         ReDim lTabSize(1 To lNumTabs) As Long
         For i = 0 To lNumTabs - 1
            lTabSize(i + 1) = tP.lTabStops(i)
         Next i
      End If
   Else
      tP2.dwMask = PFM_TABSTOPS
      tP2.cbSize = Len(tP2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tP2)
      lNumTabs = tP2.cTabCount
      If (lNumTabs > 0) Then
         iCount = tP2.cTabCount
         ReDim lTabSize(1 To lNumTabs) As Long
         ReDim lAlign(1 To lNumTabs) As Long
         ReDim lLeader(1 To lNumTabs) As Long
         For i = 0 To lNumTabs - 1
            ' First 24 bits are size:
            lTabSize(i + 1) = (tP2.lTabStops(i) And &HFFFFFF)
            ' Bits 24-27 are alignment:
            lAlign(i + 1) = (tP2.lTabStops(i) And &HF000000) \ &H1000000
            ' Bits 28-31 are leader:
            lLeader(i + 1) = (tP2.lTabStops(i) And &H70000000) \ &H10000000
         Next i
         eTabAlignment = lAlign
         eTabLeader = lLeader
      End If
   End If
        
    '<EhFooter>
    Exit Sub

GetParagraphTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetParagraphTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub SetParagraphTabs( _
      ByVal iCount As Integer, _
      ByRef lTabSize() As Long, _
      Optional ByRef eTabAlignment As Variant, _
      Optional ByRef eTabLeader As Variant _
   )
Attribute SetParagraphTabs.VB_Description = "Sets tab stops for the current paragraph."
    '<EhHeader>
    On Error GoTo SetParagraphTabs_Err
    '</EhHeader>
Dim tP As PARAFORMAT
Dim tP2 As PARAFORMAT2
Dim lR As Long
Dim lNumTabs As Long
Dim lPtrTabs As Long
Dim i As Long
   
   
   If (m_eVersion = eRICHED32) Then
      tP.dwMask = PFM_TABSTOPS
      tP.cbSize = Len(tP)
      tP.cTabCount = iCount
      If (iCount > 0) Then
         For i = 0 To iCount - 1
            tP.lTabStops(i) = lTabSize(i + 1)
         Next i
      End If
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP)
   Else
      tP2.dwMask = PFM_TABSTOPS
      tP2.cbSize = Len(tP2)
      tP2.cTabCount = iCount
      If (iCount > 0) Then
         For i = 0 To iCount - 1
            tP2.lTabStops(i) = lTabSize(i + 1)
         Next i
      End If
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tP2)
   End If
   
    '<EhFooter>
    Exit Sub

SetParagraphTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetParagraphTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub GetParagraphLineSpacing( _
      ByRef eLineSpacingStyle As ERECParagraphLineSpacingConstants, _
      ByRef ySpacing As Long _
   )
Attribute GetParagraphLineSpacing.VB_Description = "Gets the line spacing for the current paragraph."
    '<EhHeader>
    On Error GoTo GetParagraphLineSpacing_Err
    '</EhHeader>
Dim tCF2 As PARAFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED32) Then
      Unsupported
   Else
      tCF2.dwMask = PFM_LINESPACING
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tCF2)
      eLineSpacingStyle = tCF2.bLineSpacingRule
      ySpacing = tCF2.dyLineSpacing
   End If
    '<EhFooter>
    Exit Sub

GetParagraphLineSpacing_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetParagraphLineSpacing " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub SetParagraphLineSpacing( _
      ByVal eLineSpacingStyle As ERECParagraphLineSpacingConstants, _
      ByVal ySpacing As Long _
   )
Attribute SetParagraphLineSpacing.VB_Description = "Sets the line spacing for the current paragraph."
    '<EhHeader>
    On Error GoTo SetParagraphLineSpacing_Err
    '</EhHeader>
Dim tCF2 As PARAFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED32) Then
      Unsupported
   Else
      tCF2.dwMask = PFM_LINESPACING
      tCF2.cbSize = Len(tCF2)
      tCF2.bLineSpacingRule = eLineSpacingStyle
      tCF2.dyLineSpacing = ySpacing
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tCF2)
   End If
    '<EhFooter>
    Exit Sub

SetParagraphLineSpacing_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetParagraphLineSpacing " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub GetParagraphSpacing( _
      ByRef lSpaceAfter As Long, _
      ByRef lSpaceBefore As Long _
   )
Attribute GetParagraphSpacing.VB_Description = "Gets the spacing between paragraphs for the current paragraph."
    '<EhHeader>
    On Error GoTo GetParagraphSpacing_Err
    '</EhHeader>
Dim tCF2 As PARAFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED32) Then
      Unsupported
   Else
      tCF2.dwMask = PFM_SPACEBEFORE Or PFM_SPACEAFTER
      tCF2.cbSize = Len(tCF2)
      lR = W_SendMessageAnyRef(m_hWnd, EM_GETPARAFORMAT, 0, tCF2)
      lSpaceAfter = tCF2.dySpaceAfter
      lSpaceBefore = tCF2.dySpaceBefore
   End If
    '<EhFooter>
    Exit Sub

GetParagraphSpacing_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetParagraphSpacing " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub SetParagraphSpacing( _
      ByVal lSpaceAfter As Long, _
      ByVal lSpaceBefore As Long _
   )
Attribute SetParagraphSpacing.VB_Description = "Sets the spacing between paragraphs for the current paragraph."
    '<EhHeader>
    On Error GoTo SetParagraphSpacing_Err
    '</EhHeader>
Dim tCF2 As PARAFORMAT2
Dim lR As Long
   If (m_eVersion = eRICHED32) Then
      Unsupported
   Else
      tCF2.dwMask = PFM_SPACEBEFORE Or PFM_SPACEAFTER
      tCF2.cbSize = Len(tCF2)
      tCF2.dySpaceAfter = lSpaceAfter
      tCF2.dySpaceBefore = lSpaceBefore
      lR = W_SendMessageAnyRef(m_hWnd, EM_SETPARAFORMAT, 0, tCF2)
   End If
   
    '<EhFooter>
    Exit Sub

SetParagraphSpacing_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetParagraphSpacing " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Property Let UseVersion(ByVal eVersion As ERECControlVersion)
Attribute UseVersion.VB_Description = "Gets/sets which version of the RichEdit DLL to use: version 2/3 (RichEd20.DLL) or version 1 (RichEd32.DLL)"
    '<EhHeader>
    On Error GoTo UseVersion_Err
    '</EhHeader>
    If (UserControl.Ambient.UserMode) Then
        ' can't set at run time in this implementation.
        Unsupported 1
    Else
        m_eVersion = eVersion
    End If
    '<EhFooter>
    Exit Property

UseVersion_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UseVersion " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get UseVersion() As ERECControlVersion
    '<EhHeader>
    On Error GoTo UseVersion_Err
    '</EhHeader>
    UseVersion = m_eVersion
    '<EhFooter>
    Exit Property

UseVersion_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UseVersion " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get IsRtf(ByRef sFileText As String) As Boolean
Attribute IsRtf.VB_Description = "Returns whether the specified string contains RTF."
    '<EhHeader>
    On Error GoTo IsRtf_Err
    '</EhHeader>
   If (Left$(sFileText, 5) = "{\rtf") Then
      IsRtf = True
   End If
    '<EhFooter>
    Exit Property

IsRtf_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.IsRtf " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Gets/sets whether the control will redraw or not."
    '<EhHeader>
    On Error GoTo Redraw_Err
    '</EhHeader>
   Redraw = m_bRedraw
    '<EhFooter>
    Exit Property

Redraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Redraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Redraw(ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo Redraw_Err
    '</EhHeader>
   If (m_bRedraw <> bState) Then
      If (m_hWnd <> 0) Then
         If Not (bState) Then
            ' Don't redraw:
            W_SendMessage m_hWnd, WM_SETREDRAW, 0, 0
         Else
            ' Redraw again:
            W_SendMessage m_hWnd, WM_SETREDRAW, 1, 0
            InvalidateRectAsNull m_hWnd, 0, 1
            UpdateWindow m_hWnd
         End If
      End If
   End If
   m_bRedraw = bState
   PropertyChanged "Redraw"
    '<EhFooter>
    Exit Property

Redraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Redraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let Contents(ByVal eType As ERECFileTypes, ByRef sContents As String)
Attribute Contents.VB_Description = "Gets/sets the control's contents from a string in RichText or Text format."
    '<EhHeader>
    On Error GoTo Contents_Err
    '</EhHeader>
Dim tStream As EDITSTREAM
Dim lR As Long
   
   m_eProgressType = ercLoad
   
   Redraw = False
   ' Load the text:
   tStream.dwCookie = m_hWnd
   tStream.pfnCallback = plAddressOf(AddressOf LoadCallBack)
   tStream.dwError = 0
   StreamText = sContents
   RichEdit = Me
   ' The text will be streamed in though the LoadCallback function:
   lR = W_SendMessageAnyRef(m_hWnd, EM_STREAMIN, eType Or SF_UNICODE, tStream)
   ClearRichEdit
   ' Set unmodified flag
   W_SendMessage m_hWnd, EM_SETMODIFY, 0, 0
   Redraw = True
   
   m_eProgressType = ercNone
   
    '<EhFooter>
    Exit Property

Contents_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Contents " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get Contents(ByVal eType As ERECFileTypes) As String
    '<EhHeader>
    On Error GoTo Contents_Err
    '</EhHeader>
Dim tStream As EDITSTREAM
        
   m_eProgressType = ercSave
        
   tStream.dwCookie = m_hWnd
   tStream.pfnCallback = plAddressOf(AddressOf SaveCallBack)
   tStream.dwError = 0
   ' The text will be streamed out though the SaveCallback function:
   ClearStreamText
   RichEdit = Me
   W_SendMessageAnyRef m_hWnd, EM_STREAMOUT, eType Or SF_UNICODE, tStream
   ClearRichEdit
   
   Contents = StreamText()
    
   m_eProgressType = ercNone
    
    '<EhFooter>
    Exit Property

Contents_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.Contents " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Function LoadFromFile( _
      ByVal sFile As String, _
      ByVal eType As ERECFileTypes _
   ) As Boolean
Attribute LoadFromFile.VB_Description = "Loads a text or RTF file into the control."
    '<EhHeader>
    On Error GoTo LoadFromFile_Err
    '</EhHeader>
Dim hFile As Long
Dim tOF As OFSTRUCT
Dim tStream As EDITSTREAM
Dim lR As Long

   m_eProgressType = ercLoad
   
   Redraw = False

   hFile = OpenFile(sFile, tOF, OF_READ)
   If (hFile <> 0) Then
      tStream.dwCookie = hFile
      tStream.pfnCallback = plAddressOf(AddressOf LoadCallBack)
      tStream.dwError = 0
      
      RichEdit = Me
      FileMode = True
      
      ' The text will be streamed in though the LoadCallback function:
      lR = W_SendMessageAnyRef(m_hWnd, EM_STREAMIN, eType Or SF_UNICODE, tStream)
      
      LoadFromFile = (lR <> 0)
      
      FileMode = False
      ClearRichEdit
      
      CloseHandle hFile
   End If
   Redraw = True
   
   m_eProgressType = ercNone
    '<EhFooter>
    Exit Function

LoadFromFile_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.LoadFromFile " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Public Function SaveToFile( _
      ByVal sFile As String, _
      ByVal eType As ERECFileTypes _
   ) As Boolean
Attribute SaveToFile.VB_Description = "Saves the contents of the control to a text or RichText file."
    '<EhHeader>
    On Error GoTo SaveToFile_Err
    '</EhHeader>
Dim tStream As EDITSTREAM
Dim tOF As OFSTRUCT
Dim hFile As Long
Dim lR As Long
        
   m_eProgressType = ercSave
        
   hFile = OpenFile(sFile, tOF, OF_CREATE)
   If (hFile <> 0) Then
      tStream.dwCookie = hFile
      tStream.pfnCallback = plAddressOf(AddressOf SaveCallBack)
      tStream.dwError = 0
      FileMode = True
      RichEdit = Me
      
      lR = W_SendMessageAnyRef(m_hWnd, EM_STREAMOUT, eType Or SF_UNICODE, tStream)
      
      SaveToFile = (lR <> 0)
      
      FileMode = False
      ClearRichEdit
   
      CloseHandle hFile
   End If
    
   m_eProgressType = ercNone
       
    '<EhFooter>
    Exit Function

SaveToFile_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SaveToFile " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Sub RaiseLoadStatus(ByVal lAmount As Long, ByVal lTotalAmount As Long)
    '<EhHeader>
    On Error GoTo RaiseLoadStatus_Err
    '</EhHeader>
   RaiseEvent ProgressStatus(lAmount, lTotalAmount)
    '<EhFooter>
    Exit Sub

RaiseLoadStatus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.RaiseLoadStatus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub PrintDocDC( _
      ByVal lPrinterHDC As Long, _
      ByVal sDocTitle As String, _
      Optional ByVal nStartPage As Long, _
      Optional ByVal nEndPage As Long _
   )
Attribute PrintDocDC.VB_Description = "Prints the current document to a specified DC."
    '<EhHeader>
    On Error GoTo PrintDocDC_Err
    '</EhHeader>
Dim fr As FORMATRANGE
Dim lTextOut As Long, lTextAmt As Long
Dim lLastTextOut As Long
Dim hJob As Long
Dim lR As Long
Dim lMin As Long
Dim lWidth As Long, lHeight As Long
Dim lLeft As Long, lTop As Long
Dim lXOffset As Long, lYOffset As Long
Dim lPixelsX As Long, lPixelsY As Long
Dim iPage As Long
Dim rcPage As RECT, rcRender As RECT
Dim lSavedState As Long
   
   m_eProgressType = ercPrint
   
   '// Fill out the DOCINFO structure.
   Dim B() As Byte
   Dim di As DOCINFO
   di.cbSize = Len(di)
   di.lpszOutput = 0
   ' This need sorting out.
   If (sDocTitle = "") Then
       sDocTitle = "RTF Document (vbAccelerator RichEdit control)"
   End If
   B = StrConv(sDocTitle, vbFromUnicode)
   ReDim Preserve B(0 To UBound(B) + 1) As Byte
   di.lpszDocName = VarPtr(B(0))
   
   '// Fill out the FORMATRANGE structure for the RTF output.
   fr.hdc = lPrinterHDC '; // HDC
   fr.hdcTarget = fr.hdc
   fr.chrg.cpMin = 0 '; // print
   fr.chrg.cpMax = -1 '; // entire contents
    
   ' Get information about the physically printable page on the
   ' printer:
    
   ' This is the number of Pixels per inch:
   lPixelsX = GetDeviceCaps(lPrinterHDC, LOGPIXELSX)
   lPixelsY = GetDeviceCaps(lPrinterHDC, LOGPIXELSY)
    
   ' This is the number of pixels across:
   lWidth = MulDiv(GetDeviceCaps(lPrinterHDC, PHYSICALWIDTH), 1440, lPixelsX)
   ' This is the number of pixels down:
   lHeight = MulDiv(GetDeviceCaps(lPrinterHDC, PHYSICALHEIGHT), 1440, lPixelsY)
   rcPage.Right = lWidth
   rcPage.Bottom = lHeight
        
   ' Save DC so we can restore it later to the initial state:
   lSavedState = SaveDC(fr.hdc)
   ' Ensure printer DC is in text mode:
   SetMapMode fr.hdc, MM_TEXT
        
   ' Evaluate the left and right physical offsets:
   lXOffset = -GetDeviceCaps(lPrinterHDC, PHYSICALOFFSETX)
   lYOffset = -GetDeviceCaps(lPrinterHDC, PHYSICALOFFSETY)
      
   lLeft = MulDiv(m_lLeftMargin, lPixelsX, 1440)
   lLeft = lLeft + lXOffset
   If lLeft < 0 Then lLeft = 0
   lTop = MulDiv(m_lTopMargin, lPixelsY, 1440)
   lTop = lTop + lYOffset
   If lTop < 0 Then lTop = 0
   rcRender.Right = lWidth - m_lRightMargin - m_lLeftMargin
   rcRender.Bottom = lHeight - m_lBottomMargin - m_lTopMargin
    
   ' Adjust the DC left,top according to the x & y offset:
   SetViewportOrgEx fr.hdc, lLeft, lTop, ByVal 0&
      
   ' Get the text out range:
   lTextOut = 0
   lTextAmt = CharacterCount()

   ' Clear the formatting buffer:
   W_SendMessage m_hWnd, EM_FORMATRANGE, 0, 0
   '
   
   ' Get each of the pages:
   Dim tP() As FORMATRANGE
   Dim lCount As Long
   Dim bSkip As Boolean
   
   If lTextAmt > 0 Then
      fr.chrg.cpMin = 0
      fr.chrg.cpMax = -1
      lCount = 0
      Do
         ' Work out the size of text to render:
         LSet fr.rc = rcRender
         LSet fr.rcPage = rcPage
         lMin = fr.chrg.cpMin
         lTextOut = W_SendMessageAnyRef(m_hWnd, EM_FORMATRANGE, 0, fr)
         fr.chrg.cpMin = lTextOut
         If lCount > 0 Then
            ' This problem doesn't seem to get mentioned anywhere!
            ' If format range returns a smaller value than
            ' the last minimum, it has actually finished:
            If lTextOut < lMin Then
               fr.chrg.cpMin = lTextAmt
               bSkip = True
            End If
         End If
         If Not bSkip Then
            ' We cache the output rectangle and start &
            ' finish positions for subsequent printing:
            lCount = lCount + 1
            ReDim Preserve tP(1 To lCount) As FORMATRANGE
            tP(lCount).chrg.cpMin = lMin
            tP(lCount).chrg.cpMax = lTextOut - 1
            LSet tP(lCount).rc = fr.rc
         End If
      Loop While fr.chrg.cpMin <> -1 And fr.chrg.cpMin < lTextAmt
   End If
   
   RestoreDC fr.hdc, -1
   
   If nStartPage <= 0 Then
      nStartPage = 1
   ElseIf nStartPage > lCount Then
      nStartPage = lCount
   End If
   If nEndPage <= 0 Then
      nEndPage = lCount
   ElseIf nEndPage > lCount Then
      nEndPage = lCount
   End If
               
   RaiseEvent ProgressStatus(-1, -1)
   hJob = StartDoc(lPrinterHDC, di)
   If (hJob <> 0) Then
      
      ' Reset the output buffer:
      W_SendMessageAnyRef m_hWnd, EM_FORMATRANGE, 0, 0
      
      For iPage = nStartPage To nEndPage
      
         'If Not iPage = 1 Then
            StartPage fr.hdc
         'End If
         
         ' Return DC to printing condition:
         lSavedState = SaveDC(fr.hdc)
         SetMapMode fr.hdc, MM_TEXT
         SetViewportOrgEx fr.hdc, lLeft, lTop, ByVal 0&
         
         LSet fr.rc = tP(iPage).rc
         LSet fr.rcPage = rcPage
         LSet fr.chrg = tP(iPage).chrg
         
         fr.chrg.cpMin = W_SendMessageAnyRef(m_hWnd, EM_FORMATRANGE, 1, fr)
         
         RestoreDC fr.hdc, -1
         
         RaiseEvent ProgressStatus(lTextOut, lTextAmt)
         
         EndPage fr.hdc
         
      Next iPage
                        
      RaiseEvent ProgressStatus(lTextAmt, lTextAmt)

      '// Reset the formatting of the rich edit control.
      W_SendMessage m_hWnd, EM_FORMATRANGE, True, 0
    
      EndDoc fr.hdc
      
    Else
        Debug.Print "Failed to start print job"
    End If
   
    '<EhFooter>
    Exit Sub

PrintDocDC_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.PrintDocDC " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Public Sub PrintDoc( _
      ByVal sDocTitle As String _
   )
Attribute PrintDoc.VB_Description = "Prints the document after showing a Print Dialog."
    '<EhHeader>
    On Error GoTo PrintDoc_Err
    '</EhHeader>
Dim pd As PrintDlg

   '// Initialize the PRINTDLG structure.
   pd.lStructSize = Len(pd)
   pd.hWndOwner = m_hWnd
   pd.hDevMode = 0
   pd.hDevNames = 0
   pd.nFromPage = 0
   pd.nToPage = 0
   pd.nMinPage = 0
   pd.nMaxPage = 0
   pd.nCopies = 0
   pd.hInstance = App.hInstance
   pd.flags = PD_RETURNDC Or PD_NOSELECTION Or PD_PRINTSETUP
   pd.lpfnSetupHook = 0
   pd.lpSetupTemplateName = 0
   pd.lpfnPrintHook = 0
   pd.lpPrintTemplateName = 0
   
   '// Get the printer DC.
   If (PrintDlg(pd) <> 0) Then
      
      PrintDocDC pd.hdc, sDocTitle
       '// Delete the printer DC.
       DeleteDC pd.hdc
       
       m_eProgressType = ercNone
   End If

    '<EhFooter>
    Exit Sub

PrintDoc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.PrintDoc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Function plAddressOf(ByVal lAddr As Long) As Long
    ' Why do we have to write nonsense like this?
    '<EhHeader>
    On Error GoTo plAddressOf_Err
    '</EhHeader>
    plAddressOf = lAddr
    '<EhFooter>
    Exit Function

plAddressOf_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.plAddressOf " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Gets the Window handle of the control. If you want the handle of the RichEdit control itself, use RichEdithWnd instead."
    '<EhHeader>
    On Error GoTo hwnd_Err
    '</EhHeader>
   hwnd = UserControl.hwnd
    '<EhFooter>
    Exit Property

hwnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.hwnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get RichEdithWnd() As Long
Attribute RichEdithWnd.VB_Description = "Gets the Window Handle of the RichEdit control."
    '<EhHeader>
    On Error GoTo RichEdithWnd_Err
    '</EhHeader>
   RichEdithWnd = m_hWnd
    '<EhFooter>
    Exit Property

RichEdithWnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.RichEdithWnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Sets focus to the control."
    '<EhHeader>
    On Error GoTo SetFocus_Err
    '</EhHeader>
   SetFocusAPI m_hWnd
    '<EhFooter>
    Exit Sub

SetFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub pInitialise()
    '<EhHeader>
    On Error GoTo pInitialise_Err
    '</EhHeader>
Dim dwStyle As Long
Dim dwExStyle As Long
Dim lS As Long
Dim hP As Long
Dim sLib As String
Dim sClass As String

   pTerminate

   If (UserControl.Ambient.UserMode) Then
      If (m_eVersion = eRICHED20) Then
         sLib = "RICHED20.DLL"
         sClass = RICHEDIT_CLASSW
      Else
         sLib = "RICHED32.DLL"
         sClass = RICHEDIT_CLASS10A
      End If
      m_hLib = LoadLibrary(sLib)
      If m_hLib = 0 And m_eVersion = eRICHED20 Then
         ' Fall back!
         m_eVersion = eRICHED32
         sLib = "RICHED32.DLL"
         sClass = RICHEDIT_CLASS10A
        m_hLib = LoadLibrary(sLib)
      End If
     
      If m_hLib <> 0 Then
         dwStyle = WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS
         dwStyle = dwStyle Or WS_HSCROLL Or WS_VSCROLL
         dwStyle = dwStyle Or WS_TABSTOP
         dwStyle = dwStyle Or ES_MULTILINE Or ES_SAVESEL
         dwStyle = dwStyle Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
         dwStyle = dwStyle Or ES_SELECTIONBAR Or ES_NOHIDESEL
      
         If (m_bBorder) Then
            dwStyle = dwStyle Or ES_SUNKEN
            dwExStyle = WS_EX_CLIENTEDGE
         End If
      
         If (m_bTransparent) Then
            dwExStyle = dwExStyle Or WS_EX_TRANSPARENT
         End If
         
         '// Create the rich edit control.
         Set m_cTile = New cTile
         m_hWndParent = UserControl.hwnd
         m_hWndForm = GetParent(UserControl.hwnd) 'UserControl.parent.hwnd
         m_hWnd = W_CreateWindowEx( _
            dwExStyle, _
            sClass, _
            "", _
            dwStyle, _
            0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, _
            m_hWndParent, _
            0, _
            App.hInstance, _
            0)
         If (m_hWnd <> 0) Then
            EnableWindow m_hWnd, 1
            pAttachMessages
         End If
      End If
   End If
    '<EhFooter>
    Exit Sub

pInitialise_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pInitialise " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Function pTerminate()
    '<EhHeader>
    On Error GoTo pTerminate_Err
    '</EhHeader>
   If (m_hWnd <> 0) Then
      ' Remove printer DC from the
      ViewMode = ercDefault
      ' Stop subclassing:
      pDetachMessages
      ' Destroy the window:
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
      ' store that we haven't a window:
      m_hWnd = 0
      Set m_cTile = Nothing
   End If
   If (m_hLib <> 0) Then
       FreeLibrary m_hLib
       m_hLib = 0
   End If
    '<EhFooter>
    Exit Function

pTerminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pTerminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Sub pAttachMessages()
    '<EhHeader>
    On Error GoTo pAttachMessages_Err
    '</EhHeader>
Dim dwMask As Long
   m_emr = emrPreprocess
   AttachMessage Me, m_hWndForm, WM_ACTIVATE
   AttachMessage Me, m_hWndParent, WM_NOTIFY
   AttachMessage Me, m_hWndParent, WM_SETFOCUS
   AttachMessage Me, m_hWndParent, WM_PAINT
   AttachMessage Me, m_hWndParent, WM_COMMAND
   AttachMessage Me, m_hWnd, WM_ERASEBKGND
   AttachMessage Me, m_hWnd, WM_SETFOCUS
   AttachMessage Me, m_hWnd, WM_MOUSEACTIVATE
   AttachMessage Me, m_hWnd, WM_VSCROLL
   AttachMessage Me, m_hWnd, WM_HSCROLL
    
    ' Key And Mouse Events
    dwMask = ENM_KEYEVENTS Or ENM_MOUSEEVENTS
    ' Selection change
    dwMask = dwMask Or ENM_SELCHANGE
    ' Update
    dwMask = dwMask Or ENM_DROPFILES
    ' Scrolling
    dwMask = dwMask Or ENM_SCROLL
    ' Update:
    dwMask = dwMask Or ENM_UPDATE
    ' Change:
    dwMask = dwMask Or ENM_CHANGE
    
    If (m_eVersion = eRICHED20) Then
      ' Link over messages:
      dwMask = dwMask Or ENM_LINK
      ' Protected messages:
      dwMask = dwMask Or ENM_PROTECTED
    End If
    
    W_SendMessage m_hWnd, EM_SETEVENTMASK, 0, dwMask
    m_bSubClassing = True
    '<EhFooter>
    Exit Sub

pAttachMessages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pAttachMessages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub pDetachMessages()
    '<EhHeader>
    On Error GoTo pDetachMessages_Err
    '</EhHeader>
   If (m_bSubClassing) Then
      DetachMessage Me, m_hWndForm, WM_ACTIVATE
      DetachMessage Me, m_hWndParent, WM_NOTIFY
      DetachMessage Me, m_hWndParent, WM_SETFOCUS
      DetachMessage Me, m_hWndParent, WM_PAINT
      DetachMessage Me, m_hWndParent, WM_COMMAND
      DetachMessage Me, m_hWnd, WM_ERASEBKGND
      DetachMessage Me, m_hWnd, WM_SETFOCUS
      DetachMessage Me, m_hWnd, WM_MOUSEACTIVATE
      DetachMessage Me, m_hWnd, WM_VSCROLL
      DetachMessage Me, m_hWnd, WM_HSCROLL
      m_bSubClassing = False
   End If
    '<EhFooter>
    Exit Sub

pDetachMessages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pDetachMessages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Get AllowShortCut(ByVal eShortCut As ERECInbuiltShortcutConstants) As Boolean
Attribute AllowShortCut.VB_Description = "Gets/sets whether the control will respond automatically to a keyboard accelerator."
    '<EhHeader>
    On Error GoTo AllowShortCut_Err
    '</EhHeader>
   AllowShortCut = m_bAllowMethod(eShortCut)
    '<EhFooter>
    Exit Property

AllowShortCut_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.AllowShortCut " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let AllowShortCut(ByVal eShortCut As ERECInbuiltShortcutConstants, ByVal bState As Boolean)
    '<EhHeader>
    On Error GoTo AllowShortCut_Err
    '</EhHeader>
   m_bAllowMethod(eShortCut) = bState
    '<EhFooter>
    Exit Property

AllowShortCut_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.AllowShortCut " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Sub GetPageMargins( _
      ByRef lLeftMargin As Long, _
      ByRef lTopMargin As Long, _
      ByRef lRightMargin As Long, _
      ByRef lBottomMargin As Long _
   )
Attribute GetPageMargins.VB_Description = "Gets the margins of the page when it is printed, in twips."
    '<EhHeader>
    On Error GoTo GetPageMargins_Err
    '</EhHeader>
   lLeftMargin = m_lLeftMargin
   lTopMargin = m_lTopMargin
   lRightMargin = m_lRightMargin
   lBottomMargin = m_lBottomMargin
    '<EhFooter>
    Exit Sub

GetPageMargins_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.GetPageMargins " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Sub SetPageMargins( _
      Optional ByVal lLeftMargin As Long = 1800, _
      Optional ByVal lTopMargin As Long = 1800, _
      Optional ByVal lRightMargin As Long = 1440, _
      Optional ByVal lBottomMargin As Long = 1440 _
   )
Attribute SetPageMargins.VB_Description = "Sets the margins for the printed page."
    '<EhHeader>
    On Error GoTo SetPageMargins_Err
    '</EhHeader>
   m_lLeftMargin = lLeftMargin
   m_lTopMargin = lTopMargin
   m_lRightMargin = lRightMargin
   m_lBottomMargin = lBottomMargin
   If (m_eViewMode = ercWYSIWYG) Then
      ' Reset the view to account for
      ' left & right margins:
      ViewMode = ercWordWrap
      ViewMode = ercWYSIWYG
   End If
    '<EhFooter>
    Exit Sub

SetPageMargins_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.SetPageMargins " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Public Property Let ControlRightMargin(ByVal lRightMarginPixels As Long)
Attribute ControlRightMargin.VB_Description = "Gets/sets the margin from the right hand edge of the control to the RichEdit control."
    '<EhHeader>
    On Error GoTo ControlRightMargin_Err
    '</EhHeader>
   If (m_hWnd <> 0) Then
      W_SendMessage m_hWnd, EM_SETMARGINS, EC_RIGHTMARGIN, lRightMarginPixels * &H10000
      pSetViewMode m_eViewMode
   End If
   m_lRightMarginPixels = lRightMarginPixels
   PropertyChanged "ControlRightMargin"
    '<EhFooter>
    Exit Property

ControlRightMargin_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ControlRightMargin " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ControlRightMargin() As Long
    '<EhHeader>
    On Error GoTo ControlRightMargin_Err
    '</EhHeader>
   ControlRightMargin = m_lRightMarginPixels
    '<EhFooter>
    Exit Property

ControlRightMargin_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ControlRightMargin " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Let ControlLeftMargin(ByVal lLeftMarginPixels As Long)
Attribute ControlLeftMargin.VB_Description = "Gets/sets the margin from the left hand edge of the control to the RichEdit control."
    '<EhHeader>
    On Error GoTo ControlLeftMargin_Err
    '</EhHeader>
   If (m_hWnd <> 0) Then
      W_SendMessage m_hWnd, EM_SETMARGINS, EC_LEFTMARGIN, lLeftMarginPixels
      pSetViewMode m_eViewMode
   End If
   m_lLeftMarginPixels = lLeftMarginPixels
   PropertyChanged "ControlLeftMargin"
    '<EhFooter>
    Exit Property

ControlLeftMargin_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ControlLeftMargin " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get ControlLeftMargin() As Long
    '<EhHeader>
    On Error GoTo ControlLeftMargin_Err
    '</EhHeader>
   ControlLeftMargin = m_lLeftMarginPixels
    '<EhFooter>
    Exit Property

ControlLeftMargin_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ControlLeftMargin " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Function pDoDefault( _
      ByRef iKeyCode As Integer, _
      ByRef iShift As Integer, _
      ByRef bDefault As Boolean _
   )
    '<EhHeader>
    On Error GoTo pDoDefault_Err
    '</EhHeader>
Dim tCF As CHARFORMAT

   ' Debug.Print iKeyCode
   If (iShift And vbCtrlMask) = vbCtrlMask Then
      Select Case iKeyCode
      
      ' Inbuilt methods:
      Case vbKeyC
         If Not (AllowShortCut(ercCopy_CtrlC)) Then
            bDefault = False
         End If
      Case vbKeyV
         If Not (AllowShortCut(ercPaste_CtrlV)) Then
            bDefault = False
         End If
      Case vbKeyX
         If Not (AllowShortCut(ercCut_CtrlX)) Then
            bDefault = False
         End If
      Case vbKeyA
         If Not (AllowShortCut(ercSelectAll_CtrlA)) Then
            bDefault = False
         End If
      Case vbKeyZ
         If Not (AllowShortCut(ercUndo_CtrlZ)) Then
            bDefault = False
         End If
      
      ' Supplied methods:
      Case vbKeyY
         If AllowShortCut(ercRedo_CtrlY) Then
            Redo
            bDefault = False
         End If
      Case vbKeyB
         If AllowShortCut(ercBold_CtrlB) Then
            pInvertFontOption CFM_BOLD, CFE_BOLD
            bDefault = False
         End If
      Case vbKeyI
         If AllowShortCut(ercItalic_CtrlI) Then
            pInvertFontOption CFM_ITALIC, CFE_ITALIC
            bDefault = False
         End If
      Case vbKeyU
         If AllowShortCut(ercUnderline_CtrlU) Then
            pInvertFontOption CFM_UNDERLINE, CFE_UNDERLINE
            bDefault = False
         End If
      Case vbKeyAdd, 187
         If AllowShortCut(ercSubscript_CtrlMinus) Then
            ' Debug.Print "Add"
            pInvertSubScriptOption 1
            bDefault = False
         End If
      Case vbKeySubtract, 189
         If AllowShortCut(ercSuperscript_CtrlPlus) Then
            ' Debug.Print "Subtract"
            pInvertSubScriptOption -1
            bDefault = False
         End If
      Case vbKeyP
         If AllowShortCut(ercPrint_CtrlP) Then
            PrintDoc m_sFileName
            bDefault = False
         End If
      Case vbKeyN
         If AllowShortCut(ercNew_CtrlN) Then
            Contents(SF_TEXT) = ""
         End If
      
      End Select
   End If

    '<EhFooter>
    Exit Function

pDoDefault_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pDoDefault " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
Private Sub pInvertSubScriptOption(ByVal lSelItem As Long)
    '<EhHeader>
    On Error GoTo pInvertSubScriptOption_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long

   tCF.dwMask = CFM_OFFSET
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, 1, tCF)
      ' Debug.Print lR
   If (Abs(tCF.yOffset) = Abs(lSelItem)) Then
      tCF.yOffset = 0
   Else
      tCF.yOffset = Sgn(lSelItem)
   End If
   tCF.dwMask = CFM_OFFSET
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, 1, tCF)
      ' Debug.Print lR
      
    '<EhFooter>
    Exit Sub

pInvertSubScriptOption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pInvertSubScriptOption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub pInvertFontOption(ByVal lEffect As Long, ByVal lMask As Long)
    '<EhHeader>
    On Error GoTo pInvertFontOption_Err
    '</EhHeader>
Dim tCF As CHARFORMAT
Dim lR As Long

   tCF.dwEffects = lEffect
   tCF.dwMask = lMask
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_GETCHARFORMAT, (SCF_WORD Or SCF_SELECTION), tCF)
   If ((tCF.dwEffects And lEffect) = lEffect) Then
      tCF.dwEffects = 0
   Else
      tCF.dwEffects = lEffect
   End If
   tCF.dwMask = lMask
   tCF.cbSize = Len(tCF)
   lR = W_SendMessageAnyRef(m_hWnd, EM_SETCHARFORMAT, (SCF_WORD Or SCF_SELECTION), tCF)
      ' Debug.Print lR
   
    '<EhFooter>
    Exit Sub

pInvertFontOption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pInvertFontOption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub pDrawBackground(ByVal lHDC As Long, ByRef tR As RECT)
    '<EhHeader>
    On Error GoTo pDrawBackground_Err
    '</EhHeader>
Dim hBr As Long
   If Not m_cTile.Picture Is Nothing Then
      m_cTile.TileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top
   Else
      hBr = CreateSolidBrush(TranslateColor(BackColor))
      FillRect lHDC, tR, hBr
      DeleteObject hBr
   End If
    '<EhFooter>
    Exit Sub

pDrawBackground_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pDrawBackground " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub pClipScrollBars(ByRef tR As RECT)
    '<EhHeader>
    On Error GoTo pClipScrollBars_Err
    '</EhHeader>
Dim lS As Long
Dim bHorz As Boolean
Dim bVert As Boolean
Dim tWR As RECT
Dim lH As Long
Dim lW As Long

   ' This doesn't actually have the desired effect.
   ' Left in anyway in case I can work out how to do it
   ' properly.  See pRedrawScrollBars
   lS = GetWindowLong(m_hWnd, GWL_STYLE)
   bHorz = ((lS And WS_HSCROLL) = WS_HSCROLL)
   bVert = ((lS And WS_VSCROLL) = WS_VSCROLL)
   If bHorz Or bVert Then
      GetWindowRect m_hWnd, tWR
      If bHorz Then
         lH = GetSystemMetrics(SM_CYHSCROLL)
         If tR.Bottom - tR.Top > tWR.Bottom - tWR.Top - lH Then
            tR.Bottom = tR.Bottom - lH
         End If
      End If
      If bVert Then
         lW = GetSystemMetrics(SM_CXVSCROLL)
         If tR.Right - tR.Left > tWR.Right - tWR.Left - lW Then
            tR.Right = tR.Right - lW
         End If
      End If
   End If
   
    '<EhFooter>
    Exit Sub

pClipScrollBars_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pClipScrollBars " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Sub pRedrawScrollBars()
    '<EhHeader>
    On Error GoTo pRedrawScrollBars_Err
    '</EhHeader>
Dim lS As Long
Dim bHorz As Boolean
Dim bVert As Boolean
Dim tR As RECT
Dim lH As Long
Dim lW As Long

   lS = GetWindowLong(m_hWnd, GWL_STYLE)
   bHorz = ((lS And WS_HSCROLL) = WS_HSCROLL)
   bVert = ((lS And WS_VSCROLL) = WS_VSCROLL)
   If bHorz Or bVert Then
      ' unsubtle, but on deadline:
      InvalidateRectAsNull m_hWndParent, 0&, 1
      UpdateWindow m_hWndParent
      InvalidateRectAsNull m_hWnd, 0&, 1
      UpdateWindow m_hWnd
   End If

    '<EhFooter>
    Exit Sub

pRedrawScrollBars_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.pRedrawScrollBars " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    '
    '<EhHeader>
    On Error GoTo ISubclass_MsgResponse_Err
    '</EhHeader>
    '<EhFooter>
    Exit Property

ISubclass_MsgResponse_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ISubclass_MsgResponse " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    '<EhHeader>
    On Error GoTo ISubclass_MsgResponse_Err
    '</EhHeader>
   Select Case CurrentMessage
   Case WM_MOUSEACTIVATE, WM_ERASEBKGND, WM_PAINT
      ISubclass_MsgResponse = emrConsume
   Case Else
      ISubclass_MsgResponse = emrPreprocess
   End Select
    '<EhFooter>
    Exit Property

ISubclass_MsgResponse_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ISubclass_MsgResponse " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '<EhHeader>
    On Error GoTo ISubClass_WindowProc_Err
    '</EhHeader>
Dim tNMH As NMHDR_RICHEDIT
Dim tSC As SELCHANGE
Dim tEN As ENLINK
Dim tMF As MSGFILTER
Dim tPR As ENPROTECTED
Dim tP As POINTAPI
Dim tR As RECT
Dim tPS As PAINTSTRUCT
Dim X As Single, Y As Single
Dim iKeyCode As Integer, iKeyAscii As Integer, iShift As Integer
Dim iBtn As Integer
Dim bDefault As Boolean
Dim bDoIt As Boolean
Dim ID As Long
Dim bLock As Boolean
Dim iNotifyMsg As Long

   Select Case iMsg
   Case WM_COMMAND
      iNotifyMsg = (wParam And &H7FFF0000) \ &H10000
      Select Case iNotifyMsg
      Case EN_CHANGE
         RaiseEvent Change
      End Select
      
   Case WM_NOTIFY
      CopyMemory tNMH, ByVal lParam, Len(tNMH)
      If (tNMH.hwndFrom = m_hWnd) Then
         
         Select Case tNMH.code
         Case EN_UPDATE
            Debug.Print "Update"
            
         Case EN_SELCHANGE
            CopyMemory tSC, ByVal lParam, Len(tSC)
            RaiseEvent SelectionChange(tSC.chrg.cpMin, tSC.chrg.cpMax, tSC.seltyp)
            
         Case EN_LINK
            CopyMemory tEN, ByVal lParam, Len(tEN)
            RaiseEvent LinkOver(tEN.msg, tEN.chrg.cpMin, tEN.chrg.cpMax)
         
         Case EN_PROTECTED
            CopyMemory tPR, ByVal lParam, Len(tPR)
            bDoIt = False
            RaiseEvent ModifyProtected(bDoIt, tPR.chrg.cpMin, tPR.chrg.cpMax)
            If (bDoIt) Then
               ISubClass_WindowProc = 0
            Else
               ISubClass_WindowProc = 1
            End If
            
         Case EN_MSGFILTER
            bDefault = True
            CopyMemory tMF, ByVal lParam, Len(tMF)
            Select Case tMF.msg
            
            Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK
               'Debug.Print "Double click", tMF.lParam, tMF.wPad2
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent DblClick(X, Y)
            Case WM_LBUTTONDOWN
               iShift = giGetShiftState()
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent MouseDown(vbLeftButton, iShift, X, Y)
            Case WM_RBUTTONDOWN
               iShift = giGetShiftState()
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent MouseDown(vbRightButton, iShift, X, Y)
            Case WM_MBUTTONDOWN
               iShift = giGetShiftState()
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent MouseDown(vbMiddleButton, iShift, X, Y)
            Case WM_LBUTTONUP
               iShift = giGetShiftState()
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent MouseUp(vbLeftButton, iShift, X, Y)
            Case WM_RBUTTONUP
               iShift = giGetShiftState()
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent MouseUp(vbRightButton, iShift, X, Y)
            Case WM_MBUTTONUP
               iShift = giGetShiftState()
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent MouseUp(vbMiddleButton, iShift, X, Y)
            Case WM_MOUSEMOVE
               iShift = giGetShiftState()
               iBtn = giGetMouseButton()
               GetCursorPos tP
               ScreenToClient m_hWnd, tP
               X = tP.X * Screen.TwipsPerPixelX
               Y = tP.Y * Screen.TwipsPerPixelY
               RaiseEvent MouseMove(iBtn, iShift, X, Y)
            Case WM_KEYDOWN
               iShift = giGetShiftState()
               iKeyCode = tMF.wParam
               RaiseEvent KeyDown(iKeyCode, iShift)
               If Not (pDoDefault(iKeyCode, iShift, bDefault)) Then
                  If (iKeyCode <> tMF.wParam) Then
                     bDefault = False
                  End If
               End If
            Case WM_CHAR
               iShift = giGetShiftState()
               iKeyAscii = tMF.wParam
               ' Debug.Print iKeyAscii, iShift
               If Not (pDoDefault(iKeyAscii, iShift, bDefault)) Then
                  RaiseEvent KeyPress(iKeyAscii)
                  If (iKeyAscii <> tMF.wParam) Then
                     bDefault = False
                  End If
               End If
            Case WM_KEYUP
               iShift = giGetShiftState()
               iKeyCode = tMF.wParam
               RaiseEvent KeyUp(iKeyCode, iShift)
            Case Else
               'Debug.Print "Something Different:", tMF.msg, tMF.wParam, tMF.lParam, tMF.wPad1, tMF.wPad2
            End Select
            If Not bDefault Then
               ' Debug.Print "No default.."
               ISubClass_WindowProc = 1&
            End If
         End Select
         
      End If
      
   Case WM_VSCROLL
      RaiseEvent VScroll
      
   Case WM_HSCROLL
      RaiseEvent HScroll
      
   ' ------------------------------------------------------------------------------
   ' Implement focus.  Many many thanks to Mike Gainer for showing me this
   ' code.
   Case WM_SETFOCUS
      If (m_hWnd = hwnd) Then
         ' The RichEdit control:
         Dim pOleObject                  As IOleObject
         Dim pOleInPlaceSite             As IOleInPlaceSite
         Dim pOleInPlaceFrame            As IOleInPlaceFrame
         Dim pOleInPlaceUIWindow         As IOleInPlaceUIWindow
         Dim pOleInPlaceActiveObject     As VBOleGuids.IOleInPlaceActiveObject
         Dim PosRect                     As RECT
         Dim ClipRect                    As RECT
         Dim FrameInfo                   As OLEINPLACEFRAMEINFO
         Dim grfModifiers                As Long
         Dim AcceleratorMsg              As msg
         
         'Get in-place frame and make sure it is set to our in-between
         'implementation of IOleInPlaceActiveObject in order to catch
         'TranslateAccelerator calls
         Set pOleObject = Me
         Set pOleInPlaceSite = pOleObject.GetClientSite
         If Not pOleInPlaceSite Is Nothing Then
            pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, PosRect, ClipRect, FrameInfo
            If m_IPAOHookStruct.ThisPointer <> 0 Then
               CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
               If Not pOleInPlaceActiveObject Is Nothing Then
                  If Not pOleInPlaceFrame Is Nothing Then
                     pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, 0
                     If Not pOleInPlaceUIWindow Is Nothing Then
                        pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
                     End If
                  End If
               End If
               CopyMemory pOleInPlaceActiveObject, 0&, 4
            End If
         End If
      Else
         ' THe user control:
         SetFocusAPI m_hWnd
      End If
      
   Case WM_MOUSEACTIVATE
      If GetFocus() <> m_hWnd Then
         SetFocusAPI m_hWndParent
         ISubClass_WindowProc = MA_NOACTIVATE
      Else
         ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
   ' End Implement focus.
   ' ------------------------------------------------------------------------------
   
   
   Case WM_ERASEBKGND
      If m_bTransparent Then
         GetClientRect hwnd, tR
         pClipScrollBars tR
         pDrawBackground wParam, tR
         ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      Else
         ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
   
   Case WM_PAINT
      If m_bTransparent Then
         BeginPaint hwnd, tPS
         pClipScrollBars tPS.rcPaint
         pDrawBackground tPS.hdc, tPS.rcPaint
         EndPaint hwnd, tPS
         ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      Else
         ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
      
   Case WM_ACTIVATE
      If m_bTransparent Then
         If wParam > 0 Then
            pRedrawScrollBars
         End If
      End If
   
   End Select
      
    '<EhFooter>
    Exit Function

ISubClass_WindowProc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.ISubClass_WindowProc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


Private Sub UserControl_Initialize()
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>
Dim i As Long
   Debug.Print "RichEditControl:Initialise"
   ' Trap tab key:
   m_bTrapTab = True
   ' Default printing margins for an RTF file:
   m_lLeftMargin = 1800
   m_lRightMargin = 1800
   m_lTopMargin = 1440
   m_lBottomMargin = 1440
   ' Default to the real version of RichEdit:
   m_eVersion = eRICHED20
   ' Redraw the control:
   m_bRedraw = True
   ' Allow all in-built shortcuts:
   For i = ERECInbuiltShortcutConstants.[_First] To ERECInbuiltShortcutConstants.[_Last]
      m_bAllowMethod(i) = True
   Next i
   lblText.Caption = "vbAccelerator Rich Edit Control"
   ' Default text limit
   m_lLimit = 32767
   ' Enable!
   m_bEnabled = True
   
   ' Attach custom IOleInPlaceActiveObject interface
   Dim IPAO As VBOleGuids.IOleInPlaceActiveObject

   With m_IPAOHookStruct
      Set IPAO = Me
      CopyMemory .IPAOReal, IPAO, 4
      CopyMemory .TBEx, Me, 4
      .lpVTable = mIOLEInPlaceActiveObject_RE.IPAOVTable
      .ThisPointer = VarPtr(m_IPAOHookStruct)
   End With
   
    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UserControl_Initialize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_InitProperties()
    '<EhHeader>
    On Error GoTo UserControl_InitProperties_Err
    '</EhHeader>
    pInitialise
    m_eCharFormatRange = ercSetFormatAll
    Set Font = UserControl.Ambient.Font
    m_eCharFormatRange = ercSetFormatSelection
    '<EhFooter>
    Exit Sub

UserControl_InitProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UserControl_InitProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyDown_Err
    '</EhHeader>
    Debug.Print KeyCode
    '<EhFooter>
    Exit Sub

UserControl_KeyDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UserControl_KeyDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
Dim eDefault As ERECScrollBarConstants
   If (UserControl.Ambient.UserMode) Then
      m_eVersion = PropBag.ReadProperty("Version", eRICHED32)
   Else
      UseVersion = PropBag.ReadProperty("Version", eRICHED32)
   End If
   SingleLine = PropBag.ReadProperty("SingleLine", False)
   If m_bSingleLine Then
      eDefault = ercScrollBarsNone
   Else
      eDefault = ercScrollBarsBoth
   End If
   ScrollBars = PropBag.ReadProperty("ScrollBars", eDefault)
   DisableNoScroll = PropBag.ReadProperty("DisableNoScroll", False)
   HideSelection = PropBag.ReadProperty("HideSelection", False)
   PasswordChar = PropBag.ReadProperty("PasswordChar", "")
   m_bBorder = PropBag.ReadProperty("Border", True)
   TRANSPARENT = PropBag.ReadProperty("Transparent", False)
   pInitialise
   Border = m_bBorder
   m_eCharFormatRange = ercSetFormatSelection
   Dim sFnt As New StdFont
   On Error Resume Next
   Set Font = PropBag.ReadProperty("Font")
   Err.Clear
   On Error GoTo 0
   m_eCharFormatRange = ercSetFormatSelection

   BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
   ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
   Text = PropBag.ReadProperty("Text", "")
   ViewMode = PropBag.ReadProperty("ViewMode", ercWordWrap)
   ControlLeftMargin = PropBag.ReadProperty("ControlLightMargin", 0)
   ControlRightMargin = PropBag.ReadProperty("ControlRightMargin", 0)
   TextLimit = PropBag.ReadProperty("TextLimit", 32767)
   TrapTab = PropBag.ReadProperty("TrapTab", True)
   If (UserControl.Ambient.UserMode) Then
      lblText.visible = False
   Else
      lblText.visible = True
   End If
   If m_eVersion = eRICHED20 Then
      AutoURLDetect = PropBag.ReadProperty("AutoURLDetect", True)
      TextOnly = PropBag.ReadProperty("TextOnly", False)
   Else
      m_bAutoURLDetect = PropBag.ReadProperty("AutoURLDetect", True)
      m_bTextOnly = PropBag.ReadProperty("TextOnly", False)
   End If
   ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
Dim tR As RECT
   If (m_hWnd <> 0) Then
      GetClientRect m_hWndParent, tR
      MoveWindow m_hWnd, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, Abs(m_bRedraw)
      tR.Left = m_lLeftMarginPixels
      tR.Right = UserControl.ScaleWidth \ Screen.TwipsPerPixelX - m_lRightMarginPixels
      If (tR.Right < tR.Left) Then tR.Right = tR.Left
      tR.Top = 2
      tR.Bottom = UserControl.ScaleHeight \ Screen.TwipsPerPixelY
      'Redraw = False
      W_SendMessageAnyRef m_hWnd, EM_SETRECT, 0, tR
      ControlLeftMargin = m_lLeftMarginPixels
      ControlRightMargin = m_lRightMarginPixels
      'Redraw = True
   Else
      lblText.Move 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, UserControl.ScaleWidth - 4 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 4 * Screen.TwipsPerPixelY
   End If
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error GoTo UserControl_Terminate_Err
    '</EhHeader>
   
   ' Destroy the control & clear up:
   pTerminate
   
   ' Detach the custom IOleInPlaceActiveObject interface
   ' pointers.
   With m_IPAOHookStruct
      CopyMemory .IPAOReal, 0&, 4
      CopyMemory .TBEx, 0&, 4
   End With

   Debug.Print "RichEditControl:Terminate"

    '<EhFooter>
    Exit Sub

UserControl_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UserControl_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>
Dim eDefault As ERECScrollBarConstants
   ' Write properties:
   PropBag.WriteProperty "Version", UseVersion, eRICHED32
   m_eCharFormatRange = ercSetFormatAll
   PropBag.WriteProperty "Font", Font
   PropBag.WriteProperty "BackColor", BackColor, vbWindowBackground
   PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
   PropBag.WriteProperty "Text", m_sText, ""
   PropBag.WriteProperty "ViewMode", ViewMode
   PropBag.WriteProperty "Border", Border, True
   PropBag.WriteProperty "ControlLeftMargin", m_lLeftMarginPixels, 0
   PropBag.WriteProperty "ControlRightMargin", m_lRightMarginPixels, 0
   PropBag.WriteProperty "TextLimit", TextLimit, 32767
   PropBag.WriteProperty "TrapTab", TrapTab, True
   If m_eVersion = eRICHED20 Then
      PropBag.WriteProperty "AutoURLDetect", AutoURLDetect, True
      PropBag.WriteProperty "TextOnly", TextOnly, False
      PropBag.WriteProperty "Transparent", TRANSPARENT, False
   Else
      PropBag.WriteProperty "AutoURLDetect", m_bAutoURLDetect, True
      PropBag.WriteProperty "TextOnly", m_bTextOnly, False
      PropBag.WriteProperty "Transparent", m_bTransparent, False
   End If
   
   PropBag.WriteProperty "ReadOnly", ReadOnly, False
   PropBag.WriteProperty "Enabled", Enabled, True
   PropBag.WriteProperty "SingleLine", SingleLine, False
   PropBag.WriteProperty "DisableNoScroll", DisableNoScroll, False
   PropBag.WriteProperty "PasswordChar", PasswordChar, ""
   If m_bSingleLine Then
      eDefault = ercScrollBarsNone
   Else
      eDefault = ercScrollBarsBoth
   End If
   PropBag.WriteProperty "ScrollBars", ScrollBars, eDefault
   PropBag.WriteProperty "HideSelection", HideSelection, False
   
    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.vbaRichEdit.UserControl_WriteProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


