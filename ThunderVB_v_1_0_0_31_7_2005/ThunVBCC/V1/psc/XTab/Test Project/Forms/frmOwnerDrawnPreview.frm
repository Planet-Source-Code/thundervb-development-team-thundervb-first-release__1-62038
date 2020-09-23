VERSION 5.00
Object = "*\A..\..\..\..\..\MYPROJ~1\XTABCH~2\XTAB\prjXTab.vbp"
Begin VB.Form frmOwnerDrawnPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjXTab.XTab XTab1 
      Height          =   2535
      Left            =   90
      TabIndex        =   2
      Top             =   420
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   4471
      TabCaption(0)   =   "Tab 0"
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "cmdClose"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabTheme        =   4
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin VB.CommandButton cmdClose 
         Caption         =   "::Close::"
         Height          =   285
         Left            =   1260
         TabIndex        =   3
         Top             =   990
         Width           =   975
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "::OwnerDrawn::"
      Height          =   195
      Left            =   1095
      TabIndex        =   1
      Top             =   60
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "This theme is not complete. I included it in this release just to give an idea what could be done."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3150
      Width           =   3495
   End
End
Attribute VB_Name = "frmOwnerDrawnPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'===Types=============================================================================================================
Private Type Size
  cx As Long
  cy As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type LocalTabInfo
  ClickableRect As RECT
  Caption As String
End Type
'=====================================================================================================================

'===Constants=========================================================================================================
Private Const PS_SOLID As Long = 0
'=====================================================================================================================


'===Declarations======================================================================================================
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'=====================================================================================================================


'====Private Variables================================================================================================
Private m_lHwnd As Long
Private m_lHDC As Long
Private m_iActiveTabHeight As Integer
Private m_iActiveTab As Integer
Private m_iTabCount As Integer
Private m_utRect As RECT

Private m_aryTabs() As LocalTabInfo



Private Sub cmdClose_Click()
  Unload Me
End Sub

'=====================================================================================================================


'=====Event Handlers==================================================================================================
Private Sub Form_Load()
  Dim iCnt As Integer
  Dim iTabWidth As Integer
  
  Call pCacheControlProps      'get local copies of properties
  
  ReDim m_aryTabs(m_iTabCount - 1)
  
  iTabWidth = m_utRect.Right / m_iTabCount
  
  For iCnt = 0 To m_iTabCount - 1
    
    m_aryTabs(iCnt).Caption = "Tab " & iCnt
    m_aryTabs(iCnt).ClickableRect.Top = m_utRect.Bottom - m_iActiveTabHeight
    m_aryTabs(iCnt).ClickableRect.Bottom = m_aryTabs(iCnt).ClickableRect.Top + m_iActiveTabHeight
    
    m_aryTabs(iCnt).ClickableRect.Left = iTabWidth * iCnt
    m_aryTabs(iCnt).ClickableRect.Right = m_aryTabs(iCnt).ClickableRect.Left + iTabWidth
    
  Next
  m_aryTabs(iCnt - 1).ClickableRect.Right = m_utRect.Right
  
End Sub


Private Sub XTab1_DrawBackground(ByVal lhWnd As Long, ByVal lHDC As Long)
  Call pCacheControlProps      'get local copy of props
  
  Call Rectangle(m_lHDC, m_utRect.Left, m_utRect.Top, m_utRect.Right, m_utRect.Bottom)
End Sub

Private Sub XTab1_DrawOnActiveTabChange(ByVal lhWnd As Long, ByVal lHDC As Long)
  Call XTab1_DrawTabs(lhWnd, lHDC)
End Sub


Private Sub XTab1_ShowHideFocus(ByVal lhWnd As Long, ByVal lHDC As Long, ByVal bIsFocused As Boolean)
  Call XTab1_DrawBackground(lhWnd, lHDC)
  Call XTab1_DrawTabs(lhWnd, lHDC)
End Sub

Private Sub XTab1_DrawTabs(ByVal lhWnd As Long, ByVal lHDC As Long)
  Dim iCnt As Integer
  Dim utSize As Size
  Dim lPen As Long
  Dim lOldPen As Long
  
 Call pCacheControlProps  'get local copy of props
    
  For iCnt = 0 To m_iTabCount - 1
    With m_aryTabs(iCnt).ClickableRect
      Call Rectangle(m_lHDC, .Left, .Top, .Right, .Bottom)
      
      Call GetTextExtentPoint32(m_lHDC, m_aryTabs(iCnt).Caption, Len(m_aryTabs(iCnt).Caption), utSize)
      
      Call TextOut(m_lHDC, .Left + ((.Right - .Left) / 2) - (utSize.cx / 2), .Top + ((.Bottom - .Top) / 2) - (utSize.cy / 2), m_aryTabs(iCnt).Caption, Len(m_aryTabs(iCnt).Caption))
      
    End With
  Next
  
  With m_aryTabs(m_iActiveTab).ClickableRect
    lPen = CreatePen(PS_SOLID, 1, pGetRGBFromOLE(vbButtonFace))
    lOldPen = SelectObject(m_lHDC, lPen)
    MoveToEx m_lHDC, .Left + 1, .Top, 0&
    LineTo m_lHDC, m_aryTabs(m_iActiveTab).ClickableRect.Right - 1, m_aryTabs(m_iActiveTab).ClickableRect.Top
    SelectObject m_lHDC, lOldPen
    DeleteObject lPen
  End With
  
End Sub

Private Sub XTab1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim iCnt As Integer
  Dim iX As Integer
  Dim iY As Integer
  iX = CInt(x)
  iY = CInt(y)

  Call pCacheControlProps      'get props
  
  If iY < m_utRect.Bottom - m_iActiveTabHeight Then
    'Button = vbRightButton      'prevent actual code in the Xtab to execute
    Exit Sub       'if above the tab height then no need to enter teh loop
  End If

  For iCnt = 0 To m_iTabCount - 1
    If iX >= m_aryTabs(iCnt).ClickableRect.Left And iX <= m_aryTabs(iCnt).ClickableRect.Right Then
      XTab1.ActiveTab = iCnt
      Exit For
    End If
  Next
  
  Button = vbRightButton      'prevent actual code in the Xtab to execute
End Sub
'=====================================================================================================================


'===Private Functions=================================================================================================


' Convert the OLE color into equivalent RGB Combination
' i.e. Convert vbButtonFace into ==> Light Grey
Private Function pGetRGBFromOLE(lOleColor As Long) As Long
  Dim lRGBColor As Long
  Call TranslateColor(lOleColor, 0, lRGBColor)
  pGetRGBFromOLE = lRGBColor
End Function

'get client rect
Private Sub pCacheControlProps()
  m_iActiveTabHeight = XTab1.ActiveTabHeight
  m_lHwnd = XTab1.Handle
  m_lHDC = XTab1.DC
  m_iActiveTab = XTab1.ActiveTab
  m_iTabCount = XTab1.TabCount
  Call GetClientRect(m_lHwnd, m_utRect)
  m_utRect.Right = m_utRect.Right - 1
  m_utRect.Bottom = m_utRect.Bottom - 1
End Sub

'=====================================================================================================================



