VERSION 5.00
Object = "*\A..\..\prjXTab.vbp"
Begin VB.Form frmPreviewMain 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                           :: XTabs :: by Neeraj Agrawal @ PSC ::"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreviewMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "::More Themes::"
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   7620
      Width           =   2025
   End
   Begin VB.CheckBox chkShowFocusRect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Show Focus Rect"
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   7650
      Width           =   1815
   End
   Begin prjXTab.XTab XTab3 
      Height          =   1335
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab &0"
      TabAccessKey(0) =   48
      TabCaption(1)   =   "Tab &1"
      TabAccessKey(1) =   49
      TabCaption(2)   =   "Tab &2"
      TabAccessKey(2) =   50
      InActiveTabHeight=   20
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
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
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
   End
   Begin prjXTab.XTab XTab4 
      Height          =   1335
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPreviewMain.frx":000C
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPreviewMain.frx":011E
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmPreviewMain.frx":0230
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
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
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
   End
   Begin prjXTab.XTab XTab5 
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   2460
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "  Tab 0"
      TabCaption(1)   =   "  Tab 1"
      TabCaption(2)   =   "   Tab 2"
      InActiveTabHeight=   20
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   12640511
      ActiveTabBackEndColor=   12640511
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   12640511
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
      OuterBorderColor=   0
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
   End
   Begin prjXTab.XTab XTab6 
      Height          =   1335
      Left            =   3630
      TabIndex        =   3
      Top             =   2460
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16316664
      ActiveTabBackEndColor=   16316664
      InActiveTabBackStartColor=   14737632
      InActiveTabBackEndColor=   14737632
      ActiveTabForeColor=   10972496
      InActiveTabForeColor=   9474192
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
      OuterBorderColor=   0
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      HoverColorInverted=   16512
   End
   Begin prjXTab.XTab XTab7 
      Height          =   1335
      Left            =   210
      TabIndex        =   6
      Top             =   6120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0*"
      TabPicture(0)   =   "frmPreviewMain.frx":0342
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPreviewMain.frx":065C
      TabCaption(2)   =   "Tab 2"
      TabStyle        =   1
      TabTheme        =   2
      ShowFocusRect   =   0   'False
      InActiveTabBackStartColor=   -2147483626
      InActiveTabBackEndColor=   -2147483626
      InActiveTabForeColor=   -2147483631
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
      OuterBorderColor=   -2147483628
      TabStripBackColor=   -2147483626
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      HoverColorInverted=   16512
   End
   Begin prjXTab.XTab XTab8 
      Height          =   1335
      Left            =   3630
      TabIndex        =   7
      Top             =   6150
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabStyle        =   1
      TabTheme        =   2
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16761024
      ActiveTabBackEndColor=   16744576
      InActiveTabBackStartColor=   16761024
      InActiveTabBackEndColor=   16761024
      ActiveTabForeColor=   16777215
      InActiveTabForeColor=   -2147483631
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
      OuterBorderColor=   16777215
      TabStripBackColor=   -2147483626
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      HoverColorInverted=   16512
   End
   Begin prjXTab.XTab XTab1 
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   4290
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      InActiveTabHeight=   10
      TabTheme        =   3
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16316664
      ActiveTabBackEndColor=   16316664
      InActiveTabBackStartColor=   15066597
      InActiveTabBackEndColor=   14737632
      ActiveTabForeColor=   10972496
      InActiveTabForeColor=   9474192
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
      OuterBorderColor=   12632256
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      HoverColorInverted=   16512
      XRadius         =   50
      YRadius         =   50
   End
   Begin prjXTab.XTab XTab2 
      Height          =   1335
      Left            =   3600
      TabIndex        =   5
      Top             =   4290
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      InActiveTabHeight=   15
      TabStyle        =   1
      TabTheme        =   3
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   33023
      ActiveTabBackEndColor=   33023
      InActiveTabBackStartColor=   8438015
      InActiveTabBackEndColor=   8438015
      ActiveTabForeColor=   0
      InActiveTabForeColor=   0
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
      OuterBorderColor=   8421504
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      HoverColorInverted=   16512
      XRadius         =   1
      YRadius         =   1
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Round Tabs::"
      Height          =   195
      Left            =   2895
      TabIndex        =   13
      Top             =   3990
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Visual Studio .Net 2003::"
      Height          =   195
      Left            =   2490
      TabIndex        =   12
      Top             =   5850
      Width           =   2310
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Windows XP::"
      Height          =   195
      Left            =   2790
      TabIndex        =   11
      Top             =   180
      Width           =   1365
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Windows XP(Modified)::"
      Height          =   195
      Left            =   2385
      TabIndex        =   10
      Top             =   2130
      Width           =   2190
   End
End
Attribute VB_Name = "frmPreviewMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowFocusRect_Click()
  Dim bFlag As Boolean
  Dim oCtl As Control
  
  
  If chkShowFocusRect.Value = vbChecked Then
    bFlag = True
  Else
    bFlag = False
  End If
  
  'Set focus rect Property for all the XTab controls on the form
  For Each oCtl In Me.Controls
    If TypeOf oCtl Is XTab Then
      oCtl.ShowFocusRect = bFlag
    End If
  Next
End Sub

Private Sub cmdNext_Click()
  Load frmPreviewMore
  frmPreviewMore.Refresh
  
  frmPreviewMore.Left = Me.Left
  frmPreviewMore.Top = Me.Top
  
  frmPreviewMore.Show
  frmPreviewMore.Refresh
  
  Me.Hide
  Unload Me
End Sub
