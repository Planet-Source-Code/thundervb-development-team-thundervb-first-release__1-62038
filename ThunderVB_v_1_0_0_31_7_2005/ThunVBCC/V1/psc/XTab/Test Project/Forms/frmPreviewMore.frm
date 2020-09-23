VERSION 5.00
Object = "*\A..\..\prjXTab.vbp"
Begin VB.Form frmPreviewMore 
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
   Icon            =   "frmPreviewMore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreviewEvenMore 
      Caption         =   "::Preview Even More::"
      Height          =   285
      Left            =   3510
      TabIndex        =   12
      Top             =   7620
      Width           =   3315
   End
   Begin prjXTab.XTab XTab1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   540
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
      ShowFocusRect   =   0   'False
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
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "::Back::"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   7620
      Width           =   2025
   End
   Begin VB.CheckBox chkShowFocusRect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Show Focus Rect"
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   7380
      Width           =   1815
   End
   Begin prjXTab.XTab XTab2 
      Height          =   1335
      Left            =   3540
      TabIndex        =   1
      Top             =   540
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabStyle        =   1
      ShowFocusRect   =   0   'False
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
   End
   Begin prjXTab.XTab XTab3 
      Height          =   1335
      Left            =   150
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      InActiveTabHeight=   20
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16744576
      ActiveTabBackEndColor=   16744576
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   16761024
      ActiveTabForeColor=   16777215
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
   End
   Begin prjXTab.XTab XTab4 
      Height          =   1335
      Left            =   3570
      TabIndex        =   3
      Top             =   2400
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabStyle        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   0
      ActiveTabBackEndColor=   0
      InActiveTabBackStartColor=   8421504
      InActiveTabBackEndColor=   8421504
      ActiveTabForeColor=   16777215
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
      TopLeftInnerBorderColor=   0
      BottomRightInnerBorderColor=   0
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
   End
   Begin prjXTab.XTab XTab5 
      Height          =   1335
      Left            =   150
      TabIndex        =   4
      Top             =   4290
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPreviewMore.frx":000C
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPreviewMore.frx":011E
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmPreviewMore.frx":0230
      InActiveTabHeight=   20
      ShowFocusRect   =   0   'False
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
   End
   Begin prjXTab.XTab XTab6 
      Height          =   1335
      Left            =   3570
      TabIndex        =   5
      Top             =   4290
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPreviewMore.frx":0342
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPreviewMore.frx":0454
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmPreviewMore.frx":0566
      ActiveTabHeight =   22
      InActiveTabHeight=   20
      TabStyle        =   1
      ShowFocusRect   =   0   'False
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
   End
   Begin prjXTab.XTab XTab7 
      Height          =   1245
      Left            =   150
      TabIndex        =   6
      Top             =   6060
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   2196
      TabCount        =   4
      TabCaption(0)   =   "1"
      TabCaption(1)   =   "2"
      TabCaption(2)   =   "3"
      TabCaption(3)   =   "4"
      ActiveTabHeight =   50
      InActiveTabHeight=   20
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16744576
      ActiveTabBackEndColor=   16744576
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   16761024
      ActiveTabForeColor=   16777215
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
   End
   Begin prjXTab.XTab XTab8 
      Height          =   1245
      Left            =   1890
      TabIndex        =   7
      Top             =   6060
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   2196
      TabCount        =   4
      TabCaption(0)   =   "1"
      TabCaption(1)   =   "2"
      TabCaption(2)   =   "3"
      TabCaption(3)   =   "4"
      InActiveTabHeight=   50
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   8454143
      ActiveTabBackEndColor=   8454143
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   16761024
      ActiveTabForeColor=   16777215
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
   End
   Begin prjXTab.XTab XTab9 
      Height          =   1245
      Left            =   3780
      TabIndex        =   8
      Top             =   6090
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   2196
      TabCount        =   4
      TabCaption(0)   =   "1"
      TabCaption(1)   =   "2"
      TabCaption(2)   =   "3"
      TabCaption(3)   =   "4"
      InActiveTabHeight=   50
      TabTheme        =   3
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   12640511
      InActiveTabBackStartColor=   15066597
      InActiveTabBackEndColor=   -2147483626
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
      OuterBorderColor=   9474192
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      XRadius         =   50
      YRadius         =   50
   End
   Begin prjXTab.XTab XTab10 
      Height          =   1245
      Left            =   5490
      TabIndex        =   9
      Top             =   6120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   2196
      TabCount        =   4
      TabCaption(0)   =   "1"
      TabCaption(1)   =   "2"
      TabCaption(2)   =   "3"
      TabCaption(3)   =   "4"
      InActiveTabHeight=   50
      TabTheme        =   3
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   12632319
      InActiveTabBackStartColor=   15066597
      InActiveTabBackEndColor=   -2147483626
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
      OuterBorderColor=   9474192
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      XRadius         =   0
      YRadius         =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Themes Emulating a Hand::"
      Height          =   195
      Left            =   2370
      TabIndex        =   16
      Top             =   5790
      Width           =   2565
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Windows 9x (Using Icons)::"
      Height          =   195
      Left            =   2385
      TabIndex        =   15
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Windows 9x (Customized)::"
      Height          =   195
      Left            =   2340
      TabIndex        =   14
      Top             =   2070
      Width           =   2550
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Windows 9x (Standard)::"
      Height          =   195
      Left            =   2430
      TabIndex        =   13
      Top             =   180
      Width           =   2325
   End
End
Attribute VB_Name = "frmPreviewMore"
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

Private Sub cmdBack_Click()
  Load frmPreviewMain
  frmPreviewMain.Refresh
  
  frmPreviewMain.Left = Me.Left
  frmPreviewMain.Top = Me.Top
  
  frmPreviewMain.Show
  frmPreviewMain.Refresh
  
  Me.Hide
  Unload Me
End Sub

Private Sub cmdPreviewEvenMore_Click()
  Load frmPreviewEvenMore
  frmPreviewEvenMore.Refresh
  
  frmPreviewEvenMore.Left = Me.Left
  frmPreviewEvenMore.Top = Me.Top
  
  frmPreviewEvenMore.Show
  frmPreviewEvenMore.Refresh
  
  Me.Hide
  Unload Me
End Sub

