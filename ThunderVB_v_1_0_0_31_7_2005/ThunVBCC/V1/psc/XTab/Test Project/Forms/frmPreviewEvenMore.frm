VERSION 5.00
Object = "*\A..\..\prjXTab.vbp"
Begin VB.Form frmPreviewEvenMore 
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
   Icon            =   "frmPreviewEvenMore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowFocusRect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Show Focus Rect"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdPreviewProps 
      Caption         =   "::Test Control Properties::"
      Height          =   285
      Left            =   4170
      TabIndex        =   5
      Top             =   7620
      Width           =   2655
   End
   Begin VB.CommandButton cmdPreviewOwnerDrawn 
      Caption         =   "::Preview Owner Drawn Tabs::"
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Top             =   6630
      Width           =   3315
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "::Back::"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   7620
      Width           =   2025
   End
   Begin prjXTab.XTab XTab1 
      Height          =   2265
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   3995
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPreviewEvenMore.frx":000C
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPreviewEvenMore.frx":0326
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmPreviewEvenMore.frx":0640
      ActiveTabHeight =   40
      InActiveTabHeight=   38
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
      PictureAlign    =   2
      PictureSize     =   1
   End
   Begin prjXTab.XTab XTab2 
      Height          =   1905
      Left            =   210
      TabIndex        =   1
      Top             =   3510
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   3360
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPreviewEvenMore.frx":095A
      TabCaption(1)   =   "Disabled"
      TabEnabled(1)   =   0   'False
      TabPicture(1)   =   "frmPreviewEvenMore.frx":0C74
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmPreviewEvenMore.frx":0F8E
      ActiveTabHeight =   40
      InActiveTabHeight=   38
      TabStyle        =   1
      TabTheme        =   3
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16316664
      InActiveTabBackStartColor=   15066597
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
      XRadius         =   20
      YRadius         =   20
      PictureSize     =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Round Tabs(Prperty Page,Large Icon)::"
      Height          =   195
      Left            =   1815
      TabIndex        =   8
      Top             =   3240
      Width           =   3555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::Standard Tabs(Tabbed,Large Icon)::"
      Height          =   195
      Left            =   1935
      TabIndex        =   7
      Top             =   180
      Width           =   3315
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::OwnerDrawn::"
      Height          =   195
      Left            =   2940
      TabIndex        =   6
      Top             =   6300
      Width           =   1410
   End
End
Attribute VB_Name = "frmPreviewEvenMore"
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
  Load frmPreviewMore
  frmPreviewMore.Refresh
  
  frmPreviewMore.Left = Me.Left
  frmPreviewMore.Top = Me.Top
  
  frmPreviewMore.Show
  frmPreviewMore.Refresh
  
  Me.Hide
  Unload Me
End Sub

Private Sub cmdPreviewOwnerDrawn_Click()
  Load frmOwnerDrawnPreview
  frmOwnerDrawnPreview.Show vbModal, Me
End Sub

Private Sub cmdPreviewProps_Click()
  Load frmTestControlProperties
  
  Me.Hide
  frmTestControlProperties.Show vbModal, Me
  Me.Show
End Sub
