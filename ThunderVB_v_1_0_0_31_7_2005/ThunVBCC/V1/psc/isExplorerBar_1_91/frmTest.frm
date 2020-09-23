VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTest 
   Caption         =   "ISExplorerBar Test Project"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   StartUpPosition =   2  'CenterScreen
   Begin ExplorerBarLiteTest.isExplorerBar isExplorerBar1 
      Align           =   3  'Align Left
      Height          =   7485
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   13203
      FontName        =   "Arial Unicode MS"
      FontCharset     =   134
      UxThemeText     =   0   'False
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   46
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Use Theme Drawing"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   45
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdRestoreStructure 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   44
      Top             =   6480
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   20
      Top             =   7485
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   9480
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":12B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":25C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":30F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":324E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":59C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":669E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7512
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":782C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":8958
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":8C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":9584
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":96DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":9838
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":9992
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":9AEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "More"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearEvents 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   6480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdDlg 
      Left            =   9000
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox listEvents 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "frmTest.frx":9C46
      Left            =   3480
      List            =   "frmTest.frx":9C48
      TabIndex        =   0
      Top             =   4680
      Width           =   6135
   End
   Begin VB.Frame fraOption 
      Caption         =   "Add Item"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   3
      Left            =   3600
      TabIndex        =   2
      Tag             =   "AddItem"
      Top             =   1800
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   5775
         TabIndex        =   24
         Top             =   240
         Width           =   5775
         Begin VB.TextBox txtNewTextItem 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3360
            TabIndex        =   53
            Top             =   2040
            Width           =   2055
         End
         Begin VB.TextBox txtItemGroup 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            TabIndex        =   52
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtItemKey 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   49
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton cmdSetItemText 
            Caption         =   "SetItemText"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   47
            Top             =   1320
            Width           =   2055
         End
         Begin VB.PictureBox pPreview 
            BorderStyle     =   0  'None
            FillColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000F&
            Height          =   255
            Left            =   2880
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   25
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "Add Item"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   32
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox comGroup 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtNewItemCaption 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   27
            Text            =   "New Item Caption"
            Top             =   840
            Width           =   2055
         End
         Begin VB.HScrollBar hsImages 
            Height          =   255
            Left            =   2640
            TabIndex        =   26
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "New Text"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   54
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Label13 
            Caption         =   "Group"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   51
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Key"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Label11 
            Caption         =   "Set Item Text"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            BorderWidth     =   3
            X1              =   0
            X2              =   5640
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label2 
            Caption         =   "Add to Group:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Item Caption"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Select Image:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   600
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Details Group"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   4
      Left            =   3600
      TabIndex        =   6
      Tag             =   "Details"
      Top             =   1800
      Width           =   6015
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   5775
         TabIndex        =   33
         Top             =   240
         Width           =   5775
         Begin VB.CheckBox chkUseDetailsImage 
            Caption         =   "Use Image"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemoveDetailsImage 
            Caption         =   "Remove"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   42
            Top             =   1800
            Width           =   1215
         End
         Begin VB.PictureBox pDetailsImage 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   1320
            Picture         =   "frmTest.frx":9C4A
            ScaleHeight     =   1155
            ScaleWidth      =   1155
            TabIndex        =   41
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetDetailsImage 
            Caption         =   "Set Image"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   40
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkDetailsGroup 
            Caption         =   "Use Details Group"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txtDetailsDescription 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   35
            Text            =   "Description"
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtDetailsText 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   2880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Text            =   "frmTest.frx":AA32
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label9 
            Caption         =   "Details Text"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   38
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label8 
            Caption         =   "Details Caption"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   360
            Width           =   2295
         End
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Special Group"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Tag             =   "Special"
      Top             =   1800
      Width           =   6015
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   5775
         TabIndex        =   11
         Top             =   240
         Width           =   5775
         Begin VB.PictureBox pBack 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3045
            Left            =   2520
            Picture         =   "frmTest.frx":AA8A
            ScaleHeight     =   3045
            ScaleWidth      =   3585
            TabIndex        =   18
            Top             =   480
            Width           =   3585
         End
         Begin VB.PictureBox pIcon 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1560
            Picture         =   "frmTest.frx":D0E5
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   17
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtSpecialGroupCaption 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   15
            Text            =   "isExplorerBar Links"
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton cmdSetSpecialGroupImage 
            Caption         =   "Change Icon"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdChangeBackground 
            Caption         =   "Change Background"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   13
            Top             =   0
            Width           =   2175
         End
         Begin VB.CheckBox chkUseSpecialGroup 
            Caption         =   "Use Special Group"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Special Group Caption"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Add Group"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   2
      Left            =   3600
      TabIndex        =   1
      Tag             =   "AddGroup"
      Top             =   1800
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   4695
         TabIndex        =   19
         Top             =   240
         Width           =   4695
         Begin VB.TextBox txtGroupCaption 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   22
            Text            =   "New Group Caption"
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton cmdAddGroup 
            Caption         =   "Create Group"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   21
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Group Caption:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3720
      Picture         =   "frmTest.frx":D9AF
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label10 
      Caption         =   $"frmTest.frx":DCB9
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   39
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "App Demo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISExplorerBar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   2835
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDetailsGroup_Click()

    If chkDetailsGroup.Value = vbChecked Then
        isExplorerBar1.AddDetailsGroup "Details", txtDetailsDescription.Text, txtDetailsText.Text
    Else
        isExplorerBar1.HideDetailsGroup
    End If
End Sub

Private Sub chkUseDetailsImage_Click()
    If chkUseDetailsImage.Value = vbChecked Then
        isExplorerBar1.SetDetailsImage pDetailsImage.Picture
    Else
        isExplorerBar1.SetDetailsImage
    End If
End Sub

Private Sub chkUseSpecialGroup_Click()
    If chkUseSpecialGroup.Value = vbChecked Then
        'Hide Group
        isExplorerBar1.AddSpecialGroup txtSpecialGroupCaption.Text, pIcon.Picture, pBack.Picture
    Else
        isExplorerBar1.HideSpecialGroup
    End If
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdAddItem_Click()
    isExplorerBar1.AddItem comGroup.Text, txtNewItemCaption.Text, txtNewItemCaption.Text, hsImages.Value
End Sub

Private Sub cmdChangeBackground_Click()
    'Select a New Icon to show
    On Error GoTo CanceledByUser
    cdDlg.CancelError = True
    cdDlg.Filter = "Supported Images (*.bmp;*.jpg)|*.bmp;*.jpg"
    cdDlg.DialogTitle = "Select New Group Background"
    cdDlg.InitDir = App.Path
    cdDlg.ShowOpen
    pBack.Picture = LoadPicture(cdDlg.FileName)
    isExplorerBar1.AddSpecialGroup txtSpecialGroupCaption.Text, pIcon.Picture, pBack.Picture
CanceledByUser:
End Sub

Private Sub cmdClearEvents_Click()
    listEvents.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMore_Click()
    frmTest2.Show
End Sub

Private Sub cmdRemoveDetailsImage_Click()
    isExplorerBar1.SetDetailsImage
End Sub

Private Sub cmdSetDetailsImage_Click()
    'Select a New Imagen For the Details Group
    On Error GoTo CanceledByUser
    cdDlg.CancelError = True
    cdDlg.Filter = "Supported Images (*.bmp;*.jpg)|*.bmp;*.jpg"
    cdDlg.DialogTitle = "Select New Details Group Image"
    cdDlg.InitDir = App.Path
    cdDlg.ShowOpen
    Me.pDetailsImage.Picture = LoadPicture(cdDlg.FileName)
    isExplorerBar1.SetDetailsImage LoadPicture(cdDlg.FileName)
CanceledByUser:
End Sub

Private Sub cmdSetItemText_Click()
    isExplorerBar1.SetItemText txtItemGroup.Text, txtItemKey.Text, txtNewTextItem.Text
End Sub

Private Sub cmdSetSpecialGroupImage_Click()
    'Select a New Icon to show
    On Error GoTo CanceledByUser
    cdDlg.CancelError = True
    cdDlg.Filter = "Windows Icons (*.ico)|*.ico|Cursors (*.cur)|*.cur"
    cdDlg.DialogTitle = "Select New Group Icon"
    cdDlg.InitDir = App.Path
    cdDlg.ShowOpen
    pIcon.Picture = LoadPicture(cdDlg.FileName, vbLPLarge, vbLPColor, 32, 32)
    isExplorerBar1.AddSpecialGroup txtSpecialGroupCaption.Text, pIcon.Picture, pBack.Picture
CanceledByUser:
    
End Sub

Private Sub cmdAddGroup_Click()
    isExplorerBar1.AddGroup txtGroupCaption.Text, txtGroupCaption.Text, 0
    comGroup.AddItem txtGroupCaption.Text
End Sub

Private Sub cmdRestoreStructure_Click()
    'Disable redrawing, we will make lot of changes ans this prevent's useless drawings
    isExplorerBar1.DisableUpdates True
    'Clear the structure
    isExplorerBar1.ClearStructure
    'set up the image list
    isExplorerBar1.SetImageList imlSmallIcons
    'Enable the Special group
    isExplorerBar1.AddSpecialGroup txtSpecialGroupCaption.Text, pIcon.Picture, pBack.Picture
    'Add Items to Special Group
    'you can use the "name"
    isExplorerBar1.AddItem "Special Group", "Page", "View isExplorerBar Home Page ", 5
    'Or use the Index
    isExplorerBar1.AddItem -1, "About", "About isExplorerBar", 7
    isExplorerBar1.AddItem "Special Group", "PSC", "Go to Planet-source-code to download or vote", 2
    'Actions Group
    isExplorerBar1.AddGroup "Actions", "Demo Options", 0
    'Actions Items
    isExplorerBar1.AddItem "Actions", "Special", "Special  Group Options", 20
    isExplorerBar1.AddItem "Actions", "AddGroup", "Add Group ", 19
    isExplorerBar1.AddItem "Actions", "AddItem", "Add Item", 18
    isExplorerBar1.AddItem "Actions", "Details", "Details Group Options", 21
    isExplorerBar1.AddItem "Actions", "Clear", "Clear the structure", 7
    'Support For far asian languajes (Experimental)
    'I'm Using here some test Strings, I don't know what this say,
    'I copied them from VBAcceleartor and Dana seaman's examples
    ' www.vbaccelerator.com, www.cyberactivex.com
    'ÄãÕæµÄÄÜ¹»ÔÚÎÄ±¾¿òÖÐÊ¹ÓÃ±³¾°Í¼Æ¬£¬" '"&#20122;&#27954;&#35821;&#35328;&#25903;&#25345;"
    isExplorerBar1.AddGroup "Chinese", "ÄãÕæµÄÄÜ¹»"
    isExplorerBar1.AddItem "Chinese", "Item1", "µÄÄÜ¹»", 2
    isExplorerBar1.AddItem "Chinese", "Item2", "ÔÚÎÄ±¾¿òÖÐÊ", 2
    isExplorerBar1.AddItem "Chinese", "Item3", "¹ÓÃ±³¾°Í¼Æ¬", 3
    isExplorerBar1.AddItem "Chinese", "Item5", "I have not Idea what does this mean!", 7
    
    'Enable Details Group!
    chkDetailsGroup_Click
    'we have created the entire structure.
    ' Enable Redraw
    isExplorerBar1.DisableUpdates False
    isExplorerBar1.ExpandGroup -1, False
    isExplorerBar1.ExpandGroup "Actions", False
End Sub

Private Sub cmdTest_Click()
    If isExplorerBar1.UseUxThemeText Then
        isExplorerBar1.UseUxThemeText = False
        cmdTest.Caption = "Use Theme Drawing"
    Else
        isExplorerBar1.UseUxThemeText = True
        cmdTest.Caption = "Use Custom drawing"
    End If
End Sub

Private Sub Form_Load()
    Dim ni As Integer, nj As Integer
    'Add count of images to scrollbar
    hsImages.Min = 1
    hsImages.Max = imlSmallIcons.ListImages.Count
    'Build Structure
    cmdRestoreStructure_Click
    comGroup.AddItem "Special Group"
End Sub

Private Sub hsImages_Change()
    pPreview.Picture = imlSmallIcons.ListImages.item(hsImages.Value).Picture
End Sub

Private Sub isExplorerBar1_GroupClick(ByVal Group As Long, bExpanded As Boolean)
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "GroupClick. Group: " & Group & " -> Expanded = " & bExpanded
End Sub

Private Sub isExplorerBar1_GroupHover(sGroup As String)
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "GroupHover. Group: " & sGroup
End Sub

Private Sub isExplorerBar1_GroupOut(sGroup As String)
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "GroupOut: " & sGroup
End Sub

Private Sub isExplorerBar1_ItemClick(sGroup As String, sItemKey As String)
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "ItemClick:" & sGroup & " - Item:" & sItemKey
    ''''' Acording with the item that was clicked, show more options.
    Select Case LCase(sGroup)
        Case "special group"    'This is the Special Group
                                'You can usea also It's Index ( -1 )
            Select Case sItemKey
                Case "Page"
                    'View Home Page
                    isExplorerBar1.OpenLink "http://mx.geocities.com/fred_cpp/isexplorerbar.htm"
                Case "About"
                    'Show About
                    frmAbout.Show vbModal, Me
                Case "PSC"
                    'Open PSC Page
                    isExplorerBar1.OpenLink "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=53572&lngWId=1"
            End Select
        Case "actions"
            If sItemKey = "Clear" Then
                isExplorerBar1.ClearStructure
            Else
                ShowFrame sItemKey 'Show the selected Frame
            End If
    End Select
End Sub

Private Sub isExplorerBar1_ItemHover(sGroup As String, sItemKey As String)
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "ItemHover. Group: " & sGroup & " - Item:" & sItemKey
End Sub

Private Sub isExplorerBar1_ItemOut(sGroup As String, sItemKey As String)
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "ItemOut. Group: " & sGroup & " - Item:" & sItemKey
End Sub

Private Sub isExplorerBar1_MouseOut()
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "Mouse Out"
End Sub

Private Sub isExplorerBar1_MouseOver()
    ' This code is used just to notify you an avent has been fired
    AddEventNotification "Mouse Over"
End Sub

Private Sub txtDetailsDescription_Change()
    ' This Code will update the details text in the control while you type in here
    chkDetailsGroup_Click
End Sub

Private Sub txtDetailsText_Change()
    ' This Code will update the details text in the control while you type in here
    chkDetailsGroup_Click
End Sub

Private Sub txtSpecialGroupCaption_Change()
    ' This code is to update the Special Group Properties
    chkUseSpecialGroup_Click
End Sub

' Desc: Add a event description to the list.
Private Sub AddEventNotification(ByVal sDescription As String)
    listEvents.AddItem sDescription
    listEvents.Selected(listEvents.ListCount - 1) = True
End Sub

' Desc: Hides all the frames in the demo. why? to show
'       Only the frame selected to be shown.
Private Sub ShowFrame(ByVal sTag As String)
    Dim ctl
    For Each ctl In Controls
        If TypeOf ctl Is Frame Then
            If LCase(ctl.Tag) = LCase(sTag) Then
                ctl.Visible = True
            Else
                ctl.Visible = False
            End If
        End If
    Next
End Sub
