VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTest2 
   Caption         =   "Excel Data Hunter"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   740
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   1588
      ButtonWidth     =   2090
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "imlToolBar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Excel Import"
            Key             =   "Excel Import"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Build Query"
            Key             =   "Build Query"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   "About"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   7845
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlBar 
      Left            =   3840
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":12FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":1454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":1B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":20E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":272C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":3406
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":3858
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":8C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":909C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest2.frx":94EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6015
      Left            =   4080
      TabIndex        =   23
      Top             =   1680
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10610
      _Version        =   393216
      Rows            =   12
      Cols            =   6
      MousePointer    =   2
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6855
      Left            =   3960
      TabIndex        =   18
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   12091
      ImageList       =   "imlTab"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Data View"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dialog View"
            Key             =   "Dialog View"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Design View"
            Key             =   "Design View"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ExplorerBarLiteTest.isExplorerBar isExplorerBar1 
      Align           =   3  'Align Left
      Height          =   6945
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   12250
      FontCharset     =   0
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   233
         TabIndex        =   12
         Top             =   4440
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox Text5 
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
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox Text6 
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
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox Text2 
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
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
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
            Left            =   2400
            TabIndex        =   14
            Top             =   1560
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Include Sub Folders"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Containing text:"
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
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Look in:"
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
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Search for files or folders named:"
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
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1680
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   233
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   3495
         Begin VB.OptionButton OpLast 
            Caption         =   "In the last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton OpLast 
            Caption         =   "In the last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OpDate 
            Appearance      =   0  'Flat
            Caption         =   "Modified"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            MaskColor       =   &H8000000F&
            TabIndex        =   5
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OpDate 
            Caption         =   "Created"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00DB6722&
            Height          =   285
            Left            =   1320
            TabIndex        =   3
            Text            =   "1"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00DB6722&
            Height          =   285
            Left            =   1320
            TabIndex        =   2
            Text            =   "1"
            Top             =   1200
            Width           =   495
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   255
            Left            =   1800
            TabIndex        =   8
            Top             =   1200
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   255
            Left            =   1800
            TabIndex        =   9
            Top             =   840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin VB.Label Label6 
            BackColor       =   &H00F7DFD6&
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Months"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   10
            Top             =   1320
            Width           =   570
         End
      End
      Begin VB.PictureBox pQueryBuilder 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   240
         ScaleHeight     =   385
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox Check3 
            Caption         =   "And Remember this Password"
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
            Left            =   480
            TabIndex        =   39
            Top             =   4920
            Width           =   2535
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Always use this user"
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
            Left            =   240
            TabIndex        =   38
            Top             =   4560
            Width           =   2775
         End
         Begin VB.Frame Frame1 
            Caption         =   "File Name"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   170
            TabIndex        =   29
            Top             =   600
            Width           =   2895
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2175
               Left            =   120
               ScaleHeight     =   2175
               ScaleWidth      =   2655
               TabIndex        =   31
               Top             =   720
               Width           =   2655
               Begin VB.CommandButton Command2 
                  Caption         =   "..."
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
                  Left            =   1800
                  TabIndex        =   35
                  Top             =   480
                  Width           =   495
               End
               Begin VB.OptionButton Option2 
                  Caption         =   "Internet Source"
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
                  Left            =   120
                  TabIndex        =   34
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "..."
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
                  Left            =   1800
                  TabIndex        =   33
                  Top             =   120
                  Width           =   495
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Local File"
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
                  Left            =   120
                  TabIndex        =   32
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.Label Label9 
                  Caption         =   "This is a GUI Design Demo for Ilustrative porpouses only, This has NOT Data Functionality."
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   855
                  Left            =   120
                  TabIndex        =   37
                  Top             =   1200
                  Width           =   2415
               End
               Begin VB.Label Label3 
                  Caption         =   "Note:"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   36
                  Top             =   960
                  Width           =   2415
               End
            End
            Begin VB.Label Label2 
               Caption         =   "Select the Data Source"
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
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   2655
            End
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   27
            Top             =   4080
            Width           =   2775
         End
         Begin ComctlLib.TabStrip TabStrip2 
            Height          =   3495
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   6165
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   4
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Source File"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Source Table"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Advanced"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Extra Options"
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Start as user:"
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
            Left            =   240
            TabIndex        =   28
            Top             =   3840
            Width           =   2775
         End
      End
   End
   Begin ComctlLib.ImageList imlTab 
      Left            =   3840
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":9940
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":9C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":9F74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlToolBar 
      Left            =   3840
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":A28E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":A5A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":A8C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":ABDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":AEF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest2.frx":B210
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ShowClassicView
End Sub

Private Sub ShowSearchBar()
    'The following code allows you to use controls
    'inside a isExplorerBar Control.
    'follow these steps:
    '1: Insert a control in the form, on this case,
    '   For ilustrative porpouses I'll let the original
    '   name: isExplorerBar1
    '
    '2: Insert a PictureBox in the ExplorerBar. this will
    '   contain some controls, like buttons, text boxes,
    '   labels, or whatever you want. set the visible
    '   property of the picture Box to False
    '
    '3: Add the desired controls in the Picture Box, and
    '   rezise the controls and the picture box to the size
    '   you want.
    '
    '3.2 (Optional) if you want to speed up the changes,
    '   you can disable updates:
    isExplorerBar1.DisableUpdates True
    
    '3.1 READ! If toy want to delete previous Items in the
    '   control, call ClearStructure!
    isExplorerBar1.ClearStructure
    
    '4: create a group for the recently created picturebox,
    '   like this:
    isExplorerBar1.AddGroup "SearchParameters", "Search Parameters", 1
    
    'isExplorerBar1.AddSpecialGroup "Data Tasks", Me.Icon
    'isExplorerBar1.AddItem -1, "Search", "Show Search Options"
    'isExplorerBar1.AddItem -1, "Classic", "Show Normal Mode"
        
    '5: Now, Attach the picture Box to the group.
    isExplorerBar1.SetGroupChild "SearchParameters", Picture1
    
    'Optional: Add more Stuff
    isExplorerBar1.AddGroup "Date Modified", "Date Modified", 1
    isExplorerBar1.SetGroupChild "Date Modified", Picture2
    
    
    '5.1 If you disabled updates, enable Again!
    isExplorerBar1.DisableUpdates False
    
    'That's all! Easy, Isn't It?
    'Enjoy!
 
End Sub

'This sub will Create a classic Bar Structure
Private Sub ShowClassicView()
    'Prevent redrawing while loading (optional)
    'Try the code with and without this line, you will see the diference!
    isExplorerBar1.DisableUpdates True
    'Clear Previous Items and groups
    isExplorerBar1.ClearStructure
        
    'Optional: Set Imagelist
    isExplorerBar1.SetImageList imlBar
    'Add Special header with options
    isExplorerBar1.AddSpecialGroup "Data Tasks", Me.Icon
    isExplorerBar1.AddItem -1, "Save as Table", "Save as Table", 7
    isExplorerBar1.AddItem -1, "Print", "Print Selected Table", 11
        
    'add a New Group With Some Items
    isExplorerBar1.AddGroup "FileTypes", "Import Options"
    isExplorerBar1.AddItem "FileTypes", "Excel", "Excel Shet", 1
    isExplorerBar1.AddItem "FileTypes", "FrontPage", "FrontPage Document", 2
    isExplorerBar1.AddItem "FileTypes", "Bind", "Bind Document", 3
    isExplorerBar1.AddItem "FileTypes", "NonFormated", "Other Document Formats (Needs Plugins)", 4
    
    'Add another Group with items
    isExplorerBar1.AddGroup "Tools", "Tools"
    isExplorerBar1.AddItem "Tools", "Wizard", "Use Import Data Wizard", 6
    isExplorerBar1.AddItem "Tools", "Show Query Builder", "Show Query Builder", 10
    
    isExplorerBar1.DisableUpdates False

End Sub

' Desc: You Can Create Even Big Tools On the Bar, see this one:
Private Sub ShowQueryBuilder()
    'We Will Create the functionality of a dialog Like control here:
    isExplorerBar1.ClearStructure
    isExplorerBar1.AddGroup "QueryBuilder", "Query Builder"
    isExplorerBar1.SetGroupChild "QueryBuilder", pQueryBuilder
End Sub

Private Sub Form_Resize()
    'Cheap Error H.:
    On Error Resume Next
    TabStrip1.Move TabStrip1.Left, TabStrip1.Top, ScaleWidth - TabStrip1.Left - 8, ScaleHeight - TabStrip1.Top - 20
    MSHFlexGrid1.Move MSHFlexGrid1.Left, MSHFlexGrid1.Top, ScaleWidth - TabStrip1.Left - 24, ScaleHeight - MSHFlexGrid1.Top - 30
End Sub

Private Sub isExplorerBar1_GroupOut(sGroup As String)
    Debug.Print "Group Out"
End Sub

'Change the isExplorerBar Content!
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case "Search"   'Show Search Bar
            ShowSearchBar
        Case "Excel Import"
            ShowClassicView
        Case "Build Query"
            ShowQueryBuilder
        Case "About"
            isExplorerBar1.About
    End Select
End Sub

