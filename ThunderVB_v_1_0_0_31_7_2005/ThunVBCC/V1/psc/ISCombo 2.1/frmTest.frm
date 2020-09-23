VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmtest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISCombo Test"
   ClientHeight    =   5370
   ClientLeft      =   150
   ClientTop       =   525
   ClientWidth     =   8955
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   4080
      Width           =   2355
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "Go to PSC Page"
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   3720
      Width           =   2355
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   6240
      TabIndex        =   20
      Top             =   3360
      Width           =   2355
   End
   Begin VB.PictureBox pGetText 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   -4260
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   0
      Top             =   660
      Visible         =   0   'False
      Width           =   2715
      Begin VB.CommandButton cmdCancelAdd 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   915
      End
      Begin VB.PictureBox pPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   4
         Top             =   300
         Width           =   315
         Begin VB.Image imgPreview 
            Height          =   255
            Left            =   15
            Top             =   15
            Width           =   255
         End
      End
      Begin VB.VScrollBar vsbItem 
         Height          =   795
         Left            =   120
         Max             =   77
         Min             =   1
         TabIndex        =   3
         Top             =   60
         Value           =   1
         Width           =   315
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txtNewItem 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   420
         TabIndex        =   1
         Text            =   "New Item"
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.ListBox lstEvents 
      Height          =   2985
      Left            =   6240
      TabIndex        =   15
      Top             =   120
      Width           =   2475
   End
   Begin VB.CheckBox chkMouseMove 
      Caption         =   "Show Mouse Move Event"
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame fraProps 
      Caption         =   "Some Properties"
      Height          =   4875
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   6015
      Begin VB.CheckBox chkAutocomplete 
         Caption         =   "Autocomplete"
         Height          =   255
         Left            =   1680
         TabIndex        =   54
         Top             =   4440
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdClearComboList 
         Caption         =   "Clear"
         Height          =   435
         Left            =   1740
         TabIndex        =   53
         Top             =   1500
         Width           =   1515
      End
      Begin VB.CommandButton cmdSetWinXPBorderColor 
         Caption         =   "WINXPBorderColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   52
         Top             =   2040
         Width           =   1575
      End
      Begin VB.PictureBox pWINXPBorderColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   51
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Item"
         Height          =   435
         Left            =   180
         TabIndex        =   50
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "ISCombo Test"
         Height          =   975
         Left            =   60
         TabIndex        =   48
         Top             =   240
         Width           =   3435
         Begin prjTestISControls.ISCombo iscDemo 
            Height          =   435
            Left            =   180
            TabIndex        =   49
            Top             =   300
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   767
            MSOXPColor      =   8388608
            FontColor       =   -2147483640
            FontHighlightColor=   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Icon            =   "frmTest.frx":030A
            Style           =   0
         End
      End
      Begin VB.CommandButton cmdListIconsBackColor 
         Caption         =   "List Icons BackColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   47
         Top             =   3840
         Width           =   1575
      End
      Begin VB.PictureBox pListIconsBackColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   46
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton cmdListBackColor 
         Caption         =   "List BackColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   44
         Top             =   2940
         Width           =   1575
      End
      Begin VB.PictureBox pListBackColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   43
         Top             =   2940
         Width           =   495
      End
      Begin VB.PictureBox pListHoverColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   42
         Top             =   3240
         Width           =   495
      End
      Begin VB.PictureBox pListBorderColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   41
         Top             =   3540
         Width           =   495
      End
      Begin VB.CommandButton cmdListBorderColor 
         Caption         =   "List BorderColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   40
         Top             =   3540
         Width           =   1575
      End
      Begin VB.CommandButton cmdListHoverColor 
         Caption         =   "List HoverColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   39
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdRestoreColors 
         Caption         =   "Restore Colors"
         Height          =   375
         Left            =   3660
         TabIndex        =   38
         Top             =   4380
         Width           =   2175
      End
      Begin VB.CommandButton cmdFontColor 
         Caption         =   "FontColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   37
         Top             =   2340
         Width           =   1575
      End
      Begin VB.CommandButton cmdHighLightFontColor 
         Caption         =   "FontHighLighColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   36
         Top             =   2640
         Width           =   1575
      End
      Begin VB.PictureBox pHighLighFontColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   35
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox pFontColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   34
         Top             =   2340
         Width           =   495
      End
      Begin VB.PictureBox pWINXPColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   33
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox pWINXPHoverColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   32
         Top             =   1740
         Width           =   495
      End
      Begin VB.PictureBox pMSOXPColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox pMSOXPHoverColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   30
         Top             =   1140
         Width           =   495
      End
      Begin VB.CommandButton cmdWINXPHoverColor 
         Caption         =   "WINXPHoverColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   29
         Top             =   1740
         Width           =   1575
      End
      Begin VB.CommandButton cmdWINXPColor 
         Caption         =   "WINXPColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton setMSOXPHoverColor 
         Caption         =   "MSOXPHoverColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   27
         Top             =   1140
         Width           =   1575
      End
      Begin VB.CommandButton cmdSetMSOXPColor 
         Caption         =   "MSOXPColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox pHoverColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   25
         Top             =   540
         Width           =   495
      End
      Begin VB.CommandButton cmdSetHoverColor 
         Caption         =   "HoverColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   24
         Top             =   540
         Width           =   1575
      End
      Begin VB.PictureBox pBackcolor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5340
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdSetBackColor 
         Caption         =   "BackColor"
         Height          =   255
         Left            =   3660
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox pDefaultIcon 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   540
         ScaleHeight     =   375
         ScaleWidth      =   315
         TabIndex        =   18
         Top             =   3360
         Width           =   315
         Begin VB.Image imgDefIcon 
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.HScrollBar hsbDefIcon 
         Height          =   375
         Left            =   300
         Max             =   77
         Min             =   1
         TabIndex        =   17
         Top             =   3360
         Value           =   1
         Width           =   795
      End
      Begin prjTestISControls.ISCombo iscStyle 
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   2580
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         MSOXPColor      =   8388608
         Caption         =   "Style"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   3
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Left"
         Height          =   315
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4080
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Right"
         Height          =   315
         Index           =   1
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4080
         Width           =   795
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Center"
         Height          =   315
         Index           =   2
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4080
         Width           =   795
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   4440
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "SetColors to Original Value"
         Height          =   255
         Left            =   3660
         TabIndex        =   45
         Top             =   4140
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Default Icon"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Text Align"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   3840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Style"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   2280
         Width           =   2235
      End
   End
   Begin MSComctlLib.ImageList imlItems 
      Left            =   6180
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   78
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":05C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0720
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":087C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":10BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":121C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":137C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":14DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":163C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":179C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":18FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":213C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":229C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":23FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":255C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":26BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":281C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":297C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2C3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":305C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":31BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":331C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":347C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":35DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":373C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":389C
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":39FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":3F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":40DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":423C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4558
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":46B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":496C
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":4EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5038
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5194
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":52F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":544C
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":55A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5704
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5860
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":59BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":5F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6088
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":61E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6340
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":649C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":65F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7118
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7274
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6780
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":73D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":752C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":7688
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdColor 
      Left            =   5520
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      ControlName:    ISCombo.
''      Version:        2.10
''      Author:         Alfredo Córdova Pérez ( fred_cpp )
''      e-mail:         fred_cpp@hotmail.com
''                      fred_cpp@yahoo.com.mx
''
''      Description:
''
''      This is the second Release of the ISCombo Control.
''      This is a Custom ImageCombo, that supports some aditional Features
''      See ISCombo.ctl For Detailed Info.:
''      you know, you can use this freely, just give me credit.
''      Votes and suggestions are wellcome.
''


Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const STR_LINK          As String = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=35565"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  This code Is used just for this example
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetColor() As OLE_COLOR
    On Error GoTo Canceled
    cdColor.CancelError = True
    cdColor.DialogTitle = "Select a Color"
    cdColor.ShowColor
    GetColor = cdColor.Color
    Exit Function
Canceled:
    GetColor = -1
End Function



Private Sub chkAutocomplete_Click()
    iscDemo.Autocomplete = chkAutocomplete.Value
End Sub

Private Sub cmdClearComboList_Click()
    iscDemo.Clear
End Sub

Private Sub cmdClose_Click()
    Unload Me
    End
End Sub

Private Sub chkEnabled_Click()
    iscDemo.Enabled = chkEnabled.Value
End Sub

Private Sub cmdAdd_Click()
    pGetText.Left = cmdAdd.Left
    pGetText.Top = cmdAdd.Top + cmdAdd.Height + 45
    pGetText.Visible = True
    pGetText.SetFocus
End Sub

Private Sub cmdAddItem_Click()
    iscDemo.AddItem txtNewItem.text, , imlItems.ListImages(vsbItem.Value).Picture
    txtNewItem.text = "New Item"
    pGetText.Visible = False
End Sub

Private Sub cmdClear_Click()
    lstEvents.Clear
End Sub

Private Sub cmdPage_Click()
    ShellExecute 0, "open", STR_LINK, "", "", 5
End Sub

Private Sub cmdCancelAdd_Click()
    txtNewItem.text = "New Item"
    pGetText.Visible = False
End Sub

Private Sub cmdHighLightFontColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pHighLighFontColor.Backcolor = tmpColor
    iscDemo.FontHighlightColor = tmpColor
End Sub

Private Sub cmdListBackColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pListBackColor.Backcolor = tmpColor
    iscDemo.DropDownListBackColor = tmpColor
End Sub

Private Sub cmdFontColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pFontColor.Backcolor = tmpColor
    iscDemo.FontColor = tmpColor
End Sub

Private Sub cmdRestoreColors_Click()
    iscDemo.RestoreOriginalColors
End Sub

Private Sub cmdSetBackColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pBackcolor.Backcolor = tmpColor
    iscDemo.Backcolor = tmpColor
End Sub

Private Sub cmdSetHoverColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pHoverColor.Backcolor = tmpColor
    iscDemo.HoverColor = tmpColor
End Sub

Private Sub cmdSetMSOXPColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pMSOXPColor.Backcolor = tmpColor
    iscDemo.MSOXPColor = tmpColor
End Sub

Private Sub setMSOXPHoverColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pMSOXPHoverColor.Backcolor = tmpColor
    iscDemo.MSOXPHoverColor = tmpColor
End Sub

Private Sub cmdSetWinXPBorderColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pWINXPBorderColor.Backcolor = tmpColor
    iscDemo.WINXPBorderColor = tmpColor
End Sub

Private Sub cmdWINXPColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pWINXPColor.Backcolor = tmpColor
    iscDemo.WINXPColor = tmpColor
End Sub

Private Sub cmdWINXPHoverColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pWINXPHoverColor.Backcolor = tmpColor
    iscDemo.WINXPHoverColor = tmpColor
End Sub

Private Sub cmdListBorderColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pListBorderColor.Backcolor = tmpColor
    iscDemo.DropDownListBorderColor = tmpColor
End Sub

Private Sub cmdListHoverColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pListHoverColor.Backcolor = tmpColor
    iscDemo.DropDownListHoverColor = tmpColor
End Sub

Private Sub cmdListIconsBackColor_Click()
    Dim tmpColor As OLE_COLOR
    tmpColor = GetColor()
    If tmpColor = -1 Then Exit Sub
    pListIconsBackColor.Backcolor = tmpColor
    iscDemo.DropDownListIconsBackColor = tmpColor
End Sub

Private Sub Form_Load()
    Dim ni As Integer
    For ni = 1 To 4
        iscDemo.AddItem "Item " & ni & " In list", ni, imlItems.ListImages(ni).Picture
    Next ni
    iscDemo.AddItem "Planet-Source-Code.com"
    iscDemo.AddItem "Planet-Source-Code.com/vb"
    iscDemo.AddItem "Planet-Source-Code.com/c"
    iscDemo.AddItem "Planet-Source-Code.com/asp"
    iscDemo.AddItem "Planet-Source-Code.com/php"
    iscDemo.AddItem "Planet-Source-Code.com/SQL"
    iscDemo.AddItem "Planet-Source-Code.com/Delphi"
    iscDemo.AddItem "Planet-Source-Code.com/Net"
    
    iscStyle.AddItem "Normal"
    iscStyle.AddItem "Office 2000"
    iscStyle.AddItem "Office XP"
    iscStyle.AddItem "Windows XP"
    imgDefIcon.Picture = imlItems.ListImages(36).Picture
    'Set Colors to pBoxes.
    pBackcolor.Backcolor = iscDemo.Backcolor
    pHoverColor.Backcolor = iscDemo.HoverColor
    pMSOXPColor.Backcolor = iscDemo.MSOXPColor
    pMSOXPHoverColor.Backcolor = iscDemo.MSOXPHoverColor
    pWINXPColor.Backcolor = iscDemo.WINXPColor
    pWINXPHoverColor.Backcolor = iscDemo.WINXPHoverColor
    pWINXPBorderColor.Backcolor = iscDemo.WINXPBorderColor
    pFontColor.Backcolor = iscDemo.FontColor
    pHighLighFontColor.Backcolor = iscDemo.FontHighlightColor
    pListBackColor.Backcolor = iscDemo.DropDownListBackColor
    pListHoverColor.Backcolor = iscDemo.DropDownListHoverColor
    pListBorderColor.Backcolor = iscDemo.DropDownListBorderColor
    pListIconsBackColor.Backcolor = iscDemo.DropDownListIconsBackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub hsbDefIcon_Change()
    imgDefIcon.Picture = imlItems.ListImages(hsbDefIcon.Value).Picture
    Set iscDemo.Icon = imlItems.ListImages(hsbDefIcon.Value).Picture
End Sub

Private Sub iscStyle_ItemClick(iItem As Integer)
    iscStyle.Style = iItem
    iscDemo.Style = iItem
End Sub


Private Sub optAlign_Click(Index As Integer)
    iscDemo.TextAlign = Index
End Sub

Private Sub pGetText_Paint()
    pGetText.Line (0, 0)-(pGetText.ScaleWidth, 0), vb3DHighlight
    pGetText.Line (0, 0)-(0, pGetText.ScaleHeight), vb3DHighlight
    pGetText.Line (pGetText.ScaleWidth - 1, pGetText.ScaleHeight)-(pGetText.ScaleWidth - 1, 0), vb3DShadow
    pGetText.Line (pGetText.ScaleWidth, pGetText.ScaleHeight - 1)-(0, pGetText.ScaleHeight - 1), vb3DShadow
End Sub

Private Sub vsbItem_Change()
    vsbItem_Scroll
End Sub

Private Sub vsbItem_Scroll()
    Me.imgPreview.Picture = imlItems.ListImages(vsbItem.Value).Picture
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''  Events Supported By the ISCombo Version 2.0
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub iscDemo_Click()
    'Click
    lstEvents.AddItem "Click"
End Sub

Private Sub iscDemo_ItemClick(iItem As Integer)
    'ItemClick ( iitem)
    lstEvents.AddItem "ItemClick " & iItem
End Sub

Private Sub iscDemo_KeyPress(KeyAscii As Integer)
    'KeyPress (KeyAscii)
    lstEvents.AddItem "KeyPress " & KeyAscii
    ''If KeyAscii = 13 Then MsgBox "Enter Key Detected" & vbCr & "You can use event detect for multiple porpouses", vbInformation
End Sub

Private Sub iscDemo_MouseHover()
    'MouseHover
    lstEvents.AddItem "MouseHover"
End Sub

Private Sub iscDemo_MouseOut()
    'MouseOut
    lstEvents.AddItem "MouseOut"
End Sub

Private Sub iscDemo_Validate(Cancel As Boolean)
    'Validate
    lstEvents.AddItem "Validate"
End Sub

Private Sub iscDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MouseDown
    lstEvents.AddItem "MouseDown " & Button & ", " & Shift & ", " & X & ", " & Y
End Sub

Private Sub iscDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mouse Move
    If chkMouseMove.Value Then
        lstEvents.AddItem "MouseMove " & Button & ", " & Shift & ", " & X & ", " & Y
    End If
End Sub

Private Sub iscDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'MouseUp
    lstEvents.AddItem "MouseUp " & Button & ", " & Shift & ", " & X & ", " & Y
End Sub

Private Sub iscDemo_ButtonClick()
    'ButtonClick
    lstEvents.AddItem "ButtonClick"
End Sub

Private Sub iscDemo_Change()
    'Change
    lstEvents.AddItem "Change"
End Sub

Private Sub iscDemo_GotFocus()
    'GotFocus
    lstEvents.AddItem "GotFocus"
End Sub

Private Sub iscDemo_LostFocus()
    'LostFocus
    lstEvents.AddItem "LostFocus"
End Sub


