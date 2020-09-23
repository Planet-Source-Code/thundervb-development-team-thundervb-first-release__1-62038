VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmButtonTest 
   BackColor       =   &H80000005&
   Caption         =   "isButton 3.0 Test"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9000
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
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pProps 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   2280
      ScaleHeight     =   1545
      ScaleWidth      =   6465
      TabIndex        =   27
      Top             =   2880
      Width           =   6495
      Begin isButtonTest.isButton chkUseCustomColors 
         Height          =   375
         Left            =   4080
         TabIndex        =   62
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Style           =   8
         Caption         =   "Use Custom Colors"
         IconAlign       =   2
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         Value           =   -1  'True
      End
      Begin isButtonTest.isButton chkEnabled 
         Height          =   375
         Left            =   4080
         TabIndex        =   61
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Style           =   8
         Caption         =   "Enabled"
         IconAlign       =   2
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         Value           =   -1  'True
      End
      Begin VB.ComboBox comNonThemeStyle 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox comIconAlign 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Text            =   "isButton"
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox comCaptionAlign 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin isButtonTest.isButton cmdButtonType 
         Height          =   375
         Left            =   4080
         TabIndex        =   63
         Top             =   660
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Style           =   8
         Caption         =   "CheckBox"
         IconAlign       =   2
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "If theme fails, use:"
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Icon Align"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Caption"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Caption Align"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.PictureBox pProps 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   2280
      ScaleHeight     =   1545
      ScaleWidth      =   6465
      TabIndex        =   28
      Top             =   2880
      Width           =   6495
      Begin isButtonTest.isButton cmdFontHighLight 
         Height          =   300
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Icon            =   "frmButtonTest.frx":0000
         Style           =   8
         Caption         =   "Font HighLight"
         IconAlign       =   1
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin isButtonTest.isButton cmdFontColor 
         Height          =   300
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Icon            =   "frmButtonTest.frx":001C
         Style           =   8
         Caption         =   "FontColor"
         IconAlign       =   1
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin isButtonTest.isButton cmdHighLighColor 
         Height          =   300
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Icon            =   "frmButtonTest.frx":0038
         Style           =   8
         Caption         =   "HighLight Color"
         IconAlign       =   1
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin isButtonTest.isButton cmdBackColor 
         Height          =   300
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Icon            =   "frmButtonTest.frx":0054
         Style           =   8
         Caption         =   "BackColor"
         IconAlign       =   1
         iNonThemeStyle  =   0
         BackColor       =   12632256
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin isButtonTest.isButton cmdSetFont 
         Height          =   420
         Left            =   2880
         TabIndex        =   60
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   741
         Icon            =   "frmButtonTest.frx":0070
         Style           =   8
         Caption         =   "Set Font"
         IconAlign       =   1
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   15790320
         HighlightColor  =   15790320
         FontColor       =   -2147483640
         FontHighlightColor=   -2147483635
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bitstream Vera Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpBack 
         BackColor       =   &H00F0FAFA&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   120
         Width           =   255
      End
      Begin VB.Shape shpHighlight 
         BackColor       =   &H00F8FFFF&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape shpFont 
         BackColor       =   &H80000012&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape shpFontHighlight 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   255
      End
   End
   Begin VB.PictureBox pProps 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   3
      Left            =   2280
      ScaleHeight     =   1545
      ScaleWidth      =   6465
      TabIndex        =   57
      Top             =   2880
      Visible         =   0   'False
      Width           =   6495
      Begin isButtonTest.isButton cmdVote 
         Height          =   375
         Left            =   4080
         TabIndex        =   59
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Icon            =   "frmButtonTest.frx":008C
         Style           =   8
         Caption         =   "Go To PSC Page"
         IconAlign       =   2
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   15790320
         HighlightColor  =   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmButtonTest.frx":01E6
         Height          =   855
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   6255
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   8280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":02DA
      Style           =   0
      Caption         =   "Win 9X"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   4
      Left            =   4800
      TabIndex        =   4
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":02F6
      Style           =   5
      Caption         =   "Docs"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   3
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":1348
      Style           =   4
      Caption         =   "Mail"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   1
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   7
      Left            =   7320
      TabIndex        =   7
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":239A
      Style           =   9
      Caption         =   "html"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   2
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   6
      Left            =   6480
      TabIndex        =   6
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":33EC
      Style           =   8
      Caption         =   "Folder"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   1
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   5
      Left            =   5640
      TabIndex        =   5
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":443E
      Style           =   7
      Caption         =   "Drive"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   0
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   3
      Left            =   3960
      TabIndex        =   3
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":5490
      Style           =   5
      Caption         =   "Doc"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   2
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   8
      Left            =   8160
      TabIndex        =   8
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":64E2
      Style           =   10
      Caption         =   "Help"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   3
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   570
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   1170
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      Icon            =   "frmButtonTest.frx":7534
      Caption         =   "Log"
      IconAlign       =   3
      CaptionAlign    =   4
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   0
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdArray 
      Height          =   615
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Icon            =   "frmButtonTest.frx":8586
      Style           =   9
      Caption         =   "isButton"
      IconAlign       =   1
      USeCustomColors =   -1  'True
      BackColor       =   12632256
      HighlightColor  =   16777215
      FontHighlightColor=   -2147483635
      Tooltiptitle    =   "isButton Tooltip"
      ToolTipIcon     =   2
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   533
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":95D8
      Style           =   1
      Caption         =   "Soft"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   946
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":95F4
      Style           =   2
      Caption         =   "Flat"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1359
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9610
      Caption         =   "Java"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1772
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":962C
      Style           =   4
      Caption         =   "Office XP"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   2185
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9648
      Style           =   5
      Caption         =   "Windows XP"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   2598
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9664
      Style           =   5
      Caption         =   "Windows Themed"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   3011
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9680
      Style           =   7
      Caption         =   "Plastik"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   3424
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":969C
      Style           =   8
      Caption         =   "Galaxy"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":96B8
      Style           =   9
      Caption         =   "Keramik"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdStyle 
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   19
      Top             =   4260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":96D4
      Style           =   10
      Caption         =   "Mac OS X"
      IconAlign       =   1
      BackColor       =   15790320
      HighlightColor  =   8812135
      Tooltiptitle    =   "Warning"
      ToolTipIcon     =   1
      ToolTipType     =   -2147483624
      ttBackColor     =   -2147483624
      ttForeColor     =   -2147483625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pProps 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   2
      Left            =   2280
      ScaleHeight     =   1545
      ScaleWidth      =   6465
      TabIndex        =   43
      Top             =   2880
      Width           =   6495
      Begin VB.ComboBox comTTStyle 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox comTTIcon 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtToolTipText 
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Text            =   "ToolTip Text"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtTTTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Text            =   "isButton Tooltip"
         Top             =   360
         Width           =   1815
      End
      Begin isButtonTest.isButton cmdTTForeColor 
         Height          =   300
         Left            =   4200
         TabIndex        =   44
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Icon            =   "frmButtonTest.frx":96F0
         Style           =   8
         Caption         =   "Text Color"
         IconAlign       =   1
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin isButtonTest.isButton cmdTTBackColor 
         Height          =   300
         Left            =   4200
         TabIndex        =   45
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Icon            =   "frmButtonTest.frx":970C
         Style           =   8
         Caption         =   "BackColor"
         IconAlign       =   1
         iNonThemeStyle  =   0
         BackColor       =   15790320
         HighlightColor  =   8812135
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpTTBackColor 
         BackColor       =   &H80000018&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   5760
         Shape           =   3  'Circle
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "Tooltip Style:"
         Height          =   255
         Left            =   2160
         TabIndex        =   54
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Icon:"
         Height          =   255
         Left            =   2160
         TabIndex        =   51
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Title 
         BackColor       =   &H80000005&
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   1815
      End
      Begin VB.Shape shpTTForeColor 
         BackColor       =   &H80000017&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   5760
         Shape           =   3  'Circle
         Top             =   960
         Width           =   255
      End
   End
   Begin isButtonTest.isButton cmdTab 
      Height          =   375
      Index           =   2
      Left            =   5550
      TabIndex        =   48
      Top             =   4380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9728
      Style           =   8
      Caption         =   "ToolTip"
      IconAlign       =   1
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   16777215
      HighlightColor  =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdTab 
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   41
      Top             =   4380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9744
      Style           =   8
      Caption         =   "General"
      IconAlign       =   1
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   16777215
      HighlightColor  =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdTab 
      Height          =   375
      Index           =   1
      Left            =   3975
      TabIndex        =   42
      Top             =   4380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9760
      Style           =   8
      Caption         =   "Colors && Fonts"
      IconAlign       =   1
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   16777215
      HighlightColor  =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdAbout 
      Height          =   495
      Left            =   7320
      TabIndex        =   55
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "frmButtonTest.frx":977C
      Style           =   8
      Caption         =   "About"
      IconAlign       =   1
      CaptionAlign    =   3
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   16777215
      HighlightColor  =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdTab 
      Height          =   375
      Index           =   3
      Left            =   7125
      TabIndex        =   56
      Top             =   4380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmButtonTest.frx":9798
      Style           =   8
      Caption         =   "Help"
      IconAlign       =   1
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   16777215
      HighlightColor  =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000005&
      Caption         =   "Style Description"
      Height          =   615
      Left            =   5400
      TabIndex        =   24
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.0"
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "By Fred.cpp"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   720
      Width           =   2415
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   144
      X2              =   560
      Y1              =   172
      Y2              =   172
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   144
      X2              =   592
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   21
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   144
      X2              =   592
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   144
      X2              =   144
      Y1              =   8
      Y2              =   312
   End
   Begin VB.Label lblProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   375
      Index           =   1
      Left            =   2295
      TabIndex        =   25
      Top             =   2055
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "isButton - Demo App"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   26
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "isButton - Demo App"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   495
      Index           =   0
      Left            =   2535
      TabIndex        =   20
      Top             =   135
      Width           =   5415
   End
End
Attribute VB_Name = "frmButtonTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'
' Project Name: isButton Test
'
' Author:       Fred.cpp
'               fred_cpp@msn.com
'
' Page:         http://mx.geocities.com/fred_cpp/
'
'
' Description:  This form and Exe File is created to show some
'               features of the Multy Style Command Button isButton
'               I hope you like It And If you find It useful
'               please vote and leave comments and suggestions,
'               Everything Is wellcome.
'               Best Regards.
Option Explicit

Dim strStyleDescription(10) As String
Const strLinkAndUpdates As String = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=isbutton&lngWId=1&B1=Quick+Search"


Private Function GetColor() As OLE_COLOR
    On Error GoTo Cancelled
    dlgColor.CancelError = True
    dlgColor.ShowColor
    GetColor = dlgColor.Color
Exit Function
Cancelled:
    GetColor = -1
End Function

Private Sub chkEnabled_Click()
    cmdArray(0).Enabled = chkEnabled.Value
End Sub

Private Sub chkUseCustomColors_Click()
    cmdArray(0).UseCustomColors = chkUseCustomColors.Value
End Sub

Private Sub cmdAbout_Click()
    cmdAbout.About
    cmdAbout.About
End Sub

Private Sub cmdArray_Click(Index As Integer)
    Dim ni As Integer
    If Index Then Exit Sub
    For ni = 1 To 8
        cmdArray(ni).Style = ni + 2
    Next ni
End Sub

Private Sub cmdBackColor_Click()
    Dim lcolor As Long
    lcolor = GetColor
    If lcolor <> -1 Then
        shpBack.BackColor = lcolor
        cmdArray(0).BackColor = lcolor
    End If
End Sub

Private Sub cmdButtonType_Click()
    If cmdButtonType.Value Then
        cmdArray(0).ButtonType = isbCheckBox
    Else
        cmdArray(0).ButtonType = isbButton
    End If
End Sub

Private Sub cmdFontColor_Click()
    Dim lcolor As Long
    lcolor = GetColor
    If lcolor <> -1 Then
        Me.shpFont.BackColor = lcolor
        cmdArray(0).FontColor = lcolor
    End If
End Sub

Private Sub cmdFontHighLight_Click()
    Dim lcolor As Long
    lcolor = GetColor
    If lcolor <> -1 Then
        shpFontHighlight.BackColor = lcolor
        cmdArray(0).FontHighlightColor = lcolor
    End If
End Sub

Private Sub cmdHighLighColor_Click()
    Dim lcolor As Long
    lcolor = GetColor
    If lcolor <> -1 Then
        shpHighlight.BackColor = lcolor
        cmdArray(0).HighlightColor = lcolor
    End If
End Sub

Private Sub cmdSetFont_Click()
    On Error GoTo Cancelled
    With dlgColor
        .FontBold = cmdArray(0).Font.Bold
        .FontItalic = cmdArray(0).Font.Italic
        .FontName = cmdArray(0).Font.Name
        .FontSize = cmdArray(0).Font.SIZE
        .FontStrikethru = cmdArray(0).Font.Strikethrough
        .FontUnderline = cmdArray(0).Font.Underline
        .Flags = cdlCFScreenFonts
        .ShowFont
        cmdArray(0).Font.Bold = .FontBold
        cmdArray(0).Font.Italic = .FontItalic
        cmdArray(0).Font.Name = .FontName
        cmdArray(0).Font.SIZE = .FontSize
        cmdArray(0).Font.Strikethrough = .FontStrikethru
        cmdArray(0).Font.Underline = .FontUnderline
    End With
    cmdArray(0).Refresh
Exit Sub
Cancelled:
    MsgBox "UPS!"
End Sub

Private Sub cmdStyle_Click(Index As Integer)
    Dim ni As Long
    For ni = 0 To cmdArray.Count - 1
        cmdArray(ni).Style = Index
    Next ni
    lblDescription.Caption = strStyleDescription(Index)
End Sub

Private Sub cmdTab_Click(Index As Integer)
    Dim ni As Integer
    For ni = 0 To 3
        pProps(ni).Visible = False
    Next ni
    pProps(Index).Visible = True
End Sub

Private Sub cmdTTBackColor_Click()
    Dim lcolor As Long
    lcolor = GetColor
    If lcolor <> -1 Then
        shpTTBackColor.BackColor = lcolor
        cmdArray(0).ToolTipBackColor = lcolor
    End If
End Sub

Private Sub cmdTTForeColor_Click()
    Dim lcolor As Long
    lcolor = GetColor
    If lcolor <> -1 Then
        shpTTForeColor.BackColor = lcolor
        cmdArray(0).ToolTipForeColor = lcolor
    End If
End Sub

Private Sub cmdVote_Click()
    cmdVote.OpenLink strLinkAndUpdates
End Sub

Private Sub comCaptionAlign_Click()
    cmdArray(0).CaptionAlign = comCaptionAlign.ListIndex
End Sub

Private Sub comIconAlign_Click()
    cmdArray(0).IconAlign = comIconAlign.ListIndex
End Sub

Private Sub comNonThemeStyle_Change()
    Me.cmdArray(0).NonThemeStyle = comNonThemeStyle.ListIndex
End Sub

Private Sub comNonThemeStyle_Click()
    comNonThemeStyle_Change
End Sub

Private Sub comTTIcon_Change()
    cmdArray(0).ToolTipIcon = comTTIcon.ListIndex
End Sub

Private Sub comTTIcon_Click()
    comTTIcon_Change
End Sub

Private Sub comTTStyle_Change()
    cmdArray(0).ToolTipType = comTTStyle.ListIndex
End Sub

Private Sub comTTStyle_Click()
    comTTStyle_Change
End Sub

Private Sub Form_Load()
    Dim ni As Long
    'add some elements to the app demo.
    'Caption Align Options
    comCaptionAlign.AddItem "Center", 0
    comCaptionAlign.AddItem "Left", 1
    comCaptionAlign.AddItem "Right", 2
    comCaptionAlign.AddItem "Top", 3
    comCaptionAlign.AddItem "Bottom", 4
    'Text align Options
    comIconAlign.AddItem "Center", 0
    comIconAlign.AddItem "Left", 1
    comIconAlign.AddItem "Right", 2
    comIconAlign.AddItem "Top", 3
    comIconAlign.AddItem "Bottom", 4
    'Style Options
    comNonThemeStyle.AddItem "Normal", 0
    comNonThemeStyle.AddItem "Soft", 1
    comNonThemeStyle.AddItem "Flat", 2
    comNonThemeStyle.AddItem "Java", 3
    comNonThemeStyle.AddItem "[Office XP]", 4
    comNonThemeStyle.AddItem "[Windows XP]", 5
    comNonThemeStyle.AddItem "[Windows Theme]", 6
    comNonThemeStyle.AddItem "Plastik", 7
    comNonThemeStyle.AddItem "Galaxy", 8
    comNonThemeStyle.AddItem "Keramik", 9
    comNonThemeStyle.AddItem "[Mac OSX]", 10
    'Style Description
    strStyleDescription(0) = "Classic Win9X/ME Button Style, Default VB Style"
    strStyleDescription(1) = "Soft Style, I've seen this somewhere..."
    strStyleDescription(2) = "Flat Style, like the Win Me Toolbars"
    strStyleDescription(3) = "Sun Java Style, also uses system colors"
    strStyleDescription(4) = "MSOffice XP, Include the shadows,uses system colors"
    strStyleDescription(5) = "Windows XP, I think this is the best emulation for XP Luna"
    strStyleDescription(6) = "Windows Themed, use the current Installed Theme, or a default style can be set"
    strStyleDescription(7) = "Plastik, the Style of a very popular Linux Look and Feel for KDE"
    strStyleDescription(8) = "Mandrake Galaxy, the default style for Mandrake Linux 10.0"
    strStyleDescription(9) = "Keramik, the default Style for KDE 3.2"
    strStyleDescription(10) = "My favorite one, Mac OSX Style, " & vbCrLf & "the first one mimic drawn by code"
    For ni = 1 To 8
        cmdArray(ni).ToolTipText = strStyleDescription(ni + 2)
    Next ni
    Dim x As ttStyleEnum
    'Add tooltip icon options
    comTTIcon.AddItem "TTNoIcon", TTNoIcon
    comTTIcon.AddItem "TTIconInfo", TTIconInfo
    comTTIcon.AddItem "TTIconWarning", TTIconWarning
    comTTIcon.AddItem "TTIconError", TTIconError
    'add tooltip icon types
    comTTStyle.AddItem "TTStandard", TTStandard
    comTTStyle.AddItem "TTBalloon", TTBalloon
    
    lblVersion.Caption = "Version " & cmdArray(0).Version
    txtToolTipText.Text = "Version " & cmdArray(0).Version
    ''Setup properties for the demo button
    cmdArray(0).ToolTipText = "isButton " & lblVersion.Caption
    cmdArray(0).ToolTipType = TTBalloon
    cmdArray(0).BackColor = shpBack.BackColor
    cmdArray(0).HighlightColor = shpHighlight.BackColor
    cmdArray(0).FontColor = shpFont.BackColor
    cmdArray(0).FontHighlightColor = shpFontHighlight.BackColor
End Sub

Private Sub txtCaption_Change()
    cmdArray(0).Caption = txtCaption.Text
End Sub

Private Sub txtToolTipText_Change()
    cmdArray(0).ToolTip = txtToolTipText.Text
End Sub

Private Sub txtTTTitle_Change()
    cmdArray(0).ToolTipTitle = txtTTTitle.Text
End Sub
