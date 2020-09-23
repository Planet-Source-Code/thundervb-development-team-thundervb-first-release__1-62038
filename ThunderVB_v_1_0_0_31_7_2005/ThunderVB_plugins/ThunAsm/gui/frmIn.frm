VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
Begin VB.Form frmIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunLink"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   3480
      ScaleHeight     =   6015
      ScaleWidth      =   5895
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      Begin ThunVBCC_v1.XTab xTabSet 
         Height          =   5535
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   9763
         TabCaption(0)   =   "Inline Asm/C"
         TabContCtrlCnt(0)=   5
         Tab(0)ContCtrlCap(1)=   "lblAsm"
         Tab(0)ContCtrlCap(2)=   "set_ASM_CompileAsmCode"
         Tab(0)ContCtrlCap(3)=   "set_ASM_FixAsmListings"
         Tab(0)ContCtrlCap(4)=   "set_C_CompileCCode"
         Tab(0)ContCtrlCap(5)=   "lblC"
         TabCaption(1)   =   "Paths"
         TabContCtrlCnt(1)=   14
         Tab(1)ContCtrlCap(1)=   "cmdPaths_CCompiler"
         Tab(1)ContCtrlCap(2)=   "cmdPaths_ML"
         Tab(1)ContCtrlCap(3)=   "lblPath2CComp"
         Tab(1)ContCtrlCap(4)=   "lblPath2Masm"
         Tab(1)ContCtrlCap(5)=   "set_Paths_txtMasm"
         Tab(1)ContCtrlCap(6)=   "set_Paths_txtCCompiler"
         Tab(1)ContCtrlCap(7)=   "set_Paths_txtIncFiles"
         Tab(1)ContCtrlCap(8)=   "set_Paths_txtLibFiles"
         Tab(1)ContCtrlCap(9)=   "lblPath2LibFiles"
         Tab(1)ContCtrlCap(10)=   "lblPath2IncFiles"
         Tab(1)ContCtrlCap(11)=   "cmdPaths_LIBFiles"
         Tab(1)ContCtrlCap(12)=   "cmdINCFiles"
         Tab(1)ContCtrlCap(13)=   "lblAsm1"
         Tab(1)ContCtrlCap(14)=   "lblC1"
         TabCaption(2)   =   "Misc"
         TabContCtrlCnt(2)=   5
         Tab(2)ContCtrlCap(1)=   "set_Gen_chbGenAsmCHeaders"
         Tab(2)ContCtrlCap(2)=   "set_Compile_PauseAsm"
         Tab(2)ContCtrlCap(3)=   "set_Compile_PauseLink"
         Tab(2)ContCtrlCap(4)=   "set_Compile_ModifyCmdLine"
         Tab(2)ContCtrlCap(5)=   "set_Compile_SkipLinking"
         ActiveTab       =   1
         TabStyle        =   1
         TabTheme        =   2
         InActiveTabBackStartColor=   -2147483626
         InActiveTabForeColor=   -2147483631
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   -2147483638
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin ThunVBCC_v1.HzxYCheckBox set_Gen_chbGenAsmCHeaders 
            Height          =   240
            Left            =   -74760
            TabIndex        =   21
            Tag             =   "def"
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "Generate Asm/C headers"
            Pic_UncheckedNormal=   "frmIn.frx":0000
            Pic_CheckedNormal=   "frmIn.frx":0352
            Pic_MixedNormal =   "frmIn.frx":06A4
            Pic_UncheckedDisabled=   "frmIn.frx":09F6
            Pic_CheckedDisabled=   "frmIn.frx":0D48
            Pic_MixedDisabled=   "frmIn.frx":109A
            Pic_UncheckedOver=   "frmIn.frx":13EC
            Pic_CheckedOver =   "frmIn.frx":173E
            Pic_MixedOver   =   "frmIn.frx":1A90
            Pic_UncheckedDown=   "frmIn.frx":1DE2
            Pic_CheckedDown =   "frmIn.frx":2134
            Pic_MixedDown   =   "frmIn.frx":2486
         End
         Begin ThunVBCC_v1.UniLabel lblAsm 
            Height          =   240
            Left            =   -74760
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   423
            CaptionB        =   "frmIn.frx":27D8
            CaptionLen      =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.isButton cmdPaths_CCompiler 
            Height          =   330
            Left            =   4440
            TabIndex        =   15
            Top             =   2400
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":280C
            Style           =   5
            Caption         =   "..."
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
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
         Begin ThunVBCC_v1.isButton cmdPaths_ML 
            Height          =   330
            Left            =   4440
            TabIndex        =   14
            Top             =   1080
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":2828
            Style           =   5
            Caption         =   "..."
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
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
         Begin ThunVBCC_v1.UniLabel lblPath2CComp 
            Height          =   225
            Left            =   240
            Top             =   2160
            Width           =   1320
            _ExtentX        =   159
            _ExtentY        =   26
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":2844
            CaptionLen      =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.UniLabel lblPath2Masm 
            Height          =   225
            Left            =   240
            Top             =   840
            Width           =   1620
            _ExtentX        =   185
            _ExtentY        =   26
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":2888
            CaptionLen      =   21
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox set_Paths_txtMasm 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   4095
         End
         Begin VB.TextBox set_Paths_txtCCompiler 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   12
            Top             =   2400
            Width           =   4095
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_PauseAsm 
            Height          =   240
            Left            =   -74760
            TabIndex        =   9
            Top             =   1920
            Visible         =   0   'False
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "Pause before assembly"
            Pic_UncheckedNormal=   "frmIn.frx":28D2
            Pic_CheckedNormal=   "frmIn.frx":2C24
            Pic_MixedNormal =   "frmIn.frx":2F76
            Pic_UncheckedDisabled=   "frmIn.frx":32C8
            Pic_CheckedDisabled=   "frmIn.frx":361A
            Pic_MixedDisabled=   "frmIn.frx":396C
            Pic_UncheckedOver=   "frmIn.frx":3CBE
            Pic_CheckedOver =   "frmIn.frx":4010
            Pic_MixedOver   =   "frmIn.frx":4362
            Pic_UncheckedDown=   "frmIn.frx":46B4
            Pic_CheckedDown =   "frmIn.frx":4A06
            Pic_MixedDown   =   "frmIn.frx":4D58
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_PauseLink 
            Height          =   240
            Left            =   -74760
            TabIndex        =   8
            Top             =   2400
            Visible         =   0   'False
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "Pause before linking"
            Pic_UncheckedNormal=   "frmIn.frx":50AA
            Pic_CheckedNormal=   "frmIn.frx":53FC
            Pic_MixedNormal =   "frmIn.frx":574E
            Pic_UncheckedDisabled=   "frmIn.frx":5AA0
            Pic_CheckedDisabled=   "frmIn.frx":5DF2
            Pic_MixedDisabled=   "frmIn.frx":6144
            Pic_UncheckedOver=   "frmIn.frx":6496
            Pic_CheckedOver =   "frmIn.frx":67E8
            Pic_MixedOver   =   "frmIn.frx":6B3A
            Pic_UncheckedDown=   "frmIn.frx":6E8C
            Pic_CheckedDown =   "frmIn.frx":71DE
            Pic_MixedDown   =   "frmIn.frx":7530
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_ModifyCmdLine 
            Height          =   240
            Left            =   -74760
            TabIndex        =   7
            Top             =   600
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "Modify CmdLine"
            Pic_UncheckedNormal=   "frmIn.frx":7882
            Pic_CheckedNormal=   "frmIn.frx":7BD4
            Pic_MixedNormal =   "frmIn.frx":7F26
            Pic_UncheckedDisabled=   "frmIn.frx":8278
            Pic_CheckedDisabled=   "frmIn.frx":85CA
            Pic_MixedDisabled=   "frmIn.frx":891C
            Pic_UncheckedOver=   "frmIn.frx":8C6E
            Pic_CheckedOver =   "frmIn.frx":8FC0
            Pic_MixedOver   =   "frmIn.frx":9312
            Pic_UncheckedDown=   "frmIn.frx":9664
            Pic_CheckedDown =   "frmIn.frx":99B6
            Pic_MixedDown   =   "frmIn.frx":9D08
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_SkipLinking 
            Height          =   240
            Left            =   -74760
            TabIndex        =   6
            Top             =   3120
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "Skip linking"
            Pic_UncheckedNormal=   "frmIn.frx":A05A
            Pic_CheckedNormal=   "frmIn.frx":A3AC
            Pic_MixedNormal =   "frmIn.frx":A6FE
            Pic_UncheckedDisabled=   "frmIn.frx":AA50
            Pic_CheckedDisabled=   "frmIn.frx":ADA2
            Pic_MixedDisabled=   "frmIn.frx":B0F4
            Pic_UncheckedOver=   "frmIn.frx":B446
            Pic_CheckedOver =   "frmIn.frx":B798
            Pic_MixedOver   =   "frmIn.frx":BAEA
            Pic_UncheckedDown=   "frmIn.frx":BE3C
            Pic_CheckedDown =   "frmIn.frx":C18E
            Pic_MixedDown   =   "frmIn.frx":C4E0
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_ASM_CompileAsmCode 
            Height          =   240
            Left            =   -74640
            TabIndex        =   5
            Tag             =   "def"
            Top             =   840
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "Compile ASM code"
            Pic_UncheckedNormal=   "frmIn.frx":C832
            Pic_CheckedNormal=   "frmIn.frx":CB84
            Pic_MixedNormal =   "frmIn.frx":CED6
            Pic_UncheckedDisabled=   "frmIn.frx":D228
            Pic_CheckedDisabled=   "frmIn.frx":D57A
            Pic_MixedDisabled=   "frmIn.frx":D8CC
            Pic_UncheckedOver=   "frmIn.frx":DC1E
            Pic_CheckedOver =   "frmIn.frx":DF70
            Pic_MixedOver   =   "frmIn.frx":E2C2
            Pic_UncheckedDown=   "frmIn.frx":E614
            Pic_CheckedDown =   "frmIn.frx":E966
            Pic_MixedDown   =   "frmIn.frx":ECB8
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_ASM_FixAsmListings 
            Height          =   240
            Left            =   -74640
            TabIndex        =   4
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* Fix ASM listings"
            Pic_UncheckedNormal=   "frmIn.frx":F00A
            Pic_CheckedNormal=   "frmIn.frx":F35C
            Pic_MixedNormal =   "frmIn.frx":F6AE
            Pic_UncheckedDisabled=   "frmIn.frx":FA00
            Pic_CheckedDisabled=   "frmIn.frx":FD52
            Pic_MixedDisabled=   "frmIn.frx":100A4
            Pic_UncheckedOver=   "frmIn.frx":103F6
            Pic_CheckedOver =   "frmIn.frx":10748
            Pic_MixedOver   =   "frmIn.frx":10A9A
            Pic_UncheckedDown=   "frmIn.frx":10DEC
            Pic_CheckedDown =   "frmIn.frx":1113E
            Pic_MixedDown   =   "frmIn.frx":11490
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_C_CompileCCode 
            Height          =   240
            Left            =   -74640
            TabIndex        =   3
            Tag             =   "def"
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "Compile C code"
            Pic_UncheckedNormal=   "frmIn.frx":117E2
            Pic_CheckedNormal=   "frmIn.frx":11B34
            Pic_MixedNormal =   "frmIn.frx":11E86
            Pic_UncheckedDisabled=   "frmIn.frx":121D8
            Pic_CheckedDisabled=   "frmIn.frx":1252A
            Pic_MixedDisabled=   "frmIn.frx":1287C
            Pic_UncheckedOver=   "frmIn.frx":12BCE
            Pic_CheckedOver =   "frmIn.frx":12F20
            Pic_MixedOver   =   "frmIn.frx":13272
            Pic_UncheckedDown=   "frmIn.frx":135C4
            Pic_CheckedDown =   "frmIn.frx":13916
            Pic_MixedDown   =   "frmIn.frx":13C68
         End
         Begin VB.TextBox set_Paths_txtIncFiles 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            TabIndex        =   10
            Top             =   7080
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.TextBox set_Paths_txtLibFiles 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            TabIndex        =   11
            Top             =   6000
            Visible         =   0   'False
            Width           =   4095
         End
         Begin ThunVBCC_v1.UniLabel lblPath2LibFiles 
            Height          =   225
            Left            =   2400
            Top             =   6240
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   132
            _ExtentY        =   26
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":13FBA
            CaptionLen      =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.UniLabel lblPath2IncFiles 
            Height          =   225
            Left            =   2400
            Top             =   6840
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   132
            _ExtentY        =   26
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":13FFE
            CaptionLen      =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.isButton cmdPaths_LIBFiles 
            Height          =   330
            Left            =   6600
            TabIndex        =   16
            Top             =   6480
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":14042
            Style           =   5
            Caption         =   "..."
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
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
         Begin ThunVBCC_v1.isButton cmdINCFiles 
            Height          =   330
            Left            =   6600
            TabIndex        =   17
            Top             =   7080
            Visible         =   0   'False
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":1405E
            Style           =   5
            Caption         =   "..."
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
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
         Begin VB.Label lblAsm1 
            Caption         =   "Inline Asm"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblC1 
            Caption         =   "Inline C"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label lblC 
            Caption         =   "Inline C"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74760
            TabIndex        =   18
            Top             =   1800
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox pctCredits 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   -120
      ScaleHeight     =   3855
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin ThunVBCC_v1.UniLabel lblCredits 
         Height          =   975
         Left            =   240
         Top             =   1080
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1720
         CaptionB        =   "frmIn.frx":1407A
         CaptionLen      =   17
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   18
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog cdSet 
      Left            =   840
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim X As Long, Y As Long
      
    GetSettingsTabClientRect X, Y
    
    X = X * Screen.TwipsPerPixelX
    Y = Y * Screen.TwipsPerPixelY
    
    pctSettings.Move 0, 0, X, Y
    pctCredits.Move 0, 0, X, Y
    
    'center pictureboxes in the tab
    '------------------------------
    'xTabSet.Move (X - xTabSet.Width) \ 2, (Y - xTabSet.Height) \ 2
    'lblCredits.Move (X - lblCredits.Width) \ 2, (Y - lblCredits.Height) \ 2
    
    xTabSet.Move 0, 0, X, Y
    lblCredits.Move 0, 0, X, Y
    
    xTabSet.ActiveTab = 0
      
    With cdSet
      .Filter = "executable (*.exe)|*.exe|all (*.*)|*.*"
       .FileName = ""
       .InitDir = App.Path
       .flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware
       .CancelError = True
    End With
      
End Sub

'-----------------
'--- Tab Paths ---
'-----------------

Private Sub cmdPaths_ML_Click()
    Call SetPath("Path to Masm (" & Add34("ml.exe") & ")", set_Paths_txtMasm, "ml.exe")
End Sub

Private Sub cmdPaths_CCompiler_Click()
    Call SetPath("Path to C compiler (" & Add34("cl.exe") & ")", set_Paths_txtCCompiler, "cl.exe")
End Sub

Private Sub cmdPaths_LIBFiles_Click()
    Call SetDirectory("Select .LIB directory", set_Paths_txtLibFiles)
End Sub

Private Sub cmdINCFiles_Click()
    Call SetDirectory("Select .INC directory", set_Paths_txtIncFiles)
End Sub

'append "\"
Private Sub set_Paths_txtIncFiles_Validate(Cancel As Boolean)
    If Right(set_Paths_txtIncFiles.Text, 1) <> "\" Then set_Paths_txtIncFiles.Text = set_Paths_txtIncFiles.Text & "\"
End Sub

'append "\"
Private Sub set_Paths_txtLibFiles_Validate(Cancel As Boolean)
    If Right(set_Paths_txtLibFiles.Text, 1) <> "\" Then set_Paths_txtLibFiles.Text = set_Paths_txtLibFiles.Text & "\"
End Sub
