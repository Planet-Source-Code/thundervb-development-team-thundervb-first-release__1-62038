VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#24.2#0"; "ThunVBCC_v1.ocx"
Begin VB.Form frmIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunAsm"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   4200
      ScaleHeight     =   5295
      ScaleWidth      =   5055
      TabIndex        =   1
      Top             =   360
      Width           =   5055
      Begin ThunVBCC_v1.XTab xTabSet 
         Height          =   3135
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5530
         TabCaption(0)   =   "Inline Asm/C"
         TabContCtrlCnt(0)=   3
         Tab(0)ContCtrlCap(1)=   "set_ASM_CompileAsmCode"
         Tab(0)ContCtrlCap(2)=   "set_ASM_FixAsmListings"
         Tab(0)ContCtrlCap(3)=   "set_C_CompileCCode"
         TabCaption(1)   =   "Paths"
         TabContCtrlCnt(1)=   12
         Tab(1)ContCtrlCap(1)=   "cmdINCFiles"
         Tab(1)ContCtrlCap(2)=   "cmdPaths_LIBFiles"
         Tab(1)ContCtrlCap(3)=   "cmdPaths_CCompiler"
         Tab(1)ContCtrlCap(4)=   "cmdPaths_ML"
         Tab(1)ContCtrlCap(5)=   "lblPath2IncFiles"
         Tab(1)ContCtrlCap(6)=   "lblPath2LibFiles"
         Tab(1)ContCtrlCap(7)=   "lblPath2CComp"
         Tab(1)ContCtrlCap(8)=   "lblPath2Masm"
         Tab(1)ContCtrlCap(9)=   "set_Paths_txtMasm"
         Tab(1)ContCtrlCap(10)=   "set_Paths_txtCCompiler"
         Tab(1)ContCtrlCap(11)=   "set_Paths_txtLibFiles"
         Tab(1)ContCtrlCap(12)=   "set_Paths_txtIncFiles"
         TabCaption(2)   =   "Misc"
         TabContCtrlCnt(2)=   4
         Tab(2)ContCtrlCap(1)=   "set_Compile_PauseAsm"
         Tab(2)ContCtrlCap(2)=   "set_Compile_PauseLink"
         Tab(2)ContCtrlCap(3)=   "set_Compile_ModifyCmdLine"
         Tab(2)ContCtrlCap(4)=   "set_Compile_SkipLinking"
         ActiveTab       =   2
         TabStyle        =   1
         TabTheme        =   2
         InActiveTabBackStartColor=   -2147483626
         InActiveTabForeColor=   -2147483631
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Batang"
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
         Begin ThunVBCC_v1.isButton cmdINCFiles 
            Height          =   330
            Left            =   -70560
            TabIndex        =   17
            Top             =   2640
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":0000
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
         Begin ThunVBCC_v1.isButton cmdPaths_LIBFiles 
            Height          =   330
            Left            =   -70560
            TabIndex        =   16
            Top             =   2040
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":001C
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
         Begin ThunVBCC_v1.isButton cmdPaths_CCompiler 
            Height          =   330
            Left            =   -70560
            TabIndex        =   15
            Top             =   1440
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":0038
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
            Left            =   -70560
            TabIndex        =   14
            Top             =   720
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Icon            =   "frmIn.frx":0054
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
         Begin ThunVBCC_v1.UniLabel lblPath2IncFiles 
            Height          =   225
            Left            =   -74760
            Top             =   2400
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":0070
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
         Begin ThunVBCC_v1.UniLabel lblPath2LibFiles 
            Height          =   225
            Left            =   -74760
            Top             =   1800
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   397
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":00B4
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
         Begin ThunVBCC_v1.UniLabel lblPath2CComp 
            Height          =   225
            Left            =   -74760
            Top             =   1200
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   397
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":00F8
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
            Left            =   -74760
            Top             =   480
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   397
            AutoSize        =   -1  'True
            CaptionB        =   "frmIn.frx":013C
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
            Left            =   -74760
            TabIndex        =   13
            Top             =   720
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
            Left            =   -74760
            TabIndex        =   12
            Top             =   1440
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
            Left            =   -74760
            TabIndex        =   11
            Top             =   2040
            Width           =   4095
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
            Left            =   -74760
            TabIndex        =   10
            Top             =   2640
            Width           =   4095
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_PauseAsm 
            Height          =   240
            Left            =   240
            TabIndex        =   9
            Top             =   600
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
            Pic_UncheckedNormal=   "frmIn.frx":0186
            Pic_CheckedNormal=   "frmIn.frx":04D8
            Pic_MixedNormal =   "frmIn.frx":082A
            Pic_UncheckedDisabled=   "frmIn.frx":0B7C
            Pic_CheckedDisabled=   "frmIn.frx":0ECE
            Pic_MixedDisabled=   "frmIn.frx":1220
            Pic_UncheckedOver=   "frmIn.frx":1572
            Pic_CheckedOver =   "frmIn.frx":18C4
            Pic_MixedOver   =   "frmIn.frx":1C16
            Pic_UncheckedDown=   "frmIn.frx":1F68
            Pic_CheckedDown =   "frmIn.frx":22BA
            Pic_MixedDown   =   "frmIn.frx":260C
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_PauseLink 
            Height          =   240
            Left            =   240
            TabIndex        =   8
            Top             =   1080
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
            Pic_UncheckedNormal=   "frmIn.frx":295E
            Pic_CheckedNormal=   "frmIn.frx":2CB0
            Pic_MixedNormal =   "frmIn.frx":3002
            Pic_UncheckedDisabled=   "frmIn.frx":3354
            Pic_CheckedDisabled=   "frmIn.frx":36A6
            Pic_MixedDisabled=   "frmIn.frx":39F8
            Pic_UncheckedOver=   "frmIn.frx":3D4A
            Pic_CheckedOver =   "frmIn.frx":409C
            Pic_MixedOver   =   "frmIn.frx":43EE
            Pic_UncheckedDown=   "frmIn.frx":4740
            Pic_CheckedDown =   "frmIn.frx":4A92
            Pic_MixedDown   =   "frmIn.frx":4DE4
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_ModifyCmdLine 
            Height          =   240
            Left            =   240
            TabIndex        =   7
            Top             =   1560
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
            Pic_UncheckedNormal=   "frmIn.frx":5136
            Pic_CheckedNormal=   "frmIn.frx":5488
            Pic_MixedNormal =   "frmIn.frx":57DA
            Pic_UncheckedDisabled=   "frmIn.frx":5B2C
            Pic_CheckedDisabled=   "frmIn.frx":5E7E
            Pic_MixedDisabled=   "frmIn.frx":61D0
            Pic_UncheckedOver=   "frmIn.frx":6522
            Pic_CheckedOver =   "frmIn.frx":6874
            Pic_MixedOver   =   "frmIn.frx":6BC6
            Pic_UncheckedDown=   "frmIn.frx":6F18
            Pic_CheckedDown =   "frmIn.frx":726A
            Pic_MixedDown   =   "frmIn.frx":75BC
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_Compile_SkipLinking 
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   2040
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
            Pic_UncheckedNormal=   "frmIn.frx":790E
            Pic_CheckedNormal=   "frmIn.frx":7C60
            Pic_MixedNormal =   "frmIn.frx":7FB2
            Pic_UncheckedDisabled=   "frmIn.frx":8304
            Pic_CheckedDisabled=   "frmIn.frx":8656
            Pic_MixedDisabled=   "frmIn.frx":89A8
            Pic_UncheckedOver=   "frmIn.frx":8CFA
            Pic_CheckedOver =   "frmIn.frx":904C
            Pic_MixedOver   =   "frmIn.frx":939E
            Pic_UncheckedDown=   "frmIn.frx":96F0
            Pic_CheckedDown =   "frmIn.frx":9A42
            Pic_MixedDown   =   "frmIn.frx":9D94
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_ASM_CompileAsmCode 
            Height          =   240
            Left            =   -74760
            TabIndex        =   5
            Tag             =   "def"
            Top             =   600
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
            Pic_UncheckedNormal=   "frmIn.frx":A0E6
            Pic_CheckedNormal=   "frmIn.frx":A438
            Pic_MixedNormal =   "frmIn.frx":A78A
            Pic_UncheckedDisabled=   "frmIn.frx":AADC
            Pic_CheckedDisabled=   "frmIn.frx":AE2E
            Pic_MixedDisabled=   "frmIn.frx":B180
            Pic_UncheckedOver=   "frmIn.frx":B4D2
            Pic_CheckedOver =   "frmIn.frx":B824
            Pic_MixedOver   =   "frmIn.frx":BB76
            Pic_UncheckedDown=   "frmIn.frx":BEC8
            Pic_CheckedDown =   "frmIn.frx":C21A
            Pic_MixedDown   =   "frmIn.frx":C56C
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_ASM_FixAsmListings 
            Height          =   240
            Left            =   -74760
            TabIndex        =   4
            Top             =   1080
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
            Pic_UncheckedNormal=   "frmIn.frx":C8BE
            Pic_CheckedNormal=   "frmIn.frx":CC10
            Pic_MixedNormal =   "frmIn.frx":CF62
            Pic_UncheckedDisabled=   "frmIn.frx":D2B4
            Pic_CheckedDisabled=   "frmIn.frx":D606
            Pic_MixedDisabled=   "frmIn.frx":D958
            Pic_UncheckedOver=   "frmIn.frx":DCAA
            Pic_CheckedOver =   "frmIn.frx":DFFC
            Pic_MixedOver   =   "frmIn.frx":E34E
            Pic_UncheckedDown=   "frmIn.frx":E6A0
            Pic_CheckedDown =   "frmIn.frx":E9F2
            Pic_MixedDown   =   "frmIn.frx":ED44
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_C_CompileCCode 
            Height          =   240
            Left            =   -74760
            TabIndex        =   3
            Tag             =   "def"
            Top             =   1560
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
            Pic_UncheckedNormal=   "frmIn.frx":F096
            Pic_CheckedNormal=   "frmIn.frx":F3E8
            Pic_MixedNormal =   "frmIn.frx":F73A
            Pic_UncheckedDisabled=   "frmIn.frx":FA8C
            Pic_CheckedDisabled=   "frmIn.frx":FDDE
            Pic_MixedDisabled=   "frmIn.frx":10130
            Pic_UncheckedOver=   "frmIn.frx":10482
            Pic_CheckedOver =   "frmIn.frx":107D4
            Pic_MixedOver   =   "frmIn.frx":10B26
            Pic_UncheckedDown=   "frmIn.frx":10E78
            Pic_CheckedDown =   "frmIn.frx":111CA
            Pic_MixedDown   =   "frmIn.frx":1151C
         End
      End
   End
   Begin VB.PictureBox pctCredits 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   240
      ScaleHeight     =   3855
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      Begin ThunVBCC_v1.UniLabel lblCredits 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   873
         AutoSize        =   -1  'True
         CaptionB        =   "frmIn.frx":1186E
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

Private Const FORM_NAME As String = "frmIn"

Private Sub Form_Load()
          
10        Me.Caption = PLUGIN_NAMEs
20        LogMsg "Loading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, FORM_NAME, "Form_Load"
          
30        pctSettings.Move 0, 0
40        pctCredits.Move 0, 0
          
50        xTabSet.ActiveTab = 0
          
60        With cdSet
70            .Filter = "executable (*.exe)|*.exe|all (*.*)|*.*"
80            .FileName = ""
90            .InitDir = App.Path
100           .flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware
110           .CancelError = True
120       End With
          
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        LogMsg "Unloading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, FORM_NAME, "Form_Unload"
End Sub

Private Sub pctSettings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
          'set default
10        If Button = vbRightButton Then
20            SetDefaultSettings GLOBAL_, pctSettings
30            SetDefaultSettings LOCAL_, pctSettings
40        End If
          
End Sub

'-----------------
'--- Tab Paths ---
'-----------------

Private Sub cmdPaths_ML_Click()
10        Call SetPath("Path to Masm (" & Add34("ml.exe") & ")", set_Paths_txtMasm, "ml.exe")
End Sub

Private Sub cmdPaths_CCompiler_Click()
10        Call SetPath("Path to C compiler (" & Add34("cl.exe") & ")", set_Paths_txtCCompiler, "cl.exe")
End Sub

Private Sub cmdPaths_LIBFiles_Click()
10        Call SetDirectory("Select .LIB directory", set_Paths_txtLibFiles)
End Sub

Private Sub cmdINCFiles_Click()
10        Call SetDirectory("Select .INC directory", set_Paths_txtIncFiles)
End Sub

'append "\"
Private Sub set_Paths_txtIncFiles_Validate(Cancel As Boolean)
10        If Right(set_Paths_txtIncFiles.Text, 1) <> "\" Then set_Paths_txtIncFiles.Text = set_Paths_txtIncFiles.Text & "\"
End Sub

'append "\"
Private Sub set_Paths_txtLibFiles_Validate(Cancel As Boolean)
10        If Right(set_Paths_txtLibFiles.Text, 1) <> "\" Then set_Paths_txtLibFiles.Text = set_Paths_txtLibFiles.Text & "\"
End Sub
