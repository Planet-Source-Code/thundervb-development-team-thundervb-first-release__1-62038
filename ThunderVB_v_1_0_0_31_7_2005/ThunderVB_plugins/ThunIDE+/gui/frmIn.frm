VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
Begin VB.Form frmIn 
   Caption         =   "ThunIDE"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   5040
      ScaleHeight     =   4935
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      Begin ThunVBCC_v1.XTab xTabSet 
         Height          =   4695
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8281
         TabCaption(0)   =   "ASM"
         TabContCtrlCnt(0)=   9
         Tab(0)ContCtrlCap(1)=   "cdSet"
         Tab(0)ContCtrlCap(2)=   "cmdColor_DefaultASM"
         Tab(0)ContCtrlCap(3)=   "cmdColor_DeleteASM"
         Tab(0)ContCtrlCap(4)=   "cmdColor_AddASM"
         Tab(0)ContCtrlCap(5)=   "set_ASM_chbQuickWatch"
         Tab(0)ContCtrlCap(6)=   "set_ASM_chbIntelliSense"
         Tab(0)ContCtrlCap(7)=   "set_ASM_ctlColorsASM"
         Tab(0)ContCtrlCap(8)=   "set_ASM_chbAsmColoring"
         Tab(0)ContCtrlCap(9)=   "lblAsmPreview"
         TabCaption(1)   =   "C"
         TabContCtrlCnt(1)=   6
         Tab(1)ContCtrlCap(1)=   "set_C_ctlColorsC"
         Tab(1)ContCtrlCap(2)=   "cmdColor_DefaultC"
         Tab(1)ContCtrlCap(3)=   "cmdColor_DeleteC"
         Tab(1)ContCtrlCap(4)=   "cmdColor_AddC"
         Tab(1)ContCtrlCap(5)=   "set_C_chbCColoring"
         Tab(1)ContCtrlCap(6)=   "lblCPreview"
         TabCaption(2)   =   "Misc"
         TabContCtrlCnt(2)=   1
         Tab(2)ContCtrlCap(1)=   "set_Misc_chbCopyTimeColor"
         TabTheme        =   2
         InActiveTabBackStartColor=   -2147483626
         InActiveTabBackEndColor=   -2147483626
         InActiveTabForeColor=   -2147483631
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   -2147483628
         TabStripBackColor=   -2147483626
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin ThunVBCC_v1.HzxYCheckBox set_Misc_chbCopyTimeColor 
            Height          =   240
            Left            =   -74760
            TabIndex        =   15
            Tag             =   "def"
            Top             =   480
            Width           =   2205
            _ExtentX        =   3889
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
            Caption         =   "Enable CopyTime coloring"
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
         Begin ThunIDE.ColorList set_C_ctlColorsC 
            Height          =   2055
            Left            =   -74760
            TabIndex        =   14
            Top             =   840
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
         End
         Begin ThunVBCC_v1.isButton cmdColor_DefaultC 
            Height          =   300
            Left            =   -71880
            TabIndex        =   13
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            Icon            =   "frmIn.frx":27D8
            Style           =   5
            Caption         =   "Default"
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.isButton cmdColor_DeleteC 
            Height          =   300
            Left            =   -73800
            TabIndex        =   12
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            Icon            =   "frmIn.frx":27F4
            Style           =   5
            Caption         =   "Delete"
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.isButton cmdColor_AddC 
            Height          =   300
            Left            =   -74760
            TabIndex        =   11
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            Icon            =   "frmIn.frx":2810
            Style           =   5
            Caption         =   "Add"
            IconAlign       =   1
            iNonThemeStyle  =   4
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_C_chbCColoring 
            Height          =   240
            Left            =   -74760
            TabIndex        =   10
            Tag             =   "def"
            Top             =   480
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* C code coloring"
            Pic_UncheckedNormal=   "frmIn.frx":282C
            Pic_CheckedNormal=   "frmIn.frx":2B7E
            Pic_MixedNormal =   "frmIn.frx":2ED0
            Pic_UncheckedDisabled=   "frmIn.frx":3222
            Pic_CheckedDisabled=   "frmIn.frx":3574
            Pic_MixedDisabled=   "frmIn.frx":38C6
            Pic_UncheckedOver=   "frmIn.frx":3C18
            Pic_CheckedOver =   "frmIn.frx":3F6A
            Pic_MixedOver   =   "frmIn.frx":42BC
            Pic_UncheckedDown=   "frmIn.frx":460E
            Pic_CheckedDown =   "frmIn.frx":4960
            Pic_MixedDown   =   "frmIn.frx":4CB2
         End
         Begin MSComDlg.CommonDialog cdSet 
            Left            =   3720
            Top             =   3840
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ThunVBCC_v1.isButton cmdColor_DefaultASM 
            Height          =   300
            Left            =   3120
            TabIndex        =   9
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            Icon            =   "frmIn.frx":5004
            Style           =   5
            Caption         =   "Default"
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.isButton cmdColor_DeleteASM 
            Height          =   300
            Left            =   1200
            TabIndex        =   8
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            Icon            =   "frmIn.frx":5020
            Style           =   5
            Caption         =   "Delete"
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.isButton cmdColor_AddASM 
            Height          =   300
            Left            =   240
            TabIndex        =   7
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            Icon            =   "frmIn.frx":503C
            Style           =   5
            Caption         =   "Add"
            IconAlign       =   1
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_ASM_chbQuickWatch 
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   4200
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
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
            Caption         =   "Quick - Watch"
            Pic_UncheckedNormal=   "frmIn.frx":5058
            Pic_CheckedNormal=   "frmIn.frx":53AA
            Pic_MixedNormal =   "frmIn.frx":56FC
            Pic_UncheckedDisabled=   "frmIn.frx":5A4E
            Pic_CheckedDisabled=   "frmIn.frx":5DA0
            Pic_MixedDisabled=   "frmIn.frx":60F2
            Pic_UncheckedOver=   "frmIn.frx":6444
            Pic_CheckedOver =   "frmIn.frx":6796
            Pic_MixedOver   =   "frmIn.frx":6AE8
            Pic_UncheckedDown=   "frmIn.frx":6E3A
            Pic_CheckedDown =   "frmIn.frx":718C
            Pic_MixedDown   =   "frmIn.frx":74DE
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_ASM_chbIntelliSense 
            Height          =   240
            Left            =   240
            TabIndex        =   5
            Tag             =   "def"
            Top             =   3840
            Width           =   1290
            _ExtentX        =   2275
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
            Caption         =   "Intelli - Sense"
            Pic_UncheckedNormal=   "frmIn.frx":7830
            Pic_CheckedNormal=   "frmIn.frx":7B82
            Pic_MixedNormal =   "frmIn.frx":7ED4
            Pic_UncheckedDisabled=   "frmIn.frx":8226
            Pic_CheckedDisabled=   "frmIn.frx":8578
            Pic_MixedDisabled=   "frmIn.frx":88CA
            Pic_UncheckedOver=   "frmIn.frx":8C1C
            Pic_CheckedOver =   "frmIn.frx":8F6E
            Pic_MixedOver   =   "frmIn.frx":92C0
            Pic_UncheckedDown=   "frmIn.frx":9612
            Pic_CheckedDown =   "frmIn.frx":9964
            Pic_MixedDown   =   "frmIn.frx":9CB6
         End
         Begin ThunIDE.ColorList set_ASM_ctlColorsASM 
            Height          =   2055
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_ASM_chbAsmColoring 
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Tag             =   "def"
            Top             =   480
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
            Caption         =   "* ASM code coloring"
            Pic_UncheckedNormal=   "frmIn.frx":A008
            Pic_CheckedNormal=   "frmIn.frx":A35A
            Pic_MixedNormal =   "frmIn.frx":A6AC
            Pic_UncheckedDisabled=   "frmIn.frx":A9FE
            Pic_CheckedDisabled=   "frmIn.frx":AD50
            Pic_MixedDisabled=   "frmIn.frx":B0A2
            Pic_UncheckedOver=   "frmIn.frx":B3F4
            Pic_CheckedOver =   "frmIn.frx":B746
            Pic_MixedOver   =   "frmIn.frx":BA98
            Pic_UncheckedDown=   "frmIn.frx":BDEA
            Pic_CheckedDown =   "frmIn.frx":C13C
            Pic_MixedDown   =   "frmIn.frx":C48E
         End
         Begin VB.Label lblAsmPreview 
            Caption         =   "'#asm' mov eax , 12345 ; an eaxmple"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label lblCPreview 
            Caption         =   "'#c' int i = 0x80 // a sample line"
            Height          =   255
            Left            =   -74760
            TabIndex        =   16
            Top             =   3480
            Width           =   2775
         End
      End
   End
   Begin VB.PictureBox pctCredits 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   600
      ScaleHeight     =   1335
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      Begin ThunVBCC_v1.UniLabel lblCredits 
         Height          =   375
         Left            =   480
         Top             =   480
         Width           =   2250
         _ExtentX        =   265
         _ExtentY        =   53
         CaptionB        =   "frmIn.frx":C7E0
         CaptionLen      =   17
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim x As Long, y As Long
          
    GetSettingsTabClientRect x, y
    
    x = x * Screen.TwipsPerPixelX
    y = y * Screen.TwipsPerPixelY

    pctSettings.Move 0, 0, x, y
    pctCredits.Move 0, 0, x, y

    xTabSet.Move 0, 0, x, y
    lblCredits.Move 0, 0, x, y
    
    xTabSet.ActiveTab = 0
          
End Sub

'---------------
'--- Tab ASM ---
'---------------

'set ASM default colors
Private Sub cmdColor_DefaultASM_Click()
    set_ASM_ctlColorsASM.SetDefaultsAsm
End Sub

'change ASM color
Private Sub set_ASM_ctlColorsASM_ChangeColor(oldCol As Long, newcol As Long, bCancel As Boolean, bHandled As Boolean)

On Error Resume Next
       
    cdSet.ShowColor
    If Err.Number <> 0 Then
        bCancel = True
        Exit Sub
    End If
       
On Error GoTo 0
          
    newcol = cdSet.Color
    bCancel = False
    bHandled = True
          
End Sub

'add new ASM color - BUG
Private Sub cmdColor_AddASM_Click()
    set_ASM_ctlColorsASM.EditText set_ASM_ctlColorsASM.AddColor("", vbBlack)
End Sub

'delete ASM color
Private Sub cmdColor_DeleteASM_Click()
Dim oLB As ListBox
          
    Set oLB = set_ASM_ctlColorsASM.listdata
    If oLB.ListIndex = -1 Then Exit Sub
    set_ASM_ctlColorsASM.RemoveColor oLB.ListIndex
          
End Sub

'-------------
'--- Tab C ---
'-------------

'set C default colors
Private Sub cmdColor_DefaultC_Click()
    set_C_ctlColorsC.SetDefaultsC
End Sub

'delete C color
Private Sub cmdColor_DeleteC_Click()
Dim oLB As ListBox
          
    Set oLB = set_C_ctlColorsC.listdata
    If oLB.ListIndex = -1 Then Exit Sub
    set_C_ctlColorsC.RemoveColor oLB.ListIndex
          
End Sub

'add new C color - BUG
Private Sub cmdColor_AddC_Click()
    set_C_ctlColorsC.EditText set_C_ctlColorsC.AddColor("", vbBlack)
End Sub

'change C color
Private Sub set_C_ctlColorsC_ChangeColor(oldCol As Long, newcol As Long, bCancel As Boolean, bHandled As Boolean)

On Error Resume Next
          
    cdSet.ShowColor
    cdSet.CancelError = True
    If Err.Number <> 0 Then
        bCancel = True
        Exit Sub
    End If

On Error GoTo 0
          
    newcol = cdSet.Color
    bCancel = False
    bHandled = True

End Sub


Public Sub lang_UpdateGui(lang As tvb_Languages)
          
    rfile.LoadFormFromResourceFile Me, PLUGIN_NAMEs, lang
              
End Sub

