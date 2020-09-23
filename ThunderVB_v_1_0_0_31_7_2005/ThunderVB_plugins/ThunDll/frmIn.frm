VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#24.2#0"; "ThunVBCC_v1.ocx"
Begin VB.Form frmIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunDll"
   ClientHeight    =   7875
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   724
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctCredits 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   6720
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   14
      Top             =   5760
      Width           =   3855
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Credits - ThunDll"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   17.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2715
      End
   End
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   360
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      Begin ThunVBCC_v1.XTab xTabSet 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   9128
         TabCount        =   2
         TabCaption(0)   =   "StdCall Dll"
         TabContCtrlCnt(0)=   9
         Tab(0)ContCtrlCap(1)=   "set_ctlExport"
         Tab(0)ContCtrlCap(2)=   "cmdAddDllMain"
         Tab(0)ContCtrlCap(3)=   "lblEntryPointName"
         Tab(0)ContCtrlCap(4)=   "lblBaseAddress"
         Tab(0)ContCtrlCap(5)=   "lblInfo"
         Tab(0)ContCtrlCap(6)=   "set_chbExportSymbols"
         Tab(0)ContCtrlCap(7)=   "set_chbCompileDll"
         Tab(0)ContCtrlCap(8)=   "set_txtEntryPoint"
         Tab(0)ContCtrlCap(9)=   "set_txtBaseAddress"
         TabCaption(1)   =   "Pre-Loader"
         TabContCtrlCnt(1)=   9
         Tab(1)ContCtrlCap(1)=   "set_chbBPSubMain"
         Tab(1)ContCtrlCap(2)=   "set_chbBPCallThunRTMain"
         Tab(1)ContCtrlCap(3)=   "set_chbBPPreLoader"
         Tab(1)ContCtrlCap(4)=   "set_chbCallMyDllMain"
         Tab(1)ContCtrlCap(5)=   "set_chbFullLoading"
         Tab(1)ContCtrlCap(6)=   "set_chbUsePreLoader"
         Tab(1)ContCtrlCap(7)=   "lblSetBP"
         Tab(1)ContCtrlCap(8)=   "lblDebugging"
         Tab(1)ContCtrlCap(9)=   "lblGeneral"
         ActiveTab       =   1
         TabTheme        =   2
         InActiveTabBackStartColor=   -2147483626
         InActiveTabBackEndColor=   -2147483626
         InActiveTabForeColor=   -2147483631
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Begin ThunDll.ExportList set_ctlExport 
            Height          =   2415
            Left            =   -74760
            TabIndex        =   4
            Top             =   1320
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4260
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbBPSubMain 
            Height          =   240
            Left            =   360
            TabIndex        =   13
            Top             =   3480
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* Sub Main (after call to ThunRTMain)"
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
         Begin ThunVBCC_v1.HzxYCheckBox set_chbBPCallThunRTMain 
            Height          =   240
            Left            =   360
            TabIndex        =   12
            Top             =   3120
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* CallThunRTMain (before call to ThunRTMain)"
            Pic_UncheckedNormal=   "frmIn.frx":27D8
            Pic_CheckedNormal=   "frmIn.frx":2B2A
            Pic_MixedNormal =   "frmIn.frx":2E7C
            Pic_UncheckedDisabled=   "frmIn.frx":31CE
            Pic_CheckedDisabled=   "frmIn.frx":3520
            Pic_MixedDisabled=   "frmIn.frx":3872
            Pic_UncheckedOver=   "frmIn.frx":3BC4
            Pic_CheckedOver =   "frmIn.frx":3F16
            Pic_MixedOver   =   "frmIn.frx":4268
            Pic_UncheckedDown=   "frmIn.frx":45BA
            Pic_CheckedDown =   "frmIn.frx":490C
            Pic_MixedDown   =   "frmIn.frx":4C5E
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbBPPreLoader 
            Height          =   240
            Left            =   360
            TabIndex        =   11
            Top             =   2760
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* PreLoader (dll entry-point)"
            Pic_UncheckedNormal=   "frmIn.frx":4FB0
            Pic_CheckedNormal=   "frmIn.frx":5302
            Pic_MixedNormal =   "frmIn.frx":5654
            Pic_UncheckedDisabled=   "frmIn.frx":59A6
            Pic_CheckedDisabled=   "frmIn.frx":5CF8
            Pic_MixedDisabled=   "frmIn.frx":604A
            Pic_UncheckedOver=   "frmIn.frx":639C
            Pic_CheckedOver =   "frmIn.frx":66EE
            Pic_MixedOver   =   "frmIn.frx":6A40
            Pic_UncheckedDown=   "frmIn.frx":6D92
            Pic_CheckedDown =   "frmIn.frx":70E4
            Pic_MixedDown   =   "frmIn.frx":7436
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbCallMyDllMain 
            Height          =   240
            Left            =   360
            TabIndex        =   10
            Tag             =   "def"
            Top             =   1560
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* Call my DllMain"
            Pic_UncheckedNormal=   "frmIn.frx":7788
            Pic_CheckedNormal=   "frmIn.frx":7ADA
            Pic_MixedNormal =   "frmIn.frx":7E2C
            Pic_UncheckedDisabled=   "frmIn.frx":817E
            Pic_CheckedDisabled=   "frmIn.frx":84D0
            Pic_MixedDisabled=   "frmIn.frx":8822
            Pic_UncheckedOver=   "frmIn.frx":8B74
            Pic_CheckedOver =   "frmIn.frx":8EC6
            Pic_MixedOver   =   "frmIn.frx":9218
            Pic_UncheckedDown=   "frmIn.frx":956A
            Pic_CheckedDown =   "frmIn.frx":98BC
            Pic_MixedDown   =   "frmIn.frx":9C0E
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbFullLoading 
            Height          =   240
            Left            =   360
            TabIndex        =   9
            Tag             =   "def"
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* Full loading"
            Pic_UncheckedNormal=   "frmIn.frx":9F60
            Pic_CheckedNormal=   "frmIn.frx":A2B2
            Pic_MixedNormal =   "frmIn.frx":A604
            Pic_UncheckedDisabled=   "frmIn.frx":A956
            Pic_CheckedDisabled=   "frmIn.frx":ACA8
            Pic_MixedDisabled=   "frmIn.frx":AFFA
            Pic_UncheckedOver=   "frmIn.frx":B34C
            Pic_CheckedOver =   "frmIn.frx":B69E
            Pic_MixedOver   =   "frmIn.frx":B9F0
            Pic_UncheckedDown=   "frmIn.frx":BD42
            Pic_CheckedDown =   "frmIn.frx":C094
            Pic_MixedDown   =   "frmIn.frx":C3E6
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbUsePreLoader 
            Height          =   240
            Left            =   360
            TabIndex        =   8
            Tag             =   "def"
            Top             =   840
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* Use ""Pre-Loader"""
            Pic_UncheckedNormal=   "frmIn.frx":C738
            Pic_CheckedNormal=   "frmIn.frx":CA8A
            Pic_MixedNormal =   "frmIn.frx":CDDC
            Pic_UncheckedDisabled=   "frmIn.frx":D12E
            Pic_CheckedDisabled=   "frmIn.frx":D480
            Pic_MixedDisabled=   "frmIn.frx":D7D2
            Pic_UncheckedOver=   "frmIn.frx":DB24
            Pic_CheckedOver =   "frmIn.frx":DE76
            Pic_MixedOver   =   "frmIn.frx":E1C8
            Pic_UncheckedDown=   "frmIn.frx":E51A
            Pic_CheckedDown =   "frmIn.frx":E86C
            Pic_MixedDown   =   "frmIn.frx":EBBE
         End
         Begin ThunVBCC_v1.UniLabel lblSetBP 
            Height          =   255
            Left            =   360
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            CaptionB        =   "frmIn.frx":EF10
            CaptionLen      =   20
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
         Begin ThunVBCC_v1.UniLabel lblDebugging 
            Height          =   255
            Left            =   240
            Top             =   2040
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            CaptionB        =   "frmIn.frx":EF58
            CaptionLen      =   9
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
         Begin ThunVBCC_v1.UniLabel lblGeneral 
            Height          =   255
            Left            =   240
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            CaptionB        =   "frmIn.frx":EF8A
            CaptionLen      =   7
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
         Begin ThunVBCC_v1.isButton cmdAddDllMain 
            Height          =   375
            Left            =   -74760
            TabIndex        =   7
            Top             =   4680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Icon            =   "frmIn.frx":EFB8
            Style           =   5
            Caption         =   "DllMain template"
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
         Begin ThunVBCC_v1.UniLabel lblEntryPointName 
            Height          =   240
            Left            =   -74760
            Top             =   4200
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   423
            CaptionB        =   "frmIn.frx":EFD4
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
         Begin ThunVBCC_v1.UniLabel lblBaseAddress 
            Height          =   240
            Left            =   -74760
            Top             =   3840
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   423
            CaptionB        =   "frmIn.frx":F018
            CaptionLen      =   22
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
         Begin ThunVBCC_v1.UniLabel lblInfo 
            Height          =   240
            Left            =   -74760
            Top             =   960
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   423
            CaptionB        =   "frmIn.frx":F064
            CaptionLen      =   16
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
         Begin ThunVBCC_v1.HzxYCheckBox set_chbExportSymbols 
            Height          =   240
            Left            =   -72360
            TabIndex        =   3
            Top             =   600
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* Export functions"
            Pic_UncheckedNormal=   "frmIn.frx":F0A4
            Pic_CheckedNormal=   "frmIn.frx":F3F6
            Pic_MixedNormal =   "frmIn.frx":F748
            Pic_UncheckedDisabled=   "frmIn.frx":FA9A
            Pic_CheckedDisabled=   "frmIn.frx":FDEC
            Pic_MixedDisabled=   "frmIn.frx":1013E
            Pic_UncheckedOver=   "frmIn.frx":10490
            Pic_CheckedOver =   "frmIn.frx":107E2
            Pic_MixedOver   =   "frmIn.frx":10B34
            Pic_UncheckedDown=   "frmIn.frx":10E86
            Pic_CheckedDown =   "frmIn.frx":111D8
            Pic_MixedDown   =   "frmIn.frx":1152A
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbCompileDll 
            Height          =   240
            Left            =   -74760
            TabIndex        =   2
            Top             =   600
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0
            Caption         =   "* Compile DLL"
            Pic_UncheckedNormal=   "frmIn.frx":1187C
            Pic_CheckedNormal=   "frmIn.frx":11BCE
            Pic_MixedNormal =   "frmIn.frx":11F20
            Pic_UncheckedDisabled=   "frmIn.frx":12272
            Pic_CheckedDisabled=   "frmIn.frx":125C4
            Pic_MixedDisabled=   "frmIn.frx":12916
            Pic_UncheckedOver=   "frmIn.frx":12C68
            Pic_CheckedOver =   "frmIn.frx":12FBA
            Pic_MixedOver   =   "frmIn.frx":1330C
            Pic_UncheckedDown=   "frmIn.frx":1365E
            Pic_CheckedDown =   "frmIn.frx":139B0
            Pic_MixedDown   =   "frmIn.frx":13D02
         End
         Begin VB.TextBox set_txtEntryPoint 
            Height          =   285
            Left            =   -73080
            TabIndex        =   5
            Tag             =   "*DllMain"
            Top             =   4200
            Width           =   1365
         End
         Begin VB.TextBox set_txtBaseAddress 
            Height          =   285
            Left            =   -73080
            TabIndex        =   6
            Tag             =   "*11000000"
            Top             =   3840
            Width           =   1365
         End
      End
   End
   Begin VB.Menu mnuBaseAddress 
      Caption         =   "Base Address"
      Begin VB.Menu mnuBaseAddressItem 
         Caption         =   "400000"
         Index           =   1
      End
      Begin VB.Menu mnuBaseAddressItem 
         Caption         =   "11000000"
         Index           =   2
      End
   End
   Begin VB.Menu mnuDllMain 
      Caption         =   "DllMain"
      Begin VB.Menu mnuDllEntryPoint 
         Caption         =   "DllMain"
         Index           =   1
      End
      Begin VB.Menu mnuDllEntryPoint 
         Caption         =   "DummyDllMain"
         Index           =   2
      End
      Begin VB.Menu mnuDllEntryPoint 
         Caption         =   "__vbaS"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'undocumented function in vba6.dll
'we use it to detect if name of dll (entered by user) entry-point is valid function name
Private Declare Function EbIsValidIdent Lib "vba6.dll" (ByVal lpstrIdent As Long, ByRef pfIsValid As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_RBUTTONDOWN As Long = &H204

Private Const FORM_NAME As String = "frmIn"

Private Sub Form_Load()
          
10        LogMsg "Loading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, FORM_NAME, "Form_Load"
          
20        pctSettings.Move 0, 0
30        pctCredits.Move 0, 0
          
40        xTabSet.ActiveTab = 0
          
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        LogMsg "Unloading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, FORM_NAME, "Form_Unload"
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Dim b As Boolean
'
'    'TODO - BUG
'
'    set_txtEntryPoint_Validate b
'    If b = True Then Cancel = 1
'
'    set_txtBaseAddress_Validate b
'    If b = True Then Cancel = 1
'
'End Sub

Private Sub cmdAddDllMain_Click()
10        LogMsg "Add DllMain template.", PLUGIN_NAMEs, FORM_NAME, "cmdAddDllMain_Click"
20        frmDllMain.Show vbModal, Me
End Sub

Private Sub mnuBaseAddressItem_Click(Index As Integer)
10        Me.set_txtBaseAddress.Text = mnuBaseAddressItem(Index).Caption
End Sub

Private Sub mnuDllEntryPoint_Click(Index As Integer)
10        Me.set_txtEntryPoint.Text = mnuDllEntryPoint(Index).Caption
End Sub

Private Sub pctSettings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          'set default
10        If Button = vbRightButton Then
              
20            LogMsg "Set settings to default.", PLUGIN_NAMEs, FORM_NAME, "pctSettings_MouseUp"
              
30            SetDefaultSettings GLOBAL_, pctSettings
40            SetDefaultSettings LOCAL_, pctSettings
              
50        End If

End Sub

Private Sub set_txtBaseAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
10        If Button = vbRightButton Then
20            SendMessage Me.hWnd, WM_RBUTTONDOWN, 0, ByVal 0&
30            PopupMenu mnuBaseAddress
40        End If
          
End Sub

Private Sub set_txtBaseAddress_Validate(Cancel As Boolean)

          'check string
10        set_txtBaseAddress.Text = Trim(set_txtBaseAddress.Text)
20        If IsNumeric("&H" & set_txtBaseAddress.Text) = False Then
30            MsgBoxX "Invalid base address.", MSG_TITLEs
40            Cancel = True
50        End If

End Sub

Private Sub set_txtEntryPoint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        If Button = vbRightButton Then
20            SendMessage Me.hWnd, WM_RBUTTONDOWN, 0, ByVal 0&
30            PopupMenu mnuDllMain
40        End If

End Sub

Private Sub set_txtEntryPoint_Validate(Cancel As Boolean)
      Dim lRet As Long
          
10        set_txtEntryPoint.Text = Trim(set_txtEntryPoint.Text)
          
          'check string
20        If EbIsValidIdent(StrPtr(set_txtEntryPoint.Text), lRet) Or lRet = 0 Then
30            MsgBoxX "Invalid Entry-Point name.", MSG_TITLEs
40            Cancel = True
50        End If
          
End Sub
