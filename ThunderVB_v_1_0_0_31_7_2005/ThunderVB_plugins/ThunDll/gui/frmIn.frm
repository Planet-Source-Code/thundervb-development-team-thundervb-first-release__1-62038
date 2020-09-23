VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
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
      Left            =   6000
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   14
      Top             =   3240
      Width           =   3855
      Begin ThunVBCC_v1.UniLabel lblCredits 
         Height          =   615
         Left            =   480
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
         CaptionB        =   "frmIn.frx":0000
         CaptionLen      =   17
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   360
      ScaleHeight     =   5535
      ScaleWidth      =   4815
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
         Begin ThunDll.ExportList set_ctlExport 
            Height          =   2415
            Left            =   240
            TabIndex        =   4
            Top             =   1320
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4260
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbBPSubMain 
            Height          =   240
            Left            =   -74640
            TabIndex        =   13
            Top             =   3480
            Visible         =   0   'False
            Width           =   3000
            _ExtentX        =   5265
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
            Caption         =   "* Sub Main (after call to ThunRTMain)"
            Pic_UncheckedNormal=   "frmIn.frx":0042
            Pic_CheckedNormal=   "frmIn.frx":0394
            Pic_MixedNormal =   "frmIn.frx":06E6
            Pic_UncheckedDisabled=   "frmIn.frx":0A38
            Pic_CheckedDisabled=   "frmIn.frx":0D8A
            Pic_MixedDisabled=   "frmIn.frx":10DC
            Pic_UncheckedOver=   "frmIn.frx":142E
            Pic_CheckedOver =   "frmIn.frx":1780
            Pic_MixedOver   =   "frmIn.frx":1AD2
            Pic_UncheckedDown=   "frmIn.frx":1E24
            Pic_CheckedDown =   "frmIn.frx":2176
            Pic_MixedDown   =   "frmIn.frx":24C8
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbBPCallThunRTMain 
            Height          =   240
            Left            =   -74640
            TabIndex        =   12
            Top             =   3120
            Visible         =   0   'False
            Width           =   3660
            _ExtentX        =   6350
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
            Caption         =   "* CallThunRTMain (before call to ThunRTMain)"
            Pic_UncheckedNormal=   "frmIn.frx":281A
            Pic_CheckedNormal=   "frmIn.frx":2B6C
            Pic_MixedNormal =   "frmIn.frx":2EBE
            Pic_UncheckedDisabled=   "frmIn.frx":3210
            Pic_CheckedDisabled=   "frmIn.frx":3562
            Pic_MixedDisabled=   "frmIn.frx":38B4
            Pic_UncheckedOver=   "frmIn.frx":3C06
            Pic_CheckedOver =   "frmIn.frx":3F58
            Pic_MixedOver   =   "frmIn.frx":42AA
            Pic_UncheckedDown=   "frmIn.frx":45FC
            Pic_CheckedDown =   "frmIn.frx":494E
            Pic_MixedDown   =   "frmIn.frx":4CA0
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbBPPreLoader 
            Height          =   240
            Left            =   -74640
            TabIndex        =   11
            Top             =   2760
            Visible         =   0   'False
            Width           =   2220
            _ExtentX        =   4075
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
            Caption         =   "* PreLoader (dll entry-point)"
            Pic_UncheckedNormal=   "frmIn.frx":4FF2
            Pic_CheckedNormal=   "frmIn.frx":5344
            Pic_MixedNormal =   "frmIn.frx":5696
            Pic_UncheckedDisabled=   "frmIn.frx":59E8
            Pic_CheckedDisabled=   "frmIn.frx":5D3A
            Pic_MixedDisabled=   "frmIn.frx":608C
            Pic_UncheckedOver=   "frmIn.frx":63DE
            Pic_CheckedOver =   "frmIn.frx":6730
            Pic_MixedOver   =   "frmIn.frx":6A82
            Pic_UncheckedDown=   "frmIn.frx":6DD4
            Pic_CheckedDown =   "frmIn.frx":7126
            Pic_MixedDown   =   "frmIn.frx":7478
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbCallMyDllMain 
            Height          =   240
            Left            =   -74640
            TabIndex        =   10
            Tag             =   "def"
            Top             =   1560
            Width           =   1485
            _ExtentX        =   2593
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
            Caption         =   "* Call my DllMain"
            Pic_UncheckedNormal=   "frmIn.frx":77CA
            Pic_CheckedNormal=   "frmIn.frx":7B1C
            Pic_MixedNormal =   "frmIn.frx":7E6E
            Pic_UncheckedDisabled=   "frmIn.frx":81C0
            Pic_CheckedDisabled=   "frmIn.frx":8512
            Pic_MixedDisabled=   "frmIn.frx":8864
            Pic_UncheckedOver=   "frmIn.frx":8BB6
            Pic_CheckedOver =   "frmIn.frx":8F08
            Pic_MixedOver   =   "frmIn.frx":925A
            Pic_UncheckedDown=   "frmIn.frx":95AC
            Pic_CheckedDown =   "frmIn.frx":98FE
            Pic_MixedDown   =   "frmIn.frx":9C50
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbFullLoading 
            Height          =   240
            Left            =   -74640
            TabIndex        =   9
            Tag             =   "def"
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
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
            Caption         =   "* Full loading"
            Pic_UncheckedNormal=   "frmIn.frx":9FA2
            Pic_CheckedNormal=   "frmIn.frx":A2F4
            Pic_MixedNormal =   "frmIn.frx":A646
            Pic_UncheckedDisabled=   "frmIn.frx":A998
            Pic_CheckedDisabled=   "frmIn.frx":ACEA
            Pic_MixedDisabled=   "frmIn.frx":B03C
            Pic_UncheckedOver=   "frmIn.frx":B38E
            Pic_CheckedOver =   "frmIn.frx":B6E0
            Pic_MixedOver   =   "frmIn.frx":BA32
            Pic_UncheckedDown=   "frmIn.frx":BD84
            Pic_CheckedDown =   "frmIn.frx":C0D6
            Pic_MixedDown   =   "frmIn.frx":C428
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbUsePreLoader 
            Height          =   240
            Left            =   -74640
            TabIndex        =   8
            Tag             =   "def"
            Top             =   840
            Width           =   1680
            _ExtentX        =   2963
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
            Caption         =   "* Use ""Pre-Loader"""
            Pic_UncheckedNormal=   "frmIn.frx":C77A
            Pic_CheckedNormal=   "frmIn.frx":CACC
            Pic_MixedNormal =   "frmIn.frx":CE1E
            Pic_UncheckedDisabled=   "frmIn.frx":D170
            Pic_CheckedDisabled=   "frmIn.frx":D4C2
            Pic_MixedDisabled=   "frmIn.frx":D814
            Pic_UncheckedOver=   "frmIn.frx":DB66
            Pic_CheckedOver =   "frmIn.frx":DEB8
            Pic_MixedOver   =   "frmIn.frx":E20A
            Pic_UncheckedDown=   "frmIn.frx":E55C
            Pic_CheckedDown =   "frmIn.frx":E8AE
            Pic_MixedDown   =   "frmIn.frx":EC00
         End
         Begin ThunVBCC_v1.UniLabel lblSetBP 
            Height          =   255
            Left            =   -74640
            Top             =   2400
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            CaptionB        =   "frmIn.frx":EF52
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
            Left            =   -74760
            Top             =   2040
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            CaptionB        =   "frmIn.frx":EF9A
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
            Left            =   -74760
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            CaptionB        =   "frmIn.frx":EFCC
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
            Left            =   240
            TabIndex        =   7
            Top             =   4680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Icon            =   "frmIn.frx":EFFA
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
            Left            =   240
            Top             =   4200
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   423
            CaptionB        =   "frmIn.frx":F016
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
            Left            =   240
            Top             =   3840
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   423
            CaptionB        =   "frmIn.frx":F05A
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
            Left            =   240
            Top             =   960
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   423
            CaptionB        =   "frmIn.frx":F0A6
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
            Left            =   2640
            TabIndex        =   3
            Top             =   600
            Width           =   1560
            _ExtentX        =   2831
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
            Caption         =   "* Export functions"
            Pic_UncheckedNormal=   "frmIn.frx":F0E6
            Pic_CheckedNormal=   "frmIn.frx":F438
            Pic_MixedNormal =   "frmIn.frx":F78A
            Pic_UncheckedDisabled=   "frmIn.frx":FADC
            Pic_CheckedDisabled=   "frmIn.frx":FE2E
            Pic_MixedDisabled=   "frmIn.frx":10180
            Pic_UncheckedOver=   "frmIn.frx":104D2
            Pic_CheckedOver =   "frmIn.frx":10824
            Pic_MixedOver   =   "frmIn.frx":10B76
            Pic_UncheckedDown=   "frmIn.frx":10EC8
            Pic_CheckedDown =   "frmIn.frx":1121A
            Pic_MixedDown   =   "frmIn.frx":1156C
         End
         Begin ThunVBCC_v1.HzxYCheckBox set_chbCompileDll 
            Height          =   240
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   1320
            _ExtentX        =   2302
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
            Caption         =   "* Compile DLL"
            Pic_UncheckedNormal=   "frmIn.frx":118BE
            Pic_CheckedNormal=   "frmIn.frx":11C10
            Pic_MixedNormal =   "frmIn.frx":11F62
            Pic_UncheckedDisabled=   "frmIn.frx":122B4
            Pic_CheckedDisabled=   "frmIn.frx":12606
            Pic_MixedDisabled=   "frmIn.frx":12958
            Pic_UncheckedOver=   "frmIn.frx":12CAA
            Pic_CheckedOver =   "frmIn.frx":12FFC
            Pic_MixedOver   =   "frmIn.frx":1334E
            Pic_UncheckedDown=   "frmIn.frx":136A0
            Pic_CheckedDown =   "frmIn.frx":139F2
            Pic_MixedDown   =   "frmIn.frx":13D44
         End
         Begin VB.TextBox set_txtEntryPoint 
            Height          =   285
            Left            =   1920
            TabIndex        =   5
            Tag             =   "*DllMain"
            Top             =   4200
            Width           =   1365
         End
         Begin VB.TextBox set_txtBaseAddress 
            Height          =   285
            Left            =   1920
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
Dim X As Long, Y As Long
          
    GetSettingsTabClientRect X, Y
    
    X = X * Screen.TwipsPerPixelX
    Y = Y * Screen.TwipsPerPixelY

    pctSettings.Move 0, 0, X, Y
    pctCredits.Move 0, 0, X, Y

    xTabSet.Move 0, 0, X, Y
    lblCredits.Move 0, 0, X, Y
    
    xTabSet.ActiveTab = 0
          
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
    LogMsg "Add DllMain template.", FORM_NAME, "cmdAddDllMain_Click"
    frmDllMain.Show ' vbModal, Me
End Sub

Private Sub mnuBaseAddressItem_Click(Index As Integer)
    Me.set_txtBaseAddress.Text = mnuBaseAddressItem(Index).caption
End Sub

Private Sub mnuDllEntryPoint_Click(Index As Integer)
    Me.set_txtEntryPoint.Text = mnuDllEntryPoint(Index).caption
End Sub

Private Sub set_txtBaseAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        SendMessage Me.hWnd, WM_RBUTTONDOWN, 0, ByVal 0&
        PopupMenu mnuBaseAddress
    End If
    
End Sub

Private Sub set_txtBaseAddress_Validate(Cancel As Boolean)

    'check string
    set_txtBaseAddress.Text = Trim(set_txtBaseAddress.Text)
    If IsNumeric("&H" & set_txtBaseAddress.Text) = False Then
        MsgBoxX "Invalid base address.", MSG_TITLEs
        Cancel = True
    End If

End Sub

Private Sub set_txtEntryPoint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        SendMessage Me.hWnd, WM_RBUTTONDOWN, 0, ByVal 0&
        PopupMenu mnuDllMain
    End If

End Sub

Private Sub set_txtEntryPoint_Validate(Cancel As Boolean)
Dim lRet As Long
    
    set_txtEntryPoint.Text = Trim(set_txtEntryPoint.Text)
    
    'check string
    If EbIsValidIdent(StrPtr(set_txtEntryPoint.Text), lRet) Or lRet = 0 Then
        MsgBoxX "Invalid Entry-Point name.", MSG_TITLEs
        Cancel = True
    End If
    
End Sub

Private Sub xTabSet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If xTabSet.ActiveTab = 1 And Button = vbRightButton Then
        
        lblSetBP.Visible = Not lblSetBP.Visible
        lblDebugging.Visible = Not lblDebugging.Visible
        set_chbBPCallThunRTMain.Visible = Not set_chbBPCallThunRTMain.Visible
        set_chbBPPreLoader.Visible = Not set_chbBPPreLoader.Visible
        set_chbBPSubMain.Visible = Not set_chbBPSubMain.Visible
        
    End If

End Sub
