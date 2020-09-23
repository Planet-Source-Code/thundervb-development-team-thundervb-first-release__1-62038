VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmThunLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunLink"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStorno 
      Caption         =   "Storno"
      Height          =   375
      Left            =   4200
      TabIndex        =   43
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDebugEnumLocal 
      Caption         =   "Debug - Enum local controls"
      Height          =   495
      Left            =   480
      TabIndex        =   42
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdDebugEnumDefault 
      Caption         =   "Debug - Enum default controls"
      Height          =   495
      Left            =   3000
      TabIndex        =   41
      Top             =   6600
      Width           =   1575
   End
   Begin TabDlg.SSTab sstLink 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmThunLink.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "General_HideErrorDialogs"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "General_SaveOBJ"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "General_HookCompiler"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "General_ListForAllMod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "General_PopUpWindow"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Paths"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPaths_TextEditor"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdPaths_MIDL"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdPaths_ML"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Paths_MIDL"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Paths_TextEditor"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Paths_ML"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Paths_CCompiler"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Paths_LIBFiles"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Paths_INCFiles"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdPaths_CCompiler"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdPaths_LIBFiles"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdINCFiles"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblPaths_3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblPaths_2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblPaths_1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblPaths_5"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblPaths_6"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Compile"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Compile_PauseAsm"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Compile_PauseLink"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Compile_ModifyCmdLine"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Compile_SkipLinking"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Debug"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Debug_OutDebLog"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Debug_DelDebBefCom"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Debug_OutAsmToLog"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Debug_OutMapFiles"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fraDebug_Frame1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdDebug_DeleteAllFiles(0)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdGeneral_DeleteDebugDir(0)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "ASM"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ASM_FixASMListings"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "ASM_CompileASMCode"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "C"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "C_CompileCCode"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.CheckBox General_PopUpWindow 
         Caption         =   "* PopUp StdCall DLL window when compiling"
         Height          =   250
         Left            =   240
         TabIndex        =   44
         Top             =   2880
         Width           =   3975
      End
      Begin VB.CheckBox C_CompileCCode 
         Caption         =   "Compile C code"
         Height          =   250
         Left            =   -74760
         TabIndex        =   40
         Tag             =   "def"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox ASM_FixASMListings 
         Caption         =   "* Fix ASM listings"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox ASM_CompileASMCode 
         Caption         =   "Compile ASM code"
         Height          =   250
         Left            =   -74760
         TabIndex        =   38
         Tag             =   "def"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox Debug_OutDebLog 
         Caption         =   "Enable Output to DebugLog"
         Height          =   375
         Left            =   -74760
         TabIndex        =   37
         Tag             =   "def"
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox Debug_DelDebBefCom 
         Caption         =   "Delete DebugLog before compiling"
         Height          =   375
         Left            =   -74760
         TabIndex        =   36
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox Debug_OutAsmToLog 
         Caption         =   "Output Assembler messages to log"
         Height          =   375
         Left            =   -74760
         TabIndex        =   35
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CheckBox Debug_OutMapFiles 
         Caption         =   "Output detailed MASM && LINK Map files"
         Height          =   375
         Left            =   -74760
         TabIndex        =   34
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Frame fraDebug_Frame1 
         Caption         =   "Delete when UnLoading"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   31
         Top             =   3840
         Width           =   2055
         Begin VB.CheckBox Debug_DeleteLST 
            Caption         =   ".LST files"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox Debug_DeleteASM 
            Caption         =   ".ASM files"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdDebug_DeleteAllFiles 
         Caption         =   "Delete all files in ""/debug"""
         Height          =   375
         Index           =   0
         Left            =   -72600
         TabIndex        =   30
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton cmdGeneral_DeleteDebugDir 
         Caption         =   "Delete ""/debug"""
         Height          =   375
         Index           =   0
         Left            =   -72600
         TabIndex        =   29
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CheckBox Compile_PauseAsm 
         Caption         =   "Pause before assembly"
         Height          =   375
         Left            =   -74760
         TabIndex        =   28
         Top             =   900
         Width           =   2055
      End
      Begin VB.CheckBox Compile_PauseLink 
         Caption         =   "Pause before linking"
         Height          =   375
         Left            =   -74760
         TabIndex        =   27
         Top             =   1380
         Width           =   2055
      End
      Begin VB.CheckBox Compile_ModifyCmdLine 
         Caption         =   "Modify CmdLine"
         Height          =   375
         Left            =   -74760
         TabIndex        =   26
         Top             =   1860
         Width           =   1455
      End
      Begin VB.CheckBox Compile_SkipLinking 
         Caption         =   "Skip linking"
         Height          =   375
         Left            =   -74760
         TabIndex        =   25
         Top             =   2340
         Width           =   1335
      End
      Begin VB.CommandButton cmdPaths_TextEditor 
         Caption         =   "..."
         Height          =   315
         Left            =   -70680
         TabIndex        =   18
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_MIDL 
         Caption         =   "..."
         Height          =   315
         Left            =   -70680
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_ML 
         Caption         =   "..."
         Height          =   315
         Left            =   -70680
         TabIndex        =   16
         Top             =   1020
         Width           =   375
      End
      Begin VB.TextBox Paths_MIDL 
         Height          =   315
         Left            =   -74880
         TabIndex        =   15
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox Paths_TextEditor 
         Height          =   315
         Left            =   -74880
         TabIndex        =   14
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox Paths_ML 
         Height          =   315
         Left            =   -74880
         TabIndex        =   13
         Top             =   1020
         Width           =   4095
      End
      Begin VB.TextBox Paths_CCompiler 
         Height          =   315
         Left            =   -74880
         TabIndex        =   12
         Top             =   3420
         Width           =   4095
      End
      Begin VB.TextBox Paths_LIBFiles 
         Height          =   315
         Left            =   -74880
         TabIndex        =   11
         Top             =   4020
         Width           =   4095
      End
      Begin VB.TextBox Paths_INCFiles 
         Height          =   315
         Left            =   -74880
         TabIndex        =   10
         Top             =   4620
         Width           =   4095
      End
      Begin VB.CommandButton cmdPaths_CCompiler 
         Caption         =   "..."
         Height          =   315
         Left            =   -70680
         TabIndex        =   9
         Top             =   3420
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_LIBFiles 
         Caption         =   "..."
         Height          =   315
         Left            =   -70680
         TabIndex        =   8
         Top             =   4020
         Width           =   375
      End
      Begin VB.CommandButton cmdINCFiles 
         Caption         =   "..."
         Height          =   315
         Left            =   -70680
         TabIndex        =   7
         Top             =   4620
         Width           =   375
      End
      Begin VB.CheckBox General_ListForAllMod 
         Caption         =   "* Listings for all modules"
         Height          =   250
         Left            =   240
         TabIndex        =   6
         Top             =   1380
         Width           =   2295
      End
      Begin VB.CheckBox General_HookCompiler 
         Caption         =   "Hook compiler"
         Height          =   250
         Left            =   240
         TabIndex        =   5
         Tag             =   "def"
         Top             =   900
         Width           =   1935
      End
      Begin VB.CheckBox General_SaveOBJ 
         Caption         =   "* Save .OBJ files"
         Height          =   250
         Left            =   240
         TabIndex        =   4
         Top             =   1860
         Width           =   2055
      End
      Begin VB.CheckBox General_HideErrorDialogs 
         Caption         =   "* Hide error dialogs"
         Height          =   250
         Left            =   240
         TabIndex        =   3
         Top             =   2340
         Width           =   1935
      End
      Begin VB.Label lblPaths_3 
         AutoSize        =   -1  'True
         Caption         =   "Path to MIDL.exe"
         Height          =   195
         Left            =   -74880
         TabIndex        =   24
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lblPaths_2 
         AutoSize        =   -1  'True
         Caption         =   "Path to Text-Editor"
         Height          =   195
         Left            =   -74880
         TabIndex        =   23
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label lblPaths_1 
         AutoSize        =   -1  'True
         Caption         =   "Path to ML.exe"
         Height          =   195
         Left            =   -74880
         TabIndex        =   22
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label lblPaths_5 
         AutoSize        =   -1  'True
         Caption         =   "Path to C compiler"
         Height          =   195
         Left            =   -74880
         TabIndex        =   21
         Top             =   3180
         Width           =   1290
      End
      Begin VB.Label lblPaths_6 
         AutoSize        =   -1  'True
         Caption         =   "Path to .LIB files"
         Height          =   195
         Left            =   -74880
         TabIndex        =   20
         Top             =   3780
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Path to .INC files"
         Height          =   195
         Left            =   -74880
         TabIndex        =   19
         Top             =   4380
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdSet 
      Left            =   2160
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmThunLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'set Default Settings
Private Sub cmdDefault_Click()
    Call SetDefaultSettings(GLOBAL_, Me, True)
    Call SetDefaultSettings(LOCAL_, Me, True)
End Sub


'-----------------
'--- Tab Paths ---
'-----------------

Private Sub cmdPaths_MIDL_Click()
    Call SetPath("Path to " & Add34("MIDL.exe"), Paths_MIDL, "midl.exe")
End Sub

Private Sub cmdPaths_ML_Click()
    Call SetPath("Path to " & Add34("ML.exe"), Paths_ML, "ml.exe")
End Sub

Private Sub cmdPaths_Packer_Click()
    Call SetPath("Path to packer", Paths_Packer)
End Sub

Private Sub cmdPaths_TextEditor_Click()
    Call SetPath("Path to Text-Editor", Paths_TextEditor)
End Sub

Private Sub cmdPaths_CCompiler_Click()
    Call SetPath("Path to C compiler", Paths_CCompiler)
End Sub

Private Sub cmdPaths_LIBFiles_Click()
    Call SetDirectory("Select .LIB directory", Paths_LIBFiles)
End Sub

Private Sub cmdINCFiles_Click()
    Call SetDirectory("Select .INC directory", Paths_INCFiles)
End Sub

Private Sub Form_Load()
    
    'common dialog settings
    With cdSet
        .Filter = "executable (*.exe)|*.exe|all (*.*)|*.*"
        .FileName = ""
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware
        .CancelError = True
    End With
    
End Sub

'append "\"
Private Sub Paths_INCFiles_Validate(Cancel As Boolean)
    If Right(Paths_INCFiles.Text, 1) <> "\" Then Paths_INCFiles.Text = Paths_INCFiles.Text & "\"
End Sub

'append "\"
Private Sub Paths_LIBFiles_Validate(Cancel As Boolean)
    If Right(Paths_LIBFiles.Text, 1) <> "\" Then Paths_LIBFiles = Paths_LIBFiles & "\"
End Sub

'show local settings controls
Private Sub cmdDebugEnumLocal_Click()
    modPlugins.DebugLocalControls
End Sub

'show defualt settings control
Private Sub cmdDebugEnumDefault_Click()
    modPlugins.DebugDefaultControls
End Sub
