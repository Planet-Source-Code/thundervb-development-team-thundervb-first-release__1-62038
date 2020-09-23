VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmThunIDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunIDE+"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdStorno 
      Caption         =   "Storno"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin TabDlg.SSTab sstIDE 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "ASM"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblpreVasm"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdColor_DeleteASM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdColor_AddASM"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ASM_UseASMColors"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ASM_QuickWatch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ASM_IntelliSense"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "C"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdColor_DeleteC"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdColor_AddC"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "C_UseCColors"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblpreVc"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CheckBox ASM_IntelliSense 
         Caption         =   "Intelli - Sense"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox ASM_QuickWatch 
         Caption         =   "Quick - Watch"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdColor_DeleteC 
         Caption         =   "Delete"
         Height          =   255
         Left            =   -73800
         TabIndex        =   10
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdColor_AddC 
         Caption         =   "Add"
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox C_UseCColors 
         Caption         =   "* Use C code coloring"
         Height          =   195
         Left            =   -74760
         TabIndex        =   8
         Tag             =   "def"
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox ASM_UseASMColors 
         Caption         =   "* Use ASM code coloring"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Tag             =   "def"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdColor_AddASM 
         Caption         =   "Add"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdColor_DeleteASM 
         Caption         =   "Delete"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblpreVc 
         AutoSize        =   -1  'True
         Caption         =   "'#c' int i = 0x80 // a sample line"
         Height          =   195
         Left            =   -74760
         TabIndex        =   11
         Top             =   3480
         Width           =   2220
      End
      Begin VB.Label lblpreVasm 
         AutoSize        =   -1  'True
         Caption         =   "'#asm' mov eax , 12345 ; an eaxmple"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   4080
         Width           =   2625
      End
   End
End
Attribute VB_Name = "frmThunIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
