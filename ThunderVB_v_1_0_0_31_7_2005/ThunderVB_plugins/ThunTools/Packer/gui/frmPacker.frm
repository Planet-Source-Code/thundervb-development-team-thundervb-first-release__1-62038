VERSION 5.00
Begin VB.Form frmPacker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packer"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox set_chbShowPackerOutPut 
      Caption         =   "Show packer output"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton optPacker_3 
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   2835
      Width           =   255
   End
   Begin VB.OptionButton optPacker_2 
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   2355
      Width           =   255
   End
   Begin VB.OptionButton optPacker_1 
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   1875
      Width           =   255
   End
   Begin VB.TextBox set_chbDesc3 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   2820
      Width           =   1335
   End
   Begin VB.TextBox set_chbDesc2 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   2340
      Width           =   1335
   End
   Begin VB.TextBox set_chbDesc1 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   1860
      Width           =   1335
   End
   Begin VB.TextBox set_chbCmdLine3 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2820
      Width           =   2055
   End
   Begin VB.TextBox set_chbCmdLine2 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2340
      Width           =   2055
   End
   Begin VB.TextBox set_chbCmdLine1 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1860
      Width           =   2055
   End
   Begin VB.CheckBox set_chbUsePacker 
      Caption         =   "* Use packer"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Tag             =   "def"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdPaths_Packer 
      Caption         =   "..."
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox set_chbPacker 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdStorno 
      Caption         =   "Storno"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "* Use"
      Height          =   195
      Left            =   4005
      TabIndex        =   18
      Top             =   1440
      Width           =   390
   End
   Begin VB.Label lblPacker_2 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   2520
      TabIndex        =   17
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label lblPacker_1 
      AutoSize        =   -1  'True
      Caption         =   "Packer Command-Line"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   1605
   End
   Begin VB.Label lblPaths_4 
      AutoSize        =   -1  'True
      Caption         =   "Path to Packer"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "frmPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    frmPacker.Caption = PLUGIN_NAME
End Sub
