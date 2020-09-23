VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunLink"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdSet 
      Left            =   480
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctSettings 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   2280
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   2
      Top             =   2280
      Width           =   5055
      Begin VB.OptionButton set_optPacker0 
         Caption         =   "Option1"
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Tag             =   "*def"
         Top             =   1815
         Width           =   255
      End
      Begin VB.TextBox txtDesc0 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   21
         Text            =   "<nothing>"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtCmdLine0 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "<nothing>"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox set_txtPacker 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton cmdPaths_Packer 
         Caption         =   "..."
         Height          =   315
         Left            =   4440
         TabIndex        =   14
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox set_chbUsePacker 
         Caption         =   "* Use packer"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox set_txtCmdLine1 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "user defined"
         Top             =   2340
         Width           =   2055
      End
      Begin VB.TextBox set_txtCmdLine2 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "user defined"
         Top             =   2820
         Width           =   2055
      End
      Begin VB.TextBox set_txtCmdLine3 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "user defined"
         Top             =   3300
         Width           =   2055
      End
      Begin VB.TextBox set_txtDesc1 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox set_txtDesc2 
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox set_txtDesc3 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   3300
         Width           =   1335
      End
      Begin VB.OptionButton set_optPacker1 
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Tag             =   "*"
         Top             =   2355
         Width           =   255
      End
      Begin VB.OptionButton set_optPacker2 
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Tag             =   "*"
         Top             =   2835
         Width           =   255
      End
      Begin VB.OptionButton set_optPacker3 
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Tag             =   "*"
         Top             =   3315
         Width           =   255
      End
      Begin VB.CheckBox set_chbShowPackerOutPut 
         Caption         =   "Show packer output"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblPaths_4 
         AutoSize        =   -1  'True
         Caption         =   "Path to Packer"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label lblPacker_1 
         AutoSize        =   -1  'True
         Caption         =   "Packer Command-Line"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1605
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Use"
         Height          =   195
         Left            =   4005
         TabIndex        =   16
         Top             =   1440
         Width           =   390
      End
   End
   Begin VB.PictureBox pctCredits 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.Label lblCredits 
         AutoSize        =   -1  'True
         Caption         =   "Credits - Packer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2580
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

    pctCredits.Move 0, 0
    pctSettings.Move 0, 0
    
    With cdSet
        .Filter = "executable (*.exe)|*.exe|all (*.*)|*.*"
        .FileName = ""
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware
        .CancelError = True
    End With
   
End Sub

Private Sub pctSettings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'set default
    If Button = vbRightButton Then
        SetDefaultSettings GLOBAL_, pctSettings
        SetDefaultSettings LOCAL_, pctSettings
    End If
    
End Sub

Private Sub cmdPaths_Packer_Click()
    SetPath "Path to packer", set_txtPacker
End Sub

