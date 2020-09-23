VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   ClipControls    =   0   'False
   FillColor       =   &H00E0E0E0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled HzxYOption1"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin HzxYOptionTest.HzxYOption HzxYOption2 
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   1245
      _ExtentX        =   2355
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HzxYOptionTest.HzxYOption HzxYOption1 
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1245
      _ExtentX        =   2355
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648384
      Pic_FalseNormal =   "Form1.frx":0000
      Pic_TrueNormal  =   "Form1.frx":0296
      Pic_FalseDisabled=   "Form1.frx":052C
      Pic_TrueDisabled=   "Form1.frx":07C2
      Pic_FalseOver   =   "Form1.frx":0A58
      Pic_TrueOver    =   "Form1.frx":0CEE
      Pic_FalseDown   =   "Form1.frx":0F84
      Pic_TrueDown    =   "Form1.frx":121A
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Frame2"
         Height          =   2055
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   2415
         Begin HzxYOptionTest.HzxYOption opt 
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   1245
            _ExtentX        =   2355
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "HzxYOption5"
            BackColor       =   16744703
            Pic_FalseNormal =   "Form1.frx":14B0
            Pic_TrueNormal  =   "Form1.frx":1A0E
            Pic_FalseDisabled=   "Form1.frx":1F6C
            Pic_TrueDisabled=   "Form1.frx":24CA
            Pic_FalseOver   =   "Form1.frx":2A28
            Pic_TrueOver    =   "Form1.frx":2F86
            Pic_FalseDown   =   "Form1.frx":34E4
            Pic_TrueDown    =   "Form1.frx":3A42
         End
         Begin HzxYOptionTest.HzxYOption opt 
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   9
            Top             =   840
            Width           =   1245
            _ExtentX        =   2355
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "HzxYOption5"
            BackColor       =   16744703
            Pic_FalseNormal =   "Form1.frx":3FA0
            Pic_TrueNormal  =   "Form1.frx":44FE
            Pic_FalseDisabled=   "Form1.frx":4A5C
            Pic_TrueDisabled=   "Form1.frx":4FBA
            Pic_FalseOver   =   "Form1.frx":5518
            Pic_TrueOver    =   "Form1.frx":5A76
            Pic_FalseDown   =   "Form1.frx":5FD4
            Pic_TrueDown    =   "Form1.frx":6532
         End
         Begin HzxYOptionTest.HzxYOption opt 
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   10
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2355
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "HzxYOption5"
            BackColor       =   16744703
            Pic_FalseNormal =   "Form1.frx":6A90
            Pic_TrueNormal  =   "Form1.frx":6FEE
            Pic_FalseDisabled=   "Form1.frx":754C
            Pic_TrueDisabled=   "Form1.frx":7AAA
            Pic_FalseOver   =   "Form1.frx":8008
            Pic_TrueOver    =   "Form1.frx":8566
            Pic_FalseDown   =   "Form1.frx":8AC4
            Pic_TrueDown    =   "Form1.frx":9022
         End
      End
      Begin HzxYOptionTest.HzxYOption opt 
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1245
         _ExtentX        =   2355
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "HzxYOption5"
         BackColor       =   8409119
         Pic_FalseNormal =   "Form1.frx":9580
         Pic_TrueNormal  =   "Form1.frx":9ADE
         Pic_FalseDisabled=   "Form1.frx":A03C
         Pic_TrueDisabled=   "Form1.frx":A59A
         Pic_FalseOver   =   "Form1.frx":AAF8
         Pic_TrueOver    =   "Form1.frx":B056
         Pic_FalseDown   =   "Form1.frx":B5B4
         Pic_TrueDown    =   "Form1.frx":BB12
      End
      Begin HzxYOptionTest.HzxYOption opt 
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1245
         _ExtentX        =   2355
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "HzxYOption5"
         Pic_FalseNormal =   "Form1.frx":C070
         Pic_TrueNormal  =   "Form1.frx":C5CE
         Pic_FalseDisabled=   "Form1.frx":CB2C
         Pic_TrueDisabled=   "Form1.frx":D08A
         Pic_FalseOver   =   "Form1.frx":D5E8
         Pic_TrueOver    =   "Form1.frx":DB46
         Pic_FalseDown   =   "Form1.frx":E0A4
         Pic_TrueDown    =   "Form1.frx":E602
      End
      Begin HzxYOptionTest.HzxYOption opt 
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1245
         _ExtentX        =   2355
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "HzxYOption5"
         Pic_FalseNormal =   "Form1.frx":EB60
         Pic_TrueNormal  =   "Form1.frx":F0BE
         Pic_FalseDisabled=   "Form1.frx":F61C
         Pic_TrueDisabled=   "Form1.frx":FB7A
         Pic_FalseOver   =   "Form1.frx":100D8
         Pic_TrueOver    =   "Form1.frx":10636
         Pic_FalseDown   =   "Form1.frx":10B94
         Pic_TrueDown    =   "Form1.frx":110F2
      End
   End
   Begin HzxYOptionTest.HzxYOption opt 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2355
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "HzxYOption5"
      BackColor       =   16761087
      ForeColor       =   16711935
      Pic_FalseNormal =   "Form1.frx":11650
      Pic_TrueNormal  =   "Form1.frx":11BAE
      Pic_FalseDisabled=   "Form1.frx":1210C
      Pic_TrueDisabled=   "Form1.frx":1266A
      Pic_FalseOver   =   "Form1.frx":12BC8
      Pic_TrueOver    =   "Form1.frx":13126
      Pic_FalseDown   =   "Form1.frx":13684
      Pic_TrueDown    =   "Form1.frx":13BE2
   End
   Begin HzxYOptionTest.HzxYOption opt 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1245
      _ExtentX        =   2355
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "HzxYOption5"
      BackColor       =   8454016
      Pic_FalseNormal =   "Form1.frx":14140
      Pic_TrueNormal  =   "Form1.frx":1469E
      Pic_FalseDisabled=   "Form1.frx":14BFC
      Pic_TrueDisabled=   "Form1.frx":1515A
      Pic_FalseOver   =   "Form1.frx":156B8
      Pic_TrueOver    =   "Form1.frx":15C16
      Pic_FalseDown   =   "Form1.frx":16174
      Pic_TrueDown    =   "Form1.frx":166D2
   End
   Begin HzxYOptionTest.HzxYOption opt 
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1245
      _ExtentX        =   2355
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "HzxYOption5"
      BackColor       =   8454016
      Pic_FalseNormal =   "Form1.frx":16C30
      Pic_TrueNormal  =   "Form1.frx":1718E
      Pic_FalseDisabled=   "Form1.frx":176EC
      Pic_TrueDisabled=   "Form1.frx":17C4A
      Pic_FalseOver   =   "Form1.frx":181A8
      Pic_TrueOver    =   "Form1.frx":18706
      Pic_FalseDown   =   "Form1.frx":18C64
      Pic_TrueDown    =   "Form1.frx":191C2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    HzxYOption1.Value = Not HzxYOption1.Value
End Sub
Private Sub Command3_Click()
    HzxYOption1.Visible = Not HzxYOption1.Visible
End Sub
Private Sub Check1_Click()
    HzxYOption1.Enabled = (Check1.Value = Checked)
End Sub

Private Sub HzxYOption1_Click()
   HzxYOption1.Caption = ChrW$(&H6B22) & ChrW$(&H8FCE)
End Sub
