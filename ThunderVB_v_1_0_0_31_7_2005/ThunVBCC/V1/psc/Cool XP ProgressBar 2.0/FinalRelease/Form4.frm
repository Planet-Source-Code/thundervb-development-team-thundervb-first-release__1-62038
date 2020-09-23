VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Picture ProgressBar"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6675
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Use Percent Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   240
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   6
      Top             =   2640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   7
      Top             =   3240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   6
      Left            =   840
      TabIndex        =   8
      Top             =   3840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   7
      Left            =   840
      TabIndex        =   9
      Top             =   4440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   10
      Top             =   5040
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   9
      Left            =   840
      TabIndex        =   11
      Top             =   5640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   8
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2880
      TabIndex        =   0
      Top             =   6240
      Width           =   480
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MTime As Long
Dim i     As Long

Private Sub Check2_Click()
For i = XP.LBound To XP.UBound
         XP(i).ShowText = IIf(Check2.Value = 0, False, True)
Next i
End Sub

Private Sub Form_Load()
For i = XP.LBound To XP.UBound
        Set XP(i).Image = LoadPicture(App.Path & "\Pics\" & i + 1 & ".jpg")
Next i
End Sub

Private Sub Timer1_Timer()

MTime = MTime + 1

If MTime > XP(0).Max Then
    MTime = XP(0).Min
    
End If

For i = XP.LBound To XP.UBound
         XP(i).Value = MTime
Next i

LblValue = (100 * XP(0).Value) / XP(0).Max & " %"

End Sub
