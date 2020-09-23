VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Other Example"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6600
      Top             =   0
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Standard Scrolling"
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   2880
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Smooth Scrolling"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Search Style Scrolling"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   0
      Top             =   3840
      Width           =   2175
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarSilver 
      Height          =   4455
      Left            =   2640
      TabIndex        =   7
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7858
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
      Orientation     =   1
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarOlive 
      Height          =   4455
      Left            =   1680
      TabIndex        =   8
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7858
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
      Orientation     =   1
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarBlue 
      Height          =   4455
      Left            =   720
      TabIndex        =   9
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7858
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
      Orientation     =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Silver XP"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Olive XP"
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Blue XP"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MTime  As Long

Const XPBlue_ProgressBar = &H2BD228
Const XPOlive_ProgressBar = &H4A86E4
Const XPSilver_ProgressBar = &H76AE83

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

XP_ProgressBarBlue.Max = 100
XP_ProgressBarBlue.Min = 1
XP_ProgressBarBlue.Orientation = ccOrientationVertical

XP_ProgressBarOlive.Max = 100
XP_ProgressBarOlive.Min = 1
XP_ProgressBarOlive.Orientation = ccOrientationVertical

XP_ProgressBarSilver.Max = 100
XP_ProgressBarSilver.Min = 1
XP_ProgressBarSilver.Orientation = ccOrientationVertical

XP_ProgressBarBlue.Color = XPBlue_ProgressBar
XP_ProgressBarOlive.Color = XPOlive_ProgressBar
XP_ProgressBarSilver.Color = XPSilver_ProgressBar

End Sub

Private Sub Option1_Click(Index As Integer)

If Index = 2 Then
    Timer1.Interval = 20
Else                         'Time Interval yust to Show Search Style Demo.
    Timer1.Interval = 100
End If

XP_ProgressBarBlue.Scrolling = Index
XP_ProgressBarOlive.Scrolling = Index
XP_ProgressBarSilver.Scrolling = Index

End Sub

Private Sub Timer1_Timer()
MTime = MTime + 1

If MTime > XP_ProgressBarBlue.Max Then
    MTime = XP_ProgressBarBlue.Min
End If

XP_ProgressBarBlue.Value = MTime
XP_ProgressBarOlive.Value = MTime
XP_ProgressBarSilver.Value = MTime

End Sub

