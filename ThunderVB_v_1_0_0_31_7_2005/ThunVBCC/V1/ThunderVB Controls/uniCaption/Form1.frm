VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin Project1.uniCaption wCaption 
      Left            =   720
      Top             =   4200
      _ExtentX        =   2672
      _ExtentY        =   635
   End
   Begin VB.CheckBox Check2 
      Caption         =   "showVB caption"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "show in taskbar"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "color"
      Height          =   1455
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   2415
      Begin VB.OptionButton Option6 
         Caption         =   "red"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "white"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option4 
         Caption         =   "green"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "align"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Option3 
         Caption         =   "right"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "center"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "left"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "unhook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "hook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim b As Boolean

Private Sub Command1_Click()
    Me.Caption = "blaheta"
    If b = False Then
        wCaption.Caption_Start Me.hwnd, wCaption.ARIAL_UNICODE_MS 'start subclassing
        b = True
    End If
End Sub

Private Sub Command2_Click()
    If b = True Then
        wCaption.Caption_stop 'end subclassing
        b = False
    End If
End Sub

Private Sub Form_Load()
Dim s As String, i As Long
    
    b = False
    SetAlign
    SetColor
    
    s = ""
    For i = 1 To 10
        s = s & ChrW(50000 + i * 10)
    Next i
    wCaption.Caption = s 'set unicode caption
End Sub

Private Sub SetAlign()
    If Option1.Value = True Then wCaption.CaptionAlign = TEXT_ALIGN.Bottom Or TEXT_ALIGN.Left 'set align
    If Option2.Value = True Then wCaption.CaptionAlign = TEXT_ALIGN.Bottom Or TEXT_ALIGN.center
    If Option3.Value = True Then wCaption.CaptionAlign = TEXT_ALIGN.Bottom Or TEXT_ALIGN.Right
End Sub

Private Sub SetColor()
    If Option4.Value = True Then wCaption.CaptionColor = vbGreen
    If Option5.Value = True Then wCaption.CaptionColor = vbWhite
    If Option6.Value = True Then wCaption.CaptionColor = vbRed
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Command2_Click
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        wCaption.TaskBar True  'show me in taskbar
    Else
        wCaption.TaskBar False
    End If
End Sub

Private Sub Check2_Click()
wCaption.ShowVbCaption = Check2.Value
End Sub

Private Sub Option1_Click()
SetAlign
End Sub

Private Sub Option2_Click()
Option1_Click
End Sub

Private Sub Option3_Click()
Option1_Click
End Sub

Private Sub Option4_Click()
SetColor
End Sub

Private Sub Option5_Click()
Option4_Click
End Sub

Private Sub Option6_Click()
Option4_Click
End Sub

