VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "ControlContainedControls"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Border"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Image"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Caption"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enabled"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin HzxYFrameTest.HzxYFrame HzxYFrame1 
      Height          =   2775
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
      _extentx        =   6588
      _extenty        =   4895
      font            =   "Form1.frx":0000
      image           =   "Form1.frx":0030
      Begin HzxYFrameTest.HzxYFrame HzxYFrame2 
         Height          =   1095
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
         _extentx        =   4260
         _extenty        =   1931
         font            =   "Form1.frx":090A
         Begin VB.CommandButton Command8 
            Caption         =   "Cool?"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "ImageWidth"
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ImageHeight"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    HzxYFrame1.ControlContainedControls = Check1.Value
End Sub

Private Sub Command1_Click()
    HzxYFrame1.Enabled = Not HzxYFrame1.Enabled
End Sub

Private Sub Command2_Click()
    If Trim(HzxYFrame1.Caption) <> "" Then
        HzxYFrame1.Caption = ""
    Else
        HzxYFrame1.Caption = "Cool?"
    End If
End Sub

Private Sub Command3_Click()
    If HzxYFrame1.Image Is Nothing Then
        Set HzxYFrame1.Image = LoadPicture(".\media\Delete.ico")
    Else
        Set HzxYFrame1.Image = Nothing
    End If
End Sub

Private Sub Command5_Click()
    HzxYFrame1.BorderStyle = IIf(HzxYFrame1.BorderStyle = fraNone, fraFixed_Single, fraNone)
End Sub

Private Sub Command6_Click()
    HzxYFrame1.ImageHeight = IIf(HzxYFrame1.ImageHeight = 16, 24, 16)
End Sub

Private Sub Command7_Click()
    HzxYFrame1.ImageWidth = IIf(HzxYFrame1.ImageWidth = 16, 24, 16)
End Sub

Private Sub HzxYFrame1_Click()
    HzxYFrame1.Caption = ChrW$(&H6B22) & ChrW$(&H8FCE)
End Sub
