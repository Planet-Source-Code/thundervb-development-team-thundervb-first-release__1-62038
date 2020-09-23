VERSION 5.00
Begin VB.Form frmUnitest 
   Caption         =   "UniLabel test form"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   Icon            =   "frmUnitest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "AutoRedraw"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "WordWrap"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "AutoSize"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update using a string"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin UniLabelTest.UniLabel UniLabel1 
      Height          =   735
      Left            =   0
      Top             =   1440
      Width           =   4935
      _extentx        =   8705
      _extenty        =   1296
      alignment       =   2
      autoredraw      =   -1  'True
      backcolor       =   -2147483634
      captionb        =   "frmUnitest.frx":000C
      font            =   "frmUnitest.frx":003C
      forecolor       =   -2147483635
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2640
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update using byte array"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DrawTextW (to test support)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmUnitest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpArrPtr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Sub Check1_Click()
    UniLabel1.AutoSize = Check1.Value = vbChecked
    Form_Resize
End Sub
Private Sub Check2_Click()
    UniLabel1.WordWrap = Check2.Value = vbChecked
    Form_Resize
End Sub
Private Sub Check3_Click()
    UniLabel1.AutoRedraw = Check3.Value = vbChecked
End Sub
Private Sub Command1_Click()
    Dim Buffer() As Byte, MyRect As RECT
    ReDim Preserve Buffer(23)
    Buffer(0) = vbKeyI
    Buffer(2) = vbKeySpace
    Buffer(4) = vbKeyL
    Buffer(6) = vbKeyI
    Buffer(8) = vbKeyK
    Buffer(10) = vbKeyE
    Buffer(12) = vbKeySpace
    Buffer(14) = vbKeyB
    Buffer(16) = vbKeyY
    Buffer(18) = vbKeyT
    Buffer(20) = vbKeyE
    Buffer(22) = vbKeyS
    With MyRect
        .Left = 0
        .Top = 0
        .Right = Picture1.ScaleWidth
        .Bottom = Picture1.ScaleHeight
    End With
    Picture1.Cls
    DrawText Picture1.hdc, VarPtr(Buffer(0)), UBound(Buffer) \ 2 + 1, MyRect, vbNull
End Sub
Private Sub Command2_Click()
    Dim Buffer(3) As Byte
    Buffer(0) = &H41
    Buffer(1) = &H30
    Buffer(2) = &H42
    Buffer(3) = &H30
    UniLabel1.SetCaptionB Buffer
End Sub
Private Sub Command3_Click()
    UniLabel1.Caption = ChrW$(&H3043) & ChrW$(&H3044) & vbCrLf & _
        "Ooh!" & vbCrLf & "It works with very very very long lines too!"
End Sub
Private Sub Form_Load()
    Check1.Value = Abs(UniLabel1.AutoSize)
    Check2.Value = Abs(UniLabel1.WordWrap)
    Check3.Value = Abs(UniLabel1.AutoRedraw)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    UniLabel1.Move 0, UniLabel1.Top, ScaleWidth, ScaleHeight - UniLabel1.Top
End Sub

