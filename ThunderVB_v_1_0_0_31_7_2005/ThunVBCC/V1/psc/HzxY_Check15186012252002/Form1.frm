VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin HzxYCheckBoxTest.HzxYCheckBox HzxYCheckBox3 
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   4845
      _ExtentX        =   8546
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
      Value           =   0
      Caption         =   "Enable HzxYCheckBox1"
      Pic_UncheckedNormal=   "Form1.frx":0000
      Pic_CheckedNormal=   "Form1.frx":0352
      Pic_MixedNormal =   "Form1.frx":06A4
      Pic_UncheckedDisabled=   "Form1.frx":09F6
      Pic_CheckedDisabled=   "Form1.frx":0D48
      Pic_MixedDisabled=   "Form1.frx":109A
      Pic_UncheckedOver=   "Form1.frx":13EC
      Pic_CheckedOver =   "Form1.frx":173E
      Pic_MixedOver   =   "Form1.frx":1A90
      Pic_UncheckedDown=   "Form1.frx":1DE2
      Pic_CheckedDown =   "Form1.frx":2134
      Pic_MixedDown   =   "Form1.frx":2486
   End
   Begin HzxYCheckBoxTest.HzxYCheckBox HzxYCheckBox2 
      Height          =   240
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1485
      _ExtentX        =   6747
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
      Pic_UncheckedNormal=   "Form1.frx":27D8
      Pic_CheckedNormal=   "Form1.frx":2B2A
      Pic_MixedNormal =   "Form1.frx":2E7C
      Pic_UncheckedDisabled=   "Form1.frx":31CE
      Pic_CheckedDisabled=   "Form1.frx":3520
      Pic_MixedDisabled=   "Form1.frx":3872
      Pic_UncheckedOver=   "Form1.frx":3BC4
      Pic_CheckedOver =   "Form1.frx":3F16
      Pic_MixedOver   =   "Form1.frx":4268
      Pic_UncheckedDown=   "Form1.frx":45BA
      Pic_CheckedDown =   "Form1.frx":490C
      Pic_MixedDown   =   "Form1.frx":4C5E
   End
   Begin HzxYCheckBoxTest.HzxYCheckBox HzxYCheckBox1 
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1485
      _ExtentX        =   6826
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
      Enabled         =   0   'False
      Pic_UncheckedNormal=   "Form1.frx":4FB0
      Pic_CheckedNormal=   "Form1.frx":5302
      Pic_MixedNormal =   "Form1.frx":5654
      Pic_UncheckedDisabled=   "Form1.frx":59A6
      Pic_CheckedDisabled=   "Form1.frx":5CF8
      Pic_MixedDisabled=   "Form1.frx":604A
      Pic_UncheckedOver=   "Form1.frx":639C
      Pic_CheckedOver =   "Form1.frx":66EE
      Pic_MixedOver   =   "Form1.frx":6A40
      Pic_UncheckedDown=   "Form1.frx":6D92
      Pic_CheckedDown =   "Form1.frx":70E4
      Pic_MixedDown   =   "Form1.frx":7436
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HzxYCheckBox3_Click()
    HzxYCheckBox1.Enabled = (HzxYCheckBox3.Value = Checked)
    HzxYCheckBox1.Caption = ChrW$(&H6B22) & ChrW$(&H8FCE)
End Sub
