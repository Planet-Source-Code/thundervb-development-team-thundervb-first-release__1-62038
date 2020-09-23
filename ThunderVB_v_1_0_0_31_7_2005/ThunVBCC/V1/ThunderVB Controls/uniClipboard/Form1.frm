VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.uniClipboard uniClipboard1 
      Left            =   2040
      Top             =   2400
      _extentx        =   3122
      _extenty        =   635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1215
      VariousPropertyBits=   746604571
      Size            =   "2143;1720"
      FontHeight      =   165
      FontCharSet     =   238
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    uniClipboard1.SetText ChrW(51000)
    TextBox1.Text = uniClipboard1.GetText()
End Sub

