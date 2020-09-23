VERSION 5.00
Begin VB.Form frmSugestion 
   Caption         =   "Make a sugestion/Report a bug"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send Feedback"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Here you can do a sugestion / Report a bug"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmSugestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    SaveFile App.Path & "\tmpsug.txt", Text1.Text
    MsgBox "Uploading Files please wait", vbInformation Or vbOKOnly
    MsgBox UploadFiles("http://thundervb.sourceforge.net/bugreports/getfile.php", "", "", "", App.Path & "\tmpsug.txt")
    kill2 App.Path & "\tmpsug.txt"
    Me.Hide
    
End Sub

Public Sub ShowModal()

    Me.Show vbModal
    
End Sub
