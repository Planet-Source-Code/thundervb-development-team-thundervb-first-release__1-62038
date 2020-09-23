VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug Log"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AppendLog(str As String)
    
    If Me.Visible = False Then
        Me.Show
    End If

    List1.AddItem str
    On Error Resume Next
    List1.TopIndex = List1.ListCount
    
End Sub



Private Sub List1_DblClick()
    MsgBox List1.text
End Sub
