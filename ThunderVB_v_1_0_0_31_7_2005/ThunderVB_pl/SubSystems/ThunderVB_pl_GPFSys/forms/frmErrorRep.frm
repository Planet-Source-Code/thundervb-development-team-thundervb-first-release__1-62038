VERSION 5.00
Begin VB.Form frmErrorRep 
   Caption         =   "Unrecoverable Error"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Close Dialog"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSav 
      Caption         =   "Save Project(s)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"frmErrorRep.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Extended error info :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "[ext err str info]"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "[Note the VBide is/is not locked]"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   5775
   End
End
Attribute VB_Name = "frmErrorRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSav_Click()
    
    SaveProjects True
    
End Sub

Private Sub cmdStop_Click()
    
    Me.Hide
    
End Sub

Sub ShowForm(strFile As String)
On Error GoTo errH
    
    Me.Label3.caption = strFile
    Me.Label4 = "Note : The VB IDE IS NOT freesed"
    Me.Show
    Do
        Sleep 10
        DoEvents
    Loop While Me.Visible = True
Exit Sub

errH:
    Me.Label4 = "Note : The VB IDE is FREESED"
    Me.Show vbModal
    Resume Next
    
End Sub

