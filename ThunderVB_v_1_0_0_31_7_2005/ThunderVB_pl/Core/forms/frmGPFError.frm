VERSION 5.00
Begin VB.Form frmGPFError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "An unhandled error ocured...."
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSav 
      Caption         =   "Save Project(s)"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop execution"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCont 
      Caption         =   "Continue"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdRasieErr 
      Caption         =   "RaiseError"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "[Note the VBide is/is not locked]"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "[ext err str info]"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Extended error info :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   $"frmGPFError.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmGPFError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hf As GPF_actions

Public Function ShowGPF(str As String) As GPF_actions
    '<EhHeader>
    On Error GoTo ShowGPF_Err
    '</EhHeader>
On Error GoTo errH
    
    Me.Label3.caption = str
    Me.Label4 = "Note : The VB IDE IS NOT freesed"
    Me.Show
    Do
        Sleep 10
        DoEvents
    Loop While Me.Visible = True
    ShowGPF = hf
Exit Function

errH:
    Me.Label4 = "Note : The VB IDE is FREESED"
    Me.Show vbModal
    Resume Next
    
    '<EhFooter>
    Exit Function

ShowGPF_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "frmGPFError", "ShowGPF"
    '</EhFooter>
End Function

Private Sub cmdCont_Click()
    '<EhHeader>
    On Error GoTo cmdCont_Click_Err
    '</EhHeader>

    hf = GPF_Cont
    Me.Hide
    
    '<EhFooter>
    Exit Sub

cmdCont_Click_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "frmGPFError", "cmdCont_Click"
    '</EhFooter>
End Sub

Private Sub cmdRasieErr_Click()
    '<EhHeader>
    On Error GoTo cmdRasieErr_Click_Err
    '</EhHeader>

    hf = GPF_RaiseErr
    Me.Hide
    
    '<EhFooter>
    Exit Sub

cmdRasieErr_Click_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "frmGPFError", "cmdRasieErr_Click"
    '</EhFooter>
End Sub

Private Sub cmdSav_Click()
    '<EhHeader>
    On Error GoTo cmdSav_Click_Err
    '</EhHeader>
    
    SaveProjects True
    
    '<EhFooter>
    Exit Sub

cmdSav_Click_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "frmGPFError", "cmdSav_Click"
    '</EhFooter>
End Sub

Private Sub cmdStop_Click()
    '<EhHeader>
    On Error GoTo cmdStop_Click_Err
    '</EhHeader>

    hf = GPF_Stop
    Me.Hide
    
    '<EhFooter>
    Exit Sub

cmdStop_Click_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "frmGPFError", "cmdStop_Click"
    '</EhFooter>
End Sub

