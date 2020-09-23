VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   Caption         =   "Crash reports"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   3240
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Error Report"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Don't send Error Report"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   $"frmReport.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "An Error happened"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   $"frmReport.frx":0136
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TextHeader As String

Private Sub Command1_Click()
    If MsgBox("Do you want to save the error report to a file so " & vbNewLine & _
              "that you can send it to us later?", vbQuestion Or vbYesNo, _
              "Error report") = VbMsgBoxResult.vbYes Then
        'ok , save it to file..
        SaveReportToFile SelectFile(), TextHeader & vbNewLine & Text1.Text
    End If
    Me.Hide
End Sub

Private Sub Command2_Click()
    SendReportToServer Text1.Text
    Me.Hide
End Sub

Public Function SelectFile() As String

retry:
    With cd
        .Filter = ".gre files|*.gre"
        .FileName = ""
        .DialogTitle = "Select a filename to save the error report"
        .ShowSave
    End With
    
    If Len(cd.FileName) = 0 Then GoTo retry
    
    If FileExist(cd.FileName) Then
        If MsgBox("You want to overwrite the file ???", vbYesNo Or _
                  vbInformation, "Error report") = vbNo Then GoTo retry
    End If
    
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Command1_Click
End Sub

Public Sub ShowReport(info As String)
    TextHeader = info
    Me.Show vbModal
End Sub
