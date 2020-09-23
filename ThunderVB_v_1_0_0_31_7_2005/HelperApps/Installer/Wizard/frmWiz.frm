VERSION 5.00
Begin VB.Form frmWiz 
   Caption         =   "%app%"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next ->"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<-Back"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.PictureBox ctb 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   1
      Left            =   480
      ScaleHeight     =   4665
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   2040
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   120
         Top             =   1920
         Width           =   5550
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   120
         Top             =   1920
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "Copying Files"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.PictureBox ctb 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   0
      Left            =   0
      ScaleHeight     =   4665
      ScaleWidth      =   5745
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.Label Label2 
         Caption         =   "%loa%"
         Height          =   2415
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Wellcome to %app% setup program"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curTab As Long
Dim lback As Long
Public Sub SetCurTab(tb As Long)
    
    curTab = tb
    
    If curTab > ctb.UBound Then
        End
    End If
    
    If curTab = lback Then
        cmdBack.Enabled = False
    Else
        cmdBack.Enabled = True
    End If
    
    If curTab = ctb.UBound Then
        cmdNext.Caption = "Finish"
    Else
        cmdNext.Caption = "Next ->"
    End If
    
    TabChanged
    
End Sub

Private Sub cmdBack_Click()
    Dim bc As Boolean
    
    bc = True
    Setup_PrevStep bc, True
    
    If bc Then
        SetCurTab curTab - 1
        Setup_PrevStep bc, False
    End If
    
End Sub

Private Sub cmdNext_Click()
    
    Dim bc As Boolean
    
    bc = True
    Setup_NextStep bc, True
    
    If bc Then
        SetCurTab curTab + 1
        Setup_NextStep bc, False
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = ctb.LBound To ctb.UBound
        ctb(i).BorderStyle = 0
        ctb(i).Top = 0
        ctb(i).Left = 0
        ctb(i).Visible = False
    Next i
    
    Width = ctb(1).Width + 125
    Height = ctb(1).Height + 40
    
    SetCurTab 0
    
End Sub

Public Sub SetStatus(ParamArray p() As Variant)
    
    Select Case CStr(p(0))
        
        
        Case "filecopy"
        Case "filecc"
        
        
    End Select

End Sub

Sub TabChanged()
Dim i As Long

    For i = ctb.LBound To ctb.UBound
        If i <> curTab Then
            ctb(i).Visible = False
        Else
            ctb(i).Visible = True
        End If
    Next i

End Sub

Public Sub Setup_NextStep(ByRef cancel As Boolean, ByVal before As Boolean)
    If before Then
    Else
        If curTab = 1 Then
        CopyFiles
        End If
    End If
End Sub

Public Sub Setup_PrevStep(ByRef cancel As Boolean, ByVal before As Boolean)

End Sub
