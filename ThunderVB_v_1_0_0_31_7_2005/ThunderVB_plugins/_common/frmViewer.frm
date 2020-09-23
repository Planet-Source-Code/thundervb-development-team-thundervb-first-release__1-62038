VERSION 5.00
Begin VB.Form frmViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "caption"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   500
      Left            =   1750
      TabIndex        =   3
      Top             =   4900
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   500
      Left            =   3300
      TabIndex        =   2
      Top             =   4900
      Width           =   1800
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   500
      Left            =   100
      TabIndex        =   1
      Top             =   4920
      Width           =   1800
   End
   Begin VB.TextBox txtData 
      Height          =   4700
      Left            =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   100
      Width           =   5000
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sData As String, sOData As String

'sCaption - form's caption
'sText    - TextBox's text
'bViewer  - True  - only show text
'         - False - user could change text
'bCanRetNull - True - If cancel/close then return null
'            - Fasle - If cancel/close then return original value
'
'return   - ""   - when bViewer = False or user hits Cancel button
'         - text - when bViewer = True and user hits Save button

Public Function ShowViewer(sCaption As String, sText As String, Optional bViewer As Boolean = True, Optional bCanRetNull As Boolean) As String
          
       'change caption
       frmViewer.caption = sCaption
          
       'un/lock text box
       txtData.Locked = bViewer
       'set text
       txtData.Text = sText
          
       If bCanRetNull = False Then
           sOData = sText
       Else
           sOData = vbNullString
       End If
          
       'if we only want to show text
       If bViewer = True Then
          
           'show close button
          cmdClose.Visible = True
              
           'other make unvisible
          cmdSave.Visible = False
          cmdCancel.Visible = False
              
      Else
          
           'make unvisible
          cmdClose.Visible = False
              
           'make visible
          cmdSave.Visible = True
          cmdCancel.Visible = True
              
      End If
          
      sData = sOData
      'Me.Show vbModal
      Me.Show
      
      While Me.Visible = True
        Me.Show
        DoEvents
        Sleep (10)
      Wend
              
      ShowViewer = sData

End Function

'user does not want to change text
Private Sub cmdCancel_Click()
       Me.Hide
End Sub

'user close window
Private Sub cmdClose_Click()
        Me.Hide
End Sub

'user changes text
Private Sub cmdSave_Click()
       sData = txtData.Text
       Me.Hide
End Sub

Private Sub Form_Activate()
       LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Activate"
End Sub

'catch unload event
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
       Call cmdCancel_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
       LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Unload"
End Sub
