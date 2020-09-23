VERSION 5.00
Begin VB.Form frmIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunLink"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctCredits 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.Label lblCredits 
         AutoSize        =   -1  'True
         Caption         =   "Credits - CodeExplorer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3705
      End
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    pctCredits.Move 0, 0
End Sub
