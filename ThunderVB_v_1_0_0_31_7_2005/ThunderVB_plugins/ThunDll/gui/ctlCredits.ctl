VERSION 5.00
Begin VB.UserControl ctlCredits 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Credits - ThunDll"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2715
   End
End
Attribute VB_Name = "ctlCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetHwnd() As Long
    GetHwnd = UserControl.hWnd
End Function
