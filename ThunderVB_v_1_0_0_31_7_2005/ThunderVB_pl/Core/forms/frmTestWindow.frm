VERSION 5.00
Begin VB.Form frmTestWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmTestWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SM_CYCAPTION As Long = 4
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Dim confhWnd As Long, oldConfhWnd As Long

Public Sub ShowForm(caption As String, width As Long, height As Long, plugin As ThunderVB_pl_int_v1_0)
    
    With Me
        .caption = caption
        .width = width
        'PATCH - LIBOR
        'add height of caption
        .height = height + GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
    End With

    plugin.OnGuiLoad
    confhWnd = plugin.ShowConfig
    
    If confhWnd Then
        oldConfhWnd = SetParent(confhWnd, Me.hWnd)
            
        'PATCH - Libor
        'show form a modal
        Me.Show vbModal
'        Do
'            DoEvents
'            Sleep 10
'        Loop While Me.Visible = True
        SetParent confhWnd, oldConfhWnd
    End If
    
    plugin.HideCredits
    plugin.OnGuiUnLoad
    Unload Me
    
End Sub

