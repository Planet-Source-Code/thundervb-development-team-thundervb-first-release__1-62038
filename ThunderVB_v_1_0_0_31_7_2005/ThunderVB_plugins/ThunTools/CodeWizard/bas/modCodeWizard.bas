Attribute VB_Name = "modCodeWizard"
Option Explicit

Public oThunVB As ThunderVB_base
Public oMe As plugin

Public Const PLUGIN_NAME As String = "ASM Code Wizard"
Public Const MSG_TITLE As String = PLUGIN_NAME

Public Const PLUGIN_NAMEs As String = "CodeWizard"
Public Const MSG_TITLEs As String = PLUGIN_NAMEs

Public hButtonID As Long

'add plugin to menu
Public Sub AddMe2Menu()
    
    If hButtonID = 0 Then hButtonID = oThunVB.Add2ToolBar(oMe, PLUGIN_NAMEs, PLUGIN_NAMEs, frmCodeWizard.picBut.Picture)
    
End Sub

'remove icon from menu
Public Sub RemoveMeFromMenu()
    
    If hButtonID <> 0 Then
        oThunVB.RemoveButtton hButtonID
        hButtonID = 0
    End If
    
End Sub
