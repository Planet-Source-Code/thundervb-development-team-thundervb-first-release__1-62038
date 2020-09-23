Attribute VB_Name = "modInit"
Option Explicit
'For hook creation/deletion

Dim cph As Long

Public Sub Init_Hook()
    Dim t As clsCpHook
    Set t = New clsCpHook
    cph = AddCPH(t)
    LogMsg "CPH added " & cph, PLUGIN_NAMEs, "modInit", "Init_Hook"
    
End Sub

Public Sub Unload_Hook()
    
    RemoveCPH cph
    LogMsg "CPH removed " & cph, PLUGIN_NAMEs, "modInit", "Unload_Hook"
    
End Sub

