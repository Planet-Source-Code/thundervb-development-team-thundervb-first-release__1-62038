Attribute VB_Name = "modIDllt"
Option Explicit

Public Sub InstallDllTemplate()
Dim VBPath As String

    VBPath = GetRegistryKey("", "VisualBasic.Project\Shell\Make\Command", , 3)
    
    If Len(VBPath) = 0 Then
      MsgBox "Microsoft VB6 Path Could Not Be Found in the Registry:" & vbCrLf & vbCrLf & "HKEY_CLASSES_ROOT\VisualBasic.Project\Shell\Make\Command"
      Exit Sub
    End If

    VBPath = Left$(VBPath, InStr(1, VBPath, ".exe") + 3)
    VBPath = Left$(VBPath, InStrRev(VBPath, "\"))
    
    If MsgBox("You want to install dll template ?" & vbNewLine & _
              "[detected vb dir : " & VBPath & " ]", vbQuestion Or vbYesNo) = vbYes Then
        
        MsgBox "hehe not yet done"
        
    End If
              
End Sub
