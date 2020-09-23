Attribute VB_Name = "Program"
'(c) 2003 Enrico Bertozzi
' I don't remember any of the copyright of the two modules
' (registry and "LaunchAppSynchronousMod"). hopefully someone
' will find these modules and claim his copyright...

Public Sub GegisterDll(filepath As String)

        ShellExecAndWait GetSystemDir & "\regsvr32.exe", "/s """ & filepath & """, True, msg"
                
End Sub
