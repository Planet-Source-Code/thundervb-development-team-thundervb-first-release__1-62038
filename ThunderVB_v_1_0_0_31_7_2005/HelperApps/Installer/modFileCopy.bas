Attribute VB_Name = "modFileCopy"
Option Explicit


Public Sub CopyFiles()
Dim ff As Long, i As Long, DestDir As String, file As String, GroopName As String, ncount As Long

    ff = FreeFile
    
    Open iniFile For Input As ff
        ncount = getParam("files", ff)
        For i = 0 To ncount - 1
            file = getParam("file_name_" & i, ff)
            GroopName = getParam("file_groop_" & i, ff)
            DestDir = getParam(GroopName & "_destdir", ff)
            frmWiz.SetStatus "filecopy", "Groop " & GroopName & ";File " & file & " , size=" & FileLen(file), i / ncount
            FileCopy file, DestDir & "\" & file
            
            If Len(getParam(file & "_comreg", ff)) > 0 Then
                GegisterDll DestDir & "\" & file
            End If
            
        Next i
        frmWiz.SetStatus "filecopy", "Groop " & GroopName & ";File " & file & " , size=" & FileLen(file), i / ncount
        frmWiz.SetStatus "filecc"
    Close ff
    
End Sub
