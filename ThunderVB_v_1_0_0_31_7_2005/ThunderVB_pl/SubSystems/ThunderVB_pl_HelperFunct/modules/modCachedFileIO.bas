Attribute VB_Name = "modCachedFileIO"
Option Explicit

'Made by Raziel(29/8/2004[dd/mm/yyyy]) .. ten days after string builder ;)
'Based on the StringBuilder
'Simple file io using a buffer to cache data....
'use it as you wish , gime a credit

'Moved Decalres to declares_pub

Sub PrintToFile(ByRef File As file_b, ByRef Data As String)
    '<EhHeader>
    On Error GoTo PrintToFile_Err
    '</EhHeader>
    
    With File
        AppendString .buf, Data & vbNewLine
        .buflen = .buflen + Len(Data & vbNewLine)
        
        If .buflen > .maxbuflen Then
            Put #.filenum, , GetString(.buf)
            .buf.str_index = 0
            .buflen = 0
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

PrintToFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modCachedFileIO", "PrintToFile"
    '</EhFooter>
End Sub

Sub AppendToFile(ByRef File As file_b, ByRef Data As String)
    '<EhHeader>
    On Error GoTo AppendToFile_Err
    '</EhHeader>
    
    With File
        AppendString .buf, Data
        .buflen = .buflen + Len(Data)
        
        If .buflen > .maxbuflen Then
            Put .filenum, , GetString(.buf)
            .buf.str_index = 0
            .buflen = 0
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

AppendToFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modCachedFileIO", "AppendToFile"
    '</EhFooter>
End Sub


Sub FlushFile(ByRef File As file_b)
    '<EhHeader>
    On Error GoTo FlushFile_Err
    '</EhHeader>

    With File
       
        If .buflen > 0 Then
            Put .filenum, , GetString(.buf)
            .buf.str_index = 0
            .buflen = 0
        End If
        
    End With

    '<EhFooter>
    Exit Sub

FlushFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modCachedFileIO", "FlushFile"
    '</EhFooter>
End Sub

Function OpenFile(ByRef filename As String, Optional ByVal buffersize As Long = 32768) As file_b
    '<EhHeader>
    On Error GoTo OpenFile_Err
    '</EhHeader>
Dim Temp As file_b
    
    Temp.filenum = FreeFile
    Temp.maxbuflen = buffersize
    Open filename For Binary As Temp.filenum
    OpenFile = Temp
    
    '<EhFooter>
    Exit Function

OpenFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modCachedFileIO", "OpenFile"
    '</EhFooter>
End Function

Sub CloseFile(ByRef File As file_b)
    '<EhHeader>
    On Error GoTo CloseFile_Err
    '</EhHeader>
Dim nullF As file_b
    
    FlushFile File
    Close File.filenum
    File = nullF
    
    '<EhFooter>
    Exit Sub

CloseFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modCachedFileIO", "CloseFile"
    '</EhFooter>
End Sub

Public Sub FileSeek(ByRef File As file_b, ByVal pos As Long)
    
    FlushFile File
    Seek File.filenum, pos
    
End Sub

Public Function FileLength(ByRef File As file_b) As Long
    
    FileLength = LOF(File.filenum) + File.buflen
    
End Function
