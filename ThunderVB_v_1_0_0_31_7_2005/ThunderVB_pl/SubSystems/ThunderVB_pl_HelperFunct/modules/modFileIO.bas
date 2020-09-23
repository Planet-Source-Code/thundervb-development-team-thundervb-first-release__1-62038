Attribute VB_Name = "modFileIO"
Option Explicit

'loads a text file ands returns its contents as String
Public Function LoadFile_string(File As String) As String
    '<EhHeader>
    On Error GoTo LoadFile_Err
    '</EhHeader>
Dim ff As Long
    

If FileExist(File) = False Then
    ErrorBox "File does not exist (" & File & ")", "modFileIO", "LoadFile"
    Exit Function
End If
    

    
ff = FreeFile
Open File For Binary As ff
LoadFile_string = Space$(LOF(ff))
Get ff, , LoadFile_string
Close ff
    
    '<EhFooter>
    Exit Function

LoadFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "LoadFile"
    '</EhFooter>
End Function

'saves a text file
Public Sub SaveFile_string(File As String, Data As String)
    '<EhHeader>
    On Error GoTo SaveFile_Err
    '</EhHeader>
Dim ff As Long
    
    ff = FreeFile
    Open File For Output As ff
    Close ff
    Open File For Binary As ff
    Put ff, , Data
    Close ff
    
    '<EhFooter>
    Exit Sub

SaveFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "SaveFile"
    '</EhFooter>
End Sub

'loads a text file ands returns its contents as String
Public Function LoadFile_bin(File As String) As Byte()
    '<EhHeader>
    On Error GoTo LoadFile_Err
    '</EhHeader>
Dim ff As Long
    

If FileExist(File) = False Then
    ErrorBox "File dows not exist (" & File & ")", "modFileIO", "LoadFile"
    Exit Function
End If
    

    
ff = FreeFile
Open File For Binary As ff
    ReDim LoadFile_bin(LOF(ff) - 1)
    Get ff, , LoadFile_bin
Close ff
    
    '<EhFooter>
    Exit Function

LoadFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "LoadFile"
    '</EhFooter>
End Function

'saves a text file
Public Sub SaveFile_bin(File As String, Data() As Byte)
    '<EhHeader>
    On Error GoTo SaveFile_Err
    '</EhHeader>
Dim ff As Long
    
    ff = FreeFile
    Open File For Output As ff
    Close ff
    Open File For Binary As ff
    Put ff, , Data
    Close ff
    
    '<EhFooter>
    Exit Sub

SaveFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "SaveFile"
    '</EhFooter>
End Sub

'checks if a file exists
Public Function FileExist(File As String) As Boolean
    'well a crapy way to do it , but it works rather well...
    'i can't use dir$ , try dir$ with something like "c:\my.exe\" and see a wodnerfull crash..
    '<EhHeader>
    On Error GoTo FileExist_Err
    '</EhHeader>
    If A_GetFileAttributes(File) <> -1 Then FileExist = True
    
    '<EhFooter>
    Exit Function

FileExist_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "FileExist"
    '</EhFooter>
End Function

'checks if a file exists
Public Function DirExist(directory As String) As Boolean
    '<EhHeader>
    On Error GoTo DirExist_Err
    '</EhHeader>

    If Dir$(directory, vbDirectory) <> "" Then DirExist = True
    
    '<EhFooter>
    Exit Function

DirExist_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "DirExist"
    '</EhFooter>
End Function

'gets the filename from a full file path (eg "c:\windows\notepad.exe"->"notepad.exe")
Public Function GetFilename(filepath As String) As String
    '<EhHeader>
    On Error GoTo GetFilename_Err
    '</EhHeader>

    GetFilename = Split(filepath, "\")(ArrUBound(Split(filepath, "\")))
    
    '<EhFooter>
    Exit Function

GetFilename_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "GetFilename"
    '</EhFooter>
End Function

'gets the path from a full file path (eg "c:\windows\notepad.exe" ->"c:\windows\")
Public Function GetPath(filepath As String) As String
    '<EhHeader>
    On Error GoTo GetPath_Err
    '</EhHeader>

    GetPath = Replace(filepath, GetFilename(filepath), "")

    '<EhFooter>
    Exit Function

GetPath_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "GetPath"
    '</EhFooter>
End Function

'deletes a file if it exists..
Public Sub kill2(File As String)
    '<EhHeader>
    On Error GoTo kill2_Err
    '</EhHeader>

    If FileExist(File) Then Kill File

    '<EhFooter>
    Exit Sub

kill2_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modFileIO", "kill2"
    '</EhFooter>
End Sub

