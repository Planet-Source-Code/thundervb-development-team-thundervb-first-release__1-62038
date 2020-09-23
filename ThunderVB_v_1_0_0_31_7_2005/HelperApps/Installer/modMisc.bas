Attribute VB_Name = "modMisc"
Option Explicit
Public Const iniFile As String = "setup.ini"

Public Function getParam(PName As String, hfile As Long) As String
On Error GoTo Error

Dim STemp As String

'go to the first position
Seek hfile, 1

'search until the EOF for the desired string:
Do Until EOF(hfile)
    Line Input #hfile, STemp
    'if the first part of the read string equals the parameter,
    'we've found it
    If UCase(Left(STemp, Len(PName))) = UCase(PName) And Mid(STemp, Len(PName) + 1, 1) = "=" Then
        getParam = Mid(STemp, Len(PName) + 2)
        Exit Function
    End If
Loop

Error:

End Function

Public Function FileExists(Path As String) As Boolean
Dim temp As String

If Path = "" Then Exit Function

'try to Dir the current path. if the file exists, Dir returns
'its name, else it returns a null string
temp = Dir(Path)
If temp <> "" Then FileExists = True Else FileExists = False
End Function

Public Function GetSystemDir() As String
    Dim str As String * 1024
    GetSystemDirectory str, 1024
    GetSystemDir = Left(str, InStr(1, str, Chr$(0)) - 1)
End Function

