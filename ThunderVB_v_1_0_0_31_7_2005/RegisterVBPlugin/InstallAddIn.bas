Attribute VB_Name = "InstallAddIn"
Option Explicit

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Const ADDIN_NAME = "ThunderVB_pl_v1_0"

'To define the caption that appears in the Add-In Manager window go to the
'Object Browser (F2), select clsConnect, right click, select "Properties ..."
'VB's "Member Options" dialog should appear.  In the "Description" text box
'enter the caption you want to appear in the Add-Manager window.

Sub Main()
    Dim sExternalError As String
    If AddToINI(sExternalError) Then
        MsgBox "Add-In called """ & ADDIN_NAME & """ has been installed."
    Else
        MsgBox "Failed to install add-in: " & sExternalError
    End If
End Sub

'This procedure must be executed before VB's Add-In Manager will
'recognize the add-in as available.  Normally the procedure should be
'executed by the setup program.  During program development you will need
'to run it once in the immediate window to make the add-in available in
'your local environment.
Function AddToINI(sError As String) As Boolean
    Dim lngErrorCode As Long, lngErrorValue As Long
    On Error GoTo EH
    lngErrorValue = WritePrivateProfileString("Add-Ins32", ADDIN_NAME & ".Connect", "1", "vbaddin.ini")
    If lngErrorValue = 0 Then
        lngErrorCode = GetLastError
        sError = "WritePrivateProfileString generated error code: " & lngErrorCode
    Else
        AddToINI = True
    End If
    Exit Function
EH:
    sError = "Unexpected error writing private profile string to vbaddin.ini: " & Err.Description
End Function

Function GetSystemDir() As String
Dim str As String * 1024
GetSystemDirectory str, 1024
GetSystemDir = Left(str, InStr(1, str, Chr$(0)) - 1)
End Function
