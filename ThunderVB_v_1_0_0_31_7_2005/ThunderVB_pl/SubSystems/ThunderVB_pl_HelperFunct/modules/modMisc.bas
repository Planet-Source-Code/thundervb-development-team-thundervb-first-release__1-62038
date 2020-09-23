Attribute VB_Name = "modMisc"
Option Explicit

'Low Level helper functions...

'Revision history:
'
'19/8/2004[dd/mm/yyyy] : Creted by Libor
'Module created , intial version
'
'
'19/8/2004 [dd/mm/yyyy] : Edited by Raziel
'All the delcarations were made public and moved to the declaration module (\misc code\declares.bas)
'Many things added here , everything is marked
'
'
'22/9/2004[dd/mm/yyyy] : Edited by Raziel
'Fixes hre and there , mainly on string convertion
'
'
'13/4/2005[dd/mm/yyyy]  : Edited by Raziel
'Moved to new project (ThunderVB_pl_HelperFunct) , many changes

Private Declare Function CreateProcessLong2 Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CreateProcessLong Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetProcessVersion Lib "kernel32.dll" (ByVal ProcessId As Long) As Long
Private Declare Function lenCString Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function CopyCString Lib "kernel32" Alias "lstrcpynA" (ByVal lpStringDestination As String, ByVal lpStringSource As Long, ByVal lngMaxLength As Long) As Long

'Code from VbInlineASM
Public Function ExecuteCommand(ByVal CommandLine As String, ByRef sOutputText As String, Optional workdir As String, Optional ByVal eWindowState As ESW = SW_HIDE) As Boolean
'***I have patched CreateProcessLong declaration (it was wrong) - lpEnvironment As Any - should be - Byval lpEnvironment as Long *** Libor - 2004
'***I've modified this function too....JS-2002
'*** Edited by Raziel for workdir
'DOSOutpus
'Capture the outputs of a DOS command
'Author: Marco Pipino
'marcopipino@libero.it
'28/02/2002
    '<EhHeader>
    On Error GoTo ExecuteCommand_Err
    '</EhHeader>
Dim proc As PROCESS_INFORMATION     'Process info filled by CreateProcessA
Dim ret As Long                     'long variable for get the return value of the API functions
Dim start As STARTUPINFO           'StartUp Info passed to the CreateProceeeA function
Dim sa As SECURITY_ATTRIBUTES       'Security Attributes passeed to the CreateProcessA function
Dim hReadPipe As Long               'Read Pipe handle created by CreatePipe
Dim hWritePipe As Long              'Write Pite handle created by CreatePipe
Dim lngBytesread As Long            'Amount of byte read from the Read Pipe handle
Dim baBuff(255) As Byte          'String buffer reading the Pipe
        
    If Len(CommandLine) = 0 Then
        ExecuteCommand = False
        Exit Function
    End If
    LogMsg CommandLine, "modLowLevel", "ExecuteCommand"
    On Error Resume Next
    
    'Create the Pipe
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)

    If ret = 0 Then
        'If an error occur during the Pipe creation exit
        'msgboxx "CreatePipe failed. Error: " & Err.LastDllError, vbCritical
        Exit Function
    End If

    'Launch the command line application
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    start.wShowWindow = eWindowState
    
    'set the StdOutput and the StdError output to the same Write Pipe handle
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
       
    'Execute the command
    If Len(workdir) > 0 Then
        ret& = CreateProcessLong2(0&, CommandLine, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, ByVal workdir, start, proc)
    Else
        ret& = CreateProcessLong(0&, CommandLine, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, 0&, start, proc)
    End If
    
    If ret <> 1 Then
        'if the command is not found ....
        Exit Function
    End If

    'Now We can ... must close the hWritePipe
    ret = CloseHandle(hWritePipe)
    sOutputText = ""                             '*** patched LIBOR
    
    'Read the ReadPipe handle
    Do
        ret = ReadFile(ByVal hReadPipe, baBuff(0), ByVal 256, lngBytesread, ByVal 0&)
        sOutputText = sOutputText & Left$(StrConv(baBuff, vbUnicode), lngBytesread)
        'Send data to the object via ReceiveOutputs event
    Loop While ret <> 0

    'Close the opened handles
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)

    ExecuteCommand = True
    
    '<EhFooter>
    Exit Function

ExecuteCommand_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modMisc", "ExecuteCommand"
    '</EhFooter>
End Function


'***Added by Raziel [18/9/2004]
'hmm not all of em are so lowlevel , maybe change the name to common and
'move here all helper funcs??

'waits for a proccess to end
'yeh , prety much smart way ehh ?? lololololol..
'bored to replace it with waitforsignleobject ...
Public Sub WaitToEnd(proc As Long)
    '<EhHeader>
    On Error GoTo WaitToEnd_Err
    '</EhHeader>


    Do
        Sleep 100
        DoEvents
    Loop While GetProcessVersion(proc) <> 0

    '<EhFooter>
    Exit Sub

WaitToEnd_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modMisc", "WaitToEnd"
    '</EhFooter>
End Sub

'Get VB version
Function getVBVersion() As Long
    '<EhHeader>
    On Error GoTo getVBVersion_Err
    '</EhHeader>
Dim Temp As Long

    If A_GetModuleHandle("VBA5.dll") <> 0 Then Temp = 5
    If A_GetModuleHandle("VBA6.dll") <> 0 Then Temp = 6
    getVBVersion = Temp
    LogMsg "VB version=" & Temp, "modLowLevel", "getVBVersion"
    
    '<EhFooter>
    Exit Function

getVBVersion_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modMisc", "getVBVersion"
    '</EhFooter>
End Function

'From VBInlineAsm
'Copies a Cstring to a VB string - > will be replaced..
Function CStringZero(lpCString As Long) As String
    '<EhHeader>
    On Error GoTo CStringZero_Err
    '</EhHeader>
Dim lenString As Long, sBuffer As String, lpBuffer As Long, lngStringPointer As Long, refStringPointer As Long

    If lpCString = 0 Then
        CStringZero = vbNullString
    Else
        lenString = lenCString(lpCString)
        sBuffer = String$(lenString + 1, 0) 'buffer has one extra byte for terminator
        lpBuffer = CopyCString(sBuffer, lpCString, lenString + 1)
        Mid$(sBuffer, lenString + 1, 1) = " " ' to fix the 0 at the end
        CStringZero = sBuffer
    End If
    
    '<EhFooter>
    Exit Function

CStringZero_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modMisc", "CStringZero"
    '</EhFooter>
End Function


'From ansi String Pointer to vb string
Function Cstring(ByVal lpString As Long, ByVal nCount As Long) As String
    '<EhHeader>
    On Error GoTo Cstring_Err
    '</EhHeader>
Dim s As String, Temp() As Byte

    If nCount > 0 Then
        ReDim Temp(nCount)
        CopyMemory Temp(0), ByVal lpString, nCount
        Temp(nCount) = 0
        s = StrConv(Temp, vbUnicode)
    End If
    Cstring = s
    
    '<EhFooter>
    Exit Function

Cstring_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modMisc", "Cstring"
    '</EhFooter>
End Function

'From Wide String Pointer to vb string
Function CstringW(ByVal lpString As Long, ByVal nCount As Long) As String
    '<EhHeader>
    On Error GoTo CstringW_Err
    '</EhHeader>
Dim s As String, Temp() As Byte
    
    nCount = nCount * 2 - 1
    If nCount > 0 Then
        ReDim Temp(nCount)
        CopyMemory Temp(0), ByVal lpString, nCount
        Temp(nCount) = 0
        s = Temp
    End If
    CstringW = s
    
    '<EhFooter>
    Exit Function

CstringW_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modMisc", "CstringW"
    '</EhFooter>
End Function

'Form VB str to Ansi Byte array
Public Sub BstrToAnsi(str As String, ba() As Byte)
    '<EhHeader>
    On Error GoTo BstrToAnsi_Err
    '</EhHeader>

    If Len(str) = 0 Then
        ReDim ba(0)
    Else
        ba = StrConv(str, vbFromUnicode)
        ReDim Preserve ba(ArrUBound(ba) + 1)
    End If
    
    '<EhFooter>
    Exit Sub

BstrToAnsi_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modMisc", "BstrToAnsi"
    '</EhFooter>
End Sub

'Arrays..
Public Function ArrUBound(arr As Variant) As Long
    
    If IsArray(arr) = False Then
        MsgBox "Not An Array"
        ArrUBound = -1
        Exit Function
    Else
    
        On Error GoTo err
        SetEhMode Err_expected  'we are expecting an eror to happen .. it is a part of teh function xD
        ArrUBound = UBound(arr)
        RestoreEhMode           'restore error handling mode..
        Exit Function
err:
        ArrUBound = -1
        RestoreEhMode           'restore error handling mode..
    End If
    
End Function

Public Function ArrLBound(arr As Variant) As Long
        
    If IsArray(arr) = False Then
        MsgBox "Not An Array"
        ArrLBound = 2147483647
        Exit Function
    Else
        On Error GoTo err
        SetEhMode Err_expected  'we are expecting an eror to happen .. it is a part of teh function xD
        ArrLBound = LBound(arr)
        RestoreEhMode           'restore error handling mode..
        Exit Function
err:
        ArrLBound = 2147483647 'max int value on vb
        RestoreEhMode           'restore error handling mode..
    End If
    
End Function

