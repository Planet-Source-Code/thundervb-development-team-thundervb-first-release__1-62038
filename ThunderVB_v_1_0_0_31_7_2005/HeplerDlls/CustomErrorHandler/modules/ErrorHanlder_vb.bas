Attribute VB_Name = "ErrorHanlder_vb"
Option Explicit
Public th As DllHook_list
Dim vbaExH_fp As Long

     
Declare Function GetCurrentThread Lib "kernel32.dll" () As Long

Global EhModeI As Long
Global ErrHmode As EhModes
Global inited As Boolean
Global ErrHmodes() As EhModes

Public Const APP_NAME As String = "CustomErrorHandler"
Public Const APP_NAMEs As String = "CEH"
Dim CsStr As String

#Const debug_ver = False

Public Sub SetexeptH()

    ReDim ErrHmodes(0)
    inited = True

    If IsAsmOn() Then
        CreateHookList_ResumeOldList "MSVBVM60.DLL", "__vbaExceptHandler", AddressOf ExeptH, th
        InitExept th.items(0).FunctionAddress
        LogMsg "Error handling hook was inited", "ErrorHanlder_vb", "SetexeptH"
    End If

End Sub

     
Public Sub ResetexeptH()
    
    If IsAsmOn Then
        KillHookList th
        LogMsg "Error handling hook was disabled", "ErrorHanlder_vb", "SetexeptH"
    End If
    
End Sub




'__cdecl _except_handler(
'     struct _EXCEPTION_RECORD *ExceptionRecord,
'     void * EstablisherFrame,
'     struct _CONTEXT *ContextRecord,
'     void * DispatcherContext

Public Function ExeptH_bef(ByVal vbhmode As Long, ByVal val1 As Long, ByVal val2 As Long, ByVal val3 As Long, ByVal val4 As Long) As Long
Dim terr As String

    #If debug_ver Then
        LogMsg "An error was raised from " & Eip2Mod(GetEip(val3)) & " , status is " & vbhmode & ":" & ErrHmode, "ErrorHanlder_vb", "ExeptH_bef"
        LogMsg "Error Info,:" & Err.Description & ";#=" & Err.Number & "; source =" & Err.Source, "ErrorHanlder_vb", "ExeptH_bef"
    #Else
        If vbhmode <> -1 And vbhmode > 10 Then
            LogMsg "An error was raised from " & Eip2Mod(GetEip(val3)) & " , status is " & vbhmode & ":" & ErrHmode, "ErrorHanlder_vb", "ExeptH_bef"
            LogMsg "Error Info,:" & Err.Description & ";#=" & Err.Number & "; source =" & Err.Source, "ErrorHanlder_vb", "ExeptH_bef"
        End If
    #End If
    
    If ErrHmode = Err_expected Then
        ExeptH_bef = 1
        Exit Function
    ElseIf ErrHmode = Err_expected_AutoRestore Then
        ExeptH_bef = 1
        RestoreEhMode
        Exit Function
    End If
    
    If vbhmode <> 0 Then
        #If debug_ver Then
        If (vbhmode < 100) Then   'it is not comon to have > 100 error handler on
        #End If                   'the same sub hanlded ... is it ?
            ExeptH_bef = 1
            Exit Function
        #If debug_ver Then
        End If
        #End If
    Else
        
    End If
    
    #If debug_ver Then
        CallStackDump val3
        If MsgBox("[this info may be correct] Error # : " & Err.Number & vbNewLine & _
                  Err.Description & vbNewLine & _
                  Err.Source & vbNewLine & _
                  Eip2Mod(GetEip(val3)) & vbNewLine & _
                  "You want to write a dump ?", vbYesNo) = vbYes Then
            Dim t(1) As Long
            t(0) = val1: t(1) = val3
            If MsgBox("You want to write a full dump ?", vbYesNo) = vbYes Then
                WriteMinidump "c:\p1full.dmp", VarPtr(t(0)), MINIDUMP_TYPE.MiniDumpWithFullMemory Or MINIDUMP_TYPE.MiniDumpNormal, terr
            Else
                WriteMinidump "c:\p1small.dmp", VarPtr(t(0)), MINIDUMP_TYPE.MiniDumpNormal, terr
            End If
            MsgBox terr
        End If
    
    #Else
    
        Dim t(1) As Long
        t(0) = val1: t(1) = val3
        GpfSys_SendErrorReport VarPtr(t(0)), "Unhandled vb error : " & Date & " ; " & Time & "[this info may be correct] Error # : " & Err.Number & vbNewLine & _
                               Err.Description & vbNewLine & _
                               Err.Source & vbNewLine & _
                               Eip2Mod(GetEip(val3))

    #End If

    #If debug_ver Then
    ExeptH_bef = 0
    If MsgBox("You want to call vb handler too ? [not use if it works]", vbYesNo) = vbYes Then
    #End If
        ExeptH_bef = 1
    #If debug_ver Then
    End If
    #End If


End Function

Public Sub ExeptH_aft(ByVal called As Long, ByVal val1 As Long, ByVal val2 As Long, ByVal val3 As Long, ByVal val4 As Long)
    
    If MsgBox("You want to try to recover ?", vbYesNo) = vbYes Then
        Try2Restore val3
    End If
    
End Sub


Public Static Sub RestoreEhMode()

    EhModeI = EhModeI - 1
    If EhModeI < 0 Then
        ErrHmode = Err_not_expected
        EhModeI = 0
    Else
        ErrHmode = ErrHmodes(EhModeI)
    End If
    
End Sub

Public Sub CallStackStart()
    CsStr = ""
End Sub

Public Sub CallStackAdd(ByVal EIP As Long, ByVal Ebp As Long, ByVal Esp As Long)
    CsStr = CsStr & Eip2Mod(EIP) & ":" & Hex(EIP) & ";ebp=" & Hex(Ebp) & ";esp=" & Hex(Esp) & vbNewLine
End Sub

Public Sub CallStackEnd()
    MsgBox CsStr
End Sub

