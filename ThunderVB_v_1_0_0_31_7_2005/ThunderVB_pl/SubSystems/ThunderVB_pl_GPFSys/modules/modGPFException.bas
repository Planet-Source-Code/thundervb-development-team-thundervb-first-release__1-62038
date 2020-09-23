Attribute VB_Name = "modGPFException"
Option Explicit

'This code is not  mine..
'It is exteneded a by me [drkIIRaziel]
'Now , this can handle gpf's , write minudumps and optionaly report them ..
'by uploading em to the server..

'Original header:
' -------------------------------------------------------------- '
' module to handle unhandled exceptions (GPFs)
' created 25/11/02
' modified  10/12/02
' will barden
'
' 10/12/02 - added a more descriptive error message, and
'            setup the error handler to use VBs's internal
'            error bubbling to raise it.
' 5/10/2004 -Modifyed for ThunVB by Raziel
' 24/3/2005 -Minidumps to file
' 25/3/2005 -Error reporting and uploading
' -------------------------------------------------------------- '
'
' -------------------------------------------------------------- '
' apis
' -------------------------------------------------------------- '

' used to set and remove our callback
'Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

' to raise a GPF (for testing)
'Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

' to get the last GPF code
'Private Declare Function GetExceptionInformation Lib "kernel32" () As Long

' -------------------------------------------------------------- '
' consts
' -------------------------------------------------------------- '

' return values from our callback
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
Public Const EXCEPTION_CONTINUE_SEARCH = 0



' to describe the violation - defined in windows.h
Public Const EXCEPTION_CONTINUABLE              As Long = &H0
Public Const EXCEPTION_NONCONTINUABLE           As Long = &H1

Public Const EXCEPTION_ACCESS_VIOLATION         As Long = &HC0000005 ' The thread tried to read from or write to a virtual address for which it does not have the appropriate access
Public Const EXCEPTION_BREAKPOINT               As Long = &H80000003 ' A breakpoint was encountered.
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED    As Long = &HC000008C ' The thread tried to access an array element that is out of bounds and the underlying hardware supports bounds checking.
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO       As Long = &HC000008E ' The thread tried to divide a floating-point value by a floating-point divisor of zero.
Public Const EXCEPTION_FLT_INVALID_OPERATION    As Long = &HC0000090 ' This exception represents any floating-point exception not included in this list
Public Const EXCEPTION_FLT_OVERFLOW             As Long = &HC0000091 ' The exponent of a floating-point operation is greater than the magnitude allowed by the corresponding type
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO       As Long = &HC0000094 ' The thread tried to divide an integer value by an integer divisor of zero.
Public Const EXCEPTION_INT_OVERFLOW             As Long = &HC0000095 ' The result of an integer operation caused a carry out of the most significant bit of the result
Public Const EXCEPTION_ILLEGAL_INSTRUCTION      As Long = &HC000001D ' The thread tried to execute an invalid instruction
Public Const EXCEPTION_PRIV_INSTRUCTION         As Long = &HC0000096 ' The thread tried to execute an instruction whose operation is not allowed in the current machine mode

' -------------------------------------------------------------- '
' structs
' -------------------------------------------------------------- '

'GPF Interface stuff


Public GPF_action As GPF_actions
Public GPF_CodeProc As String
Public GPF_CodeMod As String
Public GPF_CodeProject As String
Public GPF_Last_Exeption As EXCEPTION_POINTERS


Public Type PMINIDUMP_EXCEPTION_INFORMATION
    ThreadId As Long
    ExceptionPointers As EXCEPTION_POINTERS
    ClientPointers As Long
End Type


Public Enum MINIDUMP_TYPE
  MiniDumpNormal = &H0
  MiniDumpWithDataSegs = &H1
  MiniDumpWithFullMemory = &H2
  MiniDumpWithHandleData = &H4
  MiniDumpFilterMemory = &H8
  MiniDumpScanMemory = &H10
  MiniDumpWithUnloadedModules = &H20
  MiniDumpWithIndirectlyReferencedMemory = &H40
  MiniDumpFilterModulePaths = &H80
  MiniDumpWithProcessThreadData = &H100
  MiniDumpWithPrivateReadWriteMemory = &H200
  MiniDumpWithoutOptionalData = &H400
  MiniDumpWithFullMemoryInfo = &H800
  MiniDumpWithThreadInfo = &H1000
  MiniDumpWithCodeSegs = &H2000
End Enum

Declare Function MiniDumpWriteDump Lib "Dbghelp.dll" (ByVal hProcess As Long, _
                                              ByVal ProcessId As Long, _
                                              ByVal hFile As Long, _
                                              ByVal DumpType As MINIDUMP_TYPE, _
                                              ByVal ExceptionParam As Long, _
                                              ByVal UserStreamParam As Long, _
                                              ByVal CallbackParam As Long) As Boolean
Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
' -------------------------------------------------------------- '
' private variables
' -------------------------------------------------------------- '
Private mlpOldProc As Long
Private pb() As gpf_pb_e, pbl As Long, pbli As Long
' -------------------------------------------------------------- '
' methods
' -------------------------------------------------------------- '
Public gpf_MinidumpPath As String

' setup the new handler
Public Function StartGPFHandler(ByRef MinidumpPath As String) As Boolean
    '<EhHeader>
    On Error GoTo StartGPFHandler_Err
    '</EhHeader>
   gpf_MinidumpPath = MinidumpPath
    
   LogMsg "Seting up GPF handler", "modGPFException", "StartGPFHandler"
   ' assume success
   StartGPFHandler = True
   
   ' if we're already handling, there's no point
   If mlpOldProc = 0 Then
   
      ' set up the handler
      mlpOldProc = SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
      ' not all systems will return a handle
      If mlpOldProc = 0 Then mlpOldProc = 1
      
   End If
   
    '<EhFooter>
    Exit Function

StartGPFHandler_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "modGPFException", "StartGPFHandler"
    '</EhFooter>
End Function

' release the new handler
Public Sub StopGPFHandler()
    '<EhHeader>
    On Error GoTo StopGPFHandler_Err
    '</EhHeader>
   LogMsg "killing GPF handler", "modGPFException", "StopGPFHandler"
   ' release the handler
   SetUnhandledExceptionFilter vbNull
   
   ' reset the variable
   mlpOldProc = 0
   
    '<EhFooter>
    Exit Sub

StopGPFHandler_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "modGPFException", "StopGPFHandler"
    '</EhFooter>
End Sub

' just for debugging - test the handler by firing a GPF
Public Sub TestGPFHandler()

   ' raise a GPF
   RaiseException EXCEPTION_ARRAY_BOUNDS_EXCEEDED, 0, 0, 0
   
End Sub

' altered on 10/12/02 by request - this function now simply raises
' an error so that VB can handle it properly, via On Error.
'Edited heavily by drkIIRaziel
Public Function ExceptionHandler(ByVal Exception As Long) As Long
Dim lTmp       As Long
Dim sType      As String
Dim lAddress   As Long
Dim sContinue  As String
Dim stra As String
'Shell "explorer """"http://sourceforge.net/tracker/?atid=710267&group_id=128073&func=browse"""", vbMaximizedFocus"
Dim strab As String, strb As String
   
   strab = CInt(Timer * 10) & "_" & "dump.dmp"
   kill2 gpf_MinidumpPath & strab
   WriteMinidump gpf_MinidumpPath & strab, Exception, MiniDumpNormal, stra
   
   strb = CInt(Timer * 10) & "_" & "report.gre"
   
   SaveReportToFile strb, strab, "GPF error"
   
   'Call frmReport.ShowReport("GPF-crash:" & Date)
   Dim uException As EXCEPTION_POINTERS
   CopyMemory uException, ByVal Exception, Len(uException)
   
   ' let's get some information about the error in order
   ' to raise a nicely defined, and explanatory error via VB
   CopyMemory lTmp, ByVal uException.pExceptionRecord.ExceptionCode, 4
   Select Case lTmp
      Case EXCEPTION_ACCESS_VIOLATION
         sType = "EXCEPTION_ACCESS_VIOLATION"
      Case EXCEPTION_BREAKPOINT
         sType = "EXCEPTION_BREAKPOINT"
      Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
         sType = "EXCEPTION_ARRAY_BOUNDS_EXCEEDED"
      Case EXCEPTION_FLT_DIVIDE_BY_ZERO
         sType = "EXCEPTION_FLT_DIVIDE_BY_ZERO"
      Case EXCEPTION_FLT_INVALID_OPERATION
         sType = "EXCEPTION_FLT_INVALID_OPERATION"
      Case EXCEPTION_FLT_OVERFLOW
         sType = "EXCEPTION_FLT_OVERFLOW"
      Case EXCEPTION_INT_DIVIDE_BY_ZERO
         sType = "EXCEPTION_INT_DIVIDE_BY_ZERO"
      Case EXCEPTION_INT_OVERFLOW
         sType = "EXCEPTION_INT_OVERFLOW"
      Case EXCEPTION_ILLEGAL_INSTRUCTION
         sType = "EXCEPTION_ILLEGAL_INSTRUCTION"
      Case EXCEPTION_PRIV_INSTRUCTION
         sType = "EXCEPTION_PRIV_INSTRUCTION"
      Case Else
         sType = "Unknown exception type 0x" & Hex$(uException.pExceptionRecord.ExceptionCode) & _
                 ".Possibly VB6 exeption that was not handled"
   End Select

   ' check for a couple of other important points..
   With uException.pExceptionRecord
      ' can we continue from this error?
      If .ExceptionFlags = EXCEPTION_CONTINUABLE Then
         sContinue = "Ok to continue."
      ElseIf .ExceptionFlags = EXCEPTION_NONCONTINUABLE Then
         sContinue = "NOT ok to continue."
      Else
         sContinue = "Probably safe to continue, but better not."
      End If
      ' and lastly, where the error occurred.
      lAddress = .ExceptionAddress
   End With
   
   Dim modfn As String
   
   modfn = modEip2Mod.Eip2Mod(uException.pExceptionRecord.ExceptionAddress)
   
   GPF_Last_Exeption = uException
    ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
    Select Case GPF_action
        Case GPF_Cont
            LogMsg "An handled GPF error (" & sType & ") " & _
                   "occurred at: " & lAddress & "(" & modfn & "). " & sContinue, GPF_CodeProject & ":" & GPF_CodeMod, GPF_CodeProc
            LogMsg "Trying to continue", GPF_CodeProject & ":" & GPF_CodeMod, GPF_CodeProc
            ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
            
        Case GPF_actions.GPF_RaiseErr
        LogMsg "An handled GPF error (" & sType & ") " & _
                   "occurred at: " & lAddress & "(" & modfn & "). " & sContinue, GPF_CodeProject & ":" & GPF_CodeMod, GPF_CodeProc
        LogMsg "Raising Error", GPF_CodeProject & ":" & GPF_CodeMod, GPF_CodeProc
        Err.Raise ThunVB_Errors.tvb_GPF_Error, _
                 "Exception Handler", _
                 "An unhandled error (" & sType & ") " & vbCrLf & _
                 "occurred at: " & lAddress & "(" & modfn & "). " & sContinue
                 ' continue with execution
                ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
                
        Case GPF_actions.GPF_Stop
        LogMsg "An handled GPF error (" & sType & ") " & _
                   "occurred at: " & lAddress & "(" & modfn & "). " & sContinue, GPF_CodeProject & ":" & GPF_CodeMod, GPF_CodeProc
            LogMsg "Killing VB proccess", GPF_CodeProject & ":" & GPF_CodeMod, GPF_CodeProc
            ExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            
        Case GPF_actions.GPF_None
             LogMsg "An unhandled error (" & sType & ") " & _
                    "occurred at: " & lAddress & "(" & modfn & "). " & sContinue, _
                    "modGPFException", "ExeptionHandler"
             Select Case frmGPFError.ShowGPF("An unhandled error (" & sType & ") " & _
                         "occurred at: " & lAddress & "(" & modfn & "). " & sContinue)
              Case GPF_actions.GPF_RaiseErr
                  Err.Raise ThunVB_Errors.tvb_GPF_Error, _
                           "Exception Handler", _
                           "An unhandled error (" & sType & ") " & vbCrLf & _
                           "occurred at: " & lAddress & "(" & modfn & "). " & sContinue
                           ' continue with execution
                           ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
              Case GPF_actions.GPF_Cont
                  'continue with execution
                  ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
              Case GPF_actions.GPF_Stop
                  'stop execution
                  ExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            End Select
    End Select

End Function

'Set the current gpf Hnadling mode
Public Sub GPF_Set(ByVal nAct As GPF_actions, ByRef fromProj As String, ByRef fromMod As String, ByRef fromProc As String)
    '<EhHeader>
    On Error GoTo GPF_Set_Err
    '</EhHeader>
Dim gpfnull As EXCEPTION_POINTERS

    If pbli >= pbl Then
        ReDim Preserve pb((pbli + 1) * 2)
        pbl = ArrUBound(pb)
    End If
    
    With pb(pbli)
        .GPF_action = GPF_action
        '.GPF_Last_Exeption = GPF_Last_Exeption
        .GPF_CodeProject = GPF_CodeProject
        .GPF_CodeMod = GPF_CodeMod
        .GPF_CodeProc = GPF_CodeProc
    End With
    
    pbli = pbli + 1
    
    GPF_Last_Exeption = gpfnull
    GPF_action = nAct
    GPF_CodeMod = fromMod
    GPF_CodeProc = fromProc
    GPF_CodeProject = fromProj
    
    '<EhFooter>
    Exit Sub

GPF_Set_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "modGPFException", "GPF_Set"
    '</EhFooter>
End Sub

'restore the previous gpf handling mode..
Public Sub GPF_Reset()
    '<EhHeader>
    On Error GoTo GPF_Reset_Err
    '</EhHeader>
Dim gpfnull As EXCEPTION_POINTERS
    
    If pbli >= 1 Then
        pbli = pbli - 1
        With pb(pbli)
            GPF_Last_Exeption = gpfnull
            GPF_action = .GPF_action
            GPF_CodeProject = .GPF_CodeProject
            GPF_CodeMod = .GPF_CodeMod
            GPF_CodeProc = .GPF_CodeProc
        End With
    Else
        GPF_Last_Exeption = gpfnull
        GPF_action = GPF_None
        GPF_CodeProject = "???"
        GPF_CodeMod = "???"
        GPF_CodeProc = "???"
    End If
    
    'unset the gpf handling data and the VB error handler

    '<EhFooter>
    Exit Sub

GPF_Reset_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "modGPFException", "GPF_Reset"
    '</EhFooter>
End Sub



'Minidumps .. YEAH
Public Function WriteMinidump(ByVal ToFile As String, _
                              ByVal ExceptionInfo As Long, _
                              ByVal DumpType As MINIDUMP_TYPE, _
                              ByRef StrError As String) As Boolean
   
    Dim hFile As Long
    hFile = A_CreateFile(ToFile, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0, _
                         CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, ByVal 0)
                         
    'if (hFile!=INVALID_HANDLE_VALUE)
    '{
    If hFile <> 0 Then
        Dim t(3) As Long
        t(0) = GetCurrentThreadId()
        t(1) = ExceptionInfo
        t(2) = 0
        '// write the dump
        'BOOL bOK = pDump( GetCurrentProcess(), GetCurrentProcessId(), hFile, MiniDumpNormal, &ExInfo, NULL, NULL );
        WriteMinidump = MiniDumpWriteDump(GetCurrentProcess(), GetCurrentProcessId(), _
                                          hFile, DumpType, VarPtr(t(0)), 0, 0)
        
        CloseHandle hFile
        If WriteMinidump <> False Then
            StrError = "No Error"
        Else
            StrError = WinApiForVb.GetLastError()
        End If
    Else
        StrError = "Can't create file " & ToFile
    End If
    
End Function



