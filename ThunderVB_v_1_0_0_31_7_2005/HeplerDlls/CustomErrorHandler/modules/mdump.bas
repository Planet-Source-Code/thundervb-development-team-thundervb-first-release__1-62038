Attribute VB_Name = "minidump"
Option Explicit


'This code is not  mine..
'It is just exteneded a bit by me [drkIIRaziel]
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





