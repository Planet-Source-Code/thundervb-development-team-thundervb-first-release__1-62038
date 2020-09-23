VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTypePub"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' holds info about a specific eception
Public Type EXCEPTION_RECORD
  ExceptionCode      As Long  ' type of exception - defined above
  ExceptionFlags     As Long  ' whether the exception is continuable or not
  pExceptionRecord   As Long  ' pointer to another EXCEPTION_RECORD struct (for nested exceptions)
  ExceptionAddress   As Long  ' the address at which the exception occurred
  NumberParameters   As Long  ' number of params in the following array
  Information(EXCEPTION_MAXIMUM_PARAMETERS - 1) As Long ' extra info.. not really needed.
End Type


' processor specific - not really needed anyway
Public Type CONTEXT
    Null As Long
End Type

' wrapper for the above types
Public Type EXCEPTION_POINTERS
  pExceptionRecord   As EXCEPTION_RECORD
  ContextRecord      As CONTEXT
End Type


Public Enum GPF_actions

    GPF_None
    GPF_RaiseErr
    GPF_Cont
    GPF_Stop
    
End Enum

Public Type gpf_pb_e

    GPF_action As GPF_actions
    GPF_CodeProc As String
    GPF_CodeMod As String
    GPF_CodeProject As String
    GPF_Last_Exeption As EXCEPTION_POINTERS

End Type


