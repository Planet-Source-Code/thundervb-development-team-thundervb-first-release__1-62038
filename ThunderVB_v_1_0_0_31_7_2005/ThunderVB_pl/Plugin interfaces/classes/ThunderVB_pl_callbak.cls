VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThunderVB_pl_cph_v1_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'For cp hook..
Option Explicit

'Called before
Public Sub CreateProcAHookBef(ByRef lpApplicationName As String, ByRef lpCommandLine As String, _
                              ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
                              ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
                              ByRef bInheritHandles As Long, ByRef dwCreationFlags As Long, _
                              ByRef lpEnvironment As Long, ByRef lpCurrentDirectory As String, _
                              ByRef lpStartupInfo As STARTUPINFO, _
                              ByRef lpProcessInformation As PROCESS_INFORMATION, _
                              ByRef bSkip As Boolean)
    
End Sub

'Called after..
Public Sub CreateProcAHookAft(ByRef lpApplicationName As String, ByRef lpCommandLine As String, _
                                   ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                   ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                   ByRef bInheritHandles As Long, ByRef dwCreationFlags As Long, _
                                   ByRef lpEnvironment As Long, ByRef lpCurrentDirectory As String, _
                                   ByRef lpStartupInfo As STARTUPINFO, _
                                   ByRef lpProcessInformation As PROCESS_INFORMATION, _
                                   ByRef ReturnCode As Long, ByRef bSkip As Boolean _
                                   )
    
End Sub

'To get the order of this one..
Public Function GetBefOrder() As tvb_CP_CallOrder

End Function

Public Function GetAftOrder() As tvb_CP_CallOrder

End Function

