VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeclaresPub"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Hooking Related..
'Create A list with Hooked dll's
Public Static Function CreateHookList(DllName As String, FunctionName As String, HookAddress As Long, Optional inModule As String) As DllHook_list
    '<EhHeader>
    On Error GoTo CreateHookList_Err
    '</EhHeader>
    
    modHook.CreateHookList_int CreateHookList, DllName, FunctionName, HookAddress, inModule
    
    '<EhFooter>
    Exit Function

CreateHookList_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "ThunderVB_base", "CreateHookList"
    '</EhFooter>
End Function

'Create A list with Hooked dll's
Public Static Sub CreateHookList_ResumeOldList(DllName As String, FunctionName As String, HookAddress As Long, ListToContinue As DllHook_list, Optional inModule As String)
    '<EhHeader>
    On Error GoTo CreateHookList_ResumeOldList_Err
    '</EhHeader>

    modHook.CreateHookList_int ListToContinue, DllName, FunctionName, HookAddress, inModule
    
    '<EhFooter>
    Exit Sub

CreateHookList_ResumeOldList_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "ThunderVB_base", "CreateHookList_ResumeOldList"
    '</EhFooter>
End Sub

'Un hook all elemets on a hooklist..
Public Static Sub KillHookList(Hooks As DllHook_list)
    '<EhHeader>
    On Error GoTo KillHookList_Err
    '</EhHeader>
    
    Call modHook.KillHookList(Hooks)
    
    '<EhFooter>
    Exit Sub

KillHookList_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "ThunderVB_base", "KillHookList"
    '</EhFooter>
End Sub

Public Static Sub UnHookEverything()
    
    Call modHook.UnHookEverything
    
End Sub

Public Static Sub HookSys_SetLogger(LoggerObj As ILogger)
    
    Call SetLogger(LoggerObj)
    
End Sub

