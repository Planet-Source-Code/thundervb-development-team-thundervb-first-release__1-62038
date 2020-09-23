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


Public Type DllToHookInfo

    DllName As String
    inModule As String 'if "" on all loaded modules , it is a list of ; seperated modules
    
End Type



Type module_entry

    Name As String
    Id As Long
    
End Type

Type module_list

    modules() As module_entry
    count As Long
    
End Type

Public Enum HookState

    hooked = 1
    unhooked = 2
    
End Enum

Public Type DllHook_entry

    ToModule As String
    DllName As String
    FunctionName As String
    FunctionAddress As Long
    HookAddress As Long
    State As HookState
    
End Type

Public Type DllHook_list

    items() As DllHook_entry
    count As Long
    Id As Long
    
End Type

Public Type DllHook_col

    items() As DllHook_list
    count As Long
    
End Type
