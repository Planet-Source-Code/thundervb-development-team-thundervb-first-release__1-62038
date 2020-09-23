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


'For String Builder

Public Type String_B

    str() As String
    str_index As Long
    str_bound As Long
    
End Type

Public Type Col_String_entry

    strlen As Long
    col As Long
    
End Type

Public Type Col_String

    str() As Col_String_entry
    str_index As Long
    str_bound As Long
    
End Type

'For buffered IO
Public Type file_b

    filenum As Long
    maxbuflen As Long
    buflen As Long
    buf As String_B
    
End Type



Public Enum ESW

    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_MAX = 10
    
End Enum

