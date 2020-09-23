VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass_Callbacks"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum subclass_mode
    wproc_replace = 1
    wproc_notify = 2
End Enum

'for wproc_replace mode
Public Sub WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByVal pOriginalPRoc As Long)

End Sub

'for wproc_notify mode
Public Sub BefWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByRef CallWProc As Boolean, ByRef CallAftProc As Boolean, ByVal pOriginalPRoc As Long)

End Sub

Public Sub AftWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long, ByVal CalledWProc As Boolean, ByVal pOriginalPRoc As Long)

End Sub


