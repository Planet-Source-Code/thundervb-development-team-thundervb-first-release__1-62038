VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim id As Long
Dim hwndOwn As Long
Dim bUn As Boolean

Public Function NewHotKey(CBhwnd As Long, modifyer As Long, key As Long, cnt As Long, Optional bunrOnExt As Boolean = False) As Boolean
 
    hwndOwn = CBhwnd
    id = cnt
    cnt = cnt + 1
    NewHotKey = RegisterHotKey(CBhwnd, id, modifyer, key)
    bUn = bunrOnExt
    
End Function

Private Sub Class_Terminate()
    
    If bUn Then
        Call UnregisterHotKey(hwndOwn, id)
    End If
    
End Sub
