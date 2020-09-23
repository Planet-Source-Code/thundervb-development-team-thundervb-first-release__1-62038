VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sce"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Static Sub SubClasshWnd(hwnd As Long, callbacks As ISubclass_Callbacks, mode As subclass_mode, Optional Unicode As Boolean = False)
    
    Call modSCeng.SubClasshWnd(hwnd, callbacks, mode, Unicode)
    
End Sub

Public Static Sub UnSubClasshWnd(hwnd As Long)

    Call modSCeng.UnSubClasshWnd(hwnd)
        
End Sub

Public Static Function CountSubClassedWindows() As Long

    CountSubClassedWindows = modSCeng.CountSubClassedWindows()
    
End Function


