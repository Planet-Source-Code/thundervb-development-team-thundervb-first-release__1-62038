VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThunderVB_pl_sch_v1_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'interface for subclassing , maybe replaced with function
'pointers if too slow..

'Called when the curent's window proc is called , before VB's one
Public Sub WindowProcBef(ByRef hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef PrevProc As Long, ByRef skipVB As Boolean, ByRef skipAft As Boolean)

    
End Sub

'Called when the curent's window proc is called , After VB's one
Public Sub WindowProcAft(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, PrevProc As Long, ByRef RetValue As Long)


End Sub

