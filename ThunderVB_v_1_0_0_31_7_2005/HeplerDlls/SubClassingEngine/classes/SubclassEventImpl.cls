VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubclassEventImpl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass_Callbacks

Public Event AftWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal CalledWProc As Boolean, ByVal pOriginalPRoc As Long)
Public Event BefWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, CallWProc As Boolean, CallAftProc As Boolean, ByVal pOriginalPRoc As Long)
Public Event WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal pOriginalPRoc As Long)

Private m_hWnd As Long, m_mode As subclass_mode

Private Sub Class_Terminate()
    UnSubClass
End Sub

Private Sub ISubclass_Callbacks_AftWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal CalledWProc As Boolean, ByVal pOriginalPRoc As Long)
    RaiseEvent AftWndProc(hwnd, uMsg, wParam, lParam, ReturnValue, CalledWProc, pOriginalPRoc)
End Sub

Private Sub ISubclass_Callbacks_BefWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, CallWProc As Boolean, CallAftProc As Boolean, ByVal pOriginalPRoc As Long)
    RaiseEvent BefWndProc(hwnd, uMsg, wParam, lParam, ReturnValue, CallWProc, CallAftProc, pOriginalPRoc)
End Sub

Private Sub ISubclass_Callbacks_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal pOriginalPRoc As Long)
    RaiseEvent WindowProc(hwnd, uMsg, wParam, lParam, ReturnValue, pOriginalPRoc)
End Sub

Public Sub SubClass(hwnd As Long, mode As subclass_mode)
    modSCeng.SubClasshWnd hwnd, Me, mode
    m_hWnd = hwnd
End Sub

Public Sub UnSubClass()
    If m_hWnd Then
        modSCeng.UnSubClasshWnd m_hWnd
        m_hWnd = 0
    End If
End Sub

Public Property Get hwnd() As Long
    hwnd = m_hWnd
End Property

Public Property Get IsSubclassed() As Boolean
    IsSubclassed = m_hWnd <> 0
End Property

Public Property Get SubClassMode() As subclass_mode
    SubClassMode = m_mode
End Property

