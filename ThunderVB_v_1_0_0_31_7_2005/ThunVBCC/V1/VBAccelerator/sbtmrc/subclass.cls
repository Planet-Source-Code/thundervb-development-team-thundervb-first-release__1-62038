VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub AttachMessage(iwp As ISubclass, ByVal hwnd As Long, _
                  ByVal iMsg As Long)
Attribute AttachMessage.VB_Description = "Start subclassing for a particular message."
    '<EhHeader>
    On Error GoTo AttachMessage_Err
    '</EhHeader>
    MSubclass.AttachMessage iwp, hwnd, iMsg
    '<EhFooter>
    Exit Sub

AttachMessage_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.GSubclass.AttachMessage " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hwnd As Long, _
                  ByVal iMsg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
    '<EhHeader>
    On Error GoTo DetachMessage_Err
    '</EhHeader>
    MSubclass.DetachMessage iwp, hwnd, iMsg
    '<EhFooter>
    Exit Sub

DetachMessage_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.GSubclass.DetachMessage " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Property Get CurrentMessage() As Long
Attribute CurrentMessage.VB_Description = "Returns the message currently being worked on.  Only valid in the MsgResponse and WindowProc items of your implemented interface."
    '<EhHeader>
    On Error GoTo CurrentMessage_Err
    '</EhHeader>
   CurrentMessage = MSubclass.CurrentMessage
    '<EhFooter>
    Exit Property

CurrentMessage_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.GSubclass.CurrentMessage " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Function CallOldWindowProc( _
      ByVal hwnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
Attribute CallOldWindowProc.VB_Description = "Calls the window procedure for this handle which was installed before you added the subclass."
    '<EhHeader>
    On Error GoTo CallOldWindowProc_Err
    '</EhHeader>
   CallOldWindowProc = MSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
    '<EhFooter>
    Exit Function

CallOldWindowProc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.GSubclass.CallOldWindowProc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

