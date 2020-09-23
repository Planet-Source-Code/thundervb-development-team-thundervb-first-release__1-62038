VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComBut"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents cbar As CommandBarEvents
Attribute cbar.VB_VarHelpID = -1
Public cbarObj As Office.CommandBarControl
Dim con As ThunderVB_pl_int_v1_0

Dim c_id As Long

Function init(ctrl As CommandBarControl, from As ThunderVB_pl_int_v1_0, IDc As Long) As Long

        c_id = IDc
        Set con = from
        Set cbarObj = ctrl
        Set cbar = VBI.Events.CommandBarEvents(cbarObj)
        init = c_id
        
End Function

Private Sub cbar_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    con.SendMessange tvbm_ButtonClick, c_id, 0, handled, CommandBarControl, CancelDefault

End Sub

Private Sub Class_Initialize()
    'nothing
End Sub

Private Sub Class_Terminate()
    
    If cbarObj Is Nothing Then Exit Sub
    cbarObj.Delete
    
End Sub
