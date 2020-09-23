VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ThunderVB_pl_int

'-------------------
'--- Description ---
'-------------------

Private Function ThunderVB_pl_int_GetName() As String
    ThunderVB_pl_int_GetName = "ASM/C RunTime"
End Function

Private Function ThunderVB_pl_int_GetDesciption() As String
    ThunderVB_pl_int_GetDesciption = "RunTime description"
End Function

Private Function ThunderVB_pl_int_GetDesciptionFull() As String
    ThunderVB_pl_int_GetDesciptionFull = "RunTime full description"
End Function

'---------------
'--- Version ---
'---------------

Private Function ThunderVB_pl_int_GetID() As Long
    ThunderVB_pl_int_GetID = -1 ' reserved for this addin by the authors , email them to get one
End Function

Private Function ThunderVB_pl_int_GetInferFaceVer() As Long
    ThunderVB_pl_int_GetInferFaceVer = 1
End Function

Private Function ThunderVB_pl_int_GetVersion() As String
    ThunderVB_pl_int_GetVersion = "1.0"
End Function

Private Function ThunderVB_pl_int_GetVersionNum() As Single()
Dim ver(2) As Single
    
    ver(0) = 1
    ver(1) = 0
    ver(2) = 0
    ThunderVB_pl_int_GetVersionNum = ver

End Function

Private Function ThunderVB_pl_int_GetWindowsVersion() As String
    ThunderVB_pl_int_GetWindowsVersion = "Win9X/ME;WinNT4;Win2000;WinXP;Win2003+"
End Function

'------------------
'--- Properties ---
'------------------

Private Function ThunderVB_pl_int_GetSpeed() As pl_Speed_Req
    ThunderVB_pl_int_GetSpeed = Idle
End Function

Private Function ThunderVB_pl_int_GetType() As pl_type
    ThunderVB_pl_int_GetType = ThunVB_plugin_CodeTool
End Function

'-----------
'--- IPC ---
'-----------

Private Function ThunderVB_pl_int_SendMessange(msg As ThunderVB_pl.tvbMsg, Optional param1 As Long, Optional param2 As Long, Optional param3 As Variant, Optional param4 As Variant, Optional param5 As Variant, Optional param6 As Variant) As Variant

    If msg = tvbm_ButtonClick Then
    
        If param1 = hButtonID Then
            frmRunTime.Show vbModal
        End If
    
    Else
        
        MsgBox "Unknown message.", vbExclamation, MSG_TITLEs
    
    End If
    
End Function

'-----------------
'--- VB Events ---
'-----------------

Private Sub ThunderVB_pl_int_OnProjectActivated(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_OnProjectAdded(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_OnProjectRemoved(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_OnProjectRenamed(project As VBIDE.VBProject, OldName As String)

End Sub

'-----------
'--- GUI ---
'-----------

Private Function ThunderVB_pl_int_ShowConfig() As Long

End Function

Private Sub ThunderVB_pl_int_HideConfig()

End Sub

'---------------
'--- Credits ---
'---------------

Private Sub ThunderVB_pl_int_HideCredits()
    frmIn.pctCredits.Visible = False
End Sub

Private Function ThunderVB_pl_int_ShowCredits() As Long
    frmIn.pctCredits.Visible = True
    ThunderVB_pl_int_ShowCredits = frmIn.pctCredits.hWnd
End Function

'------------
'--- Init ---
'------------

Private Sub ThunderVB_pl_int_OnGuiLoad()
    Load frmIn
End Sub

Private Sub ThunderVB_pl_int_OnStartUp(base As ThunderVB_base)
     
    Set oThunVB = base
    Set oMe = Me
    
    Call AddMe2Menu
        
End Sub


'--------------
'--- DeInit ---
'--------------

Private Sub ThunderVB_pl_int_OnGuiUnLoad()
    Unload frmIn
End Sub

Private Sub ThunderVB_pl_int_OnTermination()
    
    Call RemoveMeFromMenu
    
    Set oThunVB = Nothing
    Set oMe = Nothing
    
End Sub
