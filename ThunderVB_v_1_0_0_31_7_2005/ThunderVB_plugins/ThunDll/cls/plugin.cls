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

Implements ThunderVB_pl_int_v1_0

Private Const CLASS_NAME As String = "plugin"

'-------------------
'--- Description ---
'-------------------

Private Function ThunderVB_pl_int_v1_0_GetName() As String
    ThunderVB_pl_int_v1_0_GetName = PLUGIN_NAME
End Function

Private Function ThunderVB_pl_int_v1_0_GetDesciption() As String
    ThunderVB_pl_int_v1_0_GetDesciption = PLUGIN_NAMEs & " description"
End Function

Private Function ThunderVB_pl_int_v1_0_GetDesciptionFull() As String
    ThunderVB_pl_int_v1_0_GetDesciptionFull = PLUGIN_NAMEs & " full description"
End Function

'---------------
'--- Version ---
'---------------

Private Function ThunderVB_pl_int_v1_0_GetID() As ThunVB_PluginId
    ThunderVB_pl_int_v1_0_GetID = tvb_ThunDll  ' reserved for this addin by the authors , email them to get one
End Function

Private Function ThunderVB_pl_int_v1_0_GetInferFaceVer() As Long
    ThunderVB_pl_int_v1_0_GetInferFaceVer = 1
End Function

Private Function ThunderVB_pl_int_v1_0_GetVersion() As String
    ThunderVB_pl_int_v1_0_GetVersion = "1.0"
End Function

Private Function ThunderVB_pl_int_v1_0_GetVersionNum() As Single()
Dim ver(2) As Single
    
    ver(0) = 1
    ver(1) = 0
    ver(2) = 0
    ThunderVB_pl_int_v1_0_GetVersionNum = ver

End Function

Private Function ThunderVB_pl_int_v1_0_GetWindowsVersion() As String
    ThunderVB_pl_int_v1_0_GetWindowsVersion = "Win9X/ME;WinNT4;Win2000;WinXP;Win2003+"
End Function

'------------------
'--- Properties ---
'------------------

Private Function ThunderVB_pl_int_v1_0_GetSpeed() As pl_Speed_Req
    ThunderVB_pl_int_v1_0_GetSpeed = idle
End Function

Private Function ThunderVB_pl_int_v1_0_GetType() As pl_type
    ThunderVB_pl_int_v1_0_GetType = ThunVB_CPHook
End Function

'-----------
'--- IPC ---
'-----------

Private Function ThunderVB_pl_int_v1_0_SendMessange(msg As tvbMsg, Optional param1 As Long, Optional param2 As Long, Optional param3 As Variant, Optional param4 As Variant, Optional param5 As Variant, Optional param6 As Variant) As Variant

End Function

'-----------------
'--- VB Events ---
'-----------------

Private Sub ThunderVB_pl_int_v1_0_OnProjectActivated(project As VBIDE.VBProject)
    
    'Load settings for this project
    SaveSettingsToVariables

End Sub

Private Sub ThunderVB_pl_int_v1_0_OnProjectAdded(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_v1_0_OnProjectRemoved(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_v1_0_OnProjectRenamed(project As VBIDE.VBProject, OldName As String)

End Sub

'----------------
'--- Settings ---
'----------------

Private Sub ThunderVB_pl_int_v1_0_ApplySettings()
    
    'save settings
    SaveSettings GLOBAL_, frmIn.pctSettings
    SaveSettings LOCAL_, frmIn.pctSettings
    SaveSettingsToVariables

    Call modThunDll.PatchPreLoader
    
End Sub

Private Sub ThunderVB_pl_int_v1_0_CancelSettings()
    
    'Load settings
    LoadSettings GLOBAL_, frmIn.pctSettings
    LoadSettings LOCAL_, frmIn.pctSettings
    
End Sub

'-----------
'--- GUI ---
'-----------

Private Function ThunderVB_pl_int_v1_0_ShowConfig() As Long
    frmIn.pctSettings.Visible = True
    ThunderVB_pl_int_v1_0_ShowConfig = frmIn.pctSettings.hWnd
End Function

Private Sub ThunderVB_pl_int_v1_0_HideConfig()
    frmIn.pctSettings.Visible = False
End Sub

'---------------
'--- Credits ---
'---------------

Private Sub ThunderVB_pl_int_v1_0_HideCredits()
    frmIn.pctCredits.Visible = False
End Sub

Private Function ThunderVB_pl_int_v1_0_ShowCredits() As Long
    frmIn.pctCredits.Visible = True
    ThunderVB_pl_int_v1_0_ShowCredits = frmIn.pctCredits.hWnd
End Function

'------------
'--- Init ---
'------------

Private Sub ThunderVB_pl_int_v1_0_OnGuiLoad()
    
   
End Sub

Private Sub ThunderVB_pl_int_v1_0_OnStartUp()
     
    Set oMe = Me
    
    Load frmIn
    
    'Load settings
    LoadSettings GLOBAL_, frmIn.pctSettings
    LoadSettings LOCAL_, frmIn.pctSettings
    
    SaveSettingsToVariables
    
    Init_Hook
      
End Sub

'--------------
'--- DeInit ---
'--------------

Private Sub ThunderVB_pl_int_v1_0_OnGuiUnLoad()
        
End Sub

Private Sub ThunderVB_pl_int_v1_0_OnTermination()
    
    'kill objects
    Unload frmIn
    
    Unload_Hook
        
    Set oMe = Nothing
    
End Sub

'-------------
'--- Other ---
'-------------

Private Function ThunderVB_pl_int_v1_0_InitSetLang(lang As ThunderVB_pl_ResSys_v1_0.tvb_Languages) As Variant

End Function


Private Sub ThunderVB_pl_int_v1_0_SetLogger(newlogger As ThunderVB_pl_LogSys_v1_0.ILogger)
    modLog.SetLogger newlogger
End Sub
