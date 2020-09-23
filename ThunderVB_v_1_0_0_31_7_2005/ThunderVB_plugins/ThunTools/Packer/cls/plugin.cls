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

Private Sub ThunderVB_pl_int_ApplySettings()

End Sub

Private Sub ThunderVB_pl_int_CancelSettings()

End Sub

'-------------------
'--- Description ---
'-------------------

Private Function ThunderVB_pl_int_GetName() As String
    ThunderVB_pl_int_GetName = "Packer"
End Function

Private Function ThunderVB_pl_int_GetDesciption() As String
    ThunderVB_pl_int_GetDesciption = "Packer description"
End Function

Private Function ThunderVB_pl_int_GetDesciptionFull() As String
    ThunderVB_pl_int_GetDesciptionFull = "Packer full description"
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

Private Function ThunderVB_pl_int_InitSetLang(lang As ThunderVB_pl_ResSys_v1_0.tvb_Languages) As Variant

End Function

'-----------
'--- IPC ---
'-----------


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

Private Function ThunderVB_pl_int_SendMessange(msg As ThunderVB_pl_v1_0.tvbMsg, Optional param1 As Long, Optional param2 As Long, Optional param3 As Variant, Optional param4 As Variant, Optional param5 As Variant, Optional param6 As Variant) As Variant

End Function

Private Sub ThunderVB_pl_int_SetLogger(newlogger As ThunderVB_pl_LogSys_v1_0.ILogger)

End Sub

'-----------
'--- GUI ---
'-----------

Private Function ThunderVB_pl_int_ShowConfig() As Long
    ThunderVB_pl_int_ShowConfig = frmIn.pctSettings.hWnd
    frmIn.pctSettings.Visible = True
End Function

Private Sub ThunderVB_pl_int_HideConfig()
    frmIn.pctSettings.Visible = False
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
    
    'load settings
    LoadSettings GLOBAL_, frmIn.pctSettings
    LoadSettings LOCAL_, frmIn.pctSettings
    
End Sub

Private Sub ThunderVB_pl_int_OnStartUp()
     
    Set oMe = Me
    
    'save settings to variables
    SaveSettingsToVariables True
    If CBool(Get_Packer(UsePacker)) = True Then Call Init_Hook
    
End Sub


'--------------
'--- DeInit ---
'--------------

Private Sub ThunderVB_pl_int_OnGuiUnLoad()
    
    'save settings
    SaveSettings GLOBAL_, frmIn.pctSettings
    SaveSettings LOCAL_, frmIn.pctSettings
    SaveSettingsToVariables
    If CBool(Get_Packer(UsePacker)) = True Then Call Init_Hook Else Unload_Hook
    
    Unload frmIn
    
End Sub

Private Sub ThunderVB_pl_int_OnTermination()
       
    Set oThunVB = Nothing
    Set oMe = Nothing
    
    Call Unload_Hook
    
End Sub
