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

'-------------------
'--- Description ---
'-------------------

Private Function ThunderVB_pl_int_v1_0_GetName() As String
    ThunderVB_pl_int_v1_0_GetName = "ThunIDE"
              
           'Dim tvb_resfile As Resource_File
              
           'Resource_SaveFormToResourceFile tvb_resfile, frmIn, PLUGIN_NAMEs, tvb_English
           'Resource_SaveResourceFile "c:\" & "ThunIDEp.gre", tvb_resfile, ""
              
End Function

Private Function ThunderVB_pl_int_v1_0_GetDesciption() As String
    ThunderVB_pl_int_v1_0_GetDesciption = "ThunIDE description"
End Function

Private Function ThunderVB_pl_int_v1_0_GetDesciptionFull() As String
    ThunderVB_pl_int_v1_0_GetDesciptionFull = "ThunIDE full description"
End Function

'---------------
'--- Version ---
'---------------

Private Function ThunderVB_pl_int_v1_0_GetID() As ThunVB_PluginId
    ThunderVB_pl_int_v1_0_GetID = ThunVB_PluginId.tvb_ThunIDE_p ' reserved for this addin by the authors , email them to get one
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
    ThunderVB_pl_int_v1_0_GetSpeed = Hight
End Function

Private Function ThunderVB_pl_int_v1_0_GetType() As pl_type
    ThunderVB_pl_int_v1_0_GetType = ThunVB_CPHook
End Function

'-----------
'--- IPC ---
'-----------

Private Function ThunderVB_pl_int_v1_0_SendMessange(msg As tvbMsg, Optional param1 As Long, Optional param2 As Long, Optional param3 As Variant, Optional param4 As Variant, Optional param5 As Variant, Optional param6 As Variant) As Variant
          
    If msg = tvbm_ChangeLanguage_code Then
        lang_UpdateIds param1
    ElseIf msg = tvbm_ChangeLanguage_gui Then
        lang_UpdateGui param1
    ElseIf msg = tvbm_User Then
        If param1 = tidem_GetAsmColor Then
            ThunderVB_pl_int_v1_0_SendMessange = GetAsmWordColor(CStr(param3))
        ElseIf param1 = tidem_GetCColor Then
            ThunderVB_pl_int_v1_0_SendMessange = GetCWordColor(CStr(param3))
        End If
    End If
          
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
'--- SETTINGS ---
'----------------

Private Sub ThunderVB_pl_int_v1_0_ApplySettings()
Dim temp As String
    
    SaveSettings GLOBAL_, frmIn.pctSettings
    SaveSettings LOCAL_, frmIn.pctSettings
    SaveSettingsToVariables
        
    'refresh text previews
    temp = frmIn.lblAsmPreview.caption
    frmIn.lblAsmPreview.caption = ""
    frmIn.lblAsmPreview.caption = temp
    
    temp = frmIn.lblCPreview.caption
    frmIn.lblCPreview.caption = ""
    frmIn.lblCPreview.caption = temp
    
End Sub

Private Sub ThunderVB_pl_int_v1_0_CancelSettings()
          
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
    'hide intelli asm windows if they are opened :)
    IntelliAsm.HideIntelliAsm
    If asmTT.Visible Then
        asmTT.HideToolTip
        wm_tthide = True
    Else
        wm_tthide = False
    End If
End Sub

Private Sub ThunderVB_pl_int_v1_0_OnStartUp()
           
    'init
    Set oMe = Me
          
    Load frmIn
          
    LoadSettings GLOBAL_, frmIn.pctSettings
    LoadSettings LOCAL_, frmIn.pctSettings
       
    SaveSettingsToVariables
          
    on_startup
          
End Sub


'--------------
'--- DeInit ---
'--------------

Private Sub ThunderVB_pl_int_v1_0_OnGuiUnLoad()
          
End Sub

Private Sub ThunderVB_pl_int_v1_0_OnTermination()
    
    Unload frmIn
    
    'kill objects
    Set oMe = Nothing
          
    on_termination
          
End Sub

'------------
'--- Misc ---
'------------

Private Sub ThunderVB_pl_int_v1_0_SetLogger(newlogger As ThunderVB_pl_LogSys_v1_0.ILogger)
    modLog.SetLogger newlogger
End Sub

Private Function ThunderVB_pl_int_v1_0_InitSetLang(lang As ThunderVB_pl_ResSys_v1_0.tvb_Languages) As Variant

End Function
