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

Private Const CLASS_NAME As String = "plugin"

Implements ThunderVB_pl_int

'-------------------
'--- Description ---
'-------------------

Private Function ThunderVB_pl_int_GetName() As String
10        ThunderVB_pl_int_GetName = "ThunAsm"
End Function

Private Function ThunderVB_pl_int_GetDesciption() As String
          
10        ThunderVB_pl_int_GetDesciption = "ThunderVB inline Asm/C plugin"

End Function

Private Function ThunderVB_pl_int_GetDesciptionFull() As String
              
10        ThunderVB_pl_int_GetDesciptionFull = "ThunderVB inline Asm/C plugin" & vbNewLine & _
          "Adds the possibility to use inline Asm and " & vbNewLine & _
          "inmodule C in your VB 6 projects." & vbNewLine & _
          "Part of the ThunderVB project"

End Function

'---------------
'--- Version ---
'---------------

Private Function ThunderVB_pl_int_GetID() As ThunVB_PluginId
10        ThunderVB_pl_int_GetID = ThunVB_PluginId.tvb_ThunAsm ' reserved for this addin by the authors , email them to get one
End Function

Private Function ThunderVB_pl_int_GetVersion() As String
10        ThunderVB_pl_int_GetVersion = "1.0"
End Function

Private Function ThunderVB_pl_int_GetVersionNum() As Single()
      Dim ver(2) As Single
          
10        ver(0) = 1
20        ver(1) = 0
30        ver(2) = 0
40        ThunderVB_pl_int_GetVersionNum = ver

End Function

Private Function ThunderVB_pl_int_GetWindowsVersion() As String

10        ThunderVB_pl_int_GetWindowsVersion = "WinNT4;Win2000;WinXP;Win2003+"
          
End Function

'------------------
'--- Properties ---
'------------------

Private Function ThunderVB_pl_int_GetSpeed() As pl_Speed_Req
10        ThunderVB_pl_int_GetSpeed = idle
End Function

Private Function ThunderVB_pl_int_GetType() As pl_type
10        ThunderVB_pl_int_GetType = ThunVB_CPHook Or ThunVB_plugin_CodeTool
End Function

'-----------
'--- IPC ---
'-----------

Private Function ThunderVB_pl_int_SendMessange(msg As ThunderVB_pl.tvbMsg, Optional param1 As Long, Optional param2 As Long, Optional param3 As Variant, Optional param4 As Variant, Optional param5 As Variant, Optional param6 As Variant) As Variant

10        Select Case msg
          
              Case tvbMsg.tvbm_ButtonClick
              'no buttons for now so this is never recieved...
20            Case tvbMsg.tvbm_ChangeLanguage_code
30            Case tvbMsg.tvbm_ChangeLanguage_gui
              'no langige for now so this is never recieved
40            Case Is > tvbMsg.tvbm_User
50            MsgBoxX "This plugin does not support custom messages", PLUGIN_NAMEs
              
60        End Select

End Function
    
'-----------------
'--- VB Events ---
'-----------------

Private Sub ThunderVB_pl_int_OnProjectActivated(project As VBIDE.VBProject)
          
          'Load settings for this project
10        SaveSettingsToVariables True

End Sub

Private Sub ThunderVB_pl_int_OnProjectAdded(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_OnProjectRemoved(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_OnProjectRenamed(project As VBIDE.VBProject, OldName As String)

End Sub

'----------------
'--- SETTINGS ---
'----------------

Private Sub ThunderVB_pl_int_ApplySettings()
          
          'save settings
10        SaveSettings GLOBAL_, frmIn.pctSettings
20        SaveSettings LOCAL_, frmIn.pctSettings
30        SaveSettingsToVariables

End Sub

Private Sub ThunderVB_pl_int_CancelSettings()
    'what should I do? :-)
End Sub

'-----------
'--- GUI ---
'-----------

Private Function ThunderVB_pl_int_ShowConfig() As Long
10        frmIn.pctSettings.Visible = True
20        ThunderVB_pl_int_ShowConfig = frmIn.pctSettings.hWnd
End Function

Private Sub ThunderVB_pl_int_HideConfig()
10        frmIn.pctSettings.Visible = False
End Sub

'---------------
'--- Credits ---
'---------------

Private Sub ThunderVB_pl_int_HideCredits()
10        frmIn.pctCredits.Visible = False
End Sub

Private Function ThunderVB_pl_int_ShowCredits() As Long
10        frmIn.pctCredits.Visible = True
20        ThunderVB_pl_int_ShowCredits = frmIn.pctCredits.hWnd
End Function

'------------
'--- Init ---
'------------

Private Sub ThunderVB_pl_int_OnGuiLoad()
          
10        Load frmIn
          
20        LoadSettings GLOBAL_, frmIn.pctSettings
30        LoadSettings LOCAL_, frmIn.pctSettings
          
End Sub

Private Sub ThunderVB_pl_int_OnStartUp()
               
          'init objects
10        Set oMe = Me
           
          'save settings to variables
20        SaveSettingsToVariables True
          'Initilalise Everything..
          
          'bID = par.Add2ToolBar(Me, "ThunAsm Settings", "ThunderVB ASM/C Settings", frmMisc.picBut.Picture)
30        mcph = AddCPH(New cCPHook)
          

End Sub

'--------------
'--- DeInit ---
'--------------

Private Sub ThunderVB_pl_int_OnGuiUnLoad()
                   
10        Unload frmIn

End Sub

Private Sub ThunderVB_pl_int_OnTermination()
          
          'kill objects
10        Set oMe = Nothing
          'Remove CreateProccess hook
20        RemoveCPH mcph
          
End Sub

'-------------
'--- Other ---
'-------------

Private Function ThunderVB_pl_int_InitSetLang(lang As ThunderVB_pl_ResSys.tvb_Languages) As Variant

End Function
