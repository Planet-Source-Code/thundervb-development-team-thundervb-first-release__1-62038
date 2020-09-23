VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Internal "plugin" - config for Thunder VB ..
Dim MainButt As Long
Dim MainButt2 As Long

Implements ThunderVB_pl_int_v1_0


Private Sub ThunderVB_pl_int_v1_0_ApplySettings()

End Sub

Private Sub ThunderVB_pl_int_v1_0_CancelSettings()

End Sub

Private Function ThunderVB_pl_int_v1_0_GetDesciption() As String

    ThunderVB_pl_int_v1_0_GetDesciption = "ThunderVB Configurator"
    
End Function

Private Function ThunderVB_pl_int_v1_0_GetDesciptionFull() As String

    ThunderVB_pl_int_v1_0_GetDesciptionFull = "ThunderVB Configurator" & vbNewLine & _
                                         "Settings for the ThunderVB plugin interface"
    
                                   
End Function

Private Function ThunderVB_pl_int_v1_0_GetID() As ThunVB_PluginId

End Function

'Private Function ThunderVB_pl_int_v1_0_GetID() As ThunderVB_pl_v1_0.ThunVB_PluginId
    
    'ThunderVB_pl_int_v1_0_GetID = tvb_ThunderVB ' reserved for this addin by the authors , email them to get one

'End Function

Private Function ThunderVB_pl_int_v1_0_GetName() As String
    
    ThunderVB_pl_int_v1_0_GetName = "ThunderVB"

End Function

Private Function ThunderVB_pl_int_v1_0_GetSpeed() As pl_Speed_Req

    ThunderVB_pl_int_v1_0_GetSpeed = idle

End Function

Private Function ThunderVB_pl_int_v1_0_GetType() As pl_type

    ThunderVB_pl_int_v1_0_GetType = ThunVB_plugin_MiscTool

End Function

Private Function ThunderVB_pl_int_v1_0_GetVersion() As String

    ThunderVB_pl_int_v1_0_GetVersion = "1.0.0"

End Function

Private Function ThunderVB_pl_int_v1_0_GetVersionNum() As Single()
Dim Temp(2) As Single

    Temp(0) = 1
    Temp(1) = 0
    Temp(2) = 0
    
    ThunderVB_pl_int_v1_0_GetVersionNum = Temp

End Function

Private Function ThunderVB_pl_int_v1_0_GetWindowsVersion() As String
    ThunderVB_pl_int_v1_0_GetWindowsVersion = "WinNT4.0;Win2000;WinXP;Win2003+"
End Function

Private Sub ThunderVB_pl_int_v1_0_HideCredits()
On Error Resume Next
    frmInternal.Credits.Visible = False
End Sub

Private Function ThunderVB_pl_int_v1_0_InitSetLang(lang As tvb_Languages) As Variant

End Function

Private Sub ThunderVB_pl_int_v1_0_OnGuiLoad()
    
    Load frmInternal
    
End Sub

Private Sub ThunderVB_pl_int_v1_0_OnGuiUnLoad()

    Unload frmInternal
    
End Sub

Private Sub ThunderVB_pl_int_v1_0_OnProjectActivated(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_v1_0_OnProjectAdded(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_v1_0_OnProjectRemoved(project As VBIDE.VBProject)

End Sub

Private Sub ThunderVB_pl_int_v1_0_OnProjectRenamed(project As VBIDE.VBProject, OldName As String)

End Sub


Private Sub ThunderVB_pl_int_v1_0_OnStartUp()
    Load frmRes
    'MainButt = AddToAddinMenu(Me, "Thunder VB", "ThunderVB settings", frmRes.Icons(0).Image)
    'MainButt = AddToAddinMenu(Me, "Thunder VB", "ThunderVB settings", frmRes.Icons(1).Picture)
    MainButt = AddToAddinMenu(Me, "Thunder VB", "ThunderVB settings", frmRes.Icons(0).Picture)
    MainButt2 = AddButtonToToolbar(Me, "Thunder VB", "ThunderVB settings", frmRes.Icons(0).Picture)
    'MainButt = AddToAddinMenu(Me, "Thunder VB", "ThunderVB settings", frmRes.Icons(3).Picture)
    'MainButt = AddToAddinMenu(Me, "Thunder VB", "ThunderVB settings", frmRes.Icons(4).Picture)
End Sub

Private Sub ThunderVB_pl_int_v1_0_OnTermination()
    RemoveButtton MainButt
    RemoveButtton MainButt2
End Sub

Private Function ThunderVB_pl_int_v1_0_SendMessange(msg As tvbMsg, Optional param1 As Long, Optional param2 As Long, Optional param3 As Variant, Optional param4 As Variant, Optional param5 As Variant, Optional param6 As Variant) As Variant
    
    If msg = tvbm_ButtonClick Then
        frmPlugIn.RefreshLists
        frmPlugin_Show
    ElseIf msg = tvbm_ChangeLanguage_code Then
        'SaveFormToResourceFile tvb_resfile, frmPlugIn, "ThunderVB_pl", param1
        'SaveFormToResourceFile tvb_resfile, frmInternal, "ThunderVB_pl", param1
        'SaveResourceFile ThunVBPath & "tvb.gre", tvb_resfile, ""
        lang_UpdateIds param1
    ElseIf msg = tvbm_ChangeLanguage_gui Then
        lang_UpdateGui param1
    ElseIf msg > tvbm_User Then
        MsgBox "This plugin does not support custom messanges"
    End If
    
End Function


Private Sub ThunderVB_pl_int_v1_0_SetLogger(newlogger As ThunderVB_pl_LogSys_v1_0.ILogger)
    'xD nothing here . on internal pluin this is not needdeeddd XDXD
End Sub

Private Function ThunderVB_pl_int_v1_0_ShowCredits() As Long
    On Error Resume Next
    frmInternal.Credits.Visible = True
    ThunderVB_pl_int_v1_0_ShowCredits = frmInternal.Credits.hWnd
End Function

Private Function ThunderVB_pl_int_v1_0_ShowConfig() As Long
'On Error Resume Next
    frmInternal.ConfigInternal.Visible = True
    frmInternal.RefreshLists
    ThunderVB_pl_int_v1_0_ShowConfig = frmInternal.ConfigInternal.hWnd
    
End Function



Private Sub ThunderVB_pl_int_v1_0_HideConfig()
On Error Resume Next
    frmInternal.ConfigInternal.Visible = False

End Sub


