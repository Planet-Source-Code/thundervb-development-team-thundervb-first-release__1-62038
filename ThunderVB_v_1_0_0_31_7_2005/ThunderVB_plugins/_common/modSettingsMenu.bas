Attribute VB_Name = "modSettingsMenu"
Option Explicit

'
'version     - 01.00.00
'last update - 30.12.04
'
'Generic module for saving and loading settings (version for menu)
'----------------------------------------------
'   written especially for ThunderVB plugins by Libor
'   (!!!    do not change code of this module    !!!)
'
'If you have in your plugin project form that contain menu and you want to save its state,
'you can use this module that provides easy way to do this stuff.
'
'All settings menuitems must have speciall name - their name must have prefix "set".
'So if you have menuitem that has has name "mnuSomeMenu" than change its name to "setmnuSomeMenu"
'
'And if you want to set settings menu to default state
'
'


Private Const LOCAL_VALUE As String = "*"       'local setting
Private Const DEFAULT_VALUE As String = "def"   'default value
Private Const SETTINGS_NAME As String = "set"   'prefix for controls that must be saved

Private Const SETTINGS_ERROR_MENU As String = ""

Public Enum SET_SCOPE
    LOCAL_
    GLOBAL_
End Enum

Public Function SaveSettingsMenu(eScope As SET_SCOPE, frm As Form) As Long
Dim oMenu As Control

    LogMsg "Saving menu settings (" & PLUGIN_NAMEs & ") - " & IIf(eScope = GLOBAL_, "global", "local"), "modSaveLoadSettings", "SaveSettings", True, True

    For Each oMenu In frm.Controls
        If TypeOf oMenu Is Menu Then
            If Left(oMenu.Name, Len(SETTINGS_NAME)) = SETTINGS_NAME Then
                If eScope = GLOBAL_ Then
                    If Left(oMenu.Caption, Len(LOCAL_VALUE)) <> LOCAL_VALUE Then oThunVB.SaveSettingGlobal PLUGIN_NAMEs, oMenu.Name, oMenu.Checked
                Else
                    If Left(oMenu.Caption, Len(LOCAL_VALUE)) = LOCAL_VALUE Then oThunVB.SaveSettingProject PLUGIN_NAMEs, oMenu.Name, oMenu.Checked
                End If
            End If
        End If
    Next oMenu
    
    SaveSettingsMenu = True

End Function

Public Function LoadSettingsMenu(eScope As SET_SCOPE, frm As Form) As Boolean
Dim oMenu As Control

    LogMsg "Loading menu settings (" & PLUGIN_NAMEs & ") - " & IIf(eScope = GLOBAL_, "global", "local"), "modSaveLoadSettings", "SaveSettings", True, True
    
    For Each oMenu In frm.Controls
        If TypeOf oMenu Is Menu Then
            If Left(oMenu.Name, Len(SETTINGS_NAME)) = SETTINGS_NAME Then
                If eScope = GLOBAL_ Then
                    If Left(oMenu.Caption, Len(LOCAL_VALUE)) <> LOCAL_VALUE Then oMenu.Checked = oThunVB.GetSettingGlobal(PLUGIN_NAMEs, oMenu.Name, SETTINGS_ERROR_MENU)
                Else
                    If Left(oMenu.Caption, Len(LOCAL_VALUE)) = LOCAL_VALUE Then oMenu.Checked = oThunVB.GetSettingProject(PLUGIN_NAMEs, oMenu.Name, SETTINGS_ERROR_MENU)
                End If
            End If
        End If
    Next oMenu
    
    LoadSettingsMenu = True

End Function

Public Sub SetDefaultSettingsMenu(eScope As SET_SCOPE, frm As Form)
Dim oMenu As Control

    LogMsg "Setting menu to default - (" & PLUGIN_NAMEs & ") - " & IIf(eScope = GLOBAL_, "global", "local"), "modSaveLoadSettings", "SetDefaultSettings", PLUGIN_NAMEs, True

    For Each oMenu In frm.Controls
        If TypeOf oMenu Is Menu Then
            If Left(oMenu.Name, Len(SETTINGS_NAME)) = SETTINGS_NAME Then
                If Right(oMenu.Name, Len(DEFAULT_VALUE)) = DEFAULT_VALUE Then
                    If eScope = GLOBAL_ Then
                        If Left(oMenu.Caption, Len(LOCAL_VALUE)) <> LOCAL_VALUE Then oMenu.Checked = True
                    Else
                        If Left(oMenu.Caption, Len(LOCAL_VALUE)) = LOCAL_VALUE Then oMenu.Checked = True
                    End If
                Else
                    oMenu.Checked = False
                End If
            End If
        End If
    Next oMenu

End Sub
