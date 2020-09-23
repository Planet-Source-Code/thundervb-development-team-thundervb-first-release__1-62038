Attribute VB_Name = "modSettingsControls"

Option Explicit

'
'version     - 01.01.01
'last update - 03.01.04
'
'Generic module for saving and loading settings (version for common controls)
'----------------------------------------------
'   written especially for ThunderVB plugins by Libor
'   (!!!    do not change code of this module    !!!)
'
'In settings form you can use these controls
'
'CheckBox
'--------
' Local setting   - write "*" the property Caption
' Default setting - write string "def" to the Tag property
' Save            - state of CheckBox (property State)
'
'OptionBox
'---------
' Local setting   - write "*" to the Tag property
' Default setting - write string "def" to the Tag property
' Save            - state of OptionButton (property Value)
'
'TextBox
'-------
' Local setting   - write "*" to the Tag property
' Default setting - write default string to the Tag property
' Save            - text in a textbox (property Text)
'
'ListBox (you can use ListBox with MultiSelect property set to 1 or 2 too)
'-------
' Local setting   - write "*" to the Tag property
' Default setting - write default string to the Tag property
'  note : if you use MuLtiSelect ListBox use as a delimiter of default items string - "°°"
'         so Tag property of local ListBox could be .Tag = "*item1°°item2°°item3"
' Save            - all selected items
'
'ComboBox
'-------
' Local setting   - write "*" to the Tag property
' Default setting - write default string to the Tag property
' Save            - selected item (property Text)
'
'settings of these controls are saved/loaded/set to default automatically.
'Just call SaveSettings, LoadSettings, SetDefaultSettings functions.
'
'Local setting means that the setting is valid only in active project. If setting
'is not local then it is global - valid in all projets.
'
'Default setting means how the control will be initialized.
'TextBox can contain predefined text, CheckBox can be checked, OptionButton can be choosen,
'in ListBox and ComboBox some item can be preselected.
'
'note:you can combine "flags" in the Tag property.
'     If you have TextBox that is "local" and default value is "123" then Tag property should be
'     .Tag = "*123"
'     Or if some control has in Tag string "*def". This means that this control is "local" and "default".
'
'Name of a setting control must has prefix "set". So if you have TextBox called "MySuperText" then change
'its name to "setMySuperText"
'
'If you use other control that were not mentioned, you must do saving/loading/setting to default
'manually. Write your code to these functions - LoadOtherSettings, SaveOtherSettings, SetDefaultSettings.
'Headers of these functions are:
'
'Public Function SaveOtherSettings(eScope As SET_SCOPE) As Boolean
'End Function
'
'Public Function LoadOtherSettings(eScope As SET_SCOPE) As Boolean
'End Function
'
'Public Sub SetOtherDefaultSettings(eScope As SET_SCOPE)
'End Sub
'
'SaveOtherSettings, LoadOtherSettings should return true if saving/loading was successful.
'
'**********
'Procedures
'**********
'
' -> SaveSettings, LoadSettings, SetDefaultSettings
'
'All these functions have 2 parameters. The first one is a setting scope (local, global) and the second is
'is a container of settings controls. As a container you can use PictureBox (recommended), Frame or UserControl (not recommended).
'A local setting will be saved in a VBP file and a global setting in a registry.
'
' -> EnumControls
'
'Function has one parameter  - container of setting controls. It enumerates all setting controls
'and show default controls, default strings, type of controls and other useful informations.
'

Private Const LOCAL_VALUE As String = "*"       'local setting
Private Const DEFAULT_VALUE As String = "def"   'default value

Private Const SETTINGS_NAME As String = "set"   'prefix for controls that must be saved

Private Const SETTINGS_ERROR_CHECKBOX As String = "0"
Private Const SETTINGS_ERROR_OPTIONBUTTON As String = ""
Private Const SETTINGS_ERROR_TEXTBOX_LISTBOX_COMBOBOX As String = ""

Private Const MULTI_SEL_LISTBOX As String = "°°"
Private Const CONTAINER_OPTIONBUTTON As String = "_option"

Public Enum SET_SCOPE
    LOCAL_
    GLOBAL_
End Enum

Private Const MOD_NAME As String = "modSettingsControls"

'-------------------------
'--- General functions ---
'-------------------------

Public Function SaveSettings(eScope As SET_SCOPE, ByRef oContainer As Control) As Boolean
   Dim oControl As Object

   LogMsg "Saving " & Scope2Text(eScope) & " settings", MOD_NAME, "SaveSettings"

   For Each oControl In oContainer.Parent.Controls

       If CheckControl(oControl) = False Then GoTo NextControl
          
       If (TypeOf oControl Is HzxYCheckBox) Or (TypeOf oControl Is CheckBox) Then
           Call SaveCheckBox(eScope, oControl)
       ElseIf TypeOf oControl Is OptionButton Then
           Call SaveOptionButton(eScope, oControl)
       ElseIf (TypeOf oControl Is TextBox) Or (TypeOf oControl Is ComboBox) Then
          Call SaveTextBox_ComboBox(eScope, oControl)
      ElseIf (TypeOf oControl Is ListBox) Then
          Call SaveListBox(eScope, oControl)
      End If
              
NextControl:
  Next oControl
          
  SaveSettings = SaveOtherSettings(eScope)
  Exit Function

  SaveSettings = False

  LogMsg "Error during saving " & Scope2Text(eScope) & " settings", MOD_NAME, "SaveSettings"

End Function

Public Function LoadSettings(eScope As SET_SCOPE, ByRef oContainer As Control) As Boolean
   Dim oControl As Object

   LogMsg "Loading " & Scope2Text(eScope) & " settings", MOD_NAME, "LoadSettings"

    
   For Each oControl In oContainer.Parent.Controls

        If CheckControl(oControl) = False Then GoTo NextControl
        
        If (TypeOf oControl Is HzxYCheckBox) Or (TypeOf oControl Is CheckBox) Then
            Call LoadCheckBox(eScope, oControl)
        ElseIf TypeOf oControl Is OptionButton Then
            Call LoadOptionButton(eScope, oControl)
        ElseIf (TypeOf oControl Is TextBox) Then
            Call LoadTextBox(eScope, oControl)
        ElseIf (TypeOf oControl Is ComboBox) Then
            Call LoadComboBox(eScope, oControl)
        ElseIf (TypeOf oControl Is ListBox) Then
            Call LoadListBox(eScope, oControl)
        End If

NextControl:
  Next oControl

  LoadSettings = LoadOtherSettings(eScope)
  Exit Function



  LogMsg "Error during loading " & Scope2Text(eScope) & " settings", MOD_NAME, "LoadSettings"
  LoadSettings = False

End Function

Public Sub SetDefaultSettings(eScope As SET_SCOPE, ByRef oContainer As Control)
Dim oControl As Object
    
    LogMsg "Setting " & Scope2Text(eScope) & " settings to default", MOD_NAME, "SetDefaultSettings"
    
    For Each oControl In oContainer.Parent.Controls

          If CheckControl(oControl) = False Then GoTo NextControl
             
          If (TypeOf oControl Is HzxYCheckBox) Or (TypeOf oControl Is CheckBox) Then
              Call DefaultCheckBox(eScope, oControl)
          ElseIf TypeOf oControl Is OptionButton Then
              Call DefaultOptionButton(eScope, oControl)
          ElseIf (TypeOf oControl Is TextBox) Or (TypeOf oControl Is ComboBox) Then
              Call DefaultTextBox_ComboBox(eScope, oControl)
          ElseIf (TypeOf oControl Is ListBox) Then
              Call DefaultListBox(eScope, oControl)
          End If

NextControl:
    Next oControl
    
    SetOtherDefaultSettings eScope

End Sub

'-------------
'--- Debug ---
'-------------

Public Sub EnumControls(ByRef oContainer As Control)
   Dim oControl As Control, sout As String

       For Each oControl In oContainer.Parent.Controls
          
           If CheckControl(oControl) = False Then GoTo NextControl
          
           sout = sout & oControl.name & " - "
          
           If (TypeOf oControl Is HzxYCheckBox) Or (TypeOf oControl Is CheckBox) Then
               sout = sout & "CheckBox - "
               If Left$(oControl.caption, 1) = LOCAL_VALUE Then sout = sout & "Local - " Else sout = sout & "Global - "
               If oControl.Tag = DEFAULT_VALUE Then sout = sout & "Default" Else sout = sout & "Not default"
           ElseIf TypeOf oControl Is OptionButton Then
              sout = sout & "OptionButton - "
              If Left$(oControl.Tag, 1) = LOCAL_VALUE Then sout = sout & "Local - " Else sout = sout & "Global - "
              If Right$(oControl.Tag, 3) = DEFAULT_VALUE Then sout = sout & "Default" Else sout = sout & "Not default"
          Else
              If TypeOf oControl Is TextBox Then
                  sout = sout & "TextBox - "
              ElseIf TypeOf oControl Is ComboBox Then
                  sout = sout & "ComboBox - "
              ElseIf TypeOf oControl Is ListBox Then
                  sout = sout & "ListBox - "
              End If
              If Left$(oControl.Tag, 1) = LOCAL_VALUE Then sout = sout & "Local - Default string("" & Mid$(oControl.Tag, 2) & """ Else sout = sout & "Global - Default string(""" & oControl.Tag & """)"
          End If
              
          sout = sout & vbCrLf
              
          
NextControl:
      Next oControl

      MsgBoxX "Controls:" & vbCrLf & sout, vbInformation, "Debug"

End Sub

'------------------------
'--- Helper functions ---
'------------------------

Private Function CheckControl(ByRef oControl As Control) As Boolean
          
       CheckControl = False
              
       If (TypeOf oControl Is HzxYCheckBox) Or (TypeOf oControl Is CheckBox) Or (TypeOf oControl Is TextBox) Or (TypeOf oControl Is OptionButton) Or (TypeOf oControl Is ListBox) Or (TypeOf oControl Is ComboBox) Then Else Exit Function
       If Left$(oControl.name, Len(SETTINGS_NAME)) <> SETTINGS_NAME Then Exit Function
          
       CheckControl = True
          
End Function

'--- Save ---

Public Sub SaveCheckBox(eScope As SET_SCOPE, ByRef chbBox As Control)
       If eScope = GLOBAL_ Then
           If Left$(chbBox.caption, 1) <> LOCAL_VALUE Then Call SaveSettingGlobal(PLUGIN_NAME, chbBox.name, chbBox.Value)
       Else
           If Left$(chbBox.caption, 1) = LOCAL_VALUE Then Call SaveSettingProject(PLUGIN_NAME, chbBox.name, chbBox.Value)
       End If
End Sub

Public Sub SaveOptionButton(eScope As SET_SCOPE, ByRef optButton As OptionButton)
       If optButton.Value = False Then Exit Sub
       If eScope = GLOBAL_ Then
           If Left$(optButton.Tag, 1) <> LOCAL_VALUE Then Call SaveSettingGlobal(PLUGIN_NAME, optButton.Container.name & CONTAINER_OPTIONBUTTON, optButton.name)
       Else
           If Left$(optButton.Tag, 1) = LOCAL_VALUE Then Call SaveSettingProject(PLUGIN_NAME, optButton.Container.name & CONTAINER_OPTIONBUTTON, optButton.name)
       End If
End Sub

Public Sub SaveTextBox_ComboBox(eScope As SET_SCOPE, ByRef oControl As Control)
       If eScope = GLOBAL_ Then
           If Left$(oControl.Tag, 1) <> LOCAL_VALUE Then Call SaveSettingGlobal(PLUGIN_NAME, oControl.name, oControl.Text)
       Else
           If Left$(oControl.Tag, 1) = LOCAL_VALUE Then Call SaveSettingProject(PLUGIN_NAME, oControl.name, oControl.Text)
       End If
End Sub

Public Sub SaveListBox(eScope As SET_SCOPE, ByRef lstBox As ListBox)
   Dim sData As String, i As Long

       If lstBox.MultiSelect = 0 Then
           sData = lstBox.Text
       ElseIf lstBox.SelCount <> 0 Then
           For i = 0 To lstBox.ListCount - 1
               If lstBox.Selected(i) = True Then sData = sData & MULTI_SEL_LISTBOX & lstBox.List(i)
           Next i
           sData = Mid$(sData, 3)
       End If
              
      If eScope = GLOBAL_ Then
          If Left$(lstBox.Tag, 1) <> LOCAL_VALUE Then Call SaveSettingGlobal(PLUGIN_NAME, lstBox.name, sData)
      Else
          If Left$(lstBox.Tag, 1) = LOCAL_VALUE Then Call SaveSettingProject(PLUGIN_NAME, lstBox.name, sData)
      End If
End Sub

'--- Load ---

Public Sub LoadCheckBox(eScope As SET_SCOPE, ByRef chbBox As Control)

       If eScope = GLOBAL_ Then
           If Left$(chbBox.caption, 1) <> LOCAL_VALUE Then chbBox.Value = GetSettingGlobal(PLUGIN_NAME, chbBox.name, Get_DefaultCheckBox(eScope, chbBox))
       Else
           If Left$(chbBox.caption, 1) = LOCAL_VALUE Then chbBox.Value = GetSettingProject(PLUGIN_NAME, chbBox.name, Get_DefaultCheckBox(eScope, chbBox))
       End If
End Sub

Public Sub LoadOptionButton(eScope As SET_SCOPE, ByRef optButton As OptionButton)
   Dim sData As String
       If eScope = GLOBAL_ Then
           If Left$(optButton.Tag, 1) <> LOCAL_VALUE Then
               sData = GetSettingGlobal(PLUGIN_NAME, optButton.Container.name & CONTAINER_OPTIONBUTTON, SETTINGS_ERROR_OPTIONBUTTON)
               If StrComp(sData, optButton.name, vbTextCompare) = 0 Then optButton.Value = True Else optButton = False
           End If
       Else
           If Left$(optButton.Tag, 1) = LOCAL_VALUE Then
               sData = GetSettingProject(PLUGIN_NAME, optButton.Container.name & CONTAINER_OPTIONBUTTON, SETTINGS_ERROR_OPTIONBUTTON)
              If StrComp(sData, optButton.name, vbTextCompare) = 0 Then optButton.Value = True Else optButton = False
          End If
      End If
End Sub

Public Sub LoadTextBox(eScope As SET_SCOPE, ByRef txtBox As TextBox)
       If eScope = GLOBAL_ Then
           If Left$(txtBox.Tag, 1) <> LOCAL_VALUE Then txtBox.Text = GetSettingGlobal(PLUGIN_NAME, txtBox.name, SETTINGS_ERROR_TEXTBOX_LISTBOX_COMBOBOX)
       Else
           If Left$(txtBox.Tag, 1) = LOCAL_VALUE Then txtBox.Text = GetSettingProject(PLUGIN_NAME, txtBox.name, SETTINGS_ERROR_TEXTBOX_LISTBOX_COMBOBOX)
       End If
End Sub

Public Sub LoadComboBox(eScope As SET_SCOPE, ByRef cmbBox As ComboBox)
   Dim sData As String
          
       If eScope = GLOBAL_ Then
           If Left$(cmbBox.Tag, 1) <> LOCAL_VALUE Then sData = GetSettingGlobal(PLUGIN_NAME, cmbBox.name, SETTINGS_ERROR_TEXTBOX_LISTBOX_COMBOBOX)
       Else
           If Left$(cmbBox.Tag, 1) = LOCAL_VALUE Then sData = GetSettingProject(PLUGIN_NAME, cmbBox.name, SETTINGS_ERROR_TEXTBOX_LISTBOX_COMBOBOX)
       End If

       'if CombBox has Style 2 error occurrs when trying to set text that is not in combobox
   On Error Resume Next
       cmbBox.Text = sData
   On Error GoTo 0

End Sub

Public Sub LoadListBox(eScope As SET_SCOPE, ByRef lstBox As ListBox)
   Dim sData As String, sItems() As String
   Dim i As Long, j As Long
          
       If eScope = GLOBAL_ Then
           If Left$(lstBox.Tag, 1) <> LOCAL_VALUE Then sData = GetSettingGlobal(PLUGIN_NAME, lstBox.name, SETTINGS_ERROR_TEXTBOX_LISTBOX_COMBOBOX)
       Else
           If Left$(lstBox.Tag, 1) = LOCAL_VALUE Then sData = GetSettingProject(PLUGIN_NAME, lstBox.name, SETTINGS_ERROR_TEXTBOX_LISTBOX_COMBOBOX)
       End If
             
       If lstBox.MultiSelect = 0 Or Len(sData) = 0 Then
           lstBox.Text = sData
       Else
          sItems = Split(sData, MULTI_SEL_LISTBOX)
          For i = 0 To lstBox.ListCount - 1
              For j = LBound(sItems) To UBound(sItems)
                  If lstBox.List(i) = sItems(j) Then
                      lstBox.Selected(i) = True
                      Exit For
                  End If
              Next j
          Next i
      End If
          
End Sub

'--- default ---

Public Sub DefaultCheckBox(eScope As SET_SCOPE, ByRef chbBox As Control)
       If eScope = GLOBAL_ And Left$(chbBox.caption, 1) <> LOCAL_VALUE Then
           If chbBox.Tag = DEFAULT_VALUE Then chbBox.Value = 1 Else chbBox.Value = 0
       ElseIf eScope = LOCAL_ And Left$(chbBox.caption, 1) = LOCAL_VALUE Then
           If chbBox.Tag = DEFAULT_VALUE Then chbBox.Value = 1 Else chbBox.Value = 0
       End If
End Sub

Public Function Get_DefaultCheckBox(eScope As SET_SCOPE, ByRef chbBox As Control) As Long
       If eScope = GLOBAL_ And Left$(chbBox.caption, 1) <> LOCAL_VALUE Then
           If chbBox.Tag = DEFAULT_VALUE Then Get_DefaultCheckBox = 1 Else Get_DefaultCheckBox = 0
       ElseIf eScope = LOCAL_ And Left$(chbBox.caption, 1) = LOCAL_VALUE Then
           If chbBox.Tag = DEFAULT_VALUE Then Get_DefaultCheckBox = 1 Else Get_DefaultCheckBox = 0
       End If
End Function

Public Sub DefaultOptionButton(eScope As SET_SCOPE, ByRef optButton As OptionButton)
       If eScope = GLOBAL_ And Left$(optButton.Tag, 1) <> LOCAL_VALUE Then
           If optButton.Tag = DEFAULT_VALUE Then optButton.Value = True Else optButton.Value = False
       ElseIf eScope = LOCAL_ And Left$(optButton.Tag, 1) = LOCAL_VALUE Then
           If Mid$(optButton.Tag, 2) = DEFAULT_VALUE Then optButton.Value = True Else optButton.Value = False
       End If
End Sub


Public Function Get_DefaultOptionButton(eScope As SET_SCOPE, ByRef optButton As OptionButton) As Boolean
       If eScope = GLOBAL_ And Left$(optButton.Tag, 1) <> LOCAL_VALUE Then
           If optButton.Tag = DEFAULT_VALUE Then Get_DefaultOptionButton = True Else Get_DefaultOptionButton = False
       ElseIf eScope = LOCAL_ And Left$(optButton.Tag, 1) = LOCAL_VALUE Then
           If Mid$(optButton.Tag, 2) = DEFAULT_VALUE Then Get_DefaultOptionButton = True Else Get_DefaultOptionButton = False
       End If
End Function

Public Sub DefaultTextBox_ComboBox(eScope As SET_SCOPE, ByRef oControl As Control)
       If eScope = GLOBAL_ And Left$(oControl.Tag, 1) <> LOCAL_VALUE Then
           oControl.Text = oControl.Tag
       ElseIf eScope = LOCAL_ And Left$(oControl.Tag, 1) = LOCAL_VALUE Then
           oControl.Text = Mid$(oControl.Tag, 2)
       End If
End Sub


Public Function Get_DefaultTextBox_ComboBox(eScope As SET_SCOPE, ByRef oControl As Control) As String
       If eScope = GLOBAL_ And Left$(oControl.Tag, 1) <> LOCAL_VALUE Then
           Get_DefaultTextBox_ComboBox = oControl.Tag
       ElseIf eScope = LOCAL_ And Left$(oControl.Tag, 1) = LOCAL_VALUE Then
           Get_DefaultTextBox_ComboBox = Mid$(oControl.Tag, 2)
       End If
End Function

Public Sub DefaultListBox(eScope As SET_SCOPE, ByRef lstBox As ListBox)
   Dim sTag As String, sItems() As String
   Dim i As Long, j As Long
          
       If eScope = GLOBAL_ And Left$(lstBox.Tag, 1) <> LOCAL_VALUE Then
           sTag = lstBox.Tag
       ElseIf eScope = LOCAL_ And Left$(lstBox.Tag, 1) = LOCAL_VALUE Then
           sTag = Mid$(lstBox.Tag, 2)
       End If
          
       If lstBox.MultiSelect = 0 Then
           lstBox.Text = sTag
       Else
          sItems = Split(sTag, MULTI_SEL_LISTBOX)
          
          For i = 0 To lstBox.ListCount - 1
              lstBox.Selected(i) = False
              For j = LBound(sItems) To UBound(sItems)
                  If lstBox.List(i) = sItems(j) Then
                      lstBox.Selected(i) = True
                      Exit For
                  End If
              Next j
          Next i
          
      End If
          
End Sub

Public Function GetScope(ByRef oControl As Control) As SET_SCOPE
          
       If (TypeOf oControl Is TextBox) Or (TypeOf oControl Is ListBox) _
       Or (TypeOf oControl Is ComboBox) Or (TypeOf oControl Is OptionButton) Then
              
           If Left$(oControl.Tag, 1) = LOCAL_VALUE Then GetScope = LOCAL_ Else GetScope = GLOBAL_
              
       ElseIf (TypeOf oControl Is HzxYCheckBox) Or (TypeOf oControl Is CheckBox) Then
          
           If Left$(oControl.caption, 1) = LOCAL_VALUE Then GetScope = LOCAL_ Else GetScope = GLOBAL_
              
       End If

End Function

Private Function Scope2Text(eScope As SET_SCOPE) As String
       If eScope = GLOBAL_ Then
           Scope2Text = "global"
       Else
           Scope2Text = "local"
       End If
End Function
