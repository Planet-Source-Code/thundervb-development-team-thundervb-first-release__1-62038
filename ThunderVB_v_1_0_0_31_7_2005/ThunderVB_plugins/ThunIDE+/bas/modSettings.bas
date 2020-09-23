Attribute VB_Name = "modSettings"
Option Explicit

'*** ASM ***
'form Settings/Tab ASM

Private bASM_UseASMColoring As Boolean  'do ASM coloring
Public bASM_QuickWatch As Boolean
Public bASM_IntelliSense As Boolean
Private sASM_ASMColors As String        'colors

Public Enum ASM_
    UseASMColoring
    ASMColors
    QuickWatch
    IntelliSense
End Enum

'*** C ***
'form Settings/Tab C

Private bC_UseCColoring As Boolean    'do C coloring
Private sC_CColors As String          'colors

Public Enum C_
    UseCColoring
    CColors
End Enum

Private bMisc_CopyTimeColoring As Boolean
    
Public Enum Misc_
    CopyTimeColoring
End Enum

'-----------
'--- ASM ---
'-----------

'get ASM settings
'parameter - eASM - asm setting
'return -  string/true/false

Public Function Get_ASM(ByVal eASM As ASM_) As String

       Select Case eASM
           Case ASM_.ASMColors
               Get_ASM = sASM_ASMColors
           Case ASM_.UseASMColoring
               Get_ASM = bASM_UseASMColoring
           Case ASM_.QuickWatch
               Get_ASM = bASM_QuickWatch
           Case ASM_.IntelliSense
               Get_ASM = bASM_IntelliSense
      End Select
          
End Function

'change ASM flags
'parameters - eASM - flags
'           - sNewValue - new setting

Public Sub Let_ASM(ByVal eASM As ASM_, sNewValue As String)

       Select Case eASM
           Case ASM_.ASMColors
               sASM_ASMColors = sNewValue
               initAsmColors sNewValue
           Case ASM_.UseASMColoring
               bASM_UseASMColoring = CBool(sNewValue)
               AsmColoringEn bASM_UseASMColoring ' Edit -- raziel
           Case ASM_.QuickWatch
               bASM_QuickWatch = CBool(sNewValue)
          Case ASM_.IntelliSense
              bASM_IntelliSense = CBool(sNewValue)
      End Select
          
End Sub

'---------
'--- C ---
'---------

'get C settings
'parameter - eC - C setting
'return -  string/true/false

Public Function Get_C(ByVal eC As C_) As String

       Select Case eC
           Case C_.CColors
               Get_C = sC_CColors
           Case C_.UseCColoring
               Get_C = bC_UseCColoring
       End Select
          
End Function

'change Code coloring flags
'parameters - eC - flags
'           - sNewValue - new value

Public Sub Let_C(ByVal eC As C_, sNewValue As String)

       Select Case eC
           Case C_.CColors
               sC_CColors = sNewValue
               initCcolors sC_CColors
           Case C_.UseCColoring
               bC_UseCColoring = CBool(sNewValue)
               CColoringEn bC_UseCColoring
       End Select
          
End Sub

'------------
'--- Misc ---
'------------

Public Function Get_Misc(ByVal eMisc As Misc_) As Boolean

       Select Case eMisc
           Case Misc_.CopyTimeColoring
               Get_Misc = bMisc_CopyTimeColoring
       End Select
          
End Function

Public Sub Let_Misc(ByVal eMisc As Misc_, bNewValue As Boolean)

       Select Case eMisc
           Case Misc_.CopyTimeColoring
               bMisc_CopyTimeColoring = bNewValue
       End Select
          
End Sub

Public Function SaveSettingsToVariables(Optional bLoadForm As Boolean = False)
          
       If bLoadForm = True Then
           Load frmIn
           LoadSettings GLOBAL_, frmIn.pctSettings
           LoadSettings LOCAL_, frmIn.pctSettings
       End If
          
       With frmIn
              
              
          Let_ASM UseASMColoring, .set_ASM_chbAsmColoring.value
          Let_ASM ASMColors, .set_ASM_ctlColorsASM.ColorInfo
          Let_ASM IntelliSense, .set_ASM_chbIntelliSense.value
          Let_ASM QuickWatch, .set_ASM_chbQuickWatch.value
          
          Let_C UseCColoring, .set_C_chbCColoring.value
          Let_C CColors, .set_C_ctlColorsC.ColorInfo
          
          Let_Misc CopyTimeColoring, .set_Misc_chbCopyTimeColor.value
              
      End With
          
      If bLoadForm = True Then
          Unload frmIn
      End If
          
End Function

Public Function SaveOtherSettings(eScope As SET_SCOPE) As Boolean
       'other controls that must be saved
       'this function is called from SaveSettings (modSaveLoadSettings)
          
       If eScope = GLOBAL_ Then

           '--- global ---
              
           SaveSettingGlobal PLUGIN_NAME, frmIn.set_ASM_ctlColorsASM.name, frmIn.set_ASM_ctlColorsASM.ColorInfo
           SaveSettingGlobal PLUGIN_NAME, frmIn.set_C_ctlColorsC.name, frmIn.set_C_ctlColorsC.ColorInfo
              
       Else
              
           '--- local ---
              
       End If
          
       SaveOtherSettings = True
          
End Function

Public Function LoadOtherSettings(eScope As SET_SCOPE) As Boolean
       'other controls that must be inited
       'this function is called from LoadSettings (modSaveLoadSettings)
          
   Dim sData As String
          
       If eScope = GLOBAL_ Then

           '--- global ---
              
           sData = GetSettingGlobal(PLUGIN_NAME, frmIn.set_ASM_ctlColorsASM.name, "")
           If Len(sData) = 0 Then frmIn.set_ASM_ctlColorsASM.SetDefaultsAsm Else frmIn.set_ASM_ctlColorsASM.ColorInfo = sData
              
           sData = GetSettingGlobal(PLUGIN_NAME, frmIn.set_C_ctlColorsC.name, "")
           If Len(sData) = 0 Then frmIn.set_C_ctlColorsC.SetDefaultsC Else frmIn.set_C_ctlColorsC.ColorInfo = sData
              
       Else
              
           '--- local ---
              
       End If
          
       LoadOtherSettings = True
          
End Function

Public Sub SetOtherDefaultSettings(eScope As SET_SCOPE)
       'other controls that must be set to default
       'this function is called from SetDefaultSetings (modSaveLoadSettings)
             
       If eScope = GLOBAL_ Then
              
           '--- global ---
              
           frmIn.set_ASM_ctlColorsASM.SetDefaultsAsm
           frmIn.set_C_ctlColorsC.SetDefaultsC
              
       Else
              
           '--- local ---

       End If
             
End Sub
