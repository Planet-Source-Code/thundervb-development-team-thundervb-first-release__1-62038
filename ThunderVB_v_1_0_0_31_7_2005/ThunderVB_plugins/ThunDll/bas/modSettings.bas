Attribute VB_Name = "modSettings"
Option Explicit

'module only for settings

'*** StdCallDLL ***

Private bDLL_LinkAsDLL As Boolean
Private bExportSymbols As Boolean

Private lBaseAddress As Long
Private sExportedSymbols As String
Private sEntryPointName As String

Private bDLL_UsePreLoader As Boolean
Private bDLL_FullLoading As Boolean
Private bDLL_CallMyDllMain As Boolean

Private bDLL_BPSubMain As Boolean
Private bDLL_BPCallThunRTMain As Boolean
Private bDLL_BPPreLoader As Boolean

Public Enum DLL_
    
    LinkAsDll
    ExportSymbols
    
    BaseAddress
    ExportedSymbols
    EntryPointName
    
    UsePreLoader
    FullLoading
    CallMyDllMain
    
    BPSubMain
    BPCallThunRTMain
    BPPreloader
    
End Enum

Private Const MOD_NAME As String = "modSettings"

'-------------------
'--- STDCALL DLL ---
'-------------------

'get StdCall DLL settings
'parameter - eDLL - DLL setting
'return -  False/True, long, string

Public Function Get_DLL(ByVal eDLL As DLL_) As String

    Select Case eDLL
        Case DLL_.ExportSymbols
            Get_DLL = bExportSymbols
        Case DLL_.LinkAsDll
            Get_DLL = bDLL_LinkAsDLL
        
        Case DLL_.BaseAddress
            Get_DLL = lBaseAddress
        Case DLL_.ExportedSymbols
            Get_DLL = sExportedSymbols
        Case DLL_.EntryPointName
            Get_DLL = sEntryPointName
        
        Case DLL_.FullLoading
            Get_DLL = bDLL_FullLoading
        Case DLL_.UsePreLoader
            Get_DLL = bDLL_UsePreLoader
        Case DLL_.CallMyDllMain
            Get_DLL = bDLL_CallMyDllMain
            
        Case DLL_.BPCallThunRTMain
            Get_DLL = bDLL_BPCallThunRTMain
        Case DLL_.BPPreloader
            Get_DLL = bDLL_BPPreLoader
        Case DLL_.BPSubMain
            Get_DLL = bDLL_BPSubMain
    End Select

End Function

'change StdCall flags
'parameters - eDLL - flag
'           - bNewValue - new flag

Public Sub Let_DLL(ByVal eDLL As DLL_, sNewValue As String)

    Select Case eDLL
        Case DLL_.BaseAddress
            lBaseAddress = Val("&H" & sNewValue)
        Case DLL_.ExportedSymbols
            sExportedSymbols = sNewValue
        Case DLL_.EntryPointName
            sEntryPointName = sNewValue
        
        Case DLL_.ExportSymbols
            bExportSymbols = CBool(sNewValue)
        Case DLL_.LinkAsDll
            bDLL_LinkAsDLL = CBool(sNewValue)
        
        Case DLL_.FullLoading
            bDLL_FullLoading = CBool(sNewValue)
        Case DLL_.UsePreLoader
            bDLL_UsePreLoader = CBool(sNewValue)
        Case DLL_.CallMyDllMain
            bDLL_CallMyDllMain = CBool(sNewValue)
            
        Case DLL_.BPCallThunRTMain
            bDLL_BPCallThunRTMain = CBool(sNewValue)
        Case DLL_.BPPreloader
            bDLL_BPPreLoader = CBool(sNewValue)
        Case DLL_.BPSubMain
            bDLL_BPSubMain = CBool(sNewValue)
    End Select

End Sub

Public Function SaveSettingsToVariables(Optional bLoadForm As Boolean = False)
    
    If bLoadForm = True Then
        Load frmIn
        LoadSettings GLOBAL_, frmIn.pctSettings
        LoadSettings LOCAL_, frmIn.pctSettings
    End If
    
    With frmIn
      
        Let_DLL BaseAddress, .set_txtBaseAddress.Text
        Let_DLL EntryPointName, .set_txtEntryPoint.Text
        Let_DLL ExportedSymbols, .set_ctlExport.SelectedExports
  
        Let_DLL ExportSymbols, .set_chbExportSymbols.Value
        Let_DLL LinkAsDll, .set_chbCompileDll.Value
          
        Let_DLL FullLoading, .set_chbFullLoading.Value
        Let_DLL UsePreLoader, .set_chbUsePreLoader.Value
        Let_DLL CallMyDllMain, .set_chbCallMyDllMain.Value
  
        Let_DLL BPCallThunRTMain, .set_chbBPCallThunRTMain.Value
        Let_DLL BPPreloader, .set_chbBPPreLoader.Value
        Let_DLL BPSubMain, .set_chbBPSubMain.Value
      
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

    Else
        
        '--- local ---
        
        SaveSettingProject PLUGIN_NAME, frmIn.set_ctlExport.Name, frmIn.set_ctlExport.SelectedExports

    End If
    
    SaveOtherSettings = True
    
End Function

Public Function LoadOtherSettings(eScope As SET_SCOPE) As Boolean
    'other controls that must be inited
    'this function is called from LoadSettings (modSaveLoadSettings)

    If eScope = GLOBAL_ Then
        
        '--- global ---
        
    Else
        
        '--- local ---
        
        With frmIn
            .set_ctlExport.SelectedExports = GetSettingProject(PLUGIN_NAME, .set_ctlExport.Name, "")
            
            'some text must be in textboxes
            If Len(.set_txtBaseAddress.Text) = 0 Then modSettingsControls.DefaultTextBox_ComboBox GetScope(.set_txtBaseAddress), .set_txtBaseAddress
            If Len(.set_txtEntryPoint.Text) = 0 Then modSettingsControls.DefaultTextBox_ComboBox GetScope(.set_txtEntryPoint), .set_txtEntryPoint
            
        End With
        
    End If
        
    LoadOtherSettings = True
    
End Function

Public Sub SetOtherDefaultSettings(eScope As SET_SCOPE)
    'other controls that must be set to default
    'this function is called from SetDefaultSetings (modSaveLoadSettings)

    If eScope = GLOBAL_ Then
        
        '--- global ---
        
    Else
        
        '--- local ---
        
        frmIn.set_ctlExport.SelectedExports = ""
        
    End If
       
End Sub
