VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCpHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim exports() As String, expnum As Long, strCL As String, i As Long

Implements ThunderVB_pl_callbak

Private Sub ThunderVB_pl_callbak_CreateProcAHookBef(lpApplicationName As String, lpCommandLine As String, lpProcessAttributes As ThunderVB_pl.SECURITY_ATTRIBUTES, lpThreadAttributes As ThunderVB_pl.SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As ThunderVB_pl.STARTUPINFO, lpProcessInformation As ThunderVB_pl.PROCESS_INFORMATION, bSkip As Boolean)
If InStr(1, lpApplicationName, "link.exe", vbTextCompare) = 0 Then Exit Sub
    
    LogMsg "Linker Edit", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
    strCL = lpCommandLine
    LogMsg "exports strings seperation", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
    
    exports = Split(Get_DLL(ExportedSymbols), "**@@split_@@**")
    If Len(Get_DLL(ExportedSymbols)) > 0 Then
        expnum = UBound(exports)
    Else
        expnum = -1
    End If
    LogMsg expnum & " Exports", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
    
    
    If Get_DLL(LinkAsDll) Then  'create dll -> add /dll
        LogMsg "Creating Dll", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
        strCL = strCL & " /DLL "
    End If
    
    If CBool(Get_DLL(UsePreLoader)) = True Then
    
        If isInlineAsmLoaded() = False Then
            WarnBox "ThunderVB ASMC must be enabled to compile the preloader..", "modLinkerEdit", "linker_edit", PLUGIN_NAMEs
        End If
        
        If Len(Get_DLL(EntryPointName)) > 0 Then 'change entry point
            WarnBox "User Selected Entrypoint is overwriten by the" & vbNewLine & _
                    "use preloader option.The entry point will be DllMain function", "modLinkerEdit", "linker_edit", PLUGIN_NAMEs
        End If
                
        LogMsg "Changed entry point,With PreLoader", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
        strCL = Replace(strCL, "/ENTRY:__vbaS", "/ENTRY:PreLoader", , , vbTextCompare)
            
    Else
        If Len(Get_DLL(EntryPointName)) > 0 Then 'change entry point
            LogMsg "Changed entry point", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
            strCL = Replace(strCL, "/ENTRY:__vbaS", "/ENTRY:" & Get_DLL(EntryPointName), , , vbTextCompare)
        End If
    End If
    
    If Get_DLL(BaseAddress) Then  'change base address
        LogMsg "Changed dll base", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
        strCL = Replace(strCL, "/BASE:0x400000", "/BASE:0x" & Hex$(Get_DLL(BaseAddress)), , , vbTextCompare)
    End If
    
    If CBool(Get_DLL(ExportSymbols)) = True Then    'export symbols
        LogMsg "Exporting symbols", PLUGIN_NAME, "modLinkerEdit", "linker_edit"
        If expnum <> -1 Then
            For i = 0 To expnum
                If Len(exports(i)) Then strCL = strCL & " /Export:" & exports(i) & " "
            Next i
        Else
            WarnBox "ExportSymbols is set but No exports are defined", "modLinkerEdit", "LinkerEdit", PLUGIN_NAMEs
        End If
    End If
    
    lpCommandLine = strCL
End Sub


Private Sub ThunderVB_pl_callbak_CreateProcAHookAft(lpApplicationName As String, lpCommandLine As String, lpProcessAttributes As ThunderVB_pl.SECURITY_ATTRIBUTES, lpThreadAttributes As ThunderVB_pl.SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As ThunderVB_pl.STARTUPINFO, lpProcessInformation As ThunderVB_pl.PROCESS_INFORMATION, ReturnCode As Long, bSkip As Boolean)
    
    
End Sub

Private Function ThunderVB_pl_callbak_GetBefOrder() As ThunderVB_pl.tvb_CP_CallOrder
    
    ThunderVB_pl_callbak_GetBefOrder = cpo_1 'Well , it is not nedded to run first , but
                                          'since it edits some things it is good to
                                          'run as one of the firts..
                                          
End Function

Private Function ThunderVB_pl_callbak_GetAftOrder() As ThunderVB_pl.tvb_CP_CallOrder

    ThunderVB_pl_callbak_GetAftOrder = tvb_CP_CallOrder.cpo_DoNotRun 'We do not use the after hook ;)

End Function

Private Function isInlineAsmLoaded() As Boolean
Dim plugins As PlugIn_List, i As Long

    plugins = GetPlugInList
    For i = 0 To plugins.Count - 1
        With plugins.plugins(i)
            If (.used = True And .Loaded = True) And ((.interface Is Nothing) = False) Then
                If .interface.GetID = tvb_ThunAsm Then
                    isInlineAsmLoaded = True
                    Exit Function
                End If
            End If
        End With
    Next i

End Function
