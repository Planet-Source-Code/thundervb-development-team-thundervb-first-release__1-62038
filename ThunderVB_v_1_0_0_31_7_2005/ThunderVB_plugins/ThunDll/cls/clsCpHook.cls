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

Implements ThunderVB_pl_cph_v1_0

Private Const DLLMAIN As String = "PreLoader"
Private Const CLASS_NAME As String = "clsCpHook"

Private Sub ThunderVB_pl_cph_v1_0_CreateProcAHookBef(lpApplicationName As String, lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION, bSkip As Boolean)
Dim asExp() As String, sCmdLine As String, i As Long
    
    'check cmd line
    If InStr(1, lpApplicationName, "link.exe", vbTextCompare) = 0 Then Exit Sub

    LogMsg "Patching cmd-line", CLASS_NAME, "ThunderVB_pl_cph_v1_0_CreateProcAHookBef"

    'save cmd line
    sCmdLine = lpCommandLine

    'if we want to link Dll
    If Get_DLL(LinkAsDll) Then

        'change cmd line
        sCmdLine = sCmdLine & " /DLL "

        'if we want to use PreLoader
        If CBool(Get_DLL(UsePreLoader)) = True Then

            'ThunAsm must be loaded
            If isInlineAsmLoaded() = False Then
                WarnBox "ThunderASM must be enabled to compile the preloader..", CLASS_NAME, "ThunderVB_pl_cph_v1_0_CreateProcAHookBef"
                Exit Sub
            End If

            'change Entry-Point name
            sCmdLine = Replace(sCmdLine, "/ENTRY:__vbaS", "/ENTRY:" & DLLMAIN, , 1, vbTextCompare)
            'sCmdLine = Replace(sCmdLine, "/OPT:REF", "/OPT:NOREF", , 1, vbTextCompare)
            sCmdLine = sCmdLine & " " & Add34(GetThunderVBPluginsPath & "ThunDll\kernel32.lib")

        Else

            sCmdLine = Replace(sCmdLine, "/ENTRY:__vbaS", "/ENTRY:" & Get_DLL(EntryPointName), , 1, vbTextCompare)

        End If

    End If

    'set new base-address
    sCmdLine = Replace(sCmdLine, "/BASE:0x400000", "/BASE:0x" & Hex$(Get_DLL(BaseAddress)), , 1, vbTextCompare)

    'export functions
    If CBool(Get_DLL(ExportSymbols)) = True Then

        If Len(Get_DLL(ExportedSymbols)) <> 0 Then

            'TODO - pouzivam frmIn !!!!
            asExp = Split(Get_DLL(ExportedSymbols), frmIn.set_ctlExport.FuncSeparator)

            For i = ArrLBound(asExp) To ArrUBound(asExp)
                If Len(asExp(i)) <> 0 Then sCmdLine = sCmdLine & " /Export:" & asExp(i) & " "
            Next i

        Else

            WarnBox "Export symbols flag is set but no exports are defined", CLASS_NAME, "ThunderVB_pl_cph_v1_0_CreateProcAHookBef"

        End If

    End If

    LogMsg "New linker cmd-line " & sCmdLine, CLASS_NAME, "ThunderVB_pl_cph_v1_0_CreateProcAHookBef"
    lpCommandLine = sCmdLine
    
End Sub


Private Sub ThunderVB_pl_cph_v1_0_CreateProcAHookAft(lpApplicationName As String, lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION, ReturnCode As Long, bSkip As Boolean)
'
End Sub

Private Function ThunderVB_pl_cph_v1_0_GetBefOrder() As tvb_CP_CallOrder
    
    'Well , it is not nedded to run first, but since it edits some things it is good to run as one of the firts..
    ThunderVB_pl_cph_v1_0_GetBefOrder = cpo_Normal
                                          
End Function

Private Function ThunderVB_pl_cph_v1_0_GetAftOrder() As tvb_CP_CallOrder

    'We do not use the after hook ;)
    ThunderVB_pl_cph_v1_0_GetAftOrder = tvb_CP_CallOrder.cpo_DoNotRun

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
