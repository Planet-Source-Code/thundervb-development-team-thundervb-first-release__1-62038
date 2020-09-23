Attribute VB_Name = "modThunDll"
Option Explicit

'to this module put whatever you want

Public oMe As plugin

Public Const PLUGIN_NAME As String = "ThunderDll"
Public Const MSG_TITLE As String = PLUGIN_NAME

Public Const PLUGIN_NAMEs As String = "ThunDll"
Public Const MSG_TITLEs As String = PLUGIN_NAMEs

Public Const APP_NAME As String = PLUGIN_NAME
Public Const APP_NAMEs As String = PLUGIN_NAMEs


Private Const ASM_MOD_NAME As String = "modPreLoader__"
Private Const PROC_NAME As String = "ASM_BLOCK"

Private Const BP_MAIN As String = "BP_MAIN"
Private Const BP_CALLTHUNRTMAIN As String = "BP_CALLTHUNRTMAIN"
Private Const BP_PRELOADER As String = "BP_PRELOADER"

Private Const FULL_LOADING As String = "FULL_LOADING"
Private Const CALL_DLLMAIN As String = "CALL_DLLMAIN"

Private Const REPL_TRUE As String = "TRUE"
Private Const REPL_FALSE As String = "FALSE"

Private Const PRELOADER_ENTRYPOINT As String = "PreLoader"

Private Const FIND_DLL_MAIN As String = "@@AAGXXZ"
Private Const MASK_EXTERNDEF As String = "'#asm'  externdef ?_F@_M@@AAGXXZ:near"
Private Const MASK_CALL_EXTERN As String = "'#asm'      call ?_F@_M@@AAGXXZ"

Private Const REPL_DLLMAIN As String = "_F"
Private Const REPL_modDLLMAIN As String = "_M"

Private Const MOD_NAME As String = "modThunDll"

Dim cph As Long

Public Sub PatchPreLoader()

    'if PatchFlags failed (module ASM_MOD_NAME was not found) then skip all PatchFlags calls
    If PatchFlags(ASM_MOD_NAME, PROC_NAME, BP_MAIN, IIf(CBool(Get_DLL(BPSubMain)) = True, REPL_TRUE, REPL_FALSE)) = False Then GoTo NEXT_TASK1
    PatchFlags ASM_MOD_NAME, PROC_NAME, BP_CALLTHUNRTMAIN, IIf(CBool(Get_DLL(BPCallThunRTMain)) = True, REPL_TRUE, REPL_FALSE)
    PatchFlags ASM_MOD_NAME, PROC_NAME, BP_PRELOADER, IIf(CBool(Get_DLL((BPPreloader))) = True, REPL_TRUE, REPL_FALSE)
    PatchFlags ASM_MOD_NAME, PROC_NAME, FULL_LOADING, IIf(CBool(Get_DLL(FullLoading)) = True, REPL_TRUE, REPL_FALSE)
    PatchFlags ASM_MOD_NAME, PROC_NAME, CALL_DLLMAIN, IIf(CBool(Get_DLL(CallMyDllMain)) = True, REPL_TRUE, REPL_FALSE)
    
NEXT_TASK1:
    
    If PatchEntryPointName(ASM_MOD_NAME, PROC_NAME, MASK_EXTERNDEF) = False Then GoTo NEXT_TASK2
    PatchEntryPointName ASM_MOD_NAME, PRELOADER_ENTRYPOINT, MASK_CALL_EXTERN

NEXT_TASK2:

End Sub

Private Function PatchFlags(sMod As String, sFunc As String, sText As String, sNewValue As String) As Boolean
Dim sCode As String, asCode() As String, i As Long, sNewLine As String
        
On Error GoTo PatchFlags_Err
        
    sCode = GetFunctionCode(sMod, sFunc)
    If Len(sCode) = 0 Then
        LogMsg "Code of function " & Add34(sFunc) & " in module " & Add34(sMod) & " is zero-length or module " & Add34(sMod) & " is not in the project", MOD_NAME, "PatchFlags"
        PatchFlags = False
    Else
    
        asCode = Split(sCode, vbCrLf)
        For i = ArrLBound(asCode) To ArrUBound(asCode)
            If InStr(1, asCode(i), sText, vbTextCompare) <> 0 Then
                If InStr(1, asCode(i), sNewValue, vbTextCompare) = 0 Then
                    sNewLine = Trim(Mid(Trim(asCode(i)), 1, InStrRev(Trim(asCode(i)), " ")) & sNewValue)
                    LogMsg "Patching line " & i & ". Old code - " & Add34(asCode(i)) & ". New code - " & Add34(sNewLine), MOD_NAME, "PatchFlags"
                    SetFunctionLine sMod, sFunc, i, sNewLine
                End If
                Exit For
            End If
        Next i
        PatchFlags = True
    End If

    Exit Function

PatchFlags_Err:
    LogMsg "Error : " & Err.Description & "At line " & Erl, MOD_NAME, "PatchFlags"

End Function

Private Function PatchEntryPointName(sMod As String, sFunc As String, sMask As String) As Boolean
Dim sCode As String, i As Long, asCode() As String, sNewLine As String, sModule As String

On Error GoTo PatchEntryPointName_Err

    PatchEntryPointName = False

    sCode = GetFunctionCode(sMod, sFunc)
    sModule = FindModule(Get_DLL(EntryPointName))

    If Len(sModule) = 0 Then
        LogMsg "Dll entry-point (" & Add34(Get_DLL(EntryPointName)) & ") was not found in the project.", MOD_NAME, "PachEntryPointName"
        Exit Function
    ElseIf Len(sCode) = 0 Then
        LogMsg "Function " & Add34(sFunc) & " was not found in module " & sMod, MOD_NAME, "PachEntryPointName"
        Exit Function
    Else

        asCode = Split(sCode, vbCrLf)
        For i = ArrLBound(asCode) To ArrUBound(asCode)
            If InStr(1, asCode(i), FIND_DLL_MAIN, vbTextCompare) <> 0 Then
                sNewLine = Replace(sMask, REPL_DLLMAIN, Get_DLL(EntryPointName), 1, vbTextCompare)
                sNewLine = Replace(sNewLine, REPL_modDLLMAIN, sModule, 1, vbTextCompare)
                If sNewLine <> asCode(i) Then
                    LogMsg "Patching line " & i & ". Old code - " & Add34(asCode(i)) & ". New code - " & Add34(sNewLine), MOD_NAME, "PachEntryPointName"
                    SetFunctionLine sMod, sFunc, i, sNewLine
                End If
                Exit For
            End If
        Next i
        PatchEntryPointName = True
    End If

  Exit Function

PatchEntryPointName_Err:
    LogMsg "Error : " & Err.Description & "At line " & Erl, MOD_NAME, "PatchEntryPointName"
    
End Function

Private Function FindModule(sDllMain As String) As String
Dim asMod() As String, asFunc() As String, i As Long, j As Long

On Error GoTo FindModule_Err

    asMod = EnumModuleNames
    If IsStrArrayEmpty(asMod) = True Then Exit Function
  
    For i = ArrLBound(asMod) To ArrUBound(asMod)

        asFunc = EnumFunctionNames(asMod(i))
        If IsStrArrayEmpty(asFunc) = True Then Exit Function

        For j = ArrLBound(asFunc) To ArrUBound(asFunc)
            If StrComp(sDllMain, asFunc(j), vbBinaryCompare) = 0 Then
                FindModule = asMod(i)
                Exit Function
            End If
        Next j

    Next i

    FindModule = ""

  Exit Function

FindModule_Err:
    LogMsg "Error : " & Err.Description & "At line " & Erl, MOD_NAME, "FindModule"

End Function

Public Sub Init_Hook()
Dim t As clsCpHook
    
    Set t = New clsCpHook
    cph = AddCPH(t)
    LogMsg "CPH added " & cph, MOD_NAME, "Init_Hook"
    
End Sub

Public Sub Unload_Hook()
    
    RemoveCPH cph
    LogMsg "CPH removed " & cph, MOD_NAME, "Unload_Hook"
    
End Sub

Private Function IsStrArrayEmpty(ByRef asArray() As String) As Boolean
    Dim l As Long

    l = ArrLBound(asArray)
    IsStrArrayEmpty = (l = 2147483647)

End Function
