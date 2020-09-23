VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ThunderVB_pl_cph_v1_0


'edit the compiler command line if needed (the module has Asm/C code)
Private Sub ThunderVB_pl_cph_v1_0_CreateProcAHookBef(lpApplication As String, lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION, bSkip As Boolean)
On Error GoTo error_h
GPF_Set GPF_RaiseErr, "ThunAsm::cCPHook", "ThunderVB_pl_callbak_CreateProcAHookBef", PLUGIN_NAME
    If InStr(1, LCase$(lpCommandLine), "c2", vbTextCompare) <> 0 Then
        
        Fix_unnamed = CBool(Get_ASM(FixASMListings))
        file_vb = getS("-f " & Chr(34), Chr(34), lpCommandLine)
        file_obj = getS("-Fo" & Chr(34), Chr(34), lpCommandLine)  'obj file
        file_asm = Replace(file_obj, ".obj", ".asm", , , vbTextCompare) 'output asm file
    
        If FileHasAsmOrC(file_vb) Then  'module has asm code
            If CBool(Get_ASM(CompileASMCode)) = False Then
                GoTo noasm
            End If
            If (Get_General(GenerateAsmCHeaders) = True) And (donot_GenHeaders = False) Then
                donot_GenHeaders = True
                ah_path = GetPath(file_asm) & "vb_imports_ct_agen_asm"
                ch_path = GetPath(file_asm) & "vb_imports_ct_agen_c.h"
                AsmCHeadersAreAssembled = GenerateAsmCHeadersAndSave(ah_path, ch_path)
            End If
            LogMsg "Preprocesing Compiler Call for inline Asm ...", "cCPHook", "ThunderVB_pl_callbak_CreateProcAHookBef"
            'set the cmd line to create a asm file instead of an object file
            If Fix_unnamed Then
                lpCommandLine = Replace(lpCommandLine, "-Fo" & Add34(file_obj), "-FAsc -Fa" & Add34(file_asm))
            Else
                lpCommandLine = Replace(lpCommandLine, "-Fo" & Add34(file_obj), "-FAs -Fa" & Add34(file_asm))
            End If
    
            If ProjectSaved = False Then
                If MsgBoxX("Project Must be saved for asm changes to have effect" & vbNewLine & _
                       "You want to save it now?", , vbYesNo Or vbInformation) = vbYes Then
                    SaveProjects True
                End If
            End If
            has_asm = 1
        Else
noasm:
            has_asm = 0
            If CBool(Get_C(CompileCCode)) = True And CBool(Get_ASM(CompileASMCode)) = False Then
                WarnBox "Asm MUST be enabled to compile C code..", "modCreateProcHook", "c2_edit"
            End If
            If Get_Compile(ModifyCmdLine) Then
                lpCommandLine = frmViewer.ShowViewer("Modify C2 command line", lpCommandLine, False)
            End If
        End If
    ElseIf InStr(1, LCase$(lpApplication), "link", vbTextCompare) <> 0 Then
        If Get_General(GenerateAsmCHeaders) And AsmCHeadersAreAssembled Then
            donot_GenHeaders = False
            lpCommandLine = lpCommandLine & " " & Add34(ah_path & ".obj")
        End If
        AddCLibs lpCommandLine
        If Get_Compile(ModifyCmdLine) Then
            lpCommandLine = frmViewer.ShowViewer("Modify Linker command line", lpCommandLine, False)
        End If
        
    End If

GPF_Reset
Exit Sub
error_h:
    If MsgBoxX(Err.Description & vbNewLine & "Continue?", PLUGIN_NAMEs, vbYesNo) = vbYes Then
        Resume Next
    Else
        GPF_Reset
    End If
End Sub

Private Sub ThunderVB_pl_cph_v1_0_CreateProcAHookAft(lpApplicationName As String, lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION, ReturnCode As Long, bSkip As Boolean)
    
    If InStr(1, LCase$(lpApplicationName), "c2", vbTextCompare) <> 0 Then
        If has_asm Then 'well we have asm
            ProcInlineAsm lpApplicationName
            has_asm = False
        End If
    ElseIf InStr(1, LCase$(lpApplicationName), "link", vbTextCompare) <> 0 Then
        kill2 ah_path & ".obj"
        kill2 ah_path & ".asm"
        kill2 ch_path
    End If

End Sub

Private Function ThunderVB_pl_cph_v1_0_GetAftOrder() As tvb_CP_CallOrder
    
    ThunderVB_pl_cph_v1_0_GetAftOrder = cpo_First
    
End Function

Private Function ThunderVB_pl_cph_v1_0_GetBefOrder() As tvb_CP_CallOrder

    ThunderVB_pl_cph_v1_0_GetBefOrder = cpo_9

End Function
