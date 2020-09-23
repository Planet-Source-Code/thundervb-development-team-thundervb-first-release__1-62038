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

Implements ThunderVB_pl_callbak

Private Const LINK As String = "link.exe"
Private Const OUT As String = "/OUT:"""

Private Sub ThunderVB_pl_callbak_CreateProcAHookAft(lpApplicationName As String, lpCommandLine As String, lpProcessAttributes As ThunderVB_pl_v1_0.SECURITY_ATTRIBUTES, lpThreadAttributes As ThunderVB_pl_v1_0.SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As ThunderVB_pl_v1_0.STARTUPINFO, lpProcessInformation As ThunderVB_pl_v1_0.PROCESS_INFORMATION, ReturnCode As Long, bSkip As Boolean)
    
    'find link.exe in the command-line
    If StrComp(Right(Trim(lpApplicationName), Len(LINK)), LINK, vbTextCompare) = 0 Then
        
        'check paths
        If Len(Get_Packer(Path)) = 0 Then
            MsgBoxX "Path to packer is not set.", MSG_TITLEs
            Exit Sub
        ElseIf FileExist(Get_Packer(Path)) = False Then
            MsgBoxX "File " & Add34(Get_Packer(Path)) & " does not exist.", MSG_TITLEs
            Exit Sub
        End If
        
Dim sOut As String, sPathToFile As String

        'extract path to exe file
        sPathToFile = Mid(lpCommandLine, InStr(1, lpCommandLine, OUT, vbTextCompare) + Len(OUT))
        If Len(sPathToFile) = 0 Then Exit Sub
        sPathToFile = Mid(sPathToFile, 1, InStr(1, sPathToFile, """") - 1)
        If Len(sPathToFile) = 0 Then Exit Sub

        'pack it
        ExecuteCommand Add34(Get_Packer(Path)) & " " & Get_Packer(CommandLine) & " " & Add34(sPathToFile), sOut
        If CBool(Get_Packer(ShowPackerOutPut)) = True Then MsgBoxX "Packer out-put" & CrLf(2) & sOut

    End If
    
End Sub




Private Sub ThunderVB_pl_callbak_CreateProcAHookBef(lpApplicationName As String, lpCommandLine As String, lpProcessAttributes As ThunderVB_pl_v1_0.SECURITY_ATTRIBUTES, lpThreadAttributes As ThunderVB_pl_v1_0.SECURITY_ATTRIBUTES, bInheritHandles As Long, dwCreationFlags As Long, lpEnvironment As Long, lpCurrentDirectory As String, lpStartupInfo As ThunderVB_pl_v1_0.STARTUPINFO, lpProcessInformation As ThunderVB_pl_v1_0.PROCESS_INFORMATION, bSkip As Boolean)

End Sub

Private Function ThunderVB_pl_callbak_GetAftOrder() As tvb_CP_CallOrder
    ThunderVB_pl_callbak_GetAftOrder = cpo_last
End Function

Private Function ThunderVB_pl_callbak_GetBefOrder() As tvb_CP_CallOrder
    ThunderVB_pl_callbak_GetBefOrder = cpo_DoNotRun
End Function
