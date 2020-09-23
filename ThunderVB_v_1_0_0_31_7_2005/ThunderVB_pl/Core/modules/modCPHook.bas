Attribute VB_Name = "modCPHook"
Option Explicit
Dim cphl As DllHook_list
Global cph_c As cph_list

Public Sub InitCPHook()
    
    If DebugMode = False And cphl.count = 0 Then
        Call CreateHookList_ResumeOldList("kernel32.dll", "CreateProcessA", AddressOf CreateProcess_Hook, cphl, "vba6.dll")
    End If
End Sub

Public Sub KillCPHook()
    
    If DebugMode = False Then
        Call KillHookList(cphl)
    End If
    
End Sub

Public Function CreateProcess_Hook(ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, _
                                    lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                    lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                    ByVal bInheritHandles As Long, _
                                    ByVal dwCreationFlags As Long, _
                                    lpEnvironment As Long, _
                                    ByVal lpCurrentDirectory As Long, _
                                    lpStartupInfo As STARTUPINFO, _
                                    lpProcessInformation As PROCESS_INFORMATION) As Long
    Dim i As tvb_CP_CallOrder, sName As String, sLine As String, _
        sdir As String, bSkip As Boolean

    'get the needed info
    sName = CStringZero(lpApplicationName)
    sLine = CStringZero(lpCommandLine)
    sdir = CStringZero(lpCurrentDirectory)
    
    LogMsg "CreateProcess_Hook Called(" & sName & "," & sLine & "," & sdir & ")", "modCreateProcHook", "CreateProcess_Hook"
    
        
    If InStr(1, sName, "link.exe", vbTextCompare) Then
        CheckAndFixLinker sName, sLine
    End If
    
    For i = cpo_First To cpo_last
            CallBef i, sName, sLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, sdir, lpStartupInfo, lpProcessInformation, bSkip
    Next i
    
    If bSkip = False Then
        CreateProcess_Hook = CreateProcess(sName, sLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, sdir, lpStartupInfo, lpProcessInformation)
        LogMsg "Return value :" & CreateProcess_Hook, "modCreateProcHook", "CreateProcess_Hook"
        WaitForSingleObjectEx lpProcessInformation.hProcess, 100000, False
    End If

    
    For i = cpo_First To cpo_last
        CallAft i, sName, sLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, sdir, lpStartupInfo, lpProcessInformation, CreateProcess_Hook, bSkip
    Next i
    
End Function

Public Static Function AddCPH(inter As ThunderVB_pl_cph_v1_0) As Long
    Dim i As Long, BefFirst As Boolean, BefLast As Boolean
    Dim AftFirst As Boolean, AftLast As Boolean
    'look to find any one that was "opened"
    For i = 0 To cph_c.count - 1
        
        If Not (cph_c.items(i) Is Nothing) Then
            If cph_c.items(i).GetBefOrder = cpo_First Then
                BefFirst = True
            ElseIf cph_c.items(i).GetBefOrder = cpo_last Then
                BefLast = True
            End If
            
            If cph_c.items(i).GetAftOrder = cpo_First Then
                AftFirst = True
            ElseIf cph_c.items(i).GetAftOrder = cpo_last Then
                AftLast = True
            End If
        End If
        
    Next i
    
    If (inter.GetBefOrder() = cpo_First) And BefFirst = True Then
        Err.Raise ThunVB_Errors.tvb_CPH_Before_First_Exists, "AddCPH", "There is allready a before first entry"
        Exit Function
    ElseIf (inter.GetBefOrder() = cpo_last) And BefLast = True Then
        Err.Raise ThunVB_Errors.tvb_CPH_Before_Last_Exists, "AddCPH", "There is allready a before last entry"
        Exit Function
    ElseIf (inter.GetAftOrder() = cpo_First) And AftLast = True Then
        Err.Raise ThunVB_Errors.tvb_CPH_After_First_Exists, "AddCPH", "There is allready a after last entry"
        Exit Function
    ElseIf (inter.GetAftOrder() = cpo_last) And AftLast = True Then
        Err.Raise ThunVB_Errors.tvb_CPH_After_Last_Exists, "AddCPH", "There is allready a after last entry"
        Exit Function
    End If
    
    For i = 0 To cph_c.count - 1
        If cph_c.items(i) Is Nothing Then GoTo found
    Next i
    
    ReDim Preserve cph_c.items(cph_c.count) 'i= this now .. [after the for , i is tha last value for for+1]
    cph_c.count = cph_c.count + 1
found:
    Set cph_c.items(i) = inter
    AddCPH = i
End Function


Public Sub RemoveCPH(cph As Long)

    Set cph_c.items(cph) = Nothing
    
End Sub


Private Sub CallBef(ByVal Filter As tvb_CP_CallOrder, ByRef sName As String, ByRef sLine As String, _
                              ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
                              ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
                              ByRef bInheritHandles As Long, ByRef dwCreationFlags As Long, _
                              ByRef lpEnvironment As Long, ByRef sdir As String, _
                              ByRef lpStartupInfo As STARTUPINFO, _
                              ByRef lpProcessInformation As PROCESS_INFORMATION, _
                              ByRef bSkip As Boolean)
    Dim i As Long
    For i = 0 To cph_c.count - 1
        If Not (cph_c.items(i) Is Nothing) Then
            If cph_c.items(i).GetBefOrder = Filter Then
                cph_c.items(i).CreateProcAHookBef sName, sLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, sdir, lpStartupInfo, lpProcessInformation, bSkip
            End If
        End If
    Next i
    
End Sub


Private Sub CallAft(ByVal Filter As tvb_CP_CallOrder, ByRef sName As String, ByRef sLine As String, _
                                   ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                   ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                   ByRef bInheritHandles As Long, ByRef dwCreationFlags As Long, _
                                   ByRef lpEnvironment As Long, ByRef sdir As String, _
                                   ByRef lpStartupInfo As STARTUPINFO, _
                                   ByRef lpProcessInformation As PROCESS_INFORMATION, _
                                   ByRef ReturnCode As Long, ByRef bSkip As Boolean _
                                   )
    Dim i As Long
    
    For i = 0 To cph_c.count - 1
        If Not (cph_c.items(i) Is Nothing) Then
            If cph_c.items(i).GetAftOrder = Filter Then
                cph_c.items(i).CreateProcAHookAft sName, sLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, sdir, lpStartupInfo, lpProcessInformation, ReturnCode, bSkip
            End If
        End If
    Next i
    
End Sub


Public Sub CheckAndFixLinker(sName As String, sLine As String)
    
    If GetLinkerVersion(sName) < 7 Then
        sLine = Replace(sLine, "OPT:REF", "OPT:NOREF", 1, -1, vbTextCompare)
        LogMsg "CheckAndFixLinker, Adding OPT:NOREF", "modCPHook", "CheckAndFixLinker"
    End If
    
    'LogMsg "CheckAndFixLinker, Adding /MERGE:text=.text", "modCPHook", "CheckAndFixLinker"
    'sLine = sLine & " /MERGE:text=.text"
    
End Sub

Public Function GetLinkerVersion(file As String) As Long
Dim fv As String
    fv = GetFileVersion(file)
    
    If Len(fv) Then
        GetLinkerVersion = CLng(Split(fv, ".")(0))
    Else
        GetLinkerVersion = 0
    End If
    
End Function

