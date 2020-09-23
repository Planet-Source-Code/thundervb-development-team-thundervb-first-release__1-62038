Attribute VB_Name = "LaunchAppSynchronousMod"
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function NtQueryInformationProcess Lib "ntdll" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As Long, ByRef ProcessInformation As PROCESS_BASIC_INFORMATION, ByVal lProcessInformationLength As Long, ByRef lReturnLength As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 300
End Type

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260
Private Const SW_SHOW As Integer = 5
Private Const PROCESS_VM_READ = 16

Public Function ShellExecAndWait(sFileName As String, sParams As String, bWaitToEnd As Boolean, oForm As Form) As Boolean
    
    ShellExecAndWait = False
    
    Dim hSnapShot As Long
    Dim hProcess As Long
    Dim uProcess As PROCESSENTRY32
    Dim lRC As Long
    Dim lShelledProcessID As Long
    Dim lMyProcessID As Long
    Dim myOS As OSVERSIONINFO
    Dim iWinVersion As Integer
    lMyProcessID = GetCurrentProcessId
    
    lRC = ShellExecute(oForm.hwnd, "open", sFileName, sParams, CurDir$, SW_SHOW)
    
    If lRC < 32 Then
        Exit Function
    End If

    If Not bWaitToEnd Then
        Exit Function
    End If
    
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    
    iWinVersion = myOS.dwPlatformId

    Select Case iWinVersion
        Case VER_PLATFORM_WIN32_NT
            Dim cb As Long
            Dim cbNeeded As Long
            Dim NumElements As Long
            Dim ProcessIDs() As Long
            Dim lRet As Long
            Dim i As Long
            
            cb = 8
            cbNeeded = 96
            
            Do While cb <= cbNeeded
                cb = cb * 2
                ReDim ProcessIDs(cb / 4) As Long
                lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
            Loop
            NumElements = cbNeeded / 4
            
            For i = 1 To NumElements
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
                Dim lntStatus As Long
                Dim lProcessHandle As Long
                Dim lProcessBasicInfo As Long
                Dim tProcessInformation As PROCESS_BASIC_INFORMATION
                Dim lProcessInformationLength As Long
                Dim lReturnLength As Long
                lProcessHandle = hProcess
                lProcessBasicInfo = 0
                lProcessInformationLength = Len(tProcessInformation)
                lntStatus = NtQueryInformationProcess(lProcessHandle, lProcessBasicInfo, tProcessInformation, lProcessInformationLength, lReturnLength)
                If lMyProcessID = tProcessInformation.InheritedFromUniqueProcessId Then
                    lShelledProcessID = tProcessInformation.UniqueProcessId
                    lRet = CloseHandle(hProcess)
                    Exit For
                End If
                lRet = CloseHandle(hProcess)
            Next
        
        Case VER_PLATFORM_WIN32_WINDOWS
            hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
            If hSnapShot = 0 Then Exit Function
            uProcess.dwSize = Len(uProcess)
            lRC = ProcessFirst(hSnapShot, uProcess)
    
            Do While lRC
                If lMyProcessID = uProcess.th32ParentProcessID Then
                    lShelledProcessID = uProcess.th32ProcessID
                    Exit Do
                End If
                lRC = ProcessNext(hSnapShot, uProcess)
            Loop
            Call CloseHandle(hSnapShot)
        Case Else
    End Select

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lShelledProcessID)

    Do
        Call GetExitCodeProcess(hProcess, lRC)
        DoEvents
    Loop While lRC > 0
    Call CloseHandle(hProcess)
    
    ShellExecAndWait = True
    
End Function

