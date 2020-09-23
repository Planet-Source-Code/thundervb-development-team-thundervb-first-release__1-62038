Attribute VB_Name = "modHook"
Option Explicit
'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Module Created , intial version
'Coded Hook and TrogleHook functions
'
'26/8/2004[dd/mm/yyyy] : Edited by Raziel
'Added EnumerateModules
'
'22/9/2004[dd/mm/yyyy] : Edited by Raziel
'Added HookDll_entry and HookDll_List
'and the helper functions
'


Private Declare Function OpenProcess Lib "kernel32.dll" ( _
     ByVal dwDesiredAccess As Long, _
     ByVal bInheritHandle As Long, _
     ByVal dwProcessId As Long) As Long
     
Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
     ByVal hProcess As Long, _
     ByRef lphModule As Long, _
     ByVal cb As Long, _
     ByRef lpcbNeeded As Long) As Long
     
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
     ByVal hObject As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
     ByVal hProcess As Long, _
     ByVal hModule As Long, _
     ByVal lpFilename As String, _
     ByVal nSize As Long) As Long
     

Global hookcol As DllHook_col


     
Public Function TrogleHook(ToModule As String, DllName As String, _
                     EntryName As String, newFunction As Long, _
                     oldFunction As Long, errstring As String) As HookState
    '<EhHeader>
    On Error GoTo TrogleHook_Err
    '</EhHeader>
Dim Temp As Long
    Temp = Hook(ToModule, DllName, EntryName, newFunction, errstring)
    TrogleHook = hooked
    
    If Temp = 0 Then 'failed to hook , try to unhook then
        Temp = Hook(ToModule, DllName, newFunction, oldFunction, errstring)
        TrogleHook = unhooked
    End If
    If Temp Then oldFunction = Temp
    
    '<EhFooter>
    Exit Function

TrogleHook_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "TrogleHook"
    '</EhFooter>
End Function
'return OldAddress on succes and 0 on fail
Public Function Hook(ToModule As String, DllName As String, _
                     EntryName As Variant, newFunction As Long, _
                     errstring As String) As Long
    '<EhHeader>
    On Error GoTo Hook_Err
    '</EhHeader>
                     
    Dim oldFunction As Long

    If Not HookDLLImport(ToModule, DllName, EntryName, newFunction, _
                                            oldFunction, errstring) Then
    oldFunction = 0
    End If
    
    Hook = oldFunction
    
    '<EhFooter>
    Exit Function

Hook_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "Hook"
    '</EhFooter>
End Function

'This code is taken from CompilerControler

'Hooking DLL Calls
'by John Chamberlain
'
'You can use the logic in this function to hook imports in most DLLs and EXEs
'(not just in VB). It will work for most normal Win32 modules. If you use this
'function in your own code please credit its author (me!) and include this
'descriptive header so future users will know what it does.
'
'The call addresses for all implicitly linked DLLs are located in a table
'called the "Import Address Table (IAT)" (or the "Thunk" table). This table is
'generally located at module offset 0x1000 in both DLLs and EXEs and contains
'the addresses of all imported calls in a continuous list with exports from
'different modules separated by NULL (0x 0000 0000). When each DLL is loaded
'the operating system's loader patches this table with the correct addresses.
'In most PE file types an offset to the entry point (which is just past the
'IAT) is located at offset 0xDC from the PE file header which has a signature
'of 0x00004550 (="PE"). Thus the function finds the end of the IAT by scanning
'for this signature and locating the offset.
'
'This function hooks a DLL call by first getting the proc address for the
'specified call and then scanning the IAT for the address. If it is found
'the function substitutes the hook address into the table and returns the
'original address to the caller by reference (in case the caller wants to
'restore the IAT entry to its original state at a later time). If the
'return value was false then the hook could not be set and the reason will
'be returned by reference in the string sError.
'
'When you want to restore the hooked address pass the hook address as
'vCallNameOrAddress and the original address (to be restored) as lpHook.
'The function will find the hooked address in the table and replace it with
'the original address (see UnhookCreateProcess for an example).
'
Public Function HookDLLImport(sImportingModuleName As String, _
                        sExportingModuleName As String, vCallNameOrAddress As Variant, _
                        lpHook As Long, ByRef lpOriginalAddress As Long, _
                        ByRef sError As String) As Boolean
    '<EhHeader>
    On Error GoTo HookDLLImport_Err
    '</EhHeader>
    
    Dim sCallName As String, lpPEHeader As Long
    Dim lpImportingModuleHandle As Long, lpExportingModuleHandle As Long, lpProcAddress As Long
    Dim vectorIAT As Long, lenIAT As Long, lpEndIAT As Long, lpIATCallAddress As Long
    Dim lpflOldProtect As Long, lpflOldProtect2 As Long
    
    On Error GoTo EH

    'Validate the hook
    If lpHook = 0 Then sError = "Hook is null.": Exit Function

    'Get handle (address) of importing module
    lpImportingModuleHandle = GetModuleHandle(sImportingModuleName)
    If lpImportingModuleHandle = 0 Then sError = "Unable to obtain importing module handle for """ & sImportingModuleName & """.": Exit Function

    'Get the proc address of the IAT entry to be changed
    If VarType(vCallNameOrAddress) = vbString Then
    
        sCallName = CStr(vCallNameOrAddress)    'user is hooking an import
    
        'Get handle (address) of exporting module
        lpExportingModuleHandle = GetModuleHandle(sExportingModuleName)
        If lpExportingModuleHandle = 0 Then sError = "Unable to obtain exporting module handle for """ & sExportingModuleName & """.": Exit Function
    
        'Get address of call
        lpProcAddress = GetProcAddress(lpExportingModuleHandle, sCallName)
        If lpProcAddress = 0 Then sError = "Unable to obtain proc address for """ & sCallName & """.": Exit Function
    
    Else
        lpProcAddress = CLng(vCallNameOrAddress) 'user is restoring a hooked import
    End If

    'Beginning of the IAT is located at offset 0x1000 in most PE modules
    vectorIAT = lpImportingModuleHandle + &H1000

    'Scan module to find PE header by looking for header signature
    lpPEHeader = lpImportingModuleHandle
    Do
        If lpPEHeader > vectorIAT Then  'this is not a PE module
            sError = "Module """ & sImportingModuleName & """ is not a PE module."
            Exit Function
        Else
            If Deref(lpPEHeader) = IMAGE_NT_SIGNATURE Then  'we have located the module's PE header
                Exit Do
            Else
                lpPEHeader = lpPEHeader + 1 'keep searching
            End If
        End If
    Loop
    
    'Determine and validate length of the IAT. The length is at offset 0xDC in the PE header.
    lenIAT = Deref(lpPEHeader + &HDC)
    If lenIAT = 0 Or lenIAT > &HFFFFF Then 'its too big or too small to be valid
        sError = "The calculated length of the Import Address Table in """ & sImportingModuleName & """ is not valid: " & lenIAT
        Exit Function
    End If

    'Scan Import Address Table for proc address
    lpEndIAT = lpImportingModuleHandle + &H1000 + lenIAT
    Do
        If vectorIAT > lpEndIAT Then 'we have reached the end of the table
            sError = "Proc address " & Hex$(lpProcAddress) & " not found in Import Address Table of """ & sImportingModuleName & """."
            Exit Function
        Else
            lpIATCallAddress = Deref(vectorIAT)
            If lpIATCallAddress = lpProcAddress Then  'we have found the entry
                Exit Do
            Else
                vectorIAT = vectorIAT + 4   'try next entry in table
            End If
        End If
    Loop
    
    'Substitute hook for existing call address and return existing address by ref
    'We must make this memory writable to make the entry in the IAT
    If VirtualProtect(ByVal vectorIAT, 4, PAGE_EXECUTE_READWRITE, lpflOldProtect) = 0 Then
        sError = "Unable to change IAT memory to execute/read/write."
        Exit Function
    Else
        lpOriginalAddress = Deref(vectorIAT)    'save original address
        CopyMemory ByVal vectorIAT, lpHook, 4    'set the hook
        VirtualProtect ByVal vectorIAT, 4, lpflOldProtect, lpflOldProtect2  'restore memory protection
    End If

    HookDLLImport = True 'mission accomplished
Exit Function
    
EH:
    sError = "Unexpected error: " & err.Description

    '<EhFooter>
    Exit Function

HookDLLImport_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "HookDLLImport"
    '</EhFooter>
End Function

Function Deref(lngPointer As Long) As Long  'Equivalent of *lngPointer (returns the value pointed to)
    '<EhHeader>
    On Error GoTo Deref_Err
    '</EhHeader>
Dim lngValueAtPointer As Long

    CopyMemory lngValueAtPointer, ByVal lngPointer, 4
    Deref = lngValueAtPointer
    
    '<EhFooter>
    Exit Function

Deref_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "Deref"
    '</EhFooter>
End Function

Function GetAddress(adrof As Long) As Long
    '<EhHeader>
    On Error GoTo GetAddress_Err
    '</EhHeader>

    GetAddress = adrof
    
    '<EhFooter>
    Exit Function

GetAddress_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "GetAddress"
    '</EhFooter>
End Function

'Enumerate all modules...
'Usefull for hooking all module's import tables ;)

Function EnumerateModules(proc As Long) As module_list
    '<EhHeader>
    On Error GoTo EnumerateModules_Err
    '</EhHeader>
Dim hMods(1024) As Long, hProcess As Long, cbNeeded As Long, i As Long

    'Get a list of all the modules in this process.
    hProcess = OpenProcess(1040, False, proc)
    If hProcess = 0 Then Exit Function

    If EnumProcessModules(hProcess, hMods(0), 1024 * 4, cbNeeded) Then
     For i = 0 To (cbNeeded / 4)
            Dim pth As String

            'Get the full path to the module's file.
            pth = Space$(1024)
            If GetModuleFileNameEx(hProcess, hMods(i), pth, 1024) Then
                With EnumerateModules
                    ReDim Preserve .modules(.count)
                    .modules(.count).Id = hMods(i)
                    .modules(.count).Name = pth
                    .count = .count + 1
                End With
            End If
        Next i
    End If
    
    CloseHandle hProcess

    '<EhFooter>
    Exit Function

EnumerateModules_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "EnumerateModules"
    '</EhFooter>
End Function

Sub AddDllHookEntryToList(list As DllHook_list, item As DllHook_entry)
    '<EhHeader>
    On Error GoTo AddDllHookEntryToList_Err
    '</EhHeader>
    
    ReDim Preserve list.items(list.count)
    
    list.items(list.count) = item
    list.count = list.count + 1

    '<EhFooter>
    Exit Sub

AddDllHookEntryToList_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "AddDllHookEntryToList"
    '</EhFooter>
End Sub

Function NewDllHook_entry(ToModule As String, DllName As String, FunctionName As String, _
                          FunctionAddress As Long, HookAddress As Long, _
                          State As HookState) As DllHook_entry
    '<EhHeader>
    On Error GoTo NewDllHook_entry_Err
    '</EhHeader>
                          
    Dim Temp As DllHook_entry
    
        With Temp
            .ToModule = ToModule
            .DllName = DllName
            .FunctionName = FunctionName
            .FunctionAddress = FunctionAddress
            .HookAddress = HookAddress
            .State = State
        End With
        
    NewDllHook_entry = Temp

    '<EhFooter>
    Exit Function

NewDllHook_entry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "NewDllHook_entry"
    '</EhFooter>
End Function

Public Function SetHookAndRet(ByVal Todll As String, ByVal dll As String, ByVal funct As String, _
                              ByVal Address As Long, ByRef ErrorString As String) As DllHook_entry
    '<EhHeader>
    On Error GoTo SetHookAndRet_Err
    '</EhHeader>
Dim Temp As Long

    Temp = Hook(Todll, dll, funct, Address, ErrorString)
    If Temp = 0 Then
        LogMsg "Unable to set " & funct & " Hook (" & Trim$(Todll) & ") ", APP_NAME, "modHook", "SetHookAndRet"
    Else
        LogMsg funct & " Hook was set (" & Trim$(Todll) & ")", APP_NAME, "modHook", "SetHookAndRet"
        SetHookAndRet = NewDllHook_entry(Todll, dll, funct, Temp, Address, hooked)
    End If
    
    '<EhFooter>
    Exit Function

SetHookAndRet_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "SetHookAndRet"
    '</EhFooter>
End Function
  

Sub KillHookList(Hooks As DllHook_list)
    '<EhHeader>
    On Error GoTo KillHookList_Err
    '</EhHeader>
Dim Temp As Long, strtemp As String, i As Long
    hookcol.items(Hooks.Id).count = 0
    For i = Hooks.count - 1 To 0 Step -1
        With Hooks.items(i)
            
            If .State = hooked Then
                
                Temp = .FunctionAddress
                .State = TrogleHook(.ToModule, .DllName, .FunctionName, .HookAddress, Temp, strtemp)
                If Temp = 0 Then
                    WarnBox "KillHookList:" & vbNewLine & strtemp, "modHook", "KillHookList"
                    LogMsg "Unable to unset " & .FunctionName & " Hook (" & Trim$(.ToModule) & ")", APP_NAME, "modHook", "KillHookList"
                Else
                    LogMsg "Killed " & .FunctionName & " Hook (" & Trim$(.ToModule) & ")", APP_NAME, "modHook", "KillHookList"
                End If
                
            End If
            
        End With
    Next i
    
    Hooks.count = 0

    '<EhFooter>
    Exit Sub

KillHookList_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "KillHookList"
    '</EhFooter>
End Sub

Sub CreateHookList_int(Hooks As DllHook_list, DllName As String, _
                   FunctionName As String, HookAddress As Long, Optional inModule As String)
    '<EhHeader>
    On Error GoTo CreateHookList_int_Err
    '</EhHeader>

Dim Temp As DllHook_entry, strtemp As String
Dim dlls As module_list, i As Long
    If Len(inModule) = 0 Then
    
        dlls = EnumerateModules(GetCurrentProcessId)
        For i = 0 To dlls.count - 1
            With dlls.modules(i)
            
                'hook ExtTextOutA
                Temp = SetHookAndRet(.Name, DllName, FunctionName, HookAddress, strtemp)
                If Temp.FunctionAddress > 0 Then
                    AddDllHookEntryToList Hooks, Temp
                End If

            End With
        Next i
    Else
                Temp = SetHookAndRet(inModule, DllName, FunctionName, HookAddress, strtemp)
                If Temp.FunctionAddress > 0 Then
                    AddDllHookEntryToList Hooks, Temp
                End If
    End If
    AddToHookCol Hooks 'register it so that it can be unset on unload
                       'While all plugins must UNLOAD their hoosk before unloadig
                       'To be sure that everything is ok after ThunVB unload
                       'We unload any one that have been left..
    '<EhFooter>
    Exit Sub

CreateHookList_int_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "CreateHookList_int"
    '</EhFooter>
End Sub

Public Sub AddToHookCol(hklst As DllHook_list)
    '<EhHeader>
    On Error GoTo AddToHookCol_Err
    '</EhHeader>

    If hklst.count Then
        hklst.Id = hookcol.count
        ReDim Preserve hookcol.items(hookcol.count)
        hookcol.items(hookcol.count) = hklst
        hookcol.count = hookcol.count + 1
    End If
    
    '<EhFooter>
    Exit Sub

AddToHookCol_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "AddToHookCol"
    '</EhFooter>
End Sub

Public Sub UnHookEverything()
    '<EhHeader>
    On Error GoTo UnHookEverything_Err
    '</EhHeader>
Dim i As Long

    For i = 0 To hookcol.count - 1
        KillHookList hookcol.items(i)
    Next i
    
    '<EhFooter>
    Exit Sub

UnHookEverything_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modHook", "UnHookEverything"
    '</EhFooter>
End Sub
