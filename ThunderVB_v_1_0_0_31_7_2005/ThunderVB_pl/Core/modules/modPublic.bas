Attribute VB_Name = "modPublic"
Option Explicit

'this module will contain public functions for general use

'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Libor
'Module created , initial version
'
'21/8/2004[dd/mm/yyyy] : Code Edited by Libor
'added new Settings for StdCall DLL, new Save* and Read* functions: Libor
'NOTE : SaveVBP and ReadVBP needs patching
'
'26/8/2004[dd/mm/yyy]  : Code Edited by Raziel
'Added WarnBox,ErrorBox and MsgBoxX
'
'27/8/2004[dd/mm/yyyy]
'Save* and Read* functions moved to frmSettings : Libor
'
'26/8/2004[dd/mm/yyy]  : Code Edited by Raziel
'Added LogMsg
'
'7/9/2004[dd/mm/yyyy]
'Added new settings - ASM/C
'New function LoadSettings    :Libor
'
'9/9/2004[dd/mm/yyy]  : Code Edited by Raziel
'Added Putstringtocurpos
'fixed ProcStrings , now it works as it should
'
'10/9/2004[dd/mm/yyyy] :Code edited by Raziel
'Code to convert #asm_start .. #asm_end to '#asm' lines..
'
'12/9/2004[dd/mm/yyyy]
'changed settings in StdCall Tab : Libor
'
'13/9/2004[dd/mm/yyyy]
'added new option to Get_Packer, new function CrLf : Libor
'
'16/9/2004[dd/mm/yyyy]
'added "Show packer output", "Add to menu" option, better CrLf function : Libor
'added function DirExist, LoadFile function patched
'
'19/9/2004[dd/mm/yyyy]
'patched function LogMsg function, all constants are private : Libor
'
'22/9/2004[dd/mm/yyyy] : Raziel
'Minor Changes on the logging code
'
'
'1/10/2004[dd/mm/yyyy] : Raziel
'GetFunctionCode ,EnumFunctionNames,EnumModuleNames
'SetFunctionLine ,SetCurLine
'6/10/2004
'Sligth mods for GPF handling..
'Name
'7/10/2004 added save projects
'
'10/10/2004 added procStringUnderAll , GetThunVBVer and ReplWSwithSpace

'
'pfew ... that was hard..
'Rather MUCH MANY LOT changes to convert to plugined code..
'All non generic code has been removed...

'ok , started re logging changes :)
'16/7/2005 Fixed SaveProjects work with frx files ect :) (drkIIRaziel)

Public Const tvb_Error As Long = 10240

'Other
Public vb_Dll_version As Long '5/6

'Loging

'Language

'------------------------
'--- Helper Functions ---
'------------------------
Public menbut As ButArray


Function PutStringToCurCursor(str As String) As Boolean
Dim curline As Long

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveCodePane Is Nothing Then Exit Function
    If VBI.ActiveCodePane.codeModule Is Nothing Then Exit Function
    
    With VBI.ActiveCodePane
        .Window.SetFocus
        .GetSelection curline, 0, 0, 0
        .codeModule.InsertLines curline, str
    End With
    PutStringToCurCursor = True
    
End Function


Function GetFunctionCode(inMod As String, funct As String) As String
Dim pk As vbext_ProcKind, s As String
Dim pcount As Long, pline As Long, sLines As String
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strTemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveVBProject Is Nothing Then Exit Function
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Function
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.type = vbext_ct_StdModule Or _
            objComponent.type = vbext_ct_ClassModule Or _
            objComponent.type = vbext_ct_VBForm) And objComponent.Name = inMod Then
            For Each objMember In objComponent.codeModule.Members
                If objMember.type = vbext_mt_Method And objMember.Name = funct Then
                    With objComponent.codeModule
                        pline = .ProcBodyLine(funct, pk)
                        pcount = .ProcCountLines(funct, pk)
                        sLines = .Lines(pline, pcount)
                    End With
                End If
            Next
        End If
    Next

    GetFunctionCode = sLines
    
End Function

Function EnumModuleNames(Optional moduleType As vbext_ComponentType = -1) As String()
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strTemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveVBProject Is Nothing Then Exit Function
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Function
    
    'enumerate the procedures in every module file within
    'the current project
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If ((moduleType = -1) And (objComponent.type = vbext_ct_StdModule Or _
           objComponent.type = vbext_ct_ClassModule Or _
           objComponent.type = vbext_ct_VBForm)) Or (moduleType = objComponent.type) Then
           
           ReDim Preserve nams(namsC)
           nams(namsC) = objComponent.Name
           namsC = namsC + 1
           
        End If
    Next
    
    EnumModuleNames = nams
    
End Function
    

Function EnumFunctionNames(inMod As String) As String()

    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strTemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Function
    If VBI.ActiveVBProject Is Nothing Then Exit Function
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Function
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.type = vbext_ct_StdModule Or _
            objComponent.type = vbext_ct_ClassModule Or _
            objComponent.type = vbext_ct_VBForm) And objComponent.Name = inMod Then
            Dim stemp() As String
            Dim ite As Long
            
            If objComponent.codeModule.CountOfLines > 0 Then
                On Error Resume Next
                stemp = Split(objComponent.codeModule.Lines(1, objComponent.codeModule.CountOfLines), vbNewLine)
                On Error GoTo 0
            End If
            
            For ite = 0 To ArrUBound(stemp)
                On Error GoTo NextOne
                stemp(ite) = Trim$(stemp(ite))
                If GetFirstWord(LCase$(stemp(ite))) = "public" Or _
                   GetFirstWord(LCase$(stemp(ite))) = "private" Then
                   RemFisrtWord stemp(ite): stemp(ite) = Trim$(stemp(ite))
                End If
                
                If GetFirstWord(LCase$(stemp(ite))) = "sub" Or _
                   GetFirstWord(LCase$(stemp(ite))) = "function" Then
                   RemFisrtWord stemp(ite): stemp(ite) = Replace$(Trim$(stemp(ite)), "(", " ")
                    ReDim Preserve nams(namsC)
                    nams(namsC) = GetFirstWord(stemp(ite))
                    namsC = namsC + 1
                End If
NextOne:
            Next
        End If
    Next
    
    EnumFunctionNames = nams

End Function

Sub SetCurLine(inMod As String, funct As String, NewTopLine As Long)
Dim pk As vbext_ProcKind, s As String
Dim pcount As Long, pline As Long, sLines As String
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strTemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveVBProject Is Nothing Then Exit Sub
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Sub
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.type = vbext_ct_StdModule Or _
            objComponent.type = vbext_ct_ClassModule Or _
            objComponent.type = vbext_ct_VBForm) And objComponent.Name = inMod Then
            For Each objMember In objComponent.codeModule.Members
                If objMember.type = vbext_mt_Method And objMember.Name = funct Then
                    With objComponent.codeModule
                        pline = .ProcBodyLine(funct, pk)
                        .CodePane.TopLine = pline + NewTopLine
                    End With
                End If
            Next
        End If
    Next

    
End Sub

Sub SetFunctionLine(inMod As String, funct As String, LineNum As Long, newLineString As String, Optional bReplace As Boolean = True)
Dim pk As vbext_ProcKind, s As String
Dim pcount As Long, pline As Long, sLines As String
    Dim nams() As String, namsC As Long
    Dim objComponent As VBComponent
    Dim objMember As Member
    Dim strTemp As String
    Dim intTemp As Integer

    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveVBProject Is Nothing Then Exit Sub
    If VBI.ActiveVBProject.VBComponents Is Nothing Then Exit Sub
    
    For Each objComponent In VBI.ActiveVBProject.VBComponents
        If (objComponent.type = vbext_ct_StdModule Or _
            objComponent.type = vbext_ct_ClassModule Or _
            objComponent.type = vbext_ct_VBForm) And objComponent.Name = inMod Then
            For Each objMember In objComponent.codeModule.Members
                If objMember.type = vbext_mt_Method And objMember.Name = funct Then
                    With objComponent.codeModule
                        pline = .ProcBodyLine(funct, pk)
                        If bReplace Then
                            .ReplaceLine pline + LineNum, newLineString
                        Else
                            .InsertLines pline + LineNum, newLineString
                        End If
                    End With
                End If
            Next
        End If
    Next
    
End Sub

'Save all projects ...
Public Sub SaveProjects(binf As Boolean)

Dim i As Long, i2 As Long
    
    For i = 1 To VBI.VBProjects.count
    
        With VBI.VBProjects(i)
            .SaveAs .FileName
            For i2 = 1 To .VBComponents.count
            
                If .VBComponents(i2).FileCount > 1 Then
                    'For i3 = 1 To .VBComponents(i2).FileCount
                    'Finaly fixed :P
                    .VBComponents(i2).SaveAs .VBComponents(i2).FileNames(1)
                    'Next i3
                Else
                    .VBComponents(i2).SaveAs .VBComponents(i2).FileNames(1)
                End If
                
            Next i2
            If binf Then
                MsgBoxX "Project " & Add34(.Name) & " was saved", , vbInformation Or vbOKOnly
            End If
        End With
        
    Next i
    
End Sub

Public Function GetThunVBVer() As String

    GetThunVBVer = " Version " & App.Major & "." & App.Minor & "." & App.Revision
    
End Function

Function GetCodeWinParent() As Long
Dim ret As Long
    
    On Error Resume Next

    If VBI.MainWindow Is Nothing Then
        ret = GetDesktopWindow
    Else
        
        If VBI.MainWindow.hWnd <> 0 Then
            ret = VBI.MainWindow.hWnd
        Else
            ret = GetDesktopWindow
        End If
    End If
    
    GetCodeWinParent = ret

End Function

'------------------------
'- GUI Helper functions -
'------------------------
Function AddToAddinMenu(ipl As ThunderVB_pl_int_v1_0, sCaption As String, sTip As String, defimage As IPictureDisp) As Long
    
    AddToAddinMenu = AddButton(ipl, "Add-Ins", sTip, sCaption, defimage)

End Function

Function AddButtonToToolbar(ipl As ThunderVB_pl_int_v1_0, tip As String, capt As String, defimage As IPictureDisp) As Long
    
    AddButtonToToolbar = AddButton(ipl, 2, tip, capt, defimage)
    
End Function


Function AddButtonToCodeWindowMenu(ipl As ThunderVB_pl_int_v1_0, tip As String, capt As String, defimage As IPictureDisp) As Long
    
    AddButtonToCodeWindowMenu = AddButton(ipl, "Code Window", tip, capt, defimage)
    
End Function

Function AddButton(ipl As ThunderVB_pl_int_v1_0, ToBar As Variant, tip As String, capt As String, defimage As IPictureDisp) As Long
Dim oldclip As Clipboard, Temp As CommandBarControl, tempbut As New clsComBut, ClipBoardBackup As IPictureDisp

    Set Temp = VBI.CommandBars(ToBar).Controls.Add(msoControlButton)
    Temp.caption = capt
    On Error Resume Next
    If Not (defimage Is Nothing) Then
        If Clipboard.GetFormat(vbCFBitmap) Then
            Set ClipBoardBackup = Clipboard.GetData(vbCFBitmap)
        Else
            Set ClipBoardBackup = New StdPicture
        End If
        Clipboard.SetData defimage, vbCFBitmap
        Temp.PasteFace
        Clipboard.SetData ClipBoardBackup, vbCFBitmap
    End If
    On Error GoTo 0
    Temp.ToolTipText = tip
    tempbut.init Temp, ipl, menbut.count

    AddButton = AddBut(tempbut)
    LogMsg "Added button named : " & capt, "modPublic", "AddButton"
    
End Function

Public Function AddBut(cls As clsComBut) As Long

    ReDim Preserve menbut.items(menbut.count)
    Set menbut.items(menbut.count) = cls
    AddBut = menbut.count
    menbut.count = menbut.count + 1
    
End Function

Public Sub RemoveButtton(Id As Long)

    Set menbut.items(Id) = Nothing
    
End Sub

Public Function ProjectSaved() As Boolean
    If ProjectIsLoaded Then
        ProjectSaved = VBI.ActiveVBProject.Saved
    End If
End Function

Public Function ProjectIsLoaded() As Boolean
    If VBI Is Nothing Then Exit Function
    ProjectIsLoaded = Not (VBI.ActiveVBProject Is Nothing)
End Function

Public Function GetProjectObject() As VBProject
    If ProjectIsLoaded Then
        GetProjectObject = VBI.ActiveVBProject
    End If
End Function

Public Function SaveSettingProject(plugin As String, Key As String, Data As String) As Boolean
    If ProjectIsLoaded Then
        VBI.ActiveVBProject.WriteProperty "pl_" & plugin, Key, Data
        SaveSettingProject = True
    End If
End Function

Public Function GetSettingProject(plugin As String, Key As String, defdata As String) As String

    On Error GoTo err_ne
    SetEhMode Err_expected
    GetSettingProject = defdata
    If ProjectIsLoaded Then
        GetSettingProject = VBI.ActiveVBProject.ReadProperty("pl_" & plugin, Key)
    End If
err_ne:
    RestoreEhMode
End Function

Public Sub SaveSettingGlobal(plugin As String, Key As String, Data As String)
    SaveSetting APP_NAME, "pl_" & plugin, Key, Data
End Sub

Public Function GetSettingGlobal(plugin As String, Key As String, defdata As String) As String
    GetSettingGlobal = GetSetting(APP_NAME, "pl_" & plugin, Key, defdata)
End Function


Public Sub DeleteSettingGlobal(plugin As String, Key As String)
    
     DeleteSetting APP_NAME, "pl_" & plugin, Key
    
End Sub


Public Function GetActiveProject() As VBProject

    If ProjectIsLoaded Then
        GetActiveProject = VBI.ActiveVBProject
    End If
    
End Function

Public Sub GetSettingsTabClientRect(ByRef X As Long, ByRef Y As Long)

        X = frmPlugIn.curTab(0).width / Screen.TwipsPerPixelX
        Y = frmPlugIn.curTab(0).height / Screen.TwipsPerPixelY
        
End Sub


Public Function GetThunderVBDllPath() As String

    GetThunderVBDllPath = ThunVBPath
    
End Function

Public Function GetThunderVBPluginsPath() As String

    GetThunderVBPluginsPath = ThunVBPath & "plugins\"
    
End Function

Public Function GetCurrentProjectPath() As String

    If ProjectIsLoaded Then
        GetCurrentProjectPath = GetPath(GetActiveProject.FileName)
    End If
    
End Function


'Subclassing Related

'Register a sch...
Public Static Function RegisterSCH(handler As ThunderVB_pl_sch_v1_0) As Long
    Dim t As ID_List, i As Long
    
    For i = 0 To sccol.count - 1
        With sccol.item(i)
            If (.cb_int Is Nothing) = False Then
                Set .cb_int = handler
                .hWndFilter = t
                .MsgFilter = t
                RegisterSCH = i
                Exit Function
            End If
        End With
    Next i
    
    ReDim Preserve sccol.item(sccol.count)
    
    With sccol.item(sccol.count)
         Set .cb_int = handler
        .hWndFilter = t
        .MsgFilter = t
    End With
    
    RegisterSCH = sccol.count
    sccol.count = sccol.count + 1
    
End Function

'Set the list of messanges that the sch recieves..
Public Static Sub SetSCHMsgList(sch As Long, Msglist As ID_List)
    sccol.item(sch).MsgFilter = Msglist
End Sub

'Get the list of messanges that the sch recieves..
Public Static Function GetSCHMsgList(sch As Long) As ID_List
    GetSCHMsgList = sccol.item(sch).MsgFilter
End Function

'Set the list of windows that the sch recieves..
Public Static Sub SetSCHWndList(sch As Long, Wndlist As ID_List)
    sccol.item(sch).hWndFilter = Wndlist
End Sub

'Get the list of windows that the sch recieves..
Public Static Function GetSCHWndList(sch As Long) As ID_List
    GetSCHWndList = sccol.item(sch).hWndFilter
End Function

'Unregister sch
Public Static Function UnRegisterSCH(sch As Long) As Boolean

    Dim t As ID_List
    
    With sccol.item(sch)
         Set .cb_int = Nothing
        .hWndFilter = t
        .MsgFilter = t
    End With
    UnRegisterSCH = True
    
End Function

Public Function GetButton(Id As Long) As clsComBut

    Set GetButton = menbut.items(Id)
    
End Function


Public Sub DebugLog(stri As String)
    LogMsg " DebugLog =" & stri, "DebugLog", "DebugLog"
    frmDebug.AppendLog stri
    
End Sub
