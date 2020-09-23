VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThunderVB_base"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Thunder VB base class
'Class that it is used for Plugin -> Parent comunication..

'*******************
'Subclassing Related
'*******************

'Register a sch...
Public Static Function RegisterSCH(handler As ThunderVB_pl_sch_v1_0) As Long

    RegisterSCH = modPublic.RegisterSCH(handler)

End Function

'Set the list of messanges that the sch recieves..
Public Static Sub SetSCHMsgList(sch As Long, Msglist As ID_List)
    
    modPublic.SetSCHMsgList sch, Msglist
    
End Sub

'Get the list of messanges that the sch recieves..
Public Static Function GetSCHMsgList(sch As Long) As ID_List
    
    GetSCHMsgList = modPublic.GetSCHMsgList(sch)
    
End Function

'Set the list of windows that the sch recieves..
Public Static Sub SetSCHWndList(sch As Long, Wndlist As ID_List)
    
    modPublic.SetSCHWndList sch, Wndlist
    
End Sub

'Get the list of windows that the sch recieves..
Public Static Function GetSCHWndList(sch As Long) As ID_List
    
    GetSCHWndList = modPublic.GetSCHWndList(sch)
    
End Function

'Unregister sch
Public Static Function UnRegisterSCH(sch As Long) As Boolean

    UnRegisterSCH = modPublic.UnRegisterSCH(sch)

End Function


'*******************
'***GUI functions***
'*******************

'Add a item on the context menu
Public Static Function Add2ContextMenu(ipl As ThunderVB_pl_int_v1_0, tip As String, capt As String, defimage As IPictureDisp) As Long
     
     Add2ContextMenu = modPublic.AddButtonToCodeWindowMenu(ipl, tip, capt, defimage)
     
End Function

'Add a item on a menu
Public Static Function AddButton(ipl As ThunderVB_pl_int_v1_0, ToBar As Variant, tip As String, capt As String, defimage As IPictureDisp) As Long
    
    AddButton = modPublic.AddButton(ipl, ToBar, tip, capt, defimage)
    
End Function

'Add a item on the Addins menu
Public Static Function Add2AddinsMenu(ipl As ThunderVB_pl_int_v1_0, sCaption As String, sTip As String, defimage As IPictureDisp) As Long
    
    Add2AddinsMenu = modPublic.AddToAddinMenu(ipl, sCaption, sTip, defimage)
    
End Function

'Add a item on the Toolbox
Public Static Function Add2ToolBar(ipl As ThunderVB_pl_int_v1_0, tip As String, capt As String, defimage As IPictureDisp) As Long
    
    Add2ToolBar = modPublic.AddButtonToToolbar(ipl, tip, capt, defimage)
    
End Function

Public Static Function GetButton(Id As Long) As clsComBut
    
    Set GetButton = modPublic.GetButton(Id)
    
End Function

Public Static Sub RemoveButtton(Id As Long)
    
    modPublic.RemoveButtton Id
    
End Sub

Public Static Sub GetSettingsTabClientRect(ByRef X As Long, ByRef Y As Long)

    modPublic.GetSettingsTabClientRect X, Y
    
End Sub


'*******************
'****Plugin<->VB****
'*******************

Public Static Function GetVBEObject() As VBE

    Set GetVBEObject = VBI
    
End Function

Public Static Function GetPlugInList() As PlugIn_List
    GetPlugInList = plugins
End Function

'***********
'LOG
'**********
Public Static Sub DebugLog(stri As String)
    
    frmDebug.AppendLog stri
    
End Sub

'*******************
'**Ide Interaction**
'*******************

Public Static Function PutStringToCurCursor(str As String) As Boolean

    PutStringToCurCursor = modPublic.PutStringToCurCursor(str)
    
End Function

Public Static Function GetFunctionCode(inMod As String, funct As String) As String

    GetFunctionCode = modPublic.GetFunctionCode(inMod, funct)
    
End Function

Public Static Function EnumModuleNames(Optional moduleType As vbext_ComponentType = -1) As String()
    
    EnumModuleNames = modPublic.EnumModuleNames(moduleType)
    
End Function
    

Public Static Function EnumFunctionNames(inMod As String) As String()

    EnumFunctionNames = modPublic.EnumFunctionNames(inMod)

End Function

Public Static Sub SetCurLine(inMod As String, funct As String, NewTopLine As Long)

    modPublic.SetCurLine inMod, funct, NewTopLine
    
End Sub

Public Static Sub SetFunctionLine(inMod As String, funct As String, LineNum As Long, newLineString As String, Optional bReplace As Boolean = True)
    
    modPublic.SetFunctionLine inMod, funct, LineNum, newLineString, bReplace
    
End Sub

'Retuens true if project is saved , false if it is not saved/not loaded
Public Static Function ProjectSaved() As Boolean
    
    ProjectSaved = modPublic.ProjectSaved()
    
End Function

'Save all projects ...
Public Static Sub SaveProjects(binf As Boolean)

    modPublic.SaveProjects binf
    
End Sub

Public Static Function ProjectIsLoaded() As Boolean
    
    ProjectIsLoaded = modPublic.ProjectIsLoaded
    
End Function

Public Static Function GetProjectObject() As VBProject
    
    Set GetProjectObject = modPublic.GetProjectObject
    
End Function


Public Static Function GetActiveProject() As VBProject

    GetActiveProject = modPublic.GetActiveProject()
    
End Function


'*******************
'*Info On ThunderVB*
'*******************

Public Static Function GetThunVBVer() As String

    GetThunVBVer = modPublic.GetThunVBVer()
    
End Function

Public Static Function GetCodeWinParent() As Long

    GetCodeWinParent = modPublic.GetCodeWinParent()

End Function




Public Static Function GetThunderVBDllPath() As String
    
    GetThunderVBDllPath = modPublic.GetThunderVBDllPath()
    
End Function

Public Static Function GetThunderVBPluginsPath() As String

    GetThunderVBPluginsPath = modPublic.GetThunderVBPluginsPath()
    
End Function

Public Static Function GetCurrentProjectPath() As String
    
    GetCurrentProjectPath = modPublic.GetCurrentProjectPath()
    
End Function

'*******************
'*Create Proc Hooks*
'*******************

Public Static Function AddCPH(inter As ThunderVB_pl_cph_v1_0) As Long
    
    AddCPH = modCPHook.AddCPH(inter)
    
End Function


Public Static Sub RemoveCPH(cph As Long)
    
    modCPHook.RemoveCPH cph
    
End Sub


'*******************
'Save/Load  Settings
'*******************

Public Static Function SaveSettingProject(plugin As String, Key As String, Data As String) As Boolean

    SaveSettingProject = modPublic.SaveSettingProject(plugin, Key, Data)

End Function

Public Static Function GetSettingProject(plugin As String, Key As String, defdata As String) As String

    GetSettingProject = modPublic.GetSettingProject(plugin, Key, defdata)

End Function

Public Static Sub SaveSettingGlobal(plugin As String, Key As String, Data As String)
    
    modPublic.SaveSettingGlobal plugin, Key, Data
    
End Sub

Public Static Function GetSettingGlobal(plugin As String, Key As String, defdata As String) As String
    
    GetSettingGlobal = modPublic.GetSettingGlobal(plugin, Key, defdata)
    
End Function


Public Static Sub DeleteSettingGlobal(plugin As String, Key As String)
    
    modPublic.DeleteSettingGlobal plugin, Key
    
End Sub


Public Static Sub ShowSingleConfigForm(caption As String, width As Long, height As Long, plugin As ThunderVB_pl_int_v1_0)
    
    frmTestWindow.ShowForm caption, width, height, plugin
    
End Sub


