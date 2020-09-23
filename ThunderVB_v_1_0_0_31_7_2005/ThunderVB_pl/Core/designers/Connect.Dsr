VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   12315
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14130
   _ExtentX        =   24924
   _ExtentY        =   21722
   _Version        =   393216
   Description     =   "ThunderVB add-in v1.0.0 beta 10 pre2"
   DisplayName     =   "ThunderVB v1.0.0 beta 10 pre2"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ThunderVB v1.0.0"
Option Explicit


Public WithEvents prjevents As VBProjects
Attribute prjevents.VB_VarHelpID = -1
Public WithEvents BuildEvents As VBIDE.VBBuildEvents
Attribute BuildEvents.VB_VarHelpID = -1

Private Function GetDateTime() As String
    
    GetDateTime = Date$ + "  " + Time$
    GetDateTime = Replace$(GetDateTime, "/", "-")
    GetDateTime = Replace$(GetDateTime, "\", "-")
    GetDateTime = Replace$(GetDateTime, ":", "-")
    
End Function
'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
Dim sbuff As String * 260

    Resource_SetCurLang tvb_English
    
    SetLoadScreenStatus lsS_Visible
    SetLoadScreenStatus lsS_StartingUP
    
    SetLogger New LogSysLogger
    
    If DebugMode Then
        ThunVBPath = "C:\develop\vb\vb-projects\ThunderVB_plugined" & "\bin\"
    Else
        GetModuleFileName GetModuleHandle(App.EXEName), sbuff, 260
        ThunVBPath = GetPath(sbuff)
    End If
    
    SaveSetting APP_NAME, "path", "dllroot", ThunVBPath
    
    Logger.InitLogSystem ThunVBPath & "logs\thundervb_log_" & GetDateTime() & ".txt"
    HookSys_SetLogger Logger
    GPFSys_SetLogger Logger
    HelperFunct_SetLogger Logger
    prjVBErrRepl.SetLogger Logger
    
    MainhWnd = 0
    ReDim buts(0)
    'save the vb instance
    Set VBI = Application
    'Init the path
    If DebugMode = False Then
        
        StartGPFHandler ThunVBPath & "logs\"
        
    End If
    
    SetLoadScreenStatus lsS_Initing, APP_NAME
    'Get the main whnd [for subclassing]
    MainhWnd = GetCodeWinParent()
    'Init the Internal Add-in and load the other plugins ..
    init
    LoadPlugins ThunVBPath & "plugins\"

    Resource_SetCurLang tvb_None
    frmPlugIn.cmbLang.SetLanguage GetSettingGlobal(APP_NAMEs, "Language", tvb_Languages.tvb_English)

    'Start Sub class timer
    sc_ApiTimer True
        
    Call AddinLoad(Application, ConnectMode, AddInInst, custom())
        
    frmPlugIn.cmbLang_LanguageChanged GetSettingGlobal(APP_NAMEs, "Language", tvb_Languages.tvb_English)
        
    If ConnectMode = ext_cm_AfterStartup Then
        AddinInstance_OnStartupComplete custom
    End If
    
    SetLoadScreenStatus lsS_Finished
    Sleep 400
    SetLoadScreenStatus lsS_Hiden
   
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
        
    'Send terminate siglnal and unload All plugins ..
    
        
    SetLoadScreenStatus lsS_Visible
    SetLoadScreenStatus lsS_Unloading
    
    'unload all plugins...
    UnLoadPlugins
    'stop the subclassing timer
    sc_ApiTimer False

    terminate
    
    'Free All subclasses windows..
    sc_KillAllSubClasses
    
    'Remove all Buttons/Menus ect.. that may have leaked from unloading
    Dim i As Long
    For i = 0 To menbut.count - 1
        Set menbut.items(i) = Nothing
    Next i

    'Call the funtion in modMain..
    Call AddinUnload(RemoveMode, custom())
        'if not in ide
    If DebugMode = False Then
        StopGPFHandler
    End If
    
    SetLoadScreenStatus lsS_Finished
    Sleep 400
    SetLoadScreenStatus lsS_Hiden
    Logger.CloseLog
End Sub


Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

    Set prjevents = VBI.VBProjects
    'this is a nice class ., but how to get a haldle for it?
    'Set BuildEvents = Nothing
    
    Call AddinLoaded
    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveVBProject Is Nothing Then Exit Sub
    'MainhWnd = GetCodeWinPar()
    modMain.ProjectAdded VBI.ActiveVBProject
    modMain.ProjectActivated VBI.ActiveVBProject
    
End Sub

Private Sub BuildEvents_BeginCompile(ByVal VBProject As VBIDE.VBProject)

    'BuildStarted
    
End Sub

Private Sub BuildEvents_EnterDesignMode()
    
    'DesignStarted
    
End Sub

Private Sub BuildEvents_EnterRunMode()
    
    'RunStarted
    
End Sub

Private Sub prjevents_ItemActivated(ByVal VBProject As VBIDE.VBProject)

    ProjectActivated VBProject
    
End Sub

Private Sub prjevents_ItemAdded(ByVal VBProject As VBIDE.VBProject)

    ProjectAdded VBProject
    
End Sub

Private Sub prjevents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)

    ProjectRemoved VBProject
    
End Sub

Private Sub prjevents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)

    ProjectRenamed VBProject, OldName
    
End Sub


