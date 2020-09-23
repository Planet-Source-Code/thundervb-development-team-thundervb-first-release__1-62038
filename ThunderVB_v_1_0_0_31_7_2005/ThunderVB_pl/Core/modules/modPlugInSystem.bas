Attribute VB_Name = "modPlugInSystem"
Option Explicit
'Plugin Collection and Managment

Public plugins As PlugIn_List
Public Config As ThunderVB_pl_int_v1_0

Public Function pl_Speed_req2String(t As pl_Speed_Req) As String
    Select Case t
        Case pl_Speed_Req.idle
            pl_Speed_req2String = resfile.GetTextByIndex(lng_modPlSys.speedreq.idle)
        Case pl_Speed_Req.low
            pl_Speed_req2String = resfile.GetTextByIndex(lng_modPlSys.speedreq.low)
        Case pl_Speed_Req.Med
            pl_Speed_req2String = resfile.GetTextByIndex(lng_modPlSys.speedreq.medium)
        Case pl_Speed_Req.Hight
            pl_Speed_req2String = resfile.GetTextByIndex(lng_modPlSys.speedreq.high)
        Case pl_Speed_Req.Cazy
            pl_Speed_req2String = resfile.GetTextByIndex(lng_modPlSys.speedreq.veryhigh)
    End Select
End Function

Public Function pl_type2String(t As pl_type) As String
    
    If t And ThunVB_DLLHook Then
        pl_type2String = resfile.GetTextByIndex(lng_modPlSys.pltype.DLLHook) ' "Dll Hook"
    End If
    
    If t And ThunVB_CPHook Then
        pl_type2String = pl_type2String & IIf(Len(pl_type2String), ";", "") & resfile.GetTextByIndex(lng_modPlSys.pltype.CPHook) '"CPTool"
    End If
    
    If t And ThunVB_plugin_CodeTool Then
        pl_type2String = pl_type2String & IIf(Len(pl_type2String), ";", "") & resfile.GetTextByIndex(lng_modPlSys.pltype.CodeTool) '"CodeTool"
    End If
    
    If t And ThunVB_plugin_MiscTool Then
        pl_type2String = pl_type2String & IIf(Len(pl_type2String), ";", "") & resfile.GetTextByIndex(lng_modPlSys.pltype.MiscTool) '"MiscTool"
    End If
    
    If t And ThunVB_SC_Heavy Then
        pl_type2String = pl_type2String & IIf(Len(pl_type2String), ";", "") & resfile.GetTextByIndex(lng_modPlSys.pltype.SC_Heavy)  ' "Much SubClassing"
    ElseIf t And ThunVB_SC_Light Then
        pl_type2String = pl_type2String & IIf(Len(pl_type2String), ";", "") & resfile.GetTextByIndex(lng_modPlSys.pltype.SC_Light)  ' "A bit SubClassing"
    ElseIf t And ThunVB_SC_Med Then
        pl_type2String = pl_type2String & IIf(Len(pl_type2String), ";", "") & resfile.GetTextByIndex(lng_modPlSys.pltype.CPHook) ' "SubClassing"
    End If

End Function


'Loads all the plugins on this folder
Public Sub LoadPlugins(folder As String)
Dim Temp As String, temp2 As Long
    SetLoadScreenStatus lsS_LoadingPlugins, folder
    LogMsg "Loading plugins from " & Add34(folder), "modPlugInSystem", "LoadPlugins"
    Temp = Dir$(folder & "*.dll", vbNormal)
    
    Do While Len(Temp)
        temp2 = LoadPlugin(GetFilename(Temp), CBool(GetSetting(APP_NAME, "plugins", GetFilename(Temp) & "_loaded", "true")))
        If temp2 < 0 Then GoTo NextOne
        If CBool(GetSetting(APP_NAME, "plugins", plugins.plugins(temp2).dllfile & "_loaded", "true")) = False Then
            LogMsg "Found but not loaded """ + plugins.plugins(temp2).Name + """", "modPlugInSystem", "LoadPlugins"
            GoTo NextOne
        End If
        If temp2 > 0 Then
            LogMsg "Found and loaded """ + plugins.plugins(temp2).Name + """", "modPlugInSystem", "LoadPlugins"
        End If
NextOne:
        Temp = Dir$()
    Loop

End Sub

Public Sub init()

    modLanguages.lang_UpdateIds tvb_English
    plugins.count = -1
    LoadPlugin "ThunderVB_pl_v1_0" 'App.ProductName 'load the internal plugin

End Sub

Public Sub terminate()

    UnLoadPlugin 0
    
End Sub

'Unloads all the loaded plugins - EXEPT the internal one..
Public Sub UnLoadPlugins()
Dim i As Long
    
    For i = 1 To plugins.count - 1
        UnLoadPlugin i
    Next i

End Sub

'loads a plugin
Public Function LoadPlugin(dll As String, Optional bCreateInst As Boolean = True) As Long
Dim i As Long, i2 As Long
    SetLoadScreenStatus lsS_LoadingPlugin, Split(GetFilename(dll), ".")(0), dll
    i2 = -1
    For i = 0 To plugins.count - 1
        If plugins.plugins(i).used = False Then
            i2 = i
            Exit For
        End If
    Next i
    
    If i2 = -1 Then
        ReDim Preserve plugins.plugins(plugins.count + 1)
        plugins.count = ArrUBound(plugins.plugins) + 1
        i2 = plugins.count - 1
    End If
    
    'use i2 index to load plugin
    Dim dll_s As String
    dll_s = Split(GetFilename(dll), ".")(0) 'get the file (remove the .dll) , the file must not have ANY other dots
    Dim Temp As ThunderVB_pl_int_v1_0
    'A_LoadLibrary dll
    Err.Clear
    On Error Resume Next
    Set Temp = CreateObject(dll_s & ".plugin")
    Dim xxb As Object

    On Error GoTo 0
    
    If Temp Is Nothing Then
        LoadPlugin = -1
        Exit Function
    End If
    
    SetexeptH
    
    On Error GoTo unloadit:
    
    With plugins.plugins(i2)
        .used = True 'we use this slot
        .Loaded = True 'we load the addin
        Temp.InitSetLang Resource_GetCurLang()
        .Desciption = Temp.GetDesciption
        .DesciptionFull = Temp.GetDesciptionFull
        .Id = Temp.GetID
        Set .interface = Temp
        .Name = Temp.GetName
        Temp.SetLogger Logger
        .Speed = Temp.GetSpeed
        .type = Temp.GetType
        .Version = Temp.GetVersion
        .VersionNum = Temp.GetVersionNum
        If bCreateInst Then
            .interface.OnStartUp           'Load messange
        Else
            .Loaded = False
        End If
        .dllfile = dll
    End With
    'well , we loaded ;)
    LoadPlugin = i2
    
    Exit Function
unloadit:
    LoadPlugin = -1
    plugins.plugins(i2).used = False  'we use this slot
    plugins.plugins(i2).Loaded = False  'we load the addin
    
End Function

Public Sub UnLoadPlugin(index As Long)
        
        With plugins.plugins(index)
            SetLoadScreenStatus lsS_UnLoadingPlugin, .Name, .dllfile

            If IsPluginLoaded(index) Then
                .interface.OnTermination 'Unload message
            End If

            .Loaded = False
            If Not (.interface Is Nothing) Then
                Set .interface = Nothing
            End If
            'FreeLibrary GetModuleHandle(.dllfile)
            LogMsg "Unloaded " & .Name, "modPlugInSystem", "UnLoadPlugin"
        End With
        
End Sub

Public Sub lang_UpdatePluginsInfo()
Dim i As Long

    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            With plugins.plugins(i)
                .Desciption = .interface.GetDesciption
                .DesciptionFull = .interface.GetDesciptionFull
                .Version = .interface.GetVersion
            End With
        End If
    Next i
    
End Sub

Public Function IsPluginLoaded(ByVal index As Long) As Boolean

    If index >= 0 And index < plugins.count Then
        With plugins.plugins(index)
            If (.Loaded = True) And (.used = True) And Not (.interface Is Nothing) Then
                    IsPluginLoaded = True
                Else
                    IsPluginLoaded = False
            End If
        End With
    End If
    
End Function
