Attribute VB_Name = "modLanguages"
Option Explicit

Public Type lng_frmpl_info_type
    Name As Long
    Version As Long
    Speed As Long
    type As Long
    WorksOnWin As Long
    ShortDesc As Long
    FullDesc As Long
End Type


Public Type lng_frmpl_MainTab_type
    TabSettings As Long
    TabCredits As Long
    TabAbout As Long
End Type


Public Type lng_frmpl_type
    Info As lng_frmpl_info_type
    MainTab As lng_frmpl_MainTab_type
End Type

Public Type lng_modPlSys_speedreq_type
    idle As Long
    low As Long
    medium As Long
    high As Long
    veryhigh As Long
End Type

Public Type lng_modPlSys_pltype_type
    DLLHook  As Long
    CPHook As Long
    CodeTool  As Long
    MiscTool  As Long
    SC_Heavy  As Long
    SC_Light  As Long
    SC_Med  As Long
End Type

Public Type lng_modPlSys_type
    speedreq As lng_modPlSys_speedreq_type
    pltype As lng_modPlSys_pltype_type
End Type

Global lng_frmpl As lng_frmpl_type, lng_modPlSys As lng_modPlSys_type

'update all resourxe indexes..
Public Sub lang_UpdateIds(lang As tvb_Languages)

    UpdateIds_lng_frmpl lang
    UpdateIds_lng_modPlSys lang
    
End Sub

'Update all the gui for a new language;called after UpdateIds.
Public Sub lang_UpdateGui(lang As tvb_Languages)
        
        frmPlugIn.lang_UpdateGui lang
        frmInternal.lang_UpdateGui lang
        
End Sub

Private Sub UpdateIds_lng_frmpl(lang As tvb_Languages)
    
    UpdateIds_lng_frmpl_info lang
    
End Sub

Private Sub UpdateIds_lng_frmpl_info(lang As tvb_Languages)

    lng_frmpl.Info.FullDesc = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-info-fulldesc", lang)
    lng_frmpl.Info.ShortDesc = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-info-shortdesc", lang)
    lng_frmpl.Info.Name = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-info-name", lang)
    lng_frmpl.Info.Speed = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-info-speed", lang)
    lng_frmpl.Info.type = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-info-type", lang)
    lng_frmpl.Info.Version = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-info-version", lang)
    lng_frmpl.Info.WorksOnWin = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-info-worksonwin", lang)
    
    lng_frmpl.MainTab.TabSettings = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-MainTab-TabSettings", lang)
    lng_frmpl.MainTab.TabAbout = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-MainTab-TabAbout", lang)
    lng_frmpl.MainTab.TabCredits = resfile.ResourceExists("ThunderVB_pl-frmPlugIn-MainTab-TabCredits", lang)
    
End Sub

Private Sub UpdateIds_lng_modPlSys(lang As tvb_Languages)

    UpdateIds_lng_modPlSys_speedreq lang
    UpdateIds_lng_modPlSys_pltype lang
    
End Sub

Private Sub UpdateIds_lng_modPlSys_speedreq(lang As tvb_Languages)

    With lng_modPlSys.speedreq
        .high = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-speedreq-high", lang)
        .idle = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-speedreq-idle", lang)
        .low = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-speedreq-low", lang)
        .medium = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-speedreq-medium", lang)
        .veryhigh = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-speedreq-veryhigh", lang)
    End With

End Sub

Private Sub UpdateIds_lng_modPlSys_pltype(lang As tvb_Languages)

    With lng_modPlSys.pltype
        .CPHook = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-pltype-CPHook", lang)
        .DLLHook = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-pltype-DLLHook", lang)
        .CodeTool = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-pltype-plugin_CodeTool", lang)
        .MiscTool = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-pltype-plugin_MiscTool", lang)
        .SC_Heavy = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-pltype-SC_Heavy", lang)
        .SC_Light = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-pltype-SC_Light", lang)
        .SC_Med = resfile.ResourceExists("ThunderVB_pl-modPlugInSystem-pltype-SC_Med", lang)
    End With

End Sub
