Attribute VB_Name = "modfrmPluginHand"
Option Explicit

Public isVisible As Boolean


Public Sub frmPlugin_Show(Optional modal, Optional OwnerForm)
    
    If isVisible = False Then
        isVisible = True
        Form_Load_event
    End If
    
    frmPlugIn.Show modal, OwnerForm
    
End Sub

Public Sub frmPlugin_Hide()
    
    If isVisible = True Then
        isVisible = False
        Form_Unload_event
    End If
    
    frmPlugIn.Hide
    
End Sub

Public Sub Form_Load_event()

    'Dim t As Resource_File
    'Dim dl As Language_entry
    'dl.language = tvb_English
    
    't = Resource_NewFile("stef", "ThunVBTest", "ThunVBPLREs", dl)
    'SaveFormToResourceFile t, Me, "ThunderVB_pl", tvb_English, , tvb_res_Stored
    'SaveResourceFile "c:\tvb.gre", t, "", 1
    
    Dim i As Long
    frmPlugIn.oldpar = -1
    frmPlugIn.DoNotSendApp = False
    'For i = curTab.LBound To curTab.UBound
    '    curTab(i).BorderStyle = 0
    '    curTab(i).Visible = True
    'Next i
    
    
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            'MsgBox plugins.plugins(i).interface.GetName & " load"
            Call plugins.plugins(i).interface.OnGuiLoad
        End If
    Next i
    
    frmPlugIn.RefreshLists
    
    
End Sub

Public Sub Form_Unload_event()
Dim i As Long

    On Error Resume Next
    i = frmPlugIn.List1.ItemData(frmPlugIn.List1.ListIndex)
    If plugins.plugins(i).Loaded = True Then
        plugins.plugins(i).interface.HideConfig
        plugins.plugins(i).interface.HideCredits
    End If
        
    If frmPlugIn.DoNotSendApp = False Then
        For i = 0 To plugins.count - 1
            If IsPluginLoaded(i) Then
                plugins.plugins(i).interface.ApplySettings
            End If
        Next i
    End If
    
    
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.OnGuiUnLoad
        End If
    Next i
    
End Sub

