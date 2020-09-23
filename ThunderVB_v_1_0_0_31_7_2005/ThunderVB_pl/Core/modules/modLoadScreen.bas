Attribute VB_Name = "modLoadScreen"
Option Explicit

Public Enum lsStatus
    lsS_Hiden
    lsS_Visible
    lsS_StartingUP
    lsS_Initing
    lsS_LoadingPlugins
    lsS_LoadingPlugin
    lsS_UnLoadingPlugin
    lsS_Finished
    lsS_Unloading
End Enum

Dim lsst As lsStatus
Dim lsfrm As New frmLoadscreen

Public Sub SetLoadScreenStatus(ByRef lss As lsStatus, ParamArray text() As Variant)

    Select Case lss
        Case lsStatus.lsS_Hiden
            lsfrm.Hide
        Case lsStatus.lsS_Visible
            lsfrm.Show
            lsfrm.lblstatus.caption = ""
        Case lsStatus.lsS_Initing
            lsfrm.lblstatus.caption = "Initing " & text(0)
        Case lsStatus.lsS_LoadingPlugin
            lsfrm.lblstatus.caption = "Loading " & text(0) & " (in " & text(1) & ")"
        Case lsStatus.lsS_LoadingPlugin
            lsfrm.lblstatus.caption = "Unloading " & text(0) & " (in " & text(1) & ")"
        Case lsStatus.lsS_LoadingPlugins
            lsfrm.lblstatus.caption = "Loading plugins(from" & text(0) & ")"
        Case lsStatus.lsS_StartingUP
            lsfrm.lblstatus.caption = "Starting Up"
        Case lsStatus.lsS_Finished
            lsfrm.lblstatus.caption = "Finished"
        Case lsStatus.lsS_Unloading
            lsfrm.lblstatus.caption = "Unloading.."
    End Select
    
    If lsfrm.Visible = True Then
        lsfrm.Show
        DoEvents
    End If
    
End Sub



