Attribute VB_Name = "modMain"
'Load/Unload/ addin things halding

'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Module created , intial version
'
'
'2/9/2004[dd/mm/yyyy] : Edited by Raziel
'Many changes , more stable code
'
'
'9/2/2005 : Resource file loading , mutlulanguage options
'
'Notes.. This file is edited here and there all the time to
'        Support more things/better loading ect..
'        most of em are not logged cause they are too many/non important

Option Explicit

    
Global VBI As VBIDE.VBE
Global ThunVBPath As String
Global resfile As New cResFile


Sub AddinLoaded()
    

    If DebugMode = False Then
        sc_GetOldAddress MainhWnd
    End If
    
    'Init Create Process hook
    InitCPHook

    LogMsg "Startup Finished ", "modMain", "AddinLoaded"

End Sub

Sub AddinUnload(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    'UnInit Create Process hook
    KillCPHook
    
    'And UnHook All functions that ave been left..
    'While the plugins must unload all of their hooks , we check and unload
    'any remaining.. just to be sure and prevent vb crashes..
    UnHookEverything
    
End Sub

Sub AddinLoad(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    LogMsg "Starting Up", "modMain", "AddinLoad"
    LogMsg "PATH = " & ThunVBPath, "modMain", "AddinLoad"
    
    LogMsg "Loading Resource File " & ThunVBPath & "tvb.gre", "modMain", "AddinLoad"
    'tvb_resfile = Resource_LoadResourceFile(ThunVBPath & "tvb.gre", "")
    resfile.OpenFile ThunVBPath & "tvb.gre", ""
    
End Sub

Public Sub ProjectActivated(VBProject As VBProject)

    LogMsg "Project Activated , sending message to plugins", "modMain", "ProjectActivated"
    
    Dim i As Long
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.OnProjectActivated VBProject
        End If
    Next i

End Sub

Public Sub ProjectAdded(VBProject As VBProject)

    Dim i As Long
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.OnProjectAdded VBProject
        End If
    Next i
    
    If Not (VBI.ActiveVBProject Is Nothing) Then
        ProjectActivated VBI.ActiveVBProject
    End If
    
End Sub

Public Sub ProjectRemoved(VBProject As VBProject)

    Dim i As Long
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.OnProjectRemoved VBProject
        End If
    Next i

End Sub

Public Sub ProjectRenamed(VBProject As VBProject, OldName As String)

    Dim i As Long
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.OnProjectRenamed VBProject, OldName
        End If
    Next i
    
End Sub
