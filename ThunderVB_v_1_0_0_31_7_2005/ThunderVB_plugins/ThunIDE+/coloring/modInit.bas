Attribute VB_Name = "modInit"
Option Explicit

Dim sch_id As Long
Global vbi As VBE

Public Sub on_startup()
          
       Set vbi = GetVBEObject()
       If DebugMode = False Then InitAsmColorHook
          
       LogMsg "Initing CopyTime Coloring ", "modInit", "on_startup"
       InitKeyWords
          
       Dim temp As String
       temp = GetThunderVBPluginsPath() & "ThunIDE\asmdefs.isml"
            
       If Mid$(temp, 2, 2) <> ":\" Then
           temp = "c:\asmdefs.isml"
       End If
          
      LogMsg "Loading ISML file [assembly intellisense database] from " + temp, "modInit", "on_startup"
      dat = LoadIsmlFile(temp)
      LogMsg "Loaded ISML file [assembly intellisense database] from " + temp, "modInit", "on_startup"
      LogMsg "# of ISML keywords :" & dat.kw_count, "modInit", "on_startup"
      LogMsg "# of ISML lists :" & dat.ListCount, "modInit", "on_startup"
          
      temp = GetThunderVBPluginsPath() & "ThunIDE\ThunIDEp.gre"
          
      If Mid$(temp, 2, 2) <> ":\" Then
          temp = "c:\ThunIDEp.gre"
      End If
      LogMsg "Loading Resource File " & temp, "modInit", "on_startup"
       'resfile = Resource_LoadResourceFile(temp, "")
      rfile.OpenFile temp, ""
          
      LogMsg "Registing SubClasser", "modInit", "on_startup"
      sch_id = RegisterSCH(New cls_subclass)
      LogMsg "Registed SubClasser ; id=" & sch_id, "modInit", "on_startup"
          
End Sub


Public Sub on_termination()

       If (DebugMode = False) And App.LogMode = 1 Then KillAsmColorHook
       LogMsg "UnRegisting SubClasser;id=" & sch_id, "modInit", "on_startup"
       UnRegisterSCH sch_id
          
End Sub

