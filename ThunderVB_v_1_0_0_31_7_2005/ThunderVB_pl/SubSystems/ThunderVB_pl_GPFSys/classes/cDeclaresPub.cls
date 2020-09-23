VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeclaresPub"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Static Function StartGPFHandler(ByRef MinidumpPath As String) As Boolean

    StartGPFHandler = modGPFException.StartGPFHandler(MinidumpPath)

End Function

' release the new handler
Public Static Sub StopGPFHandler()
    
    modGPFException.StopGPFHandler
    
End Sub


Public Static Sub GPF_Reset()
    
    Call modGPFException.GPF_Reset
    
End Sub

Public Static Sub GPF_Set(ByVal nAct As GPF_actions, ByRef fromProj As String, ByRef fromMod As String, ByRef fromProc As String)

    Call modGPFException.GPF_Set(nAct, fromProj, fromMod, fromProc)

End Sub

Public Static Sub GPFSys_SetLogger(LoggerObj As ILogger)
    
    Call SetLogger(LoggerObj)
    
End Sub

Public Static Sub GpfSys_SendErrorReport(ByVal Exception As Long, ByRef info As String)
   
   Dim stra As String, strb As String
   
   stra = CInt(Timer * 10) & "_" & "dump.dmp"
   kill2 gpf_MinidumpPath & stra
   WriteMinidump gpf_MinidumpPath & stra, Exception, MiniDumpNormal, ""
   
   strb = CInt(Timer * 10) & "_" & "report.gre"
   
   SaveReportToFile strb, stra, info
   
   frmErrorRep.ShowForm strb
   
End Sub

Public Static Sub GpfSys_SendSugestion()
   
   Shell "explorer mailto:stef.mp@gmail.com", vbMaximizedFocus
   
End Sub

