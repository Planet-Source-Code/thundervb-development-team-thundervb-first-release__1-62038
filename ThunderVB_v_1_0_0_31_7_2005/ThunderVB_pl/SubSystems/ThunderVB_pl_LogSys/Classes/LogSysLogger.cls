VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogSysLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ILogger

Dim outFile As String

Private Function ILogger_GetLogFile() As String
    
    ILogger_GetLogFile = outFile
    
End Function

Public Sub ILogger_InitLogSystem(ByRef File As String, Optional ByVal bAppendData As Boolean = True)
    
    outFile = File
    Call modLogSystem.InitLogSystem(File, bAppendData)
    LogMsg "Log SubSystem Version " & app.Major & "." & app.Minor & "." & app.Revision & " was inited", APP_NAME, "LogSysLogger.cls", "ILogger_InitLogSystem"
    
End Sub

Public Sub ILogger_LogMsg(str As String, ProjectName As String, codeModule As String, codePosition As String, Optional bLogTime As Boolean = True)
 
    Call modLogSystem.LogMsg(str, ProjectName, codeModule, codePosition, bLogTime)
   
End Sub


Public Function ILogger_WarnBox(str As String, codeModule As String, codePosition As String, ProjectName As String, Optional style As VbMsgBoxStyle = vbOKOnly Or vbExclamation) As VbMsgBoxResult
    
    ILogger_WarnBox = modLogSystem.WarnBox(str, codeModule, codePosition, style, ProjectName)
    
End Function

Public Function ILogger_ErrorBox(ByRef str As String, ByRef codeModule As String, ByRef codePosition As String, ByRef ProjectName As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly Or vbCritical) As VbMsgBoxResult
    
    ILogger_ErrorBox = modLogSystem.ErrorBox(str, codeModule, codePosition, style, ProjectName)
    
End Function

Public Sub ILogger_CloseLog()
 
    Call modLogSystem.CloseLog
   
End Sub

Public Function ILogger_MsgBoxX(ByRef str As String, Optional ByRef caption As String = APP_NAME, Optional ByVal style As VbMsgBoxStyle = vbOKOnly Or vbInformation) As VbMsgBoxResult
    
    ILogger_MsgBoxX = modLogSystem.MsgBoxX(str, caption, style)
    
End Function

