VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Static Sub InitLogSystem(ByRef File As String, Optional ByVal bAppendData As Boolean = True)

End Sub

Public Static Sub LogMsg(str As String, ProjectName As String, codeModule As String, codePosition As String, Optional bLogTime As Boolean = True)

End Sub

Public Static Function WarnBox(str As String, codeModule As String, codePosition As String, ProjectName As String, Optional style As VbMsgBoxStyle = vbOKOnly Or vbExclamation) As VbMsgBoxResult

End Function

Public Static Function ErrorBox(ByRef str As String, ByRef codeModule As String, ByRef codePosition As String, ByRef ProjectName As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly Or vbCritical) As VbMsgBoxResult

End Function

Public Static Sub CloseLog()

End Sub

Public Static Function MsgBoxX(ByRef str As String, Optional ByRef caption As String = APP_NAME, Optional ByVal style As VbMsgBoxStyle = vbOKOnly Or vbInformation) As VbMsgBoxResult

End Function


Public Static Function GetLogFile() As String

End Function
