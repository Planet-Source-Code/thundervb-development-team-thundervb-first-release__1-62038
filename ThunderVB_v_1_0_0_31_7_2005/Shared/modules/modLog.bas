Attribute VB_Name = "modLog"
Option Explicit

Public Logger As ILogger

Public Sub SetLogger(new_Logger As ILogger)
    
    Set Logger = new_Logger

End Sub

Public Sub ResetLogger()
    
    Set Logger = Nothing
    
End Sub

Public Sub LogMsg(str As String, codeModule As String, codePosition As String, Optional bLogTime As Boolean = True)
    
    If Not (Logger Is Nothing) Then
    
        Call Logger.LogMsg(str, APP_NAME, codeModule, codePosition, bLogTime)
    
    End If
    
End Sub

Public Function WarnBox(str As String, codeModule As String, codePosition As String, Optional style As VbMsgBoxStyle = vbOKOnly Or vbExclamation) As VbMsgBoxResult
    
    If Not (Logger Is Nothing) Then
    
        WarnBox = Logger.WarnBox(str, codeModule, codePosition, APP_NAME, style)
     
    End If
        
End Function

Public Function ErrorBox(ByRef str As String, ByRef codeModule As String, ByRef codePosition As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly Or vbCritical) As VbMsgBoxResult
    
    If Not (Logger Is Nothing) Then
    
        ErrorBox = Logger.ErrorBox(str, codeModule, codePosition, APP_NAME, style)
    
    End If

End Function

Public Function MsgBoxX(ByRef str As String, Optional ByRef caption As String = APP_NAME, Optional ByVal style As VbMsgBoxStyle = vbOKOnly Or vbInformation) As VbMsgBoxResult
    
    If Not (Logger Is Nothing) Then
    
        MsgBoxX = Logger.MsgBoxX(str, caption, style)
    
    End If

End Function

