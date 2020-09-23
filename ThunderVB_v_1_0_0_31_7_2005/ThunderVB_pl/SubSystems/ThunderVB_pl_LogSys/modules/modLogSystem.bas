Attribute VB_Name = "modLogSystem"
Option Explicit

Global log_file As file_b

'Added by Raziel , 25/8/2004
'Used to display messages to the user
'According to the pluing settings (Log msg's , hide them ect)

Public Function WarnBox(str As String, codeModule As String, codePosition As String, Optional style As VbMsgBoxStyle = vbOKOnly Or vbExclamation, Optional app As String = APP_NAME) As VbMsgBoxResult
    '<EhHeader>
    On Error GoTo WarnBox_Err
    '</EhHeader>
    
    'If Get_General(HideErrorDialogs) = False Then
        WarnBox = MsgBox("Warning : " & vbNewLine & str, style, codeModule & ":" & codePosition)
    'End If
    
    LogMsg "(From " & codeModule & ":" & codePosition & ") " & str, app, "modPublic", "WarnBox"
    
    '<EhFooter>
    Exit Function

WarnBox_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "ThunderVB_pl", "modPublic", "WarnBox"
    '</EhFooter>
End Function

Public Function ErrorBox(str As String, codeModule As String, codePosition As String, Optional style As VbMsgBoxStyle = vbOKOnly Or vbCritical, Optional app As String = APP_NAME) As VbMsgBoxResult
    '<EhHeader>
    On Error GoTo ErrorBox_Err
    '</EhHeader>
    
    'If Get_General(HideErrorDialogs) = False Then
        ErrorBox = MsgBox("Error : " & vbNewLine & str, style, codeModule & ":" & codePosition)
    'End If
    
    LogMsg "(From " & codeModule & ":" & codePosition & ") " & str, app, "modPublic", "ErrorBox"
    
    '<EhFooter>
    Exit Function

ErrorBox_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "ThunderVB_pl", "modPublic", "ErrorBox"
    '</EhFooter>
End Function

Public Function MsgBoxX(str As String, Optional caption As String = APP_NAME, Optional style As VbMsgBoxStyle = vbOKOnly Or vbInformation) As VbMsgBoxResult
    '<EhHeader>
    On Error GoTo MsgBoxX_Err
    '</EhHeader>

    MsgBoxX = MsgBox(str, style, caption)
    
    '<EhFooter>
    Exit Function

MsgBoxX_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "ThunderVB_pl", "modPublic", "MsgBoxX"
    '</EhFooter>
End Function


Public Sub InitLogSystem(ByRef File As String, Optional ByVal bAppendData As Boolean = True)
    log_file = OpenFile(File, 512)
    
    If bAppendData Then
        FileSeek log_file, FileLength(log_file) + 1
    Else
        kill2 File
    End If
    
    AppendToFile log_file, "*******************************************************" & vbNewLine
End Sub

'Log format : [Time : ]\[codeproject::codemodule:codePosition\] str
Public Static Sub LogMsg(str As String, ProjectName As String, codeModule As String, codePosition As String, Optional bLogTime As Boolean = True)
    '<EhHeader>
    On Error GoTo LogMsg_Err
    '</EhHeader>
    Dim temp As String
    
    'If Get_Debug(EnableOutPutToDebugLog) = False Then Exit Sub
    
    'Get_Paths(Debug_Directory)
    
    If bLogTime Then
        AppendToFile log_file, Date$ & " " & Time & " : "
    End If
     
    AppendToFile log_file, "[" & ProjectName & "::" & codeModule & ":" & codePosition & "]" & vbTab & str & vbNewLine
    FlushFile log_file
    
    '<EhFooter>
    Exit Sub

LogMsg_Err:
    LogMsg "Error : " & Err.Description & " , At " & Add34(Err.Source) & ":" & Erl, "ThunderVB_pl", "modPublic", "LogMsg"
    '</EhFooter>
End Sub

Public Sub CloseLog()
    
    If log_file.maxbuflen > 0 Then
        CloseFile log_file
    End If
    
End Sub

