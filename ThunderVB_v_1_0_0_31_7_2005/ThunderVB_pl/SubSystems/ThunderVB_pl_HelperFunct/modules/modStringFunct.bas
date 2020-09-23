Attribute VB_Name = "modStringFunct"
Option Explicit

'adds "  on the start and end of a string if they do not exist
Public Function Add34(ByRef sText As String) As String
    '<EhHeader>
    On Error GoTo Add34_Err
    '</EhHeader>
Dim Temp As String

    Add34 = sText
    If Len(Add34) > 1 Then
        If Mid$(Add34, 1, 1) <> Chr$(34) Then Add34 = Chr$(34) & Add34
        If Mid$(Add34, Len(Add34), 1) <> Chr$(34) Then Add34 = Add34 & Chr$(34)
    End If
    
    '<EhFooter>
    Exit Function

Add34_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "Add34"
    '</EhFooter>
End Function

'removes "  on the start and end of a string if they do exist
Public Function Rem34(ByRef sText As String) As String
    '<EhHeader>
    On Error GoTo Rem34_Err
    '</EhHeader>
Dim Temp As String

    Rem34 = sText
    If Len(Rem34) > 1 Then
        If Mid$(Rem34, 1, 1) = Chr$(34) Then Rem34 = Left$(Rem34, Len(Rem34))
        If Mid$(Rem34, Len(Rem34), 1) = Chr$(34) Then Rem34 = Right$(Rem34, Len(Rem34))
    End If
    
    '<EhFooter>
    Exit Function

Rem34_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "Rem34"
    '</EhFooter>
End Function

'get the string between the find1 and find2 strings (without containing them)
Public Function getS(ByRef find1 As String, ByRef find2 As String, ByRef str As String, Optional ByRef start As Long = 1) As String
    '<EhHeader>
    On Error GoTo getS_Err
    '</EhHeader>
Dim i As Long, i2 As Long

    i = InStr(start, str, find1, vbTextCompare) + Len(find1)
    i2 = InStr(i, str, find2, vbTextCompare)
    If i2 > i Then
        getS = Mid$(str, i, i2 - i)
        start = i2
    Else
        start = 0
    End If
    
    '<EhFooter>
    Exit Function

getS_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "getS"
    '</EhFooter>
End Function

'instr with white space
Public Function InStrWS(ByVal start As Long, ByRef string1 As String, ByRef string2 As String, ByVal cm As VbCompareMethod) As Long
    '<EhHeader>
    On Error GoTo InStrWS_Err
    '</EhHeader>
Dim Temp As Long, i As Long

    Temp = InStr(start, string1, string2, cm)
    If Temp Then
        Do
            If Mid$(string1, Temp + i, 1) = " " Then
                i = i + 1
                If i > Len(string1) Then i = Len(string1): Exit Do
            Else
                Exit Do
            End If
        Loop
    Temp = Temp + i
    End If
    InStrWS = Temp
    
    '<EhFooter>
    Exit Function

InStrWS_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "InStrWS"
    '</EhFooter>
End Function


'Splits all words to array ..
Public Function GetAllWordsToArr(ByRef StrFrom As String) As String()
    '<EhHeader>
    On Error GoTo GetAllWordsToArr_Err
    '</EhHeader>
Dim toS As String_B, Temp As Long, t2 As Long, str As String
    
    If Mid$(str, Len(StrFrom), 1) = " " Then
        t2 = 1
    End If
    
    str = Trim$(StrFrom)
    Do
        Temp = InStr(1, str, " ")
        
        If Temp > 0 Then
            AppendString toS, Left$(str, Temp - 1)
            str = Trim$(Right$(str, Len(str) - Temp))
        Else
            AppendString toS, str
            str = ""
        End If
    Loop While Temp > 0
    
    If t2 = 1 Then
        AppendString toS, " "
        'frmDConsole.AppendLog str & "|"
    End If
    
    FinaliseString toS
    
    GetAllWordsToArr = toS.str
    
    '<EhFooter>
    Exit Function

GetAllWordsToArr_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "GetAllWordsToArr"
    '</EhFooter>
End Function

'removes the first word from a string
Public Sub RemFisrtWord(ByRef str As String) 'remove fisrt word
    '<EhHeader>
    On Error GoTo RemFisrtWord_Err
    '</EhHeader>
Dim Temp As Long

    Temp = InStr(1, str, " ")
    If Temp Then
        str = Right$(str, Len(str) - Temp)
    Else
        str = ""
    End If
    
    '<EhFooter>
    Exit Sub

RemFisrtWord_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "RemFisrtWord"
    '</EhFooter>
End Sub



'Changes all strings spaces in the s string to "_"
'eg string "this is 'a simple' "string cont"aining a string" is changed to
'          "this is 'a_simple' "string_cont"aining a string"
Function ProcStrings(ByRef s As String) As String
    '<EhHeader>
    On Error GoTo ProcStrings_Err
    '</EhHeader>
Dim Temp As String, st As Long, stold As Long
    
    st = 1
    stold = 0
    Do While st > stold
        stold = st
        Temp = Add34(getS(Chr$(34), Chr$(34), s, st))
        
        If st > stold And st > 0 Then
            s = Replace$(s, Temp, Replace$(Replace$(Replace$(Temp, " ", "_"), ";", "_"), "'", "_"))
            st = st + 1
        End If
    Loop

    st = 1
    stold = 0
    Do While st > stold
        stold = st
        Temp = "'" & (getS("'", "'", s, st)) & "'"
        
        If st > stold And st > 0 Then
            s = Replace$(s, Temp, Replace$(Replace$(Temp, " ", "_"), ";", "_"))
            st = st + 1
        End If
    Loop
    
    ProcStrings = s
    
    '<EhFooter>
    Exit Function

ProcStrings_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "ProcStrings"
    '</EhFooter>
End Function

'Replaces any queted string with "______" ...
Function ProcStringsUnderAll(ByRef s As String) As String
    '<EhHeader>
    On Error GoTo ProcStringsUnderAll_Err
    '</EhHeader>
Dim Temp As String, st As Long, stold As Long
    
    st = 1
    stold = 0
    Do While st > stold
        stold = st
        Temp = Add34(getS(Chr$(34), Chr$(34), s, st))
        
        If st > stold And st > 0 Then
            s = Replace$(s, Temp, String$(Len(Temp), "_"))
            st = st + 1
        End If
    Loop
    
    ProcStringsUnderAll = s
    
    '<EhFooter>
    Exit Function

ProcStringsUnderAll_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "ProcStringsUnderAll"
    '</EhFooter>
End Function


'make string that contains vbCrLf characters
'parameter - lCount - number of vbCrLfs
'return - string - string of vbCrLfs

Public Function CrLf(Optional ByVal lCount As Long = 1) As String
    '<EhHeader>
    On Error GoTo CrLf_Err
    '</EhHeader>
Dim i As Long
    For i = 1 To lCount
        CrLf = CrLf & vbCrLf
    Next i
    '<EhFooter>
    Exit Function

CrLf_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "CrLf"
    '</EhFooter>
End Function



'get the fisrt word of a string , with the space folowing it
Public Function GetFirstWordWithSpace(str As String) As String
    '<EhHeader>
    On Error GoTo GetFirstWordWithSpace_Err
    '</EhHeader>
Dim Temp As Long
    
    Temp = InStrWS(1, str, " ", vbBinaryCompare)
    If Temp Then
        GetFirstWordWithSpace = Left$(str, Temp - 1)
    Else
        GetFirstWordWithSpace = str
    End If
    
    '<EhFooter>
    Exit Function

GetFirstWordWithSpace_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "GetFirstWordWithSpace"
    '</EhFooter>
End Function

'removes the first word from a string and the space after it
Public Sub RemFirstWordWithSpace(ByRef str As String) 'remove fisrt word
    '<EhHeader>
    On Error GoTo RemFirstWordWithSpace_Err
    '</EhHeader>
Dim Temp As Long

    Temp = InStrWS(1, str, " ", vbBinaryCompare)
    If Temp Then
        str = Right$(str, Len(str) - Temp + 1)
    Else
        str = ""
    End If
    
    '<EhFooter>
    Exit Sub

RemFirstWordWithSpace_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "RemFirstWordWithSpace"
    '</EhFooter>
End Sub

'get the fisrt word of a string
Public Function GetFirstWord(ByRef str As String) As String
    '<EhHeader>
    On Error GoTo GetFirstWord_Err
    '</EhHeader>
Dim Temp As Long

    Temp = InStr(1, str, " ")
    If Temp Then
        GetFirstWord = Left$(str, Temp - 1)
    Else
        GetFirstWord = str
    End If
    
    '<EhFooter>
    Exit Function

GetFirstWord_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "modStringFunct", "GetFirstWord"
    '</EhFooter>
End Function


