Attribute VB_Name = "modStringBuilder"
Option Explicit

'Made by Raziel(19/8/2004[dd/mm/yyyy])
'A Simple but yet effective String Builder
'Giving a speed boost on big strings creation...
'Also , this code is used to create colored strings
'use it as you wish , gime  a credit

'Moved declares to declares_pub

'For normal stringz
Sub AppendString(ByRef toString As String_B, ByRef Data As String)
    '<EhHeader>
    On Error GoTo AppendString_Err
    '</EhHeader>

    With toString
        If .str_index >= .str_bound Then
            If .str_bound = 0 Then .str_bound = 1
            ReDim Preserve .str(.str_bound * 2)
            .str_bound = ArrUBound(.str)
        End If
        .str(.str_index) = Data
        .str_index = .str_index + 1
    End With
    
    '<EhFooter>
    Exit Sub

AppendString_Err:
    '</EhFooter>
End Sub

Sub FinaliseString(ByRef toString As String_B)
    '<EhHeader>
    On Error GoTo FinaliseString_Err
    '</EhHeader>

    With toString
        ReDim Preserve .str(.str_index - 1)
        .str_bound = ArrUBound(.str)
    End With
    
    '<EhFooter>
    Exit Sub

FinaliseString_Err:
    '</EhFooter>
End Sub

Function GetString(ByRef fromString As String_B) As String
    '<EhHeader>
    On Error GoTo GetString_Err
    '</EhHeader>

    With fromString
        ReDim Preserve .str(.str_index - 1)
        GetString = Join$(.str, "")
        .str_bound = ArrUBound(.str)
    End With
    
    '<EhFooter>
    Exit Function

GetString_Err:
    '</EhFooter>
End Function


'for color code...
Sub AppendColString(ByRef toString As Col_String, ByVal Data As Long, ByVal col As Long)
    '<EhHeader>
    On Error GoTo AppendColString_Err
    '</EhHeader>
Dim Temp As Col_String_entry

    With toString
        If .str_index >= .str_bound Then
            If .str_bound = 0 Then .str_bound = 1
            ReDim Preserve .str(.str_bound * 2)
            .str_bound = ArrUBound(.str)
        End If
        Temp.strlen = Data
        Temp.col = col
        .str(.str_index) = Temp
        .str_index = .str_index + 1
    End With
    
    '<EhFooter>
    Exit Sub

AppendColString_Err:
    '</EhFooter>
End Sub

Sub FinaliseColString(ByRef toString As Col_String)
    '<EhHeader>
    On Error GoTo FinaliseColString_Err
    '</EhHeader>

    With toString
        ReDim Preserve .str(.str_index - 1)
    End With
    
    '<EhFooter>
    Exit Sub

FinaliseColString_Err:
    '</EhFooter>
End Sub
