Attribute VB_Name = "modSciHelp"
Option Explicit
'Helper Functions/Const for Scintilla Control
'All functions/Consts are declared here

Public Declare Function CopyCStringA Lib "kernel32" Alias "lstrcpynA" (ByVal lpStringDestination As Long, ByVal lpStringSource As Long, ByVal lngMaxLength As Long) As Long
Public Function sci_Deref_SCI_SCNotification(ptr As Long) As sci_SCNotification
    '<EhHeader>
    On Error GoTo sci_Deref_SCI_SCNotification_Err
    '</EhHeader>
Dim Temp As sci_SCNotification_Un, tempw As String, temp_u_m As sci_SCNotification_Un
Dim tempM As sci_SCNotification, tempMF As sci_SCNotification

    If ptr = 0 Then Exit Function
    CopyMemory Temp, ByVal ptr, Len(Temp)
    tempw = sci_CStringZero(Temp.Text)
    CopyMemory temp_u_m, tempM, Len(tempM)
    Temp.Text = temp_u_m.Text 'change to teh tempM's string pointer.. (is usualy 0 = unallocated but to be sure..)
    CopyMemory tempM, Temp, Len(Temp)
    tempM.Text = tempw
    sci_Deref_SCI_SCNotification = tempM
        
    '<EhFooter>
    Exit Function

sci_Deref_SCI_SCNotification_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.modSciHelp.sci_Deref_SCI_SCNotification " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Note : the sci prefix here , these exist and in the general declaration module
'(without the prefix) They are redeclared here to make the scintilla control independatn
'from the declaration module and the modLowLevel module

'Copies a Zero terminated Cstring to a VB string , cstring of max len of MaxLen
Public Function sci_CStringZero(ByVal lpString As Long, Optional MaxLen As Long = 4096) As String
    '<EhHeader>
    On Error GoTo sci_CStringZero_Err
    '</EhHeader>
Dim s As String, Temp() As Byte, sz As Long, i As Long
ReDim Temp(MaxLen)
    
    If lpString = 0 Then
        sci_CStringZero = vbNullString
        Exit Function
    End If
    
    sz = CopyCStringA(VarPtr(Temp(0)), lpString, MaxLen)
    
    If Temp(0) = 0 Then
        sci_CStringZero = vbNullString
        Exit Function
    End If
    
    For i = 0 To UBound(Temp)
        If Temp(i) = 0 Then ReDim Preserve Temp(i - 1): Exit For
    Next i
    
    sci_CStringZero = StrConv(Temp, vbUnicode)
    
    '<EhFooter>
    Exit Function

sci_CStringZero_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.modSciHelp.sci_CStringZero " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Sub sci_BstrToAnsi(Str As String, ba() As Byte)
    '<EhHeader>
    On Error GoTo sci_BstrToAnsi_Err
    '</EhHeader>

    If Len(Str) = 0 Then
        ReDim ba(0)
    Else
        ba = StrConv(Str, vbFromUnicode)
        ReDim Preserve ba(UBound(ba) + 1)
    End If
    
    '<EhFooter>
    Exit Sub

sci_BstrToAnsi_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.modSciHelp.sci_BstrToAnsi " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'adds "  on the start and end of a string if they do not exist
Public Function Add34(sText As String) As String
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
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.modSciHelp.Add34 " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

