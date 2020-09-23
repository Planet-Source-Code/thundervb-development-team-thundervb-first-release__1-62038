Attribute VB_Name = "modArithComp"
Option Explicit


'hehehe , using zlib....
Private Declare Function compress Lib "zlibwapi_tvb.dll" (ByRef dest As Byte, ByRefdestLen As Long, _
        ByRef source As Byte, ByVal sourceLen As Long) As Long
Private Declare Function uncompress Lib "zlibwapi_tvb.dll" (ByRef dest As Byte, ByRef destLen As Long, _
        ByRef source As Byte, ByVal sourceLen As Long) As Long


Public Sub Compress_Arr(ByteArray() As Byte)
Dim Temp() As Byte, sz As Long
    ReDim Temp(ArrUBound(ByteArray) * 2)
    
    sz = ArrUBound(Temp) - 4
    Call compress(Temp(4), sz, ByteArray(0), ArrUBound(ByteArray))
    ReDim Preserve Temp(4 + sz)
    sz = ArrUBound(ByteArray)
    
    CopyMemory Temp(0), sz, 4
    ByteArray = Temp
    
End Sub

Public Sub DeCompress_Arr(ByteArray() As Byte)
    Dim Temp() As Byte, sz As Long
    
    CopyMemory sz, ByteArray(0), 4
    
    ReDim Temp(sz)
      
    Call uncompress(Temp(0), sz, ByteArray(4), ArrUBound(ByteArray) - 4)
    If (sz <> ArrUBound(Temp)) Then MsgBox "Error on resource decompresion"
    ByteArray = Temp

End Sub
