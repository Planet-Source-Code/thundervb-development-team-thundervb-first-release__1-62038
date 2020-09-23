Attribute VB_Name = "modFileVersion"

Option Explicit

'this code isn't ours, it was written by J. Rongen
'and could be found here
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11550&lngWId=1

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
   "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
   dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
   "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias _
   "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
   lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
   "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As _
   Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long
    
Public Function GetFileVersion(ByVal PathWithFilename As String) As String
Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(7) As String
Dim strTemp As String
Dim intTemp As Integer
       
    ' size
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen > 0 Then
       ReDim bytBuffer(lngBufferlen)
       lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
       If lngRc <> 0 Then
          lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
                   lngVerPointer, lngBufferlen)
          If lngRc <> 0 Then
             'lngVerPointer is a pointer to four 4 bytes of Hex number,
             'first two bytes are language id, and last two bytes are code
             'page. However, strLangCharset needs a  string of
             '4 hex digits, the first two characters correspond to the
             'language id and last two the last two character correspond
             'to the code page id.
             MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
             lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                    bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
             strLangCharset = Hex(lngHexNumber)
             'now we change the order of the language id and code page
             'and convert it into a string representation.
             'For example, it may look like 040904E4
             'Or to pull it all apart:
             '04------        = SUBLANG_ENGLISH_USA
             '--09----        = LANG_ENGLISH
             ' ----04E4 = 1252 = Codepage for Windows:Multilingual
             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop
             ' assign propertienames
             strVersionInfo(0) = "CompanyName"
             strVersionInfo(1) = "FileDescription"
             strVersionInfo(2) = "FileVersion"
             strVersionInfo(3) = "InternalName"
             strVersionInfo(4) = "LegalCopyright"
             strVersionInfo(5) = "OriginalFileName"
             strVersionInfo(6) = "ProductName"
             strVersionInfo(7) = "ProductVersion"
             ' loop and get fileproperties
             For intTemp = 0 To 7
                strBuffer = String$(255, 0)
                strTemp = "\StringFileInfo\" & strLangCharset _
                   & "\" & strVersionInfo(intTemp)
                lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                      lngVerPointer, lngBufferlen)
                If lngRc <> 0 Then
                   ' get and format data
                   lstrcpy strBuffer, lngVerPointer
                   strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                   strVersionInfo(intTemp) = strBuffer
                 Else
                   ' property not found
                   strVersionInfo(intTemp) = "?"
                End If
             Next intTemp
          End If
       End If
    End If

    ' assign array to user-defined-type
'    FileInfo.CompanyName = strVersionInfo(0)
'    FileInfo.FileDescription = strVersionInfo(1)
'    FileInfo.FileVersion = strVersionInfo(2)
'    FileInfo.InternalName = strVersionInfo(3)
'    FileInfo.LegalCopyright = strVersionInfo(4)
'    FileInfo.OrigionalFileName = strVersionInfo(5)
'    FileInfo.ProductName = strVersionInfo(6)
'    FileInfo.ProductVersion = strVersionInfo(7)

    GetFileVersion = strVersionInfo(7)

End Function



