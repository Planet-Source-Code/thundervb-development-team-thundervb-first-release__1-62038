Attribute VB_Name = "modErrReport"

'You need reference to Microsoft WinHTTP Services 5.0 or 5.1 to use this example

'For uploading code :
'Credit to Joseph Z. Xu (jzxu@napercom.com)
'Modified by Mohd Idzuan Alias (iklan2k@yahoo.com) August 18, 2004
'Editided by drkIIRaziel

Dim WinHttpReq As WinHttp.WinHttpRequest
Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0
Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
Const BOUNDARY = "Xu02=$"
Const HEADER = "--Xu02=$"
Const FOOTER = "--Xu02=$--"


'Report an error , upload the file ot the server
Public Sub ReportError_create(ByRef userinfo As String, target As String, ParamArray strFiles() As Variant)
Dim fl As cResFile, i As Long
    Set fl = New cResFile
    fl.Create "ThunderVB - GPF report System", "Crash report", "Crash report"
    
    For i = 0 To UBound(strFiles)
        fl.AddEntry Resource_NewDataEntry(LoadFile_bin(CStr(strFiles(i))), i & "_" & Timer & "-" & GetFilename(CStr(strFiles(i))), tvb_res_Compressed, tvb_None)
    Next i
    fl.AddEntry Resource_NewTextEntry(userinfo, "userinfo", tvb_res_Compressed, tvb_None)
    
    fl.SaveFileAs target, ""

End Sub

Private Function getFile(strFileName As String) As String

    Dim strFile As String
    
    ' Grap the file
    nFile = FreeFile
    Open strFileName For Binary As #nFile
    strFile = String(LOF(nFile), " ")
    Get #nFile, , strFile
    Close #nFile
    
    getFile = strFile
    
End Function

Public Sub SaveReportToFile(strFile As String, strDumpFile As String, strUserInfo As String)
    
    If Not (Logger Is Nothing) Then
        ReportError_create strUserInfo, strFile, strDumpFile, Logger.GetLogFile
    Else
        ReportError_create strUserInfo, strFile, strDumpFile
    End If
    
End Sub
