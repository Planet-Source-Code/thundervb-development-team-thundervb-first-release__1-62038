Attribute VB_Name = "modResFileIO_v0_9"
Option Explicit
'file format version 0.9 [ResSys v1 beta 2]
'Suport for this file format is removed as of 11/5/2005[dd/mm/yyyy]
'due to possible out of memory errors that happen if a file is not a gre file

Private Const tvb_resform_version As Long = 1

Public Function Resource_LoadResourceFile_v0_9(file As String, key As String, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Resource_File
    Dim Temp As Resource_File, i As Long
    Dim ff As Long
    
    If FileExist(file) Then
        ff = FreeFile
        Open file For Binary As ff
        Temp.header = Resource_ReadHeader(ff)
    Else
        Close ff
        Err.Raise ThunVB_Errors.tvb_File_Does_Not_Exist, "LoadResourceFile", "File " & file & " was not found..."
        Exit Function
    End If
    Get ff, , Temp.numEntrys
    If Temp.numEntrys > 0 Then
        ReDim Temp.headers(Temp.numEntrys - 1)
        For i = 0 To Temp.numEntrys - 1
            Temp.headers(i) = Resource_LoadEntry(ff, key, 1)
        Next i
    Else
        Close ff
        Err.Raise ThunVB_Errors.tvb_Res_NoDataEntrys, "LoadResourceFile", "Resource_File.numEntrys" _
                  & " is invalid [" & Temp.numEntrys & "]"
        Exit Function
    End If
    
    Close ff
    Temp.header.fftag = fftag
    Temp.header.Version = modResFileIO.tvb_resform_version
    
    Resource_LoadResourceFile_v0_9 = Temp
    
End Function

Private Function Resource_ReadHeader(ff As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As tvb_res_main_header

    Dim Temp As tvb_res_main_header, i As Long
    Dim sz As Long, tmpsNull As String, tmplNull As Long
    With Temp
        'Read all the static data..
        Get ff, , .Version
        Get ff, , sz: .Name = Space$(sz)
        Get ff, , .Name
        
        Get ff, , sz: .author = Space$(sz)
        Get ff, , .author
        
        Get ff, , sz: .CreatedWith = Space$(sz)
        Get ff, , .CreatedWith
        
        Get ff, , sz: .Description = Space$(sz)
        Get ff, , .Description
        
        Get ff, , tmplNull
        'Read and check the number of suported languages..
        If tmplNull > 0 Then
            'Load them ;)
            For i = 0 To tmplNull - 1
                Call Resource_ReadLanguageEntry(ff)
            Next i
        End If
    End With
    
    'Ok, done :):):)
    'Check if header is valid :)
    If Resource_MainHeaderIsSupported(Temp) Then
        'yay , we can return it :)
        Resource_ReadHeader = Temp
    Else
        'Too bad .. all this was for nothing :(
        Err.Raise ThunVB_Errors.tvb_Res_Header_Corupted, "Resource_ReadHeader", "Header is not supported/Corupted"
        Exit Function
    End If
    

End Function

Private Function Resource_LoadEntry(ff As Long, key As String, DictionarySize As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As tvb_res_entry

    Dim Temp As tvb_res_entry
    With Temp
        .header = Resource_Read_tvb_res_entry_header(ff)
        If Resource_EntryHeaderIsValid(.header) Then
            
            Get ff, , .Length
            If .Length > 0 Then
                ReDim .Data(.Length - 1)
                Get ff, , .Data
                Resource_DecodeData Temp, key, DictionarySize 'Decompress it ect..
           End If
        Else
            Err.Raise ThunVB_Errors.tvb_Res_Header_Corupted, "Resource_LoadEntry", "Error on tvb_res_entry.header"
            Exit Function
        End If
    End With
    Resource_LoadEntry = Temp

End Function

Private Function Resource_MainHeaderIsSupported(header As tvb_res_main_header) As Boolean

    If header.Version <> tvb_resform_version Then
        Resource_MainHeaderIsSupported = False
        Exit Function
    End If
    
    Resource_MainHeaderIsSupported = True
    

End Function

Private Function Resource_EntryHeaderIsValid(header As tvb_res_entry_header, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Boolean

    'Yeah , toooo much checking :D .. this is for future use..
    With header
        
    End With
    
    Resource_EntryHeaderIsValid = True
    

End Function

'Decodes /decompresses the data
Private Sub Resource_DecodeData(Data As tvb_res_entry, key As String, DictionarySize As Long)


    If Data.header.PackMode And tvb_res_Encrypted Then
        Dim Temp As New clsSimpleXOR
        Temp.DecryptByte Data.Data, key
    End If
    
    If Data.header.PackMode And tvb_res_Compressed Then
        DeCompress_Arr Data.Data
    End If
    Data.Length = ArrUBound(Data.Data) + 1
    
End Sub


Private Sub Resource_ReadLanguageEntry(ff As Long)

Dim sz As Long
Dim tmplNull As Long, tmpsNull As String

    
    Get ff, , tmplNull
    Get ff, , tmplNull: tmpsNull = Space$(tmplNull)
    Get ff, , tmpsNull

    
End Sub


Private Function Resource_Read_tvb_res_entry_header(ff As Long) As tvb_res_entry_header

    Dim sz As Long, tmpsNull As String
    With Resource_Read_tvb_res_entry_header
        Get ff, , sz: .Id = Space$(sz)
        Get ff, , .Id
        Get ff, , .DataType
        Get ff, , .language
        Get ff, , sz: tmpsNull = Space$(sz)
        Get ff, , tmpsNull
        Get ff, , .PackMode
        Get ff, , .PackInfo
    End With
    
End Function
