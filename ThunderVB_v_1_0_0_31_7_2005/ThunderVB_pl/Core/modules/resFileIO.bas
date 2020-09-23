Attribute VB_Name = "modResFileIO"
Option Explicit

Public Const tvb_resform_version As Long = 1


Public Function LoadImage( _
   ImageBytes As tvb_res_Data) As StdPicture
    '<EhHeader>
    On Error GoTo LoadImage_Err
    '</EhHeader>
Dim oPersist As IPersistStream
Dim oStream As IStream
Dim lSize As Long
  If ImageBytes.Length < 1 Then Exit Function
   ' Calculate the array size
   lSize = UBound(ImageBytes.Data) - LBound(ImageBytes.Data) + 1
   If lSize = 1 Then Exit Function
   ' Create a stream object
   ' in global memory
   Set oStream = CreateStreamOnHGlobal(0, True)
   
   ' Write the header to the stream
   oStream.Write &H746C&, 4&
   
   ' Write the array size
   oStream.Write lSize, 4&
   
   ' Write the image data
   oStream.Write ImageBytes.Data(LBound(ImageBytes.Data)), lSize
   
   ' Move the stream position to
   ' the start of the stream
   oStream.Seek 0, STREAM_SEEK_SET
      
   ' Create a new empty picture object
   Set LoadImage = New StdPicture
   
   ' Get the IPersistStream interface
   ' of the picture object
   Set oPersist = LoadImage
   
   ' Load the picture from the stream
   oPersist.Load oStream
    
   ' Release the streamobject
   Set oStream = Nothing
   
    '<EhFooter>
    Exit Function

LoadImage_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "LoadImage"
    '</EhFooter>
End Function

Public Function SaveImage( _
   ByVal image As StdPicture) As tvb_res_Data
    '<EhHeader>
    On Error GoTo SaveImage_Err
    '</EhHeader>
Dim abData() As Byte
Dim oPersist As IPersistStream
Dim oStream As IStream
Dim lSize As Long
Dim tStat As STATSTG
   If image Is Nothing Then
       SaveImage.Length = 0
       Exit Function
   End If
   ' Get the image IPersistStream interface
   Set oPersist = image
   
   ' Create a stream on global memory
   Set oStream = CreateStreamOnHGlobal(0, True)
   
   ' Save the picture in the stream
   oPersist.Save oStream, True
      
   ' Get the stream info
   oStream.Stat tStat, STATFLAG_NONAME
      
   ' Get the stream size
   lSize = tStat.cbSize * 10000 - 8
   
   If lSize = 0 Then
       SaveImage.Length = 0
       Exit Function
   Else
   ' Initialize the array
   ReDim abData(0 To lSize - 1)
   End If
   ' Move the stream position to
   ' the start of the stream
   oStream.Seek 0.0008, STREAM_SEEK_SET
   
   ' Read all the stream in the array
   oStream.Read abData(0), lSize
   
   ' Return the array
   SaveImage.Data = abData
   SaveImage.Length = UBound(abData) + 1
   
   ' Release the stream object
   Set oStream = Nothing

    '<EhFooter>
    Exit Function

SaveImage_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "SaveImage"
    '</EhFooter>
End Function


Public Function Resource_LoadResourceFile(file As String, Key As String) As Resource_File
    '<EhHeader>
    On Error GoTo LoadResourceFile_Err
    '</EhHeader>
    Dim temp As Resource_File, i As Long
    Dim ff As Long
    
    If FileExist(file) Then
        ff = FreeFile
        Open file For Binary As ff
        temp.header = Resource_ReadHeader(ff)
    Else
        Close ff
        err.Raise ThunVB_Errors.tvb_File_Does_Not_Exist, "LoadResourceFile", "File " & file & " was not found..."
        Exit Function
    End If
    Get ff, , temp.numEntrys
    If temp.numEntrys > 0 Then
        ReDim temp.headers(temp.numEntrys - 1)
        For i = 0 To temp.numEntrys - 1
            temp.headers(i) = Resource_LoadEntry(ff, Key, 1)
        Next i
    Else
        Close ff
        err.Raise ThunVB_Errors.tvb_Res_NoDataEntrys, "LoadResourceFile", "Resource_File.numEntrys" _
                  & " is invalid [" & temp.numEntrys & "]"
        Exit Function
    End If
    
    Close ff
    Resource_LoadResourceFile = temp
    
    '<EhFooter>
    Exit Function

LoadResourceFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "LoadResourceFile"
    '</EhFooter>
End Function

Public Sub Resource_SaveResourceFile(file As String, Data As Resource_File, Key As String)
    '<EhHeader>
    On Error GoTo SaveResourceFile_Err
    '</EhHeader>
    Dim i As Long
    Dim ff As Long
    

    ff = FreeFile
    kill2 file
    Open file For Binary As ff
    Resource_SaveHeader ff, Data.header

    Put ff, , Data.numEntrys
    If Data.numEntrys > 0 Then
        For i = 0 To Data.numEntrys - 1
            Resource_SaveEntry ff, Key, Data.headers(i), 1
        Next i
    Else
        Close ff
        err.Raise ThunVB_Errors.tvb_Res_NoDataEntrys, "LoadResourceFile", "Resource_File.numEntrys" _
                  & " is invalid [" & Data.numEntrys & "]"
        Exit Sub
    End If
    Close ff
    
    '<EhFooter>
    Exit Sub

SaveResourceFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "SaveResourceFile"
    '</EhFooter>
End Sub

Public Function Resource_ReadHeader(ff As Long) As tvb_res_main_header
    '<EhHeader>
    On Error GoTo Resource_ReadHeader_Err
    '</EhHeader>
    Dim temp As tvb_res_main_header, i As Long
    Dim sz As Long
    With temp
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
        
        Get ff, , .LanguagesCount
        'Read and check the number of suported languages..
        If .LanguagesCount > 0 Then
            ReDim .Languages(.LanguagesCount - 1)
        Else
            err.Raise ThunVB_Errors.tvb_Res_No_Languages, "Resource_ReadHeader", "tvb_res_main_header.LanguagesCount" _
                      & " is invalid [" & .LanguagesCount & "]"
            Exit Function
        End If
        'Load them ;)
        For i = 0 To .LanguagesCount - 1
            .Languages(i) = Resource_ReadLanguageEntry(ff)
        Next i
    End With
    
    'Ok, done :):):)
    'Check if header is valid :)
    If Resource_MainHeaderIsSupported(temp) Then
        'yay , we can return it :)
        Resource_ReadHeader = temp
    Else
        'Too bad .. all this was for nothing :(
        err.Raise ThunVB_Errors.tvb_Res_Header_Corupted, "Resource_ReadHeader", "Header is not supported/Corupted"
        Exit Function
    End If
    
    '<EhFooter>
    Exit Function

Resource_ReadHeader_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_ReadHeader"
    '</EhFooter>
End Function

Public Sub Resource_SaveHeader(ff As Long, header As tvb_res_main_header)
    '<EhHeader>
    On Error GoTo Resource_SaveHeader_Err
    '</EhHeader>
    Dim i As Long
    
    With header
        'Save all the static data..
        Put ff, , .Version
        Put ff, , Len(.Name)
        Put ff, , .Name
        
        Put ff, , Len(.author)
        Put ff, , .author
        
        Put ff, , Len(.CreatedWith)
        Put ff, , .CreatedWith
        
        Put ff, , Len(.Description)
        Put ff, , .Description
        
        Put ff, , .LanguagesCount
        
        'check the number of suported languages..
        If .LanguagesCount <= 0 Then
            err.Raise ThunVB_Errors.tvb_Res_No_Languages, "Resource_SaveHeader", "header.LanguagesCount" _
                      & " is invalid [" & .LanguagesCount & "]"
            Exit Sub
        End If
        'Save them ;)
        For i = 0 To .LanguagesCount - 1
            Resource_SaveLanguageEntry ff, .Languages(i)
        Next i
        'Ok, done :):):)
    End With
    
    '<EhFooter>
    Exit Sub

Resource_SaveHeader_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_SaveHeader"
    '</EhFooter>
End Sub

Public Function Resource_LoadEntry(ff As Long, Key As String, DictionarySize As Long) As tvb_res_entry
    '<EhHeader>
    On Error GoTo Resource_LoadEntry_Err
    '</EhHeader>
    Dim temp As tvb_res_entry
    With temp
        .header = Resource_Read_tvb_res_entry_header(ff)
        If Resource_EntryHeaderIsValid(.header) Then
            
            Get ff, , .Length
            If .Length > 0 Then
                ReDim .Data(.Length - 1)
                Get ff, , .Data
                Resource_DecodeData temp, Key, DictionarySize 'Decompress it ect..
           End If
        Else
            err.Raise ThunVB_Errors.tvb_Res_Header_Corupted, "Resource_LoadEntry", "Error on tvb_res_entry.header"
            Exit Function
        End If
    End With
    Resource_LoadEntry = temp
    '<EhFooter>
    Exit Function

Resource_LoadEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_LoadEntry"
    '</EhFooter>
End Function

Public Sub Resource_SaveEntry(ff As Long, Key As String, Data As tvb_res_entry, DictionarySize As Long)
    '<EhHeader>
    On Error GoTo Resource_SaveEntry_Err
    '</EhHeader>
    Dim temp As tvb_res_entry
    temp = Data
    With temp
        If Resource_EntryHeaderIsValid(.header) Then
        
            Resource_Save_tvb_res_entry_header ff, .header
            Resource_EncodeData temp, Key, DictionarySize
            Put ff, , .Length
            Put ff, , .Data
            
        Else
            err.Raise ThunVB_Errors.tvb_Res_Header_Corupted, "Resource_LoadEntry", "Error on tvb_res_entry.header"
            Exit Sub
        End If
    End With

    '<EhFooter>
    Exit Sub

Resource_SaveEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_SaveEntry"
    '</EhFooter>
End Sub

Public Function Resource_MainHeaderIsSupported(header As tvb_res_main_header) As Boolean
    '<EhHeader>
    On Error GoTo Resource_MainHeaderIsSupported_Err
    '</EhHeader>
    
    If header.Version <> tvb_resform_version Then
        Resource_MainHeaderIsSupported = False
        Exit Function
    End If
    
    Resource_MainHeaderIsSupported = True
    
    '<EhFooter>
    Exit Function

Resource_MainHeaderIsSupported_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_MainHeaderIsSupported"
    '</EhFooter>
End Function

Public Function Resource_EntryHeaderIsValid(header As tvb_res_entry_header) As Boolean
    '<EhHeader>
    On Error GoTo Resource_EntryHeaderIsValid_Err
    '</EhHeader>
    
    'Yeah , toooo much checking :D .. this is for future use..
    With header
        
    End With
    
    Resource_EntryHeaderIsValid = True
    
    '<EhFooter>
    Exit Function

Resource_EntryHeaderIsValid_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_EntryHeaderIsValid"
    '</EhFooter>
End Function

'Decodes /decompresses the data
Public Sub Resource_DecodeData(Data As tvb_res_entry, Key As String, DictionarySize As Long)
    '<EhHeader>
    On Error GoTo Resource_DecodeData_Err
    '</EhHeader>

    If Data.header.PackMode And tvb_res_Encrypted Then
        Dim temp As New clsSimpleXOR
        temp.DecryptByte Data.Data, Key
    End If
    
    If Data.header.PackMode And tvb_res_Compressed Then
        DeCompress_LZSSLazy Data.Data, DictionarySize
    End If
    Data.Length = UBound(Data.Data) + 1
    
    '<EhFooter>
    Exit Sub

Resource_DecodeData_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_DecodeData"
    '</EhFooter>
End Sub

'encodes /compresses the data
Public Sub Resource_EncodeData(Data As tvb_res_entry, Key As String, dictSiz As Long)
    '<EhHeader>
    On Error GoTo Resource_EncodeData_Err
    '</EhHeader>

    If Data.header.PackMode And tvb_res_Compressed Then
        Compress_LZSSLazy Data.Data, dictSiz
    End If
    If Data.header.PackMode And tvb_res_Encrypted Then
        Dim temp As New clsSimpleXOR
        temp.EncryptByte Data.Data, Key
    End If
    Data.Length = UBound(Data.Data) + 1
    
    '<EhFooter>
    Exit Sub

Resource_EncodeData_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_EncodeData"
    '</EhFooter>
End Sub

Public Function Resource_ReadLanguageEntry(ff As Long) As Language_entry
    '<EhHeader>
    On Error GoTo Resource_ReadLanguageEntry_Err
    '</EhHeader>
Dim sz As Long

    With Resource_ReadLanguageEntry
        Get ff, , .language
        Get ff, , sz: .LanguageString = Space$(sz)
        Get ff, , .LanguageString
    End With
    
    '<EhFooter>
    Exit Function

Resource_ReadLanguageEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_ReadLanguageEntry"
    '</EhFooter>
End Function

Public Sub Resource_SaveLanguageEntry(ff As Long, Data As Language_entry)
    '<EhHeader>
    On Error GoTo Resource_SaveLanguageEntry_Err
    '</EhHeader>
    
    With Data
        Put ff, , .language
        Put ff, , Len(.LanguageString)
        Put ff, , .LanguageString
    End With
    
    '<EhFooter>
    Exit Sub

Resource_SaveLanguageEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_SaveLanguageEntry"
    '</EhFooter>
End Sub

Public Function Resource_Read_tvb_res_entry_header(ff As Long) As tvb_res_entry_header
    '<EhHeader>
    On Error GoTo Resource_Read_tvb_res_entry_header_Err
    '</EhHeader>
    Dim sz As Long
    With Resource_Read_tvb_res_entry_header
        Get ff, , sz: .Id = Space$(sz)
        Get ff, , .Id
        Get ff, , .DataType
        Get ff, , .language
        Get ff, , sz: .LanguageString = Space$(sz)
        Get ff, , .LanguageString
        Get ff, , .PackMode
        Get ff, , .PackInfo
    End With
    
    '<EhFooter>
    Exit Function

Resource_Read_tvb_res_entry_header_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_Read_tvb_res_entry_header"
    '</EhFooter>
End Function

Public Sub Resource_Save_tvb_res_entry_header(ff As Long, Data As tvb_res_entry_header)
    '<EhHeader>
    On Error GoTo Resource_Save_tvb_res_entry_header_Err
    '</EhHeader>
    
    With Data
        Put ff, , Len(.Id)
        Put ff, , .Id
        Put ff, , .DataType
        Put ff, , .language
        Put ff, , Len(.LanguageString)
        Put ff, , .LanguageString
        Put ff, , .PackMode
        Put ff, , .PackInfo
    End With
    
    '<EhFooter>
    Exit Sub

Resource_Save_tvb_res_entry_header_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_Save_tvb_res_entry_header"
    '</EhFooter>
End Sub

Public Sub Resource_AddEntry(toFile As Resource_File, NewEntry As tvb_res_entry)
    '<EhHeader>
    On Error GoTo Resource_AddEntry_Err
    '</EhHeader>
    
    If NewEntry.Length < 1 Then Exit Sub
    
    With toFile
        ReDim Preserve .headers(.numEntrys)
        .headers(.numEntrys) = NewEntry
        .numEntrys = .numEntrys + 1
    End With
    
    '<EhFooter>
    Exit Sub

Resource_AddEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_AddEntry"
    '</EhFooter>
End Sub

Public Sub Resource_RemoveEntry(toFile As Resource_File, Id As Long)
    '<EhHeader>
    On Error GoTo Resource_RemoveEntry_Err
    '</EhHeader>
    Dim i As Long
    With toFile
        For i = Id To .numEntrys - 2
            .headers(i) = .headers(i + 1)
        Next i
        .numEntrys = .numEntrys - 1
        ReDim Preserve .headers(.numEntrys)
    End With
    '<EhFooter>
    Exit Sub

Resource_RemoveEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_RemoveEntry"
    '</EhFooter>
End Sub

Public Function Resource_NewTextEntry(TextData As String, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages, Optional LanguageString As String = "") As tvb_res_entry
    '<EhHeader>
    On Error GoTo Resource_NewTextEntry_Err
    '</EhHeader>

    If Len(TextData) = 0 Then Exit Function
    With Resource_NewTextEntry
        .header.Id = Id
        .header.DataType = tvb_res_Text
        .header.language = language
        .header.LanguageString = LanguageString
        .header.PackMode = PackMode
        .Data = TextData
        .Length = UBound(.Data) + 1
    End With
    
    '<EhFooter>
    Exit Function

Resource_NewTextEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_NewTextEntry"
    '</EhFooter>
End Function

Public Function Resource_NewDataEntry(Data() As Byte, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages, Optional LanguageString As String = "") As tvb_res_entry
    '<EhHeader>
    On Error GoTo Resource_NewDataEntry_Err
    '</EhHeader>

    With Resource_NewDataEntry
        .header.Id = Id
        .header.DataType = tvb_res_Data
        .header.language = language
        .header.LanguageString = LanguageString
        .header.PackMode = PackMode
        .Data = Data
        .Length = UBound(Data) + 1
    End With
    
    '<EhFooter>
    Exit Function

Resource_NewDataEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_NewDataEntry"
    '</EhFooter>
End Function

Public Function Resource_NewImageEntry(image As StdPicture, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages, Optional LanguageString As String = "") As tvb_res_entry
    '<EhHeader>
    On Error GoTo Resource_NewImageEntry_Err
    '</EhHeader>
Dim temp As tvb_res_Data

    With Resource_NewImageEntry
        .header.Id = Id
        .header.DataType = tvb_res_Image
        .header.language = language
        .header.LanguageString = LanguageString
        .header.PackMode = PackMode
        temp = SaveImage(image)
        .Data = temp.Data
        .Length = temp.Length
    End With
    
    '<EhFooter>
    Exit Function

Resource_NewImageEntry_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_NewImageEntry"
    '</EhFooter>
End Function

Public Function Resource_Exists(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional LanguageString As String = "") As Long
    '<EhHeader>
    On Error GoTo Resource_Exists_Err
    '</EhHeader>
Dim i As Long, temp As Language_entry

    If lang = 0 Then
        temp = Cur_Language
    Else
        temp.language = lang
        temp.LanguageString = LanguageString
    End If
    
    With file
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language Then
                    Resource_Exists = i
                    Exit Function
                End If
            End If
        Next i
        
        'Try with the default language
        If .header.LanguagesCount > 0 Then
            temp = .header.Languages(0)
        Else ' if not default lang specified , default to english
            temp.language = tvb_English
            temp.LanguageString = ""
        End If
        
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language Then
                    Resource_Exists = i
                    Exit Function
                End If
            End If
        Next i
    End With
    
    Resource_Exists = -1
    
    '<EhFooter>
    Exit Function

Resource_Exists_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_Exists"
    '</EhFooter>
End Function

Public Function Resource_GetText(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional LanguageString As String = "") As String
    '<EhHeader>
    On Error GoTo Resource_GetText_Err
    '</EhHeader>
Dim i As Long, temp As Language_entry
    If lang = 0 Then
        temp = Cur_Language
    Else
        temp.language = lang
        temp.LanguageString = LanguageString
    End If
    With file
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language Then
                    Resource_GetText = .headers(i).Data
                    Exit Function
                End If
            End If
        Next i
        
        'Try with the default language
        If .header.LanguagesCount > 0 Then
            temp = .header.Languages(0)
        Else ' if not default lang specified , default to english
            temp.language = tvb_English
            temp.LanguageString = ""
        End If
        
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language Then
                    Resource_GetText = .headers(i).Data
                    Exit Function
                End If
            End If
        Next i
    End With
    
    err.Raise ThunVB_Errors.tvb_Res_NoDataEntrys, "ThunderVB_pl::modResFileIO:Resource_GetText", "The specifyed entry " & Id & "was not found..."
    
    '<EhFooter>
    Exit Function

Resource_GetText_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_GetText"
    '</EhFooter>
End Function


Public Function Resource_GetTextByIndex(file As Resource_File, index As Long) As String
    '<EhHeader>
    On Error GoTo Resource_GetTextByIndex_Err
    '</EhHeader>
Dim i As Long, temp As Language_entry
    
    Resource_GetTextByIndex = file.headers(index).Data
    
    '<EhFooter>
    Exit Function

Resource_GetTextByIndex_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_GetTextByIndex"
    '</EhFooter>
End Function

Public Function Resource_GetImage(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional LanguageString As String = "") As StdPicture
    '<EhHeader>
    On Error GoTo Resource_GetImage_Err
    '</EhHeader>
Dim i As Long, temp As Language_entry, temp2 As tvb_res_Data
    If lang = 0 Then
        temp = Cur_Language
    Else
        temp.language = lang
        temp.LanguageString = LanguageString
    End If
    With file
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language And .headers(i).header.LanguageString = temp.LanguageString Then
                    temp2.Length = .headers(i).Length
                    temp2.Data = .headers(i).Data
                    Set Resource_GetImage = LoadImage(temp2)
                    Exit Function
                End If
            End If
        Next i
        
        'Try with the default language
        temp = .header.Languages(0)
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language And .headers(i).header.LanguageString = temp.LanguageString Then
                    temp2.Length = .headers(i).Length
                    temp2.Data = .headers(i).Data
                    Set Resource_GetImage = LoadImage(temp2)
                    Exit Function
                End If
            End If
        Next i
    End With
    
    '<EhFooter>
    Exit Function

Resource_GetImage_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_GetImage"
    '</EhFooter>
End Function


Public Function Resource_GetImageByIndex(file As Resource_File, index As Long) As StdPicture
    '<EhHeader>
    On Error GoTo Resource_GetImageByIndex_Err
    '</EhHeader>
Dim i As Long, temp As Language_entry, temp2 As tvb_res_Data

    temp2.Length = file.headers(index).Length
    temp2.Data = file.headers(index).Data
    Set Resource_GetImageByIndex = LoadImage(temp2)

Resource_GetImageByIndex_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_GetImageByIndex"
    '</EhFooter>
End Function

Public Function Resource_GetData(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional LanguageString As String = "") As tvb_res_Data
    '<EhHeader>
    On Error GoTo Resource_GetData_Err
    '</EhHeader>
Dim i As Long, temp As Language_entry
    If lang = 0 Then
        temp = Cur_Language
    Else
        temp.language = lang
        temp.LanguageString = LanguageString
    End If
    With file
        For i = 0 To .numEntrys
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language And .headers(i).header.LanguageString = temp.LanguageString Then
                    Resource_GetData.Data = .headers(i).Data
                    Resource_GetData.Length = .headers(i).Length
                    Exit Function
                End If
            End If
        Next i
        
        'Try with the default language
        temp = .header.Languages(0)
        For i = 0 To .numEntrys
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = temp.language And .headers(i).header.LanguageString = temp.LanguageString Then
                    Resource_GetData.Data = .headers(i).Data
                    Resource_GetData.Length = .headers(i).Length
                    Exit Function
                End If
            End If
        Next i
    End With
    
    '<EhFooter>
    Exit Function

Resource_GetData_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_GetData"
    '</EhFooter>
End Function

Public Sub Resource_LoadFormFromResourceFile(file As Resource_File, Form As Object, Optional prefix As String, Optional lang As tvb_Languages, Optional LanguageString As String = "")
    '<EhHeader>
    On Error GoTo LoadFormFromResourceFile_Err
    '</EhHeader>
    Dim ctrl As Object, t As String, erln As Long, i As Long
    
    Dim temp As Long
    temp = Resource_Exists(file, prefix & "-" & Form.Name & "-caption", lang, LanguageString)
    If temp <> -1 Then
        Form.caption = Resource_GetTextByIndex(file, temp)
    End If
    
    t = prefix & "-" & Form.Name & "-"
    For Each ctrl In Form
    
        res_load_text file, ctrl, t, 0, lang, LanguageString
        res_load_caption file, ctrl, t, 0, lang, LanguageString
        res_load_image file, ctrl, t, 0, lang, LanguageString
        res_load_picture file, ctrl, t, 0, lang, LanguageString
        res_load_ToolTipText file, ctrl, t, 0, lang, LanguageString
        res_load_List file, ctrl, t, 0, lang, LanguageString

    Next ctrl
    '<EhFooter>
    Exit Sub

LoadFormFromResourceFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "LoadFormFromResourceFile"
    '</EhFooter>
End Sub

Private Sub res_load_text(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_load_text_Err
    '</EhHeader>
    On Error GoTo ext
    Dim temp As Long
    temp = Resource_Exists(file, t & ctrl.Name & "-text", lang, LanguageString)
    If temp <> -1 Then
        ctrl.Text = Resource_GetTextByIndex(file, temp)
    End If
ext:
    '<EhFooter>
    Exit Sub

res_load_text_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_load_text"
    '</EhFooter>
End Sub

Private Sub res_load_caption(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_load_caption_Err
    '</EhHeader>
    On Error GoTo ext
    
    Dim temp As Long
    temp = Resource_Exists(file, t & ctrl.Name & "-caption", lang, LanguageString)
    If temp <> -1 Then
        ctrl.caption = Resource_GetTextByIndex(file, temp)
    End If
ext:
    '<EhFooter>
    Exit Sub

res_load_caption_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_load_caption"
    '</EhFooter>
End Sub

Private Sub res_load_image(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_load_image_Err
    '</EhHeader>
    On Error GoTo ext
    
    Dim temp As Long
    temp = Resource_Exists(file, t & ctrl.Name & "-image", lang, LanguageString)
    If temp <> -1 Then
        If Not (ctrl.image Is Nothing) Then
            Set ctrl.image = Resource_GetTextByIndex(file, temp)
        End If
    End If
ext:
    '<EhFooter>
    Exit Sub

res_load_image_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_load_image"
    '</EhFooter>
End Sub

Private Sub res_load_picture(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_load_picture_Err
    '</EhHeader>
    On Error GoTo ext
    
    Dim temp As Long
    temp = Resource_Exists(file, t & ctrl.Name & "-picture", lang, LanguageString)
    If temp <> -1 Then
        If Not (ctrl.Picture Is Nothing) Then
            Set ctrl.Picture = Resource_GetTextByIndex(file, temp)
        End If
    End If
ext:
    '<EhFooter>
    Exit Sub

res_load_picture_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_load_picture"
    '</EhFooter>
End Sub

Private Sub res_load_ToolTipText(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_load_ToolTipText_Err
    '</EhHeader>
    
    Dim temp As Long
    On Error GoTo ext
    
    temp = Resource_Exists(file, t & ctrl.Name & "-image", lang, LanguageString)
    If temp <> -1 Then
        ctrl.ToolTipText = Resource_GetTextByIndex(file, temp)
    End If
ext:
    '<EhFooter>
    Exit Sub

res_load_ToolTipText_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_load_ToolTipText"
    '</EhFooter>
End Sub

Private Sub res_load_List(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_load_List_Err
    '</EhHeader>
    On Error GoTo ext

    Dim i As Long
    If Resource_Exists(file, t & ctrl.Name & "-list-" & i, lang, LanguageString) <> -1 Then
        For i = 0 To ctrl.ListCount
            ctrl.AddItem Resource_GetText(file, t & ctrl.Name & "-list-" & i, lang, LanguageString), i
        Next i
    End If
ext:
    '<EhFooter>
    Exit Sub

res_load_List_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_load_List"
    '</EhFooter>
End Sub

Public Sub Resource_SaveFormToResourceFile(file As Resource_File, Form As Object, Optional prefix As String, Optional lang As tvb_Languages, Optional LanguageString As String = "", Optional PackMode As tvb_res_pack_mode = tvb_res_Stored)
    '<EhHeader>
    On Error GoTo Resource_SaveFormToResourceFile_Err
    '</EhHeader>
Dim ctrl As Object, t As String, erln As Long, i As Long

    Resource_AddEntry file, Resource_NewTextEntry(Form.caption, prefix & "-" & Form.Name & "-caption", PackMode, lang, LanguageString)

    t = prefix & "-" & Form.Name & "-"
    For Each ctrl In Form
        res_save_text file, ctrl, t, PackMode, lang, LanguageString
        res_save_caption file, ctrl, t, PackMode, lang, LanguageString
        res_save_image file, ctrl, t, PackMode, lang, LanguageString
        res_save_picture file, ctrl, t, PackMode, lang, LanguageString
        res_save_ToolTipText file, ctrl, t, PackMode, lang, LanguageString
        res_save_list file, ctrl, t, PackMode, lang, LanguageString
    Next ctrl
    
    '<EhFooter>
    Exit Sub

Resource_SaveFormToResourceFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_SaveFormToResourceFile"
    '</EhFooter>
End Sub

Private Sub res_save_text(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_save_text_Err
    '</EhHeader>
    On Error GoTo ext
    Resource_AddEntry file, Resource_NewTextEntry(ctrl.Text, t & ctrl.Name & "-text", PackMode, lang, LanguageString)
ext:
    '<EhFooter>
    Exit Sub

res_save_text_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_save_text"
    '</EhFooter>
End Sub

Private Sub res_save_caption(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_save_caption_Err
    '</EhHeader>
    On Error GoTo ext
    Resource_AddEntry file, Resource_NewTextEntry(ctrl.caption, t & ctrl.Name & "-caption", PackMode, lang, LanguageString)
ext:
    '<EhFooter>
    Exit Sub

res_save_caption_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_save_caption"
    '</EhFooter>
End Sub


Private Sub res_save_image(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_save_image_Err
    '</EhHeader>
    On Error GoTo ext
    If Not (ctrl.image Is Nothing) Then
        Resource_AddEntry file, Resource_NewImageEntry(ctrl.image, t & ctrl.Name & "-image", PackMode, lang, LanguageString)
    End If
ext:
    '<EhFooter>
    Exit Sub

res_save_image_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_save_image"
    '</EhFooter>
End Sub

Private Sub res_save_picture(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_save_picture_Err
    '</EhHeader>
    On Error GoTo ext
    If Not (ctrl.Picture Is Nothing) Then
        Resource_AddEntry file, Resource_NewImageEntry(ctrl.Picture, t & ctrl.Name & "-picture", PackMode, lang, LanguageString)
    End If
ext:
    '<EhFooter>
    Exit Sub

res_save_picture_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_save_picture"
    '</EhFooter>
End Sub

Private Sub res_save_ToolTipText(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_save_ToolTipText_Err
    '</EhHeader>
    On Error GoTo ext
    Resource_AddEntry file, Resource_NewTextEntry(ctrl.ToolTipText, t & ctrl.Name & "-tooltiptext", PackMode, lang, LanguageString)
ext:
    '<EhFooter>
    Exit Sub

res_save_ToolTipText_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_save_ToolTipText"
    '</EhFooter>
End Sub


Private Sub res_save_list(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages, LanguageString As String)
    '<EhHeader>
    On Error GoTo res_save_list_Err
    '</EhHeader>
    On Error GoTo ext
    Dim i As Long
    For i = 0 To ctrl.ListCount
        Resource_AddEntry file, Resource_NewTextEntry(ctrl.list(i), t & ctrl.Name & "-list-" & i, PackMode, lang, LanguageString)
    Next i
ext:
    '<EhFooter>
    Exit Sub

res_save_list_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "res_save_list"
    '</EhFooter>
End Sub


Public Function Resource_LanguageIdToString(language As tvb_Languages) As String
    '<EhHeader>
    On Error GoTo Resource_LanguageIdToString_Err
    '</EhHeader>
    Select Case language
        Case tvb_Languages.tvb_English
            Resource_LanguageIdToString = "English"
        Case tvb_Languages.tvb_Czech
            Resource_LanguageIdToString = "Czech"
        Case tvb_Languages.tvb_Greek
            Resource_LanguageIdToString = "Greek"
        Case Else
            Resource_LanguageIdToString = "Other[ " & language & "]"
    End Select
    '<EhFooter>
    Exit Function

Resource_LanguageIdToString_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_LanguageIdToString"
    '</EhFooter>
End Function

Public Function Resource_NewFile(author As String, desc As String, Name As String, deflangn As Language_entry) As Resource_File
    '<EhHeader>
    On Error GoTo Resource_NewFile_Err
    '</EhHeader>
Dim temp As Resource_File

    With temp.header
        .author = author
        .CreatedWith = "ThunderVB_pl"
        .Description = desc
        .LanguagesCount = 1
        ReDim .Languages(0)
        .Languages(0) = deflangn
        .Name = Name
        .Version = 1
    End With
    
    Resource_NewFile = temp
    
    '<EhFooter>
    Exit Function

Resource_NewFile_Err:
    LogMsg "Error : " & err.Description & " , At " & Add34(err.Source) & ":" & Erl, "ThunderVB_pl", "modResFileIO", "Resource_NewFile"
    '</EhFooter>
End Function

