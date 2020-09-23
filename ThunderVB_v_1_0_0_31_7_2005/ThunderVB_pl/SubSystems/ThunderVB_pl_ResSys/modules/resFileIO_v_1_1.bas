Attribute VB_Name = "modResFileIO"
Option Explicit
'File format version 1.0 [ResSys v1.0]
'This is the first release ..
'newer versions will be able to be converted from this format

Public Const tvb_resform_ver_maj  As Long = 1
Public Const tvb_resform_ver_min  As Long = 1
Public Const tvb_resform_version As Long = tvb_resform_ver_maj * 65536 Or tvb_resform_ver_min

Public Const defLang As Long = tvb_Languages.tvb_English
Public Const fftag  As Long = 1163020078 '".GRE"

Global Cur_Language As tvb_Languages

Public Function LoadImage( _
   ImageBytes As tvb_res_Data) As StdPicture
Dim oPersist As IPersistStream
Dim oStream As IStream
Dim lSize As Long
  If ImageBytes.Length < 1 Then Exit Function
   ' Calculate the array size
   lSize = ArrUBound(ImageBytes.Data) - ArrLBound(ImageBytes.Data) + 1
   If lSize = 1 Then Exit Function
   ' Create a stream object
   ' in global memory
   Set oStream = CreateStreamOnHGlobal(0, True)
   
   ' Write the header to the stream
   oStream.Write &H746C&, 4&
   
   ' Write the array size
   oStream.Write lSize, 4&
   
   ' Write the image data
   oStream.Write ImageBytes.Data(ArrLBound(ImageBytes.Data)), lSize
   
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
   
End Function

Public Function SaveImage( _
   ByVal image As StdPicture) As tvb_res_Data
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
   SaveImage.Length = ArrUBound(abData) + 1
   
   ' Release the stream object
   Set oStream = Nothing

End Function


Public Function Resource_LoadResourceFile(file As String, key As String, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Resource_File
    Dim Temp As Resource_File, i As Long
    Dim ff As Long
    ErrorCode = NoError
    
    If FileExist(file) Then
        ff = FreeFile
        Open file For Binary As ff
        Temp.header = Resource_ReadHeader(ff, ErrorCode, ErrorString)
        If ErrorCode <> NoError Then
            Exit Function
        End If
    Else
        Close ff
        ErrorCode = FileDoesNotExist
        ErrorString = "File " & file & " was not found..."
        Exit Function
    End If
    Get ff, , Temp.numEntrys
    If Temp.numEntrys > 0 Then
        ReDim Temp.headers(Temp.numEntrys - 1)
        For i = 0 To Temp.numEntrys - 1
            Temp.headers(i) = Resource_LoadEntry(ff, key, 1, ErrorCode, ErrorString)
            If ErrorCode <> NoError Then Exit Function
        Next i
    Else
        Close ff
        ErrorCode = NoDataEntrys
        ErrorString = "Resource_File.numEntrys is invalid [" & Temp.numEntrys & "]"
        Exit Function
    End If
    
    Close ff
    Resource_LoadResourceFile = Temp
    
End Function

Public Function Resource_SaveResourceFile(file As String, Data As Resource_File, key As String, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Boolean
    Dim i As Long
    Dim ff As Long
    ErrorCode = NoError

    ff = FreeFile
    kill2 file
    Open file For Binary As ff
    Resource_SaveHeader ff, Data.header

    Put ff, , Data.numEntrys
    If Data.numEntrys > 0 Then
        For i = 0 To Data.numEntrys - 1
            Resource_SaveEntry ff, key, Data.headers(i), 1
        Next i
    Else
        Close ff
        ErrorCode = NoDataEntrys
        ErrorString = "Resource_File.numEntrys is invalid [" & Data.numEntrys & "]"
        Resource_SaveResourceFile = False
        Exit Function
    End If
    Close ff
    
    Resource_SaveResourceFile = True
    
End Function

Public Function Resource_ReadHeader(ff As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As tvb_res_main_header
    Dim Temp As tvb_res_main_header, i As Long
    Dim sz As Long
    ErrorCode = NoError
    
    With Temp
        'Read all the static data..
        Get ff, , .fftag
        '.fftag = fftag
        Get ff, , .Version
        'Check if the header is comppitable
        
        If Resource_MainHeaderIsSupported(Temp) = False Then
            'Too bad .. all this was for nothing :(
            'On Error GoTo 0
            ErrorCode = UnsuportedFile
            ErrorString = "File is Corupted/not Supported"
            Exit Function
        End If
        
        'yay , we can continue it :)
        Get ff, , sz: .Name = Space$(sz)
        Get ff, , .Name
        
        Get ff, , sz: .author = Space$(sz)
        Get ff, , .author
        
        Get ff, , sz: .CreatedWith = Space$(sz)
        Get ff, , .CreatedWith
        
        Get ff, , sz: .Description = Space$(sz)
        Get ff, , .Description
        
    End With
    
    'Ok, done :):):)
    Resource_ReadHeader = Temp
    
End Function

Public Sub Resource_SaveHeader(ff As Long, header As tvb_res_main_header)
    Dim i As Long
    
    With header
        'Save all the static data..
        Put ff, , .fftag
        Put ff, , .Version
        
        'Save the custom info..
        Put ff, , Len(.Name)
        Put ff, , .Name
        
        Put ff, , Len(.author)
        Put ff, , .author
        
        Put ff, , Len(.CreatedWith)
        Put ff, , .CreatedWith
        
        Put ff, , Len(.Description)
        Put ff, , .Description
        
        'Ok, done :):):)
    End With
    
End Sub

Public Function Resource_LoadEntry(ff As Long, key As String, DictionarySize As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As tvb_res_entry
    Dim Temp As tvb_res_entry
    ErrorCode = NoError
    
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
            ErrorString = "Error on tvb_res_entry.header"
            ErrorCode = InvalidEntryHeader
            Exit Function
        End If
    End With
    Resource_LoadEntry = Temp
End Function

Public Function Resource_SaveEntry(ff As Long, key As String, Data As tvb_res_entry, DictionarySize As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Boolean
    Dim Temp As tvb_res_entry
    ErrorCode = NoError
    
    Temp = Data
    With Temp
        If Resource_EntryHeaderIsValid(.header) Then
        
            Resource_Save_tvb_res_entry_header ff, .header
            Resource_EncodeData Temp, key, DictionarySize
            Put ff, , .Length
            Put ff, , .Data
        Else
            ErrorCode = InvalidEntryHeader
            ErrorString = "Error on tvb_res_entry.header"
            Resource_SaveEntry = False
            Exit Function
        End If
    End With
    
    Resource_SaveEntry = True
End Function

Public Function Resource_MainHeaderIsSupported(header As tvb_res_main_header) As Boolean
    
    'Is a resource file ?
    If header.fftag <> fftag Then
        'who &^*^&$%^ tried to open a non resource file ?!?!
        Resource_MainHeaderIsSupported = False
        Exit Function
    End If
    
    'Is it a suported version ?
    If header.Version <> tvb_resform_version Then
        Resource_MainHeaderIsSupported = False
        Exit Function
    End If
    
    Resource_MainHeaderIsSupported = True
    
End Function

Public Function Resource_EntryHeaderIsValid(header As tvb_res_entry_header) As Boolean
    
    'Yeah , toooo much checking :D .. this is for future use..
    With header
        
    End With
    
    Resource_EntryHeaderIsValid = True
    
End Function

'Decodes /decompresses the data
Public Sub Resource_DecodeData(Data As tvb_res_entry, key As String, DictionarySize As Long)

    If Data.header.PackMode And tvb_res_Encrypted Then
        Dim Temp As New clsSimpleXOR
        Temp.DecryptByte Data.Data, key
    End If
    
    If Data.header.PackMode And tvb_res_Compressed Then
        DeCompress_Arr Data.Data
    End If
    Data.Length = ArrUBound(Data.Data) + 1
    
End Sub

'encodes /compresses the data
Public Sub Resource_EncodeData(Data As tvb_res_entry, key As String, dictSiz As Long)

    If Data.header.PackMode And tvb_res_Compressed Then
        Compress_Arr Data.Data
    End If
    If Data.header.PackMode And tvb_res_Encrypted Then
        Dim Temp As New clsSimpleXOR
        Temp.EncryptByte Data.Data, key
    End If
    Data.Length = ArrUBound(Data.Data) + 1
    
End Sub

Public Function Resource_ReadLanguageEntry(ff As Long) As tvb_Languages
Dim sz As Long

    Get ff, , Resource_ReadLanguageEntry
    
End Function

Public Sub Resource_SaveLanguageEntry(ff As Long, Data As tvb_Languages)
    
    Put ff, , Data

End Sub

Public Function Resource_Read_tvb_res_entry_header(ff As Long) As tvb_res_entry_header
    Dim sz As Long
    With Resource_Read_tvb_res_entry_header
        Get ff, , sz: .Id = Space$(sz)
        Get ff, , .Id
        Get ff, , .DataType
        Get ff, , .language
        Get ff, , .PackMode
        Get ff, , .PackInfo
    End With
    
End Function

Public Sub Resource_Save_tvb_res_entry_header(ff As Long, Data As tvb_res_entry_header)
    
    With Data
        Put ff, , Len(.Id)
        Put ff, , .Id
        Put ff, , .DataType
        Put ff, , .language
        Put ff, , .PackMode
        Put ff, , .PackInfo
    End With
    
End Sub

Public Sub Resource_AddEntry(toFile As Resource_File, NewEntry As tvb_res_entry)
    
    If NewEntry.Length < 1 Then Exit Sub
    
    With toFile
        ReDim Preserve .headers(.numEntrys)
        .headers(.numEntrys) = NewEntry
        .numEntrys = .numEntrys + 1
    End With
    
End Sub

Public Sub Resource_RemoveEntry(toFile As Resource_File, Id As Long)
    Dim i As Long
    With toFile
        For i = Id To .numEntrys - 2
            .headers(i) = .headers(i + 1)
        Next i
        .numEntrys = .numEntrys - 1
        ReDim Preserve .headers(.numEntrys)
    End With
End Sub

Public Function Resource_NewTextEntry(TextData As String, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages) As tvb_res_entry

    If Len(TextData) = 0 Then Exit Function
    With Resource_NewTextEntry
        .header.Id = Id
        .header.DataType = tvb_res_Text
        .header.language = language
        .header.PackMode = PackMode
        .Data = TextData
        .Length = ArrUBound(.Data) + 1
    End With
    
End Function

Public Function Resource_NewDataEntry(Data() As Byte, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages) As tvb_res_entry

    With Resource_NewDataEntry
        .header.Id = Id
        .header.DataType = tvb_res_Data
        .header.language = language
        .header.PackMode = PackMode
        .Data = Data
        .Length = ArrUBound(Data) + 1
    End With
    
End Function

Public Function Resource_NewImageEntry(image As StdPicture, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages) As tvb_res_entry
Dim Temp As tvb_res_Data

    With Resource_NewImageEntry
        .header.Id = Id
        .header.DataType = tvb_res_Image
        .header.language = language
        .header.PackMode = PackMode
        Temp = SaveImage(image)
        .Data = Temp.Data
        .Length = Temp.Length
    End With
    
End Function

Public Function Resource_Exists(file As Resource_File, Id As String, Optional lang As tvb_Languages) As Long
Dim i As Long, language As tvb_Languages

    If lang = 0 Then
        language = Cur_Language
    Else
        language = lang
    End If
    
    With file
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = language Then
                    Resource_Exists = i
                    Exit Function
                End If
            End If
        Next i
        

        ' if not found with req lang ,try with default
        language = defLang
        
        
        For i = 0 To .numEntrys - 1
            If LCase$(.headers(i).header.Id) = LCase$(Id) Then
                If .headers(i).header.language = language Then
                    Resource_Exists = i
                    Exit Function
                End If
            End If
        Next i
    End With
    
    Resource_Exists = -1
    
End Function

Public Function Resource_GetText(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As String
Dim i As Long
    ErrorCode = NoError
    
    i = Resource_Exists(file, Id, lang)
    If i >= 0 Then
        Resource_GetText = file.headers(i).Data
    Else
        ErrorCode = EntryDoesNotExist
        ErrorString = "The specifyed entry " & Id & "was not found..."
    End If
    
End Function


Public Function Resource_GetTextByIndex(file As Resource_File, index As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As String
    Dim i As Long
    ErrorCode = NoError
    
    If index >= 0 And index < file.numEntrys Then
        Resource_GetTextByIndex = file.headers(index).Data
    Else
        ErrorCode = EntryDoesNotExist
        ErrorString = "The specifyed entry " & index & "was not found..."
    End If
    
End Function

Public Function Resource_GetImage(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As StdPicture
    Dim i As Long, temp2 As tvb_res_Data
    ErrorCode = NoError
    
    i = Resource_Exists(file, Id, lang)
    If i >= 0 Then
        temp2.Length = file.headers(i).Length
        temp2.Data = file.headers(i).Data
        Set Resource_GetImage = LoadImage(temp2)
    Else
        ErrorCode = EntryDoesNotExist
        ErrorString = "The specifyed entry " & Id & "was not found..."
    End If
    
End Function


Public Function Resource_GetImageByIndex(file As Resource_File, index As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As StdPicture
    Dim i As Long, temp2 As tvb_res_Data
    ErrorCode = NoError
    
    If index >= 0 And index < file.numEntrys Then
        temp2.Length = file.headers(index).Length
        temp2.Data = file.headers(index).Data
        Set Resource_GetImageByIndex = LoadImage(temp2)
    Else
        ErrorCode = EntryDoesNotExist
        ErrorString = "The specifyed entry " & index & "was not found..."
    End If
    
End Function

Public Function Resource_GetData(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As tvb_res_Data
    Dim i As Long
    ErrorCode = NoError
   
    i = Resource_Exists(file, Id, lang)
    If i >= 0 Then
        Resource_GetData.Data = file.headers(i).Data
        Resource_GetData.Length = file.headers(i).Length
    Else
        ErrorCode = EntryDoesNotExist
        ErrorString = "The specifyed entry " & Id & "was not found..."
    End If
    
End Function

Public Sub Resource_LoadFormFromResourceFile(file As Resource_File, Form As Object, Optional prefix As String, Optional lang As tvb_Languages)
    Dim ctrl As Object, t As String, erln As Long, i As Long
    
    Dim Temp As Long
    Temp = Resource_Exists(file, prefix & "-" & Form.Name & "-caption", lang)
    If Temp <> -1 Then
        Form.Caption = Resource_GetTextByIndex(file, Temp)
    End If
    
    t = prefix & "-" & Form.Name & "-"
    For Each ctrl In Form
    
        res_load_text file, ctrl, t, 0, lang
        res_load_caption file, ctrl, t, 0, lang
        res_load_image file, ctrl, t, 0, lang
        res_load_picture file, ctrl, t, 0, lang
        res_load_ToolTipText file, ctrl, t, 0, lang
        res_load_List file, ctrl, t, 0, lang

    Next ctrl
End Sub

Private Sub res_load_text(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    Dim Temp As Long
    Temp = Resource_Exists(file, t & ctrl.Name & "-text", lang)
    If Temp <> -1 Then
        ctrl.Text = Resource_GetTextByIndex(file, Temp)
    End If
ext:
End Sub

Private Sub res_load_caption(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    
    Dim Temp As Long
    Temp = Resource_Exists(file, t & ctrl.Name & "-caption", lang)
    If Temp <> -1 Then
        ctrl.Caption = Resource_GetTextByIndex(file, Temp)
    End If
ext:
End Sub

Private Sub res_load_image(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    
    Dim Temp As Long
    Temp = Resource_Exists(file, t & ctrl.Name & "-image", lang)
    If Temp <> -1 Then
        If Not (ctrl.image Is Nothing) Then
            Set ctrl.image = Resource_GetTextByIndex(file, Temp)
        End If
    End If
ext:
End Sub

Private Sub res_load_picture(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    
    Dim Temp As Long
    Temp = Resource_Exists(file, t & ctrl.Name & "-picture", lang)
    If Temp <> -1 Then
        If Not (ctrl.Picture Is Nothing) Then
            Set ctrl.Picture = Resource_GetTextByIndex(file, Temp)
        End If
    End If
ext:
End Sub

Private Sub res_load_ToolTipText(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    
    Dim Temp As Long
    On Error GoTo ext
    
    Temp = Resource_Exists(file, t & ctrl.Name & "-image", lang)
    If Temp <> -1 Then
        ctrl.ToolTipText = Resource_GetTextByIndex(file, Temp)
    End If
ext:
End Sub

Private Sub res_load_List(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext

    Dim i As Long
    If Resource_Exists(file, t & ctrl.Name & "-list-" & i, lang) <> -1 Then
        For i = 0 To ctrl.ListCount
            ctrl.AddItem Resource_GetText(file, t & ctrl.Name & "-list-" & i, lang), i
        Next i
    End If
ext:
End Sub

Public Sub Resource_SaveFormToResourceFile(file As Resource_File, Form As Object, Optional prefix As String, Optional lang As tvb_Languages, Optional PackMode As tvb_res_pack_mode = tvb_res_Stored)
Dim ctrl As Object, t As String, erln As Long, i As Long

    Resource_AddEntry file, Resource_NewTextEntry(Form.Caption, prefix & "-" & Form.Name & "-caption", PackMode, lang)

    t = prefix & "-" & Form.Name & "-"
    For Each ctrl In Form
        res_save_text file, ctrl, t, PackMode, lang
        res_save_caption file, ctrl, t, PackMode, lang
        res_save_image file, ctrl, t, PackMode, lang
        res_save_picture file, ctrl, t, PackMode, lang
        res_save_ToolTipText file, ctrl, t, PackMode, lang
        res_save_list file, ctrl, t, PackMode, lang
    Next ctrl
    
End Sub

Private Sub res_save_text(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    Resource_AddEntry file, Resource_NewTextEntry(ctrl.Text, t & ctrl.Name & "-text", PackMode, lang)
ext:
End Sub

Private Sub res_save_caption(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    Resource_AddEntry file, Resource_NewTextEntry(ctrl.Caption, t & ctrl.Name & "-caption", PackMode, lang)
ext:
End Sub


Private Sub res_save_image(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    If Not (ctrl.image Is Nothing) Then
        Resource_AddEntry file, Resource_NewImageEntry(ctrl.image, t & ctrl.Name & "-image", PackMode, lang)
    End If
ext:
End Sub

Private Sub res_save_picture(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    If Not (ctrl.Picture Is Nothing) Then
        Resource_AddEntry file, Resource_NewImageEntry(ctrl.Picture, t & ctrl.Name & "-picture", PackMode, lang)
    End If
ext:
End Sub

Private Sub res_save_ToolTipText(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    Resource_AddEntry file, Resource_NewTextEntry(ctrl.ToolTipText, t & ctrl.Name & "-tooltiptext", PackMode, lang)
ext:
End Sub


Private Sub res_save_list(file As Resource_File, ctrl As Object, t As String, PackMode As tvb_res_pack_mode, lang As tvb_Languages)
    On Error GoTo ext
    Dim i As Long
    For i = 0 To ctrl.ListCount
        Resource_AddEntry file, Resource_NewTextEntry(ctrl.List(i), t & ctrl.Name & "-list-" & i, PackMode, lang)
    Next i
ext:
End Sub


Public Function Resource_LanguageIdToString(language As tvb_Languages) As String
    Select Case language
    
        Case tvb_Languages.tvb_English
            Resource_LanguageIdToString = "English"
            
        Case tvb_Languages.tvb_Czech
            Resource_LanguageIdToString = "Czech"
            
        Case tvb_Languages.tvb_Greek
            Resource_LanguageIdToString = "Greek"
            
        Case tvb_Languages.tvb_Russian
            Resource_LanguageIdToString = "Russian"
    
        Case tvb_Languages.tvb_Arabic
            Resource_LanguageIdToString = "Arabic"
            
        Case tvb_Languages.tvb_Bulgarian
            Resource_LanguageIdToString = "Bulgarian"
            
        Case tvb_Languages.tvb_Croatian
            Resource_LanguageIdToString = "Croatian"
            
        Case tvb_Languages.tvb_Danish
            Resource_LanguageIdToString = "Danish"
            
        Case tvb_Languages.tvb_Dutch
            Resource_LanguageIdToString = "Dutch"
            
        Case tvb_Languages.tvb_Estonian
            Resource_LanguageIdToString = "Estonian"
            
        Case tvb_Languages.tvb_Finnish
            Resource_LanguageIdToString = "Finnish"
            
        Case tvb_Languages.tvb_French
            Resource_LanguageIdToString = "French"
            
        Case tvb_Languages.tvb_German
            Resource_LanguageIdToString = "German"
            
        Case tvb_Languages.tvb_Hebrew
            Resource_LanguageIdToString = "Hebrew"
            
        Case tvb_Languages.tvb_Hungarian
            Resource_LanguageIdToString = "Hungarian"
            
        Case tvb_Languages.tvb_Italian
            Resource_LanguageIdToString = "Italian"
            
        Case tvb_Languages.tvb_Japanese
            Resource_LanguageIdToString = "Japanese"
            
        Case tvb_Languages.tvb_Korean
            Resource_LanguageIdToString = "Korean"
            
        Case tvb_Languages.tvb_Latvian
            Resource_LanguageIdToString = "Latvian"
            
        Case tvb_Languages.tvb_Lithuanian
            Resource_LanguageIdToString = "Lithuanian"
            
        Case tvb_Languages.tvb_Norwegian
            Resource_LanguageIdToString = "Norwegian"
            
        Case tvb_Languages.tvb_Polish
            Resource_LanguageIdToString = "Polish"
            
        Case tvb_Languages.tvb_Portuguese_Brazil
            Resource_LanguageIdToString = "Portuguese (Brazil)"
            
        Case tvb_Languages.tvb_Portuguese_Standard
            Resource_LanguageIdToString = "Portuguese (Standard)"
            
        Case tvb_Languages.tvb_Romanian
            Resource_LanguageIdToString = "Romanian"
            
        Case tvb_Languages.tvb_Simplified_Chinese
            Resource_LanguageIdToString = "Simplified Chinese"
            
        Case tvb_Languages.tvb_Spanish
            Resource_LanguageIdToString = "Spanish"
            
        Case tvb_Languages.tvb_Slovak
            Resource_LanguageIdToString = "Slovak"
            
        Case tvb_Languages.tvb_Slovenian
            Resource_LanguageIdToString = "Slovenian"
            
        Case tvb_Languages.tvb_Swedish
            Resource_LanguageIdToString = "Swedish"
            
        Case tvb_Languages.tvb_Thai
            Resource_LanguageIdToString = "Thai"
            
        Case tvb_Languages.tvb_Traditional_Chinese
            Resource_LanguageIdToString = "Traditional Chinese"
            
        Case tvb_Languages.tvb_Turkish
            Resource_LanguageIdToString = "Turkish"
            
        Case tvb_Languages.tvb_None
            Resource_LanguageIdToString = "None"
            
        Case Else
            Resource_LanguageIdToString = "Not Recognised ;#=" & language
                
    End Select
    
End Function

Public Function Resource_NewFile(author As String, desc As String, Name As String) As Resource_File
Dim Temp As Resource_File

    With Temp.header
        .author = author
        .CreatedWith = "ThunnderVB_pl_ResSys_v1.0 final"
        .Description = desc
        .Name = Name
        
        'init version and file format tag
        .Version = tvb_resform_version
        .fftag = fftag
    End With
    
    Resource_NewFile = Temp
    
End Function

