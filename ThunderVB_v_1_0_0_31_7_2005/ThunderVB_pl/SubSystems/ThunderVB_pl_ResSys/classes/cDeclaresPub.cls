VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeclaresPub"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************
'Resource Subsystem
'*******************

Public Static Function Resource_LanguageIdToString(language As tvb_Languages) As String

    Resource_LanguageIdToString = modResFileIO.Resource_LanguageIdToString(language)
    
End Function

Public Static Function Resource_LoadResourceFile(file As String, key As String, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Resource_File

    Resource_LoadResourceFile = modResFileIO.Resource_LoadResourceFile(file, key, ErrorCode, ErrorString)
    
End Function

Public Static Sub Resource_SaveResourceFile(file As String, Data As Resource_File, key As String, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String)
    
    modResFileIO.Resource_SaveResourceFile file, Data, key, ErrorCode, ErrorString
    
End Sub

Public Static Sub Resource_AddEntry(toFile As Resource_File, NewEntry As tvb_res_entry)
    
    modResFileIO.Resource_AddEntry toFile, NewEntry
    
End Sub

Public Static Sub Resource_RemoveEntry(toFile As Resource_File, Id As Long)
    
    modResFileIO.Resource_RemoveEntry toFile, Id
    
End Sub

Public Static Function Resource_NewTextEntry(TextData As String, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages) As tvb_res_entry
    
    Resource_NewTextEntry = modResFileIO.Resource_NewTextEntry(TextData, Id, PackMode, language)
    
End Function

Public Static Function Resource_NewDataEntry(Data() As Byte, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages) As tvb_res_entry
    
    Resource_NewDataEntry = modResFileIO.Resource_NewDataEntry(Data, Id, PackMode, language)
    
End Function

Public Static Function Resource_NewImageEntry(image As StdPicture, Id As String, PackMode As tvb_res_pack_mode, language As tvb_Languages) As tvb_res_entry
    'Set image = frmRes.Picture
    Resource_NewImageEntry = modResFileIO.Resource_NewImageEntry(image, Id, PackMode, language)
    
End Function

Public Function Resource_Exists(file As Resource_File, Id As String, Optional lang As tvb_Languages) As Long
    
    Resource_Exists = modResFileIO.Resource_Exists(file, Id, lang)
    
End Function

Public Static Function Resource_GetText(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As String
    
    Resource_GetText = modResFileIO.Resource_GetText(file, Id, lang, ErrorCode, ErrorString)
    
End Function

Public Static Function Resource_GetTextByIndex(file As Resource_File, index As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As String
    
    Resource_GetTextByIndex = modResFileIO.Resource_GetTextByIndex(file, index, ErrorCode, ErrorString)
    
End Function


Public Static Function Resource_GetImage(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As StdPicture
    
    Set Resource_GetImage = modResFileIO.Resource_GetImage(file, Id, lang, ErrorCode, ErrorString)
    
End Function


Public Static Function Resource_GetImageByIndex(file As Resource_File, index As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As StdPicture
    
    Set Resource_GetImageByIndex = modResFileIO.Resource_GetImageByIndex(file, index, ErrorCode, ErrorString)
    
End Function

Public Static Function Resource_GetData(file As Resource_File, Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As tvb_res_Data
    
    Resource_GetData = modResFileIO.Resource_GetData(file, Id, lang, ErrorCode, ErrorString)
    
End Function


Public Static Sub Resource_LoadFormFromResourceFile(file As Resource_File, Form As Object, Optional prefix As String, Optional lang As tvb_Languages)
    
    modResFileIO.Resource_LoadFormFromResourceFile file, Form, prefix, lang
    
End Sub

Public Static Sub Resource_SaveFormToResourceFile(file As Resource_File, Form As Object, Optional prefix As String, Optional lang As tvb_Languages, Optional PackMode As tvb_res_pack_mode = tvb_res_Stored)
    
    modResFileIO.Resource_SaveFormToResourceFile file, Form, prefix, lang, PackMode
    
End Sub

Public Static Function Resource_NewFile(author As String, desc As String, Name As String) As Resource_File
    
    Resource_NewFile = modResFileIO.Resource_NewFile(author, desc, Name)
    
End Function

Public Static Function Resource_GetCurLang() As tvb_Languages
 
    Resource_GetCurLang = Cur_Language
    
End Function

Public Static Function Resource_SetCurLang(value As tvb_Languages)

    Cur_Language = value
    
End Function

Public Function Resource_SaveImageToResData(ByVal image As StdPicture) As tvb_res_Data
   
   Resource_SaveImageToResData = modResFileIO.SaveImage(image)
   
End Function

Public Function Resource_LoadImageFromResData(ImageBytes As tvb_res_Data) As StdPicture

   Set Resource_LoadImageFromResData = modResFileIO.LoadImage(ImageBytes)
   
End Function

