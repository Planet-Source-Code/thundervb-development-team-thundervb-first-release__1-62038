VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResFile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim t As Resource_File

Public FileName As String
Public FileKey As String
Dim m_isFileopen As Boolean

Public Function OpenFile(ByRef file As String, ByRef key As String, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Boolean
    
    If TryFF10(file, key) = False Then
        ErrorString = "File " & file & " is a not suported file format" & vbNewLine & "Suported formats are 1.0[tvbrs v1.0]"
        ErrorCode = tvb_res_errorcodes.UnsuportedFile
        OpenFile = False
        Exit Function
    End If
    
    FileName = file
    FileKey = key
    m_isFileopen = True
    OpenFile = True

End Function

Public Sub Create(ByRef author As String, ByRef desc As String, ByRef Name As String)
    
    t = Resource_NewFile(author, desc, Name)
    m_isFileopen = True
    
End Sub

Public Sub SaveFile()

    Resource_SaveResourceFile FileName, t, FileKey
    
End Sub

Public Sub SaveFileAs(newFilename As String, newFileKey As String)

    FileName = newFilename: FileKey = newFileKey
    Resource_SaveResourceFile FileName, t, FileKey
    
End Sub

Public Sub CloseFile()

    SaveFile
    m_isFileopen = False
    
End Sub


Private Function TryFF10(ByRef strfile As String, key As String, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As Boolean

    t = modResFileIO.Resource_LoadResourceFile(strfile, key, ErrorCode, ErrorString)
    If ErrorCode = NoError Then
        TryFF10 = True
    Else
        TryFF10 = False
    End If

End Function

Public Property Get IsFileOpen() As Boolean

    IsFileOpen = m_isFileopen
    
End Property

Public Function ResourceExists(Id As String, Optional lang As tvb_Languages) As Long
    
    ResourceExists = modResFileIO.Resource_Exists(t, Id, lang)
    
End Function



Public Function GetText(Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As String
    
    GetText = modResFileIO.Resource_GetText(t, Id, lang, ErrorCode, ErrorString)
    
End Function


Public Function GetTextByIndex(index As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As String
    Dim i As Long
    
    GetTextByIndex = modResFileIO.Resource_GetTextByIndex(t, index, ErrorCode, ErrorString)
    
End Function

Public Function GetImage(Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As StdPicture
    
    Set GetImage = modResFileIO.Resource_GetImage(t, Id, lang, ErrorCode, ErrorString)
    
End Function


Public Function GetImageByIndex(index As Long, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As StdPicture

    Set GetImageByIndex = modResFileIO.Resource_GetImageByIndex(t, index, ErrorCode, ErrorString)

End Function

Public Function GetData(Id As String, Optional lang As tvb_Languages, Optional ByRef ErrorCode As tvb_res_errorcodes, Optional ByRef ErrorString As String) As tvb_res_Data
        
    GetData = modResFileIO.Resource_GetData(t, Id, lang, ErrorCode, ErrorString)
        
End Function

Public Sub LoadFormFromResourceFile(ToForm As Object, Optional prefix As String, Optional lang As tvb_Languages)
    SetEhMode Err_expected
    
    Call modResFileIO.Resource_LoadFormFromResourceFile(t, ToForm, prefix, lang)
    
    RestoreEhMode
End Sub

Public Sub SaveFormToResourceFile(FromForm As Object, Optional prefix As String, Optional lang As ThunderVB_pl_ResSys_v1_0.tvb_Languages, Optional PackMode As ThunderVB_pl_ResSys_v1_0.tvb_res_pack_mode = tvb_res_Stored)
    SetEhMode Err_expected

    Call modResFileIO.Resource_SaveFormToResourceFile(t, FromForm, prefix, lang, PackMode)
    
    RestoreEhMode
End Sub


Public Function LanguageIdToString(language As ThunderVB_pl_ResSys_v1_0.tvb_Languages) As String
    
    LanguageIdToString = modResFileIO.Resource_LanguageIdToString(language)
    
End Function

Public Property Get file() As Resource_File

    file = t
    
End Property

Public Property Let file(value As Resource_File)

     t = value
    
End Property

Public Property Get Items(i As Long) As tvb_res_entry

    Items = t.headers(i)
    
End Property


Public Property Let Items(i As Long, value As tvb_res_entry)

    t.headers(i) = value
    
End Property

Public Property Get ItemCount() As Long
    
    ItemCount = t.numEntrys
    
End Property


Public Sub AddEntry(NewEntry As tvb_res_entry)
    
    Call modResFileIO.Resource_AddEntry(t, NewEntry)
    
End Sub

Public Sub RemoveEntry(Id As Long)
    
    Call modResFileIO.Resource_RemoveEntry(t, Id)
    
End Sub

