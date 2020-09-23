VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTypePub"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ToDo : add some others here
Public Enum tvb_Languages

    tvb_English = 1
    tvb_Czech = 2
    tvb_Greek = 3
    tvb_Russian = 4
    tvb_Arabic = 5
    tvb_Bulgarian = 6
    tvb_Croatian = 7
    tvb_Danish = 8
    tvb_Dutch = 9
    tvb_Estonian = 10
    tvb_Finnish = 11
    tvb_French = 12
    tvb_German = 13
    tvb_Hebrew = 14
    tvb_Hungarian = 15
    tvb_Italian = 16
    tvb_Japanese = 17
    tvb_Korean = 18
    tvb_Latvian = 19
    tvb_Lithuanian = 20
    tvb_Norwegian = 21
    tvb_Polish = 22
    tvb_Portuguese_Brazil = 23
    tvb_Portuguese_Standard = 24
    tvb_Romanian = 25
    tvb_Simplified_Chinese = 26
    tvb_Spanish = 27
    tvb_Slovak = 28
    tvb_Slovenian = 29
    tvb_Swedish = 30
    tvb_Thai = 31
    tvb_Traditional_Chinese = 32
    tvb_Turkish = 33

    'misc..
    tvb_null = 0  'null language - is replaced with the current
    tvb_min = 1
    tvb_max = 33
    tvb_None = -1 '-- no comment
    

End Enum


'How the data is stored..
Public Enum tvb_res_pack_mode

    tvb_res_Stored = 1
    tvb_res_Compressed = 2 'UZL compression
    tvb_res_Encrypted = 4  'SXEA custom key size..
    tvb_res_EncryptedCompressed = 6 'it is a logical OR ..
    
End Enum

'What type of data?
Public Enum tvb_res_Data_Type

    tvb_res_Text = 0        'unicode text
    tvb_res_Image = 1       'image
    tvb_res_Data = 2        'binary data
    
End Enum


'Header for each file
Public Type tvb_res_main_header
    
    fftag As Long
    Version As Long 'Version of the format
    Name As String 'name for this file
    Description As String 'description for this file
    CreatedWith As String 'name of the tool used to create this file
    author As String 'name of the person that created this file..

End Type

'Header for every entry
Public Type tvb_res_entry_header
    
    Id As String ' id for this entry , must be unique per language..
    language As tvb_Languages 'Language of this entry
    PackMode As tvb_res_pack_mode
    PackInfo As Long         'Settings for the packer..
    DataType As tvb_res_Data_Type
    
End Type

Public Type tvb_res_Data
    
    Length As Long 'The underlaying data structure
    Data() As Byte
    
End Type

'A simple entry structure
Public Type tvb_res_entry

    header As tvb_res_entry_header
    Length As Long
    Data() As Byte
    
End Type

'Represeting a resource file
Public Type Resource_File

    header As tvb_res_main_header
    numEntrys As Long
    headers() As tvb_res_entry
    
End Type



Public Enum tvb_res_errorcodes
    NoError = 0
    UnsuportedFile = 1
    FileDoesNotExist = 2
    NoDataEntrys = 3
    InvalidEntryHeader = 4
    EntryDoesNotExist = 5
End Enum
