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

'Functions "Exports" (static public functions)


'*******************
'***String helper***
'*******************

'get the fisrt word of a string , with the space folowing it
Public Static Function GetFirstWordWithSpace(ByRef str As String) As String

    GetFirstWordWithSpace = modStringFunct.GetFirstWordWithSpace(str)
    
End Function

'removes the first word from a string and the space after it
Public Static Sub RemFirstWordWithSpace(ByRef str As String)  'remove fisrt word

    Call modStringFunct.RemFirstWordWithSpace(str)
    
End Sub

'get the fisrt word of a string
Public Static Function GetFirstWord(ByRef str As String) As String

    GetFirstWord = modStringFunct.GetFirstWord(str)
    
End Function

'removes the first word from a string
Public Static Sub RemFisrtWord(ByRef str As String)  'remove fisrt word

    Call modStringFunct.RemFisrtWord(str)
    
End Sub

'Splits all words to array ..
Public Static Function GetAllWordsToArr(ByRef str As String) As String()

    GetAllWordsToArr = modStringFunct.GetAllWordsToArr(str)
    
End Function

'adds "  on the start and end of a string if they do not exist
Public Static Function Add34(ByRef sText As String) As String

    Add34 = modStringFunct.Add34(sText)

End Function

'removes "  on the start and end of a string if they do exist
Public Static Function Rem34(ByRef sText As String) As String

    Rem34 = modStringFunct.Rem34(sText)
    
End Function

'get the string between the find1 and find2 strings (without containing them)
Public Static Function getS(ByRef find1 As String, ByRef find2 As String, ByRef str As String, Optional ByRef start As Long = 1) As String

    getS = modStringFunct.getS(find1, find2, str, start)
    
End Function

'instr with white space
Public Static Function InStrWS(ByVal start As Long, ByRef string1 As String, ByRef string2 As String, ByVal cm As VbCompareMethod) As Long

    InStrWS = modStringFunct.InStrWS(start, string1, string2, cm)
    
End Function


'Replaces any queted string with "______" ...
Public Static Function ProcStringsUnderAll(ByRef s As String) As String

    ProcStringsUnderAll = modStringFunct.ProcStringsUnderAll(s)
    
End Function

Public Static Function CrLf(Optional ByVal lCount As Long = 1) As String

    CrLf = modStringFunct.CrLf(lCount)

End Function

Public Static Sub AppendString(ByRef toString As String_B, ByRef Data As String)

    modStringBuilder.AppendString toString, Data
    
End Sub

Public Static Sub FinaliseString(ByRef toString As String_B)

    modStringBuilder.FinaliseString toString
    
End Sub

Public Static Function GetString(ByRef fromString As String_B) As String

    GetString = modStringBuilder.GetString(fromString)
    
End Function


'for color code...
Public Static Sub AppendColString(ByRef toString As Col_String, ByVal Data As Long, ByVal col As Long)

    modStringBuilder.AppendColString toString, Data, col
    
End Sub

Public Static Sub FinaliseColString(ByRef toString As Col_String)

    modStringBuilder.FinaliseColString toString
    
End Sub

'*******************
'******File IO******
'*******************

'loads a text file ands returns its contents as String
Public Static Function LoadFile(File As String) As String

    LoadFile = modFileIO.LoadFile_string(File)
    
End Function

'saves a text file
Public Static Sub SaveFile(File As String, Data As String)

    modFileIO.SaveFile_string File, Data
    
End Sub

'loads a bin file ands returns its contents as byte()
Public Static Function LoadFile_bin(File As String) As Byte()

    LoadFile_bin = modFileIO.LoadFile_bin(File)
    
End Function

'saves a bin file
Public Static Sub SaveFile_bin(File As String, Data() As Byte)

    modFileIO.SaveFile_bin File, Data
    
End Sub

'checks if a file exists
Public Static Function FileExist(File As String) As Boolean
    
    FileExist = modFileIO.FileExist(File)
    
End Function

'checks if a file exists
Public Static Function DirExist(directory As String) As Boolean

    DirExist = modFileIO.DirExist(directory)
    
End Function

'gets the filename from a full file path (eg "c:\windows\notepad.exe"->"notepad.exe")
Public Static Function GetFilename(filepath As String) As String

    GetFilename = modFileIO.GetFilename(filepath)
    
End Function

'gets the path from a full file path (eg "c:\windows\notepad.exe" ->"c:\windows\")
Public Static Function GetPath(filepath As String) As String

    GetPath = modFileIO.GetPath(filepath)

End Function

'deletes a file if it exists..
Public Static Sub kill2(File As String)

    modFileIO.kill2 File

End Sub


Public Static Sub PrintToFile(File As file_b, Data As String)
    
    modCachedFileIO.PrintToFile File, Data

End Sub

Public Static Sub AppendToFile(File As file_b, Data As String)
    
    modCachedFileIO.AppendToFile File, Data
    
End Sub


Public Static Sub FlushFile(File As file_b)

    modCachedFileIO.FlushFile File

End Sub

Public Static Function OpenFile(filename As String, Optional buffersize As Long = 32768) As file_b
    
    OpenFile = modCachedFileIO.OpenFile(filename, buffersize)
    
End Function

Public Static Sub CloseFile(File As file_b)

    modCachedFileIO.CloseFile File
    
End Sub

Public Static Sub FileSeek(ByRef File As file_b, ByVal pos As Long)
    
    Call modCachedFileIO.FileSeek(File, pos)
    
End Sub

Public Static Function FileLength(ByRef File As file_b) As Long
    
   FileLength = modCachedFileIO.FileLength(File)
    
End Function

'*******************
'****Misc Things****
'*******************

Public Static Function ExecuteCommand(ByVal CommandLine As String, ByRef sOutputText As String, Optional workdir As String, Optional ByVal eWindowState As ESW = SW_HIDE) As Boolean

    ExecuteCommand = modMisc.ExecuteCommand(CommandLine, sOutputText, workdir, eWindowState)
    
End Function


Public Static Function getVBVersion() As Long

    getVBVersion = modMisc.getVBVersion()
    
End Function

Public Static Function CStringZero(lpCString As Long) As String

    CStringZero = modMisc.CStringZero(lpCString)
    
End Function


'From ansi String Pointer to vb string
Public Static Function Cstring(ByVal lpString As Long, ByVal nCount As Long) As String

    Cstring = modMisc.Cstring(lpString, nCount)
    
End Function

'From Wide String Pointer to vb string
Public Static Function CstringW(ByVal lpString As Long, ByVal nCount As Long) As String

    CstringW = modMisc.CstringW(lpString, nCount)

End Function

'Form VB str to Ansi Byte array
Public Static Sub BstrToAnsi(str As String, ba() As Byte)

     modMisc.BstrToAnsi str, ba
    
End Sub

Public Static Function ArrUBound(arr As Variant) As Long
    
    ArrUBound = modMisc.ArrUBound(arr)
    
End Function



Public Static Function ArrLBound(arr As Variant) As Long
    
    ArrLBound = modMisc.ArrLBound(arr)
    
End Function



Public Static Sub HelperFunct_SetLogger(LoggerObj As ILogger)
    
    Call SetLogger(LoggerObj)
    
End Sub

