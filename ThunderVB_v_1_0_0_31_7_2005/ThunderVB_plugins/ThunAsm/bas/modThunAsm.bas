Attribute VB_Name = "modThunAsm"
Option Explicit

Public oMe As plugin

Global mcph As Long
'Global bID As Long
Global has_asm As Boolean, file_obj As String, file_asm As String, file_vb As String
Global Fix_unnamed As Boolean, Cancel_compile As Boolean
Global asm As Boolean, file_include As String

Global Const strIncFileData As String = ";; LISTING.INC" & vbNewLine & ";; This file contains assembler macros and is included by the files created" & vbNewLine & ";; with the -FA compiler switch to be assembled by MASM (Microsoft Macro" & vbNewLine & ";; Assembler)." & vbNewLine & ";; Copyright (c) 1993, Microsoft Corporation. All rights reserved." & vbNewLine & _
";; non destructive nops" & vbNewLine & "npad macro size" & vbNewLine & "if size eq 1" & vbNewLine & "  nop" & vbNewLine & "else" & vbNewLine & " if size eq 2" & vbNewLine & "   mov edi, edi" & vbNewLine & " else" & vbNewLine & "  if size eq 3" & vbNewLine & "    ; lea ecx, [ecx+00]" & vbNewLine & "    DB 8DH, 49H, 00H" & vbNewLine & "  else" & vbNewLine & "   if size eq 4" & vbNewLine & "     ; lea esp, [esp+00]" & vbNewLine & "     DB 8DH, 64H, 24H, 00H" & vbNewLine & _
"   else" & vbNewLine & "    if size eq 5" & vbNewLine & "      add eax, DWORD PTR 0" & vbNewLine & "    else" & vbNewLine & "     if size eq 6" & vbNewLine & "       ; lea ebx, [ebx+00000000]" & vbNewLine & "       DB 8DH, 9BH, 00H, 00H, 00H, 00H" & vbNewLine & "     else" & vbNewLine & "      if size eq 7" & vbNewLine & "    ; lea esp, [esp+00000000]" & vbNewLine & "    DB 8DH, 0A4H, 24H, 00H, 00H, 00H, 00H" & vbNewLine & _
"      else" & vbNewLine & "    %out error: unsupported npad size" & vbNewLine & "    .Err" & vbNewLine & "      endif" & vbNewLine & "     endif" & vbNewLine & "    endif" & vbNewLine & "   endif" & vbNewLine & "  endif" & vbNewLine & " endif" & vbNewLine & "endif" & vbNewLine & "endm" & vbNewLine & _
";; destructive nops" & vbNewLine & "dpad macro size, reg" & vbNewLine & "if size eq 1" & vbNewLine & "  inc reg" & vbNewLine & "else" & vbNewLine & "  %out error: unsupported dpad size" & vbNewLine & "  .Err" & vbNewLine & "endif" & vbNewLine & "endm" & vbNewLine


Public Const PLUGIN_NAME As String = "ThunderAsm"
Public Const MSG_TITLE As String = PLUGIN_NAME

Public Const PLUGIN_NAMEs As String = "ThunAsm"
Public Const MSG_TITLEs As String = PLUGIN_NAMEs


Public Const APP_NAME As String = PLUGIN_NAME
Public Const APP_NAMEs As String = PLUGIN_NAMEs

'module for whatever you want

'choose directory dialog
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

'Global Const AsmCHeadersEnabled As Boolean = True
Public AsmCHeadersAreAssembled As Boolean

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'------------------------
'--- helper functions ---
'------------------------

'set path to the program
'- sDialogTitle - open dialog Title
'-txtTarget     - textbox where path will be stored
'-sAppName      - app name (eg. ml.exe or midl.exe)

Public Sub SetPath(sDialogTitle As String, txtTarget As TextBox, Optional sAppName As String = "")

'set dialog title
frmIn.cdSet.DialogTitle = sDialogTitle
frmIn.cdSet.FileName = ""

'set new init directory
If Len(txtTarget.Text) <> 0 Then frmIn.cdSet.InitDir = Left(txtTarget.Text, InStrRev(txtTarget.Text, "\")) Else frmIn.cdSet.InitDir = App.Path & "\"

On Error Resume Next

'select file
frmIn.cdSet.ShowOpen
'cancel was pressed
If Err.Number = 32755 Then Exit Sub

On Error GoTo 0

'check predefined app name
If Len(sAppName) = 0 Then GoTo 10

'check filename
If StrComp(Right(frmIn.cdSet.FileName, Len(sAppName)), sAppName, vbTextCompare) <> 0 Then
10:
    MsgBoxX "Select " & Add34(sAppName) & " file.", vbInformation, "Settings"
Else

    'store path to the textbox
    txtTarget.Text = frmIn.cdSet.FileName
End If

End Sub

'code from API-GUIDE
'parameters - sPrompt - prompt
'           - txtTarget - textbox, where path to directory will be stored
'KPD-Team 1998
'URL: http://www.allapi.net/
'KPDTeam@Allapi.net
Public Sub SetDirectory(ByVal sPrompt As String, txtTarget As TextBox)
Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hwndOwner = frmIn.pctSettings.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(sPrompt, "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
            'append \
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        End If
    End If

    txtTarget.Text = sPath

End Sub

Public Function ExecuteCommand(ByVal CommandLine As String, sOutputText As String, Optional workdir As String, Optional eWindowState As ESW) As Boolean
        If Get_Compile(ModifyCmdLine) Then
            CommandLine = frmViewer.ShowViewer("Modify Command Line", CommandLine, False)
        End If
    ExecuteCommand = ThunderVB_pl_HFunct_v1_0.ExecuteCommand(CommandLine, sOutputText, workdir, eWindowState)
End Function
