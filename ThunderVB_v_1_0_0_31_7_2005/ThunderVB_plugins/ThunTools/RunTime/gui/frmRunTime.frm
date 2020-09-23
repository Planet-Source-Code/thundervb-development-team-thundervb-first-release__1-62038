VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRunTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASM - C RunTime"
   ClientHeight    =   3945
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBut 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   1800
      Picture         =   "frmRunTime.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdToCursor 
      Caption         =   "Cursor"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdToClipboard 
      Caption         =   "Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwRT 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paste to"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   585
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Se&lect"
      Begin VB.Menu mnuAll 
         Caption         =   "All"
      End
      Begin VB.Menu mnuNo 
         Caption         =   "No"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlyASM 
         Caption         =   "Only ASM"
      End
      Begin VB.Menu mnuOnlyC 
         Caption         =   "Only C"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuFileType 
         Caption         =   "File type"
         Begin VB.Menu set_mnuASM_def 
            Caption         =   "ASM"
         End
         Begin VB.Menu set_mnuC_def 
            Caption         =   "C"
         End
      End
      Begin VB.Menu mnuGenerate 
         Caption         =   "Generate"
         Begin VB.Menu set_mnuRunTime_def 
            Caption         =   "RunTime engine"
         End
         Begin VB.Menu set_mnuProcedures_def 
            Caption         =   "Procedures"
         End
      End
      Begin VB.Menu mnuScope 
         Caption         =   "Procedures scope"
         Begin VB.Menu set_mnuPublic_def 
            Caption         =   "Public"
         End
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuHeaders 
         Caption         =   "Headers"
      End
      Begin VB.Menu mnuEditRuntimeEngine 
         Caption         =   "RunTime engine"
      End
      Begin VB.Menu mnuProcedure 
         Caption         =   "Procedure"
      End
   End
End
Attribute VB_Name = "frmRunTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************
'*** ASM/C RunTime ***
'*********************

'08.10. 2004 - basic GUI and code
'09.10. 2004 - code improvement
'10.10. 2004 - supports even C language, better GUI
'16.10. 2004 - code improved
'10.12. 2004 - upgrade of code and gui, now it is a plugin for ThunVB


Private Const ENUM_ As String = "#ENUM#"        'members in enum
Private Const NUMBER_ As String = "#NUMBER#"    'dimension of the array of pointers
Private Const INIT_ As String = "#INIT#"        'place for init. code

Private Const VB_ As String = "#VBNAME#"        'VB name
Private Const RT_ As String = "#RTNAME#"        'RunTime name

'fill the array with pointers
Private Const INIT_POINTER As String = "apVB(eVB.#VBNAME#) = GetPointer(AddressOf #RTNAME#)"
Private Const INIT_IAT As String = "apIAT(eVB.#VBNAME#) = EnumIAT(App.hInstance, DLL, Enum2String(#VBNAME#))"

Private Const ENUM2VB_ As String = "#ENUM2VB#"  'convert enum to VB name
Private Const ENUM2RT_ As String = "#ENUM2RT#"  'convert enum to RunTime name

'convert enum to names
Private Const ENUM_VB As String = "Enum2String = Choose(eProcedure,#VBNAME#)"
Private Const ENUM_RT As String = "Enum2String = Choose(eProcedure,#RTNAME#)"

'unknow header
Private Const UNKNOWN_HEADER As String = "??? unknown header ???"

Private Const INCLUDE_ASM As String = "'#asm' include "  'include - asm
Private Const INCLUDE_C As String = "'#c' include "      'include - c

Private Enum eFileType
    ASM = 1
    C = 2
End Enum

Private Const SEPARATOR_FILE As String = "_"

Private Const RUNTIME_FILE As String = "modRunTimeEngine.bas"  'name of module where runtime is stored
Private Const RUNTIME_DIR As String = "runtime"               'name of directory where asm/c files are stored
Private Const RUNTIME_HEADERS As String = "headers.txt"

'----------------------
'--- CONTROL EVENTS ---
'----------------------

Private Sub Form_Load()
    
On Error Resume Next
    Me.Left = oThunVB.GetSettingGlobal(PLUGIN_NAMEs, "FormLeft", "0")
    Me.Top = oThunVB.GetSettingGlobal(PLUGIN_NAMEs, "FormTop", "0")
On Error GoTo 0
    
    frmRunTime.Caption = PLUGIN_NAME
    LogMsg "Loading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, "Form_Load", True, True
       
    LoadSettingsMenu GLOBAL_, Me
    LoadSettingsMenu LOCAL_, Me
    'init list view
    Call InitListView
    'filter files
    Call FilterFiles
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set default
    If Button = vbRightButton Then
        SetDefaultSettingsMenu GLOBAL_, Me
        SetDefaultSettingsMenu LOCAL_, Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingsMenu GLOBAL_, Me
    SaveSettingsMenu LOCAL_, Me
    
    oThunVB.SaveSettingGlobal PLUGIN_NAMEs, "FormLeft", Me.Left
    oThunVB.SaveSettingGlobal PLUGIN_NAMEs, "FormTop", Me.Top
    
    LogMsg "Unloading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, "Form_Unload", True, True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim sCode As String

    'get all code
    sCode = GenerateAllCode
    If Len(sCode) = 0 Then Exit Sub

    'show it
    frmViewer.ShowViewer PLUGIN_NAMEs, sCode

End Sub

'List View sorting
Private Sub lvwRT_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwRT
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

'------------
'--- MENU ---
'------------

'select all
Private Sub mnuAll_Click()
Dim i As Long
    For i = 1 To lvwRT.ListItems.Count
        lvwRT.ListItems.Item(i).Checked = True
    Next i
End Sub

'select no
Private Sub mnuNo_Click()
Dim i As Long
    For i = 1 To lvwRT.ListItems.Count
        lvwRT.ListItems.Item(i).Checked = False
    Next i
End Sub

'select only ASM files
Private Sub mnuOnlyASM_Click()
Dim i As Long
    With lvwRT.ListItems
        For i = 1 To .Count
            .Item(i).Checked = (.Item(i).ListSubItems(2).Text = Enum2String(ASM))
        Next i
    End With
End Sub

'select only C files
Private Sub mnuOnlyC_Click()
Dim i As Long
    With lvwRT.ListItems
        For i = 1 To .Count
            .Item(i).Checked = (.Item(i).ListSubItems(2).Text = Enum2String(C))
        Next i
    End With
End Sub

'file filter - ASM
Private Sub set_mnuASM_def_Click()
    set_mnuASM_def.Checked = Not set_mnuASM_def.Checked
    'reload list view
    Call FilterFiles
End Sub

'file filter - C
Private Sub set_mnuC_def_Click()
    set_mnuC_def.Checked = Not set_mnuC_def.Checked
    'reload list view
    Call FilterFiles
End Sub

'add procedures
Private Sub set_mnuProcedures_def_Click()
    set_mnuProcedures_def.Checked = Not set_mnuProcedures_def.Checked
End Sub

'add runtime
Private Sub set_mnuRunTime_def_Click()
    set_mnuRunTime_def.Checked = Not set_mnuRunTime_def.Checked
End Sub

'scope of procedures
Private Sub set_mnuPublic_def_Click()
    set_mnuPublic_def.Checked = Not set_mnuPublic_def.Checked
End Sub

'edit headers
Private Sub mnuHeaders_Click()
Dim sFile As String, sTemp As String

    'load file
    sFile = LoadFile(oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & RUNTIME_HEADERS)
    If Len(sFile) = 0 Then
        MsgBoxX "Error during reading headers file.", PLUGIN_NAMEs
        Exit Sub
    End If

    sTemp = frmViewer.ShowViewer(PLUGIN_NAMEs, sFile, False, False)
    If Len(sTemp) = 0 Or sTemp = sFile Then Exit Sub
    
    SaveFile oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & RUNTIME_HEADERS, sTemp
    
End Sub

'edit selected procedure
Private Sub mnuProcedure_Click()
Dim sFile As String, sTemp As String
    
    If lvwRT.ListItems.Count = 0 Then Exit Sub
    
    'load file
    sFile = LoadFile(oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & lvwRT.SelectedItem.Tag)
    If Len(sFile) = 0 Then
        MsgBoxX "Error during procedure file.", PLUGIN_NAMEs
        Exit Sub
    End If

    sTemp = frmViewer.ShowViewer(PLUGIN_NAMEs, sFile, False, False)
    If Len(sTemp) = 0 Or sTemp = sFile Then Exit Sub

    SaveFile oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & lvwRT.SelectedItem.Tag, sTemp
    
End Sub

'edit runtime engine
Private Sub mnuEditRuntimeEngine_Click()
Dim sFile As String, sTemp As String

    'load file
    sFile = LoadFile(oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & RUNTIME_FILE)
    If Len(sFile) = 0 Then
        MsgBoxX "Error during reading RunTime engine file.", PLUGIN_NAMEs
        Exit Sub
    End If

    sTemp = frmViewer.ShowViewer(PLUGIN_NAMEs, sFile, False, False)
    If Len(sTemp) = 0 Or sTemp = sFile Then Exit Sub
    
    SaveFile oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & RUNTIME_FILE, sTemp

End Sub


'------------------------
'--- HELPER FUNCTIONS ---
'------------------------

'fill list view
'parameters - eType - C or ASM files
Private Sub AddFilesToListView(eType As eFileType)
Dim sPath As String, sVB As String, sRT As String
Dim lPos As Long, sFileType As String

    'enum all files in the directory
    sFileType = "*." & Enum2String(eType)
    sPath = Dir(oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\*." & Enum2String(eType), vbNormal)

    Do While Len(sPath) <> 0

        'look for separator
        lPos = InStr(1, sPath, SEPARATOR_FILE, vbTextCompare)
        If lPos <> 0 Then

            'extract VB and RT name from file name
            sVB = Left(sPath, lPos - 1)
            sRT = Mid(sPath, lPos + 1, Len(sPath) - lPos - Len(sFileType) + 1)

            With lvwRT

                'add VB name
                .ListItems.Add , , sVB
                'add RT name
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , sRT
                'add type of language (C or ASM)
                .ListItems.Item(.ListItems.Count).ListSubItems.Add , , Enum2String(eType)

                'in Tag save file name
                .ListItems.Item(.ListItems.Count).Tag = sPath

            End With

        End If

        sPath = Dir

    Loop

End Sub

'enum to string
Private Function Enum2String(eType As eFileType)
    Enum2String = Choose(eType, "ASM", "C")
End Function

'return declare of RunTime function
'parameters - sRTName - run time name of function
Private Function GetHeader(sRTName As String) As String
Dim asHeaders() As String, i As Long

    'load file
    asHeaders = Split(LoadFile(oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & RUNTIME_HEADERS), vbCrLf)

On Error Resume Next
    i = LBound(asHeaders)
    If Err.Number <> 0 Then Exit Function
On Error GoTo 0

    'enum lines
    For i = LBound(asHeaders) To UBound(asHeaders)

        'try to find line where function is declared
        If InStr(1, asHeaders(i), sRTName, vbTextCompare) <> 0 Then

            'OK
            GetHeader = asHeaders(i)
            Exit Function

        End If

    Next i

End Function

'init. list view
Private Sub InitListView()

    With lvwRT

        'clear all
        .ColumnHeaders.Clear
        .ListItems.Clear

        .View = lvwReport
        .LabelEdit = lvwManual
        .Checkboxes = True
        .HideSelection = False

        'add columns
        .ColumnHeaders.Add 1, , "VB name", .Width * 1 / 3
        .ColumnHeaders.Add 2, , "RunTime name", .Width * 1 / 3
        .ColumnHeaders.Add 3, , "Language", .Width * 1 / 3 - 60

    End With

End Sub

'create basic code for hooking engine
Private Function CreateBasicCode() As String
Dim sMod As String, i As Long, sEnum As String, lCount As Long, sInit As String
Dim sEnum2VB As String, sEnum2RT As String

    'load bas file
    sMod = LoadFile(oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & RUNTIME_FILE)
    If Len(sMod) = 0 Then Exit Function

    lCount = 0

    With lvwRT.ListItems
        For i = 1 To .Count
            If .Item(i).Checked = True Then

                'create enum
                lCount = lCount + 1
                sEnum = sEnum & vbCrLf & Space(4) & .Item(i).Text & "_" & .Item(i).ListSubItems(2).Text & " = " & lCount

                'create init (fill array with pointers)
                sInit = sInit & vbCrLf & Space(4) & INIT_POINTER & vbCrLf & Space(4) & INIT_IAT & vbCrLf
                sInit = Replace(sInit, VB_, .Item(i).Text & "_" & .Item(i).ListSubItems(2).Text)
                sInit = Replace(sInit, RT_, .Item(i).ListSubItems(1).Text)

                'create enum2string function
                sEnum2VB = sEnum2VB & ", " & .Item(i).Text & "_" & .Item(i).ListSubItems(2).Text
                sEnum2RT = sEnum2RT & ", " & .Item(i).ListSubItems(1).Text

            End If
        Next i
    End With

    If lCount = 0 Then Exit Function

    'add enum
    sMod = Replace(sMod, ENUM_, Mid(sEnum, Len(vbCrLf & Space(4)) + 1), , 1, vbTextCompare)
    'add constant
    sMod = Replace(sMod, NUMBER_, lCount, , 1, vbTextCompare)
    'add init
    sMod = Replace(sMod, INIT_, Mid(sInit, Len(vbCrLf & Space(4)) + 1), , 1)

    'path enum to string line
    sEnum2VB = Replace(ENUM_VB, VB_, Mid(sEnum2VB, Len(" ,")), , 1)
    sEnum2RT = Replace(ENUM_RT, RT_, Mid(sEnum2RT, Len(" ,")), , 1)

    'add enum 2 string
    sMod = Replace(sMod, ENUM2VB_, sEnum2VB, , 1)
    sMod = Replace(sMod, ENUM2RT_, sEnum2RT, , 1)

    CreateBasicCode = sMod

End Function

'create procedures for hooking
Private Function CreateHookProcedures(bPublic As Boolean) As String
Dim i As Long, sCode As String, sHeader As String
Dim lPos As Long, sNewHeader As String, bSub As Boolean

    With lvwRT.ListItems
        For i = 1 To .Count
            If .Item(i).Checked = True Then

                'get declare of function
                sHeader = GetHeader(.Item(i).ListSubItems(1).Text)

                If Len(sHeader) = 0 Then
                    'declare was not found - use "unknown" header
NOHEADER:
                    sNewHeader = UNKNOWN_HEADER & vbCrLf & IIf(.Item(i).ListSubItems(2).Text = Enum2String(ASM), INCLUDE_ASM, INCLUDE_C) & oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & .Item(i).Tag & vbCrLf & UNKNOWN_HEADER
                Else

                    'add public/name
                    sNewHeader = IIf(bPublic = True, "Public", "Private") & " Function " & .Item(i).ListSubItems(1).Text & "_" & .Item(i).ListSubItems(2).Text

                    'add parameters
                    lPos = InStr(1, sHeader, "(")
                    If lPos <> 0 Then sNewHeader = sNewHeader & Mid(sHeader, lPos) Else GoTo NOHEADER

                    'add "include" line
                    If .Item(i).ListSubItems(2).Text = Enum2String(ASM) Then sNewHeader = sNewHeader & vbCrLf & INCLUDE_ASM & oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & .Item(i).Tag Else sNewHeader = sNewHeader & vbCrLf & INCLUDE_C & oThunVB.GetThunderVBPluginsPath & RUNTIME_DIR & "\" & .Item(i).Tag

                    'add end function
                    sNewHeader = sNewHeader & vbCrLf & "End Function"

                    'if it is Sub replace string "Function" with string "Sub"
                    If InStr(1, sHeader, " Sub ", vbTextCompare) <> 0 Then sNewHeader = Replace(sNewHeader, " Function", " Sub")

                End If

                sCode = sCode & vbCrLf & vbCrLf & sNewHeader

            End If
        Next i
    End With

    'trim vbcrlf
    CreateHookProcedures = Mid(sCode, Len(vbCrLf & vbCrLf) + 1)

End Function

'generate all (engine, procedures) code
Private Function GenerateAllCode() As String

    'generate engine code
    If set_mnuRunTime_def.Checked = True Then GenerateAllCode = CreateBasicCode
    'generate hooks
    If set_mnuProcedures_def.Checked = True Then GenerateAllCode = GenerateAllCode & vbCrLf & vbCrLf & CreateHookProcedures(set_mnuPublic_def.Checked)

End Function

Private Sub FilterFiles()
    
    'init listview
    Call InitListView
    
    'add to the listview C files
    If set_mnuC_def.Checked = True Then Call AddFilesToListView(C)
    'add to the listview ASM files
    If set_mnuASM_def.Checked = True Then AddFilesToListView (ASM)
    
End Sub
