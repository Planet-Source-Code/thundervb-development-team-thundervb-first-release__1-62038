VERSION 5.00
Begin VB.Form frmCodeExplorer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASM Explorer"
   ClientHeight    =   3000
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin CodeExplorer.MsFlexGridEdit ctlFG 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3201
   End
   Begin VB.PictureBox picBut 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   4080
      Picture         =   "frmCodeExplorer.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbMod 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      Left            =   5400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ComboBox cmbProc 
      Height          =   315
      Left            =   2760
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      Caption         =   "Module"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   630
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Procedure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   885
   End
   Begin VB.Menu mnuRestore 
      Caption         =   "&Restore"
      Begin VB.Menu mnuRestoreItem 
         Caption         =   "Restore Item"
      End
      Begin VB.Menu mnuRestoreAll 
         Caption         =   "Restore All"
      End
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save"
      Begin VB.Menu mnuSaveItem 
         Caption         =   "Save Item"
      End
      Begin VB.Menu mnuSaveAll 
         Caption         =   "Save All"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFind 
         Caption         =   "Find all"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace all"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
         Begin VB.Menu mnuClearItems 
            Caption         =   "All ASM code"
            Index           =   1
         End
         Begin VB.Menu mnuClearItems 
            Caption         =   "All comments"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "frmCodeExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************
'*** ASM/C Code Explorer ***
'***************************

'plugin for ThunderVB
'   by Libor & drkIIraziel

'13.09. 2004 - initial version, GUI
'18.09. 2004 - better GUI, code for filter
'19.09. 2004 - better filter and code, filter for labels
'14.10. 2004 - implemented filtering code, added new filter "All ASM code"
'10.12. 2004 - upgrade od code and gui, now it is a plugin for ThunVB
'24.12. 2004 - added editable flexgrid
'30.12. 2004 - code restoring and saving, better editing
'01.01. 2005 - upgrade of flexgrid. menu Edit

'beginning of the code
'*********************
Private Const F_LOCAL As String = "local"
Private Const F_RET As String = "ret"
Private Const F_INT As String = "int"
Private Const F_CALL As String = "call"

'middle of the code
'******************
Private Const F_EQU As String = " equ "
Private Const F_DD As String = " dd "
Private Const F_DW As String = " dw "
Private Const F_DB As String = " db "

'end of the code
'***************
Private Const F_LABEL = ":"

Private Const COMMENT As String = ";"

'filter type
Private Enum eFilterType
    LOCAL_ = 1
    RET_ = 2
    Label_ = 3
    EQU_ = 4
    DWORD_ = 5
    BYTE_ = 6
    WORD_ = 7
    CALL_ = 8
    INT_ = 9
    All_ = 10
End Enum

'position of filter
Private Enum eCode
    code_beginning
    code_middle
    code_end
    code_all
End Enum

Private Enum eGetAsmInfo
    ASM_Code
    ASM_Comment
End Enum

'structure for storing info about 1 line of asm code
Private Type tInfoASM
    lLine As Long        'number of line
    sASM As String       'asm code
    sComment As String   'comment
    sOriginal As String  'real line of code
End Type

Private Const ASM_PREFIX As String = "'#asm'"
Dim ASM_NOTHING As String

Private Const VBP_KEY As String = "module"
Private Const SEPARATOR As String = ";"

Private bLoaded As Boolean

Private WithEvents fg As MSFlexGrid
Attribute fg.VB_VarHelpID = -1

Dim atInfoASM() As tInfoASM

'--------------
'--- EVENTS ---
'--------------

Private Sub fg_DblClick()
    'check selected item
    If fg.Rows <= 1 Or cmbProc.ListIndex = -1 Or cmbMod.ListIndex = -1 Or cmbFilter.ListIndex = -1 Then Exit Sub
    'set current line
    With fg
        If .MouseCol = 0 And .MouseRow >= 1 Then Call SetCurLine(cmbMod.Text, cmbProc.Text, .TextMatrix(.MouseRow, .MouseCol))
    End With
End Sub


Private Sub cmbFilter_Click()

    'check comboboxes
    If bLoaded = False Or cmbMod.ListIndex = -1 Or cmbProc.ListIndex = -1 Or cmbFilter.ListIndex = -1 Then Exit Sub
    'init flexgrid
    Call InitFlexGrid
    
    'select filter
    Select Case cmbFilter.ItemData(cmbFilter.ListIndex)
        Case eFilterType.LOCAL_
            Call AddToFlexGrid(Filter(LOCAL_, code_beginning))
        Case eFilterType.CALL_
            Call AddToFlexGrid(Filter(CALL_, code_beginning))
        Case eFilterType.RET_
            Call AddToFlexGrid(Filter(RET_, code_beginning))
        Case eFilterType.INT_
            Call AddToFlexGrid(Filter(INT_, code_beginning))
        Case eFilterType.EQU_
            Call AddToFlexGrid(Filter(EQU_, code_middle))
        Case eFilterType.BYTE_
            Call AddToFlexGrid(Filter(BYTE_, code_middle))
        Case eFilterType.DWORD_
            Call AddToFlexGrid(Filter(DWORD_, code_middle))
        Case eFilterType.WORD_
            Call AddToFlexGrid(Filter(WORD_, code_middle))
        Case eFilterType.Label_
            Call AddToFlexGrid(Filter(Label_, code_end))
        Case eFilterType.All_
            Call AddToFlexGrid(Filter(All_, code_all))
    End Select
    
End Sub

'combobox - list of all Modules
Private Sub cmbMod_Click()
    Call InitProc         'get all procedures in Module
    Call cmbFilter_Click  'init filter
End Sub

'combobox - list of all Procedures
Private Sub cmbProc_Click()
    Call cmbFilter_Click   'init filter
    cmbFilter.SetFocus
End Sub

Private Sub Form_Activate()
Dim sData As String, asSel() As String

    'init all
    Call InitMod        'modules
    Call InitProc       'procedures
    Call InitFlexGrid   'flexgrid
    Call InitFilter     'filter
    
    'set last selected filter
    sData = oThunVB.GetSettingProject(PLUGIN_NAME, VBP_KEY, "")
    If Len(sData) <> 0 Then
On Error Resume Next
        asSel = Split(sData, SEPARATOR)
        cmbMod.Text = asSel(0)     'module
        cmbProc.Text = asSel(1)    'procedure
        cmbFilter.Text = asSel(2)  'filter
On Error GoTo 0
    End If
        
    cmbFilter.SetFocus
    
End Sub

'initialize
Private Sub Form_Load()
    
On Error Resume Next
    Me.Left = oThunVB.GetSettingGlobal(PLUGIN_NAMEs, "FormLeft", "0")
    Me.Top = oThunVB.GetSettingGlobal(PLUGIN_NAMEs, "FormTop", "0")
On Error GoTo 0
    
    Me.Caption = PLUGIN_NAME
    ASM_NOTHING = "<nothing>" & Chr(160)  'chr(160) is invisible char
    LogMsg "Loading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, "Form_Load", True, True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oThunVB.SaveSettingProject PLUGIN_NAME, VBP_KEY, cmbMod.Text & SEPARATOR & cmbProc.Text & SEPARATOR & cmbFilter.Text
    oThunVB.SaveSettingGlobal PLUGIN_NAMEs, "FormLeft", Me.Left
    oThunVB.SaveSettingGlobal PLUGIN_NAMEs, "FormTop", Me.Top
    LogMsg "Unloading " & Add34(Me.Caption) & " window", PLUGIN_NAMEs, "Form_Load", True, True
End Sub

'user is going to edit text in a cell
Private Sub ctlFG_BeforeTextEdit(FlexGrid As MSFlexGridLib.MSFlexGrid, CellText As String, DefaultEditText As String, bEdit As Boolean, ByVal eReason As eSTART_EDIT)
    'there is a no text in a cell, clear it
    If CellText = ASM_NOTHING And FlexGrid.CellForeColor = vbRed Then DefaultEditText = ""
End Sub

'user finished editing the cell
Private Sub ctlFG_AfterTextEdit(FlexGrid As MSFlexGridLib.MSFlexGrid, ByVal CellText As String, NewCellText As String, bContinue As Boolean, ByVal eReason As eSTOP_EDIT)
    
    NewCellText = Trim(NewCellText)

    'if uses changes asm code or comment then change forecolor in the cell
    Select Case FlexGrid.col
        Case 1
            If NewCellText <> atInfoASM(FlexGrid.Row).sASM Then FlexGrid.CellForeColor = vbRed Else FlexGrid.CellForeColor = vbBlack
        Case 2
            If NewCellText <> atInfoASM(FlexGrid.Row).sComment Then FlexGrid.CellForeColor = vbRed Else FlexGrid.CellForeColor = vbBlack
    End Select

    'add info that there is no text in a cell
    If Len(NewCellText) = 0 And FlexGrid.CellForeColor = vbRed Then NewCellText = ASM_NOTHING
    
End Sub

'--- menu ---

'restore all changed text (comment+asm code)
Private Sub mnuRestoreAll_Click()
Dim i As Long, lPrevCol As Long, lPrevRow As Long

    'save selected cell
    lPrevCol = fg.col
    lPrevRow = fg.Row

    'check number of rows
    If fg.Rows <= 1 Then Exit Sub
    
    'iterate throw all rows
    For i = 1 To fg.Rows - 1

        fg.Row = i
        
        'first column (asm code)
        fg.col = 1
        If fg.CellForeColor = vbRed Then
            fg.Text = atInfoASM(fg.Row).sASM
            fg.CellForeColor = vbBlack
        End If
        
        'second column (comment)
        fg.col = 2
        If fg.CellForeColor = vbRed Then
            fg.Text = atInfoASM(fg.Row).sComment
            fg.CellForeColor = vbBlack
        End If

    Next i
    
    'restore selected cell
    fg.col = lPrevCol
    fg.Row = lPrevRow
    
End Sub

'restore selected item in the flexgrid
Private Sub mnuRestoreItem_Click()
    
    'check selected item
    If fg.col >= 1 And fg.Row >= 1 And fg.CellForeColor = vbRed Then
        If fg.col = 1 Then
            fg.Text = atInfoASM(fg.Row).sASM     'asm code
        ElseIf fg.col = 2 Then
            fg.Text = atInfoASM(fg.Row).sComment 'comment
        End If
        fg.CellForeColor = vbBlack               'set new color
    End If
    
End Sub

'save all changed items
Private Sub mnuSaveAll_Click()
Dim i As Long, lPrevCol As Long, lPrevRow As Long

    'save selected cell
    lPrevCol = fg.col
    lPrevRow = fg.Row

    'check number of rows
    If fg.Rows <= 1 Then Exit Sub
    
    'iterate throw all rows
    For i = 1 To fg.Rows - 1

        fg.Row = i
        
        'first column (asm code)
        fg.col = 1
        If fg.CellForeColor = vbRed Then
            'clear text
            If fg.Text = ASM_NOTHING Then fg.Text = ""
            ChangeLineAndUpdateAsmInfo ASM_Code, atInfoASM(fg.Row)
            fg.CellForeColor = vbBlack
        End If
        
        'second column (comment)
        fg.col = 2
        If fg.CellForeColor = vbRed Then
            'clear text
            If fg.Text = ASM_NOTHING Then fg.Text = ""
            ChangeLineAndUpdateAsmInfo ASM_Comment, atInfoASM(fg.Row)
            fg.CellForeColor = vbBlack
        End If

    Next i
    
    'restore selected cell
    fg.col = lPrevCol
    fg.Row = lPrevRow
    
End Sub

'save selected item
Private Sub mnuSaveItem_Click()
    
    If fg.col >= 1 And fg.Row >= 1 And fg.CellForeColor = vbRed Then
        
        'asm code
        If fg.col = 1 Then
            ChangeLineAndUpdateAsmInfo ASM_Code, atInfoASM(fg.Row)
        'comment
        ElseIf fg.col = 2 Then
            ChangeLineAndUpdateAsmInfo ASM_Comment, atInfoASM(fg.Row)
        End If
        
        fg.CellForeColor = vbBlack
        
    End If
    
End Sub

'find some text
Private Sub mnuFind_Click()
Dim sFind As String, sOut As String
Dim i As Long

    If fg.Rows <= 1 Then Exit Sub

    sFind = Trim(InputBox("Find what :", PLUGIN_NAMEs))
    If Len(sFind) = 0 Then Exit Sub
    
    For i = LBound(atInfoASM) To UBound(atInfoASM)
        If InStr(atInfoASM(i).sOriginal, sFind) <> 0 Then sOut = sOut & ", " & atInfoASM(i).lLine
    Next i
    
    MsgBoxX "Text " & Add34(sFind) & " was found on lines - " & Mid(sOut, 3), PLUGIN_NAMEs
    
End Sub

'replace
Private Sub mnuReplace_Click()
Dim sWhat As String, sWith As String
Dim i As Long, lCol As String, lRow As Long
Dim sItem As String, j As Long

    If fg.Rows <= 1 Then Exit Sub

    sWhat = Trim(InputBox("Replace :", PLUGIN_NAMEs))
    sWith = Trim(InputBox("With :", PLUGIN_NAMEs))
    
    lRow = fg.Row
    lCol = fg.col
    
    For i = LBound(atInfoASM) To UBound(atInfoASM)
            
            'try to replace text
            sItem = Trim(Replace(atInfoASM(i).sASM, sWhat, sWith))
            'check it
            If sItem <> atInfoASM(i).sASM Then
                fg.Row = i: fg.col = 1
                If Len(sItem) <> 0 Then fg.Text = sItem Else fg.Text = ASM_NOTHING
                fg.CellForeColor = vbRed
            End If
            
            'try to replace text
            sItem = Trim(Replace(atInfoASM(i).sComment, sWhat, sWith))
            'check it
            If sItem <> atInfoASM(i).sComment Then
                fg.Row = i: fg.col = 2
                If Len(sItem) <> 0 Then fg.Text = sItem Else fg.Text = ASM_NOTHING
                fg.CellForeColor = vbRed
            End If
        
    Next i
    
    fg.col = lCol
    fg.Row = lRow
    
End Sub

'clear - asm code/comments
' index - 1 - all asm code
'       - 2 - all comments
Private Sub mnuClearItems_Click(Index As Integer)
Dim lRow As Long, lCol As Long, i As Long

    If fg.Rows <= 1 Then Exit Sub
    
    lRow = fg.Row
    lCol = fg.col
    
    fg.col = Index
    
    For i = 1 To fg.Rows - 1
        fg.Row = i
        fg.Text = ASM_NOTHING
        fg.CellForeColor = vbRed
    Next i
    
    fg.Row = lRow
    fg.col = lCol

End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

'------------------------
'--- HELPER FUNCTIONS ---
'------------------------

'get info from line of code
'parameters - eInfo - type of info
'           - sAsmLine - line of code
'return     - string

Private Function GetInfo(eInfo As eGetAsmInfo, sAsmLine As String) As String
Dim lComment As Long

    'look for comment
    lComment = InStr(sAsmLine, COMMENT)
    
    'choose type of info
    If eInfo = ASM_Code Then
        
        If lComment = 0 Then
            'no comment
            GetInfo = sAsmLine
        Else
            'extract line of code
            GetInfo = Left(sAsmLine, lComment - 1)
        End If
    
    Else

        If lComment = 0 Then
            'no comment
            GetInfo = ""
        Else
            'extract comment
            GetInfo = Right(sAsmLine, Len(sAsmLine) - lComment)
        End If
        
    End If

    GetInfo = Trim(GetInfo)

End Function

'add tInfoASM structures to flexgrid
'parameters - tInfoASM() - array of tInfoASM structures

Private Sub AddToFlexGrid(tInfoASM() As tInfoASM)
Dim i As Long
    
    'clear array
    Erase atInfoASM
    
    With fg
    
        'clear flexgrid
        .Rows = 1
        
        'check array
On Error Resume Next
        i = LBound(tInfoASM)
        If Err.Number <> 0 Then Exit Sub
On Error GoTo 0
    
        .Rows = UBound(tInfoASM) + 1
    
        'add items
        For i = LBound(tInfoASM) To UBound(tInfoASM)
            .TextMatrix(i, 0) = tInfoASM(i).lLine
            .TextMatrix(i, 1) = tInfoASM(i).sASM
            .TextMatrix(i, 2) = tInfoASM(i).sComment
        Next i

    End With
    
    'save it in public array
    atInfoASM = tInfoASM

End Sub

'get all asm lines that suit selected filter
'parameters - eFilter - type of filter
'           - eType - type of code
'return - array of tInfoASM structures

Private Function Filter(eFilter As eFilterType, eType As eCode) As tInfoASM()
Dim i As Long, sLine As String, lCount As Long, sFind As String
Dim bContinue As Boolean, atASM() As tInfoASM
Dim asLines() As String, sOrigLine As String

    'check items
    If cmbMod.ListIndex = -1 Or cmbProc.ListIndex = -1 Then Exit Function
    
    'get code of function and split it
    asLines = Split(GetFunctionCode(cmbMod.Text, cmbProc.Text), vbCrLf)
    'check the array
    On Error Resume Next
        i = LBound(asLines)
        If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    lCount = 0

    'choose filter
    Select Case eFilter
        Case eFilterType.CALL_
            sFind = F_CALL
        Case eFilterType.INT_
            sFind = F_INT
        Case eFilterType.LOCAL_
            sFind = F_LOCAL
        Case eFilterType.RET_
            sFind = F_RET
        Case eFilterType.EQU_
            sFind = F_EQU
        Case eFilterType.BYTE_
            sFind = F_DB
        Case eFilterType.DWORD_
            sFind = F_DD
        Case eFilterType.WORD_
            sFind = F_DW
        Case eFilterType.Label_
            sFind = F_LABEL
        Case eFilterType.All_
            sFind = "*."
    End Select
    
    'check filter
    If sFind = "" Then Exit Function
    
    'enum all lines
    For i = LBound(asLines) To UBound(asLines)
    
        sLine = Trim(asLines(i))
        'is it line with ASM code?
        If StrComp(Left(sLine, Len(ASM_PREFIX)), ASM_PREFIX, vbTextCompare) <> 0 Then GoTo 10
            
        'trim ASM prefix
        sOrigLine = sLine
        sLine = Trim(Mid(sLine, Len(ASM_PREFIX) + 1))
        bContinue = False
        
        'type of code
        Select Case eType
            Case eCode.code_beginning
                If StrComp(Left(sLine, Len(sFind)), sFind, vbTextCompare) = 0 Then bContinue = True
            Case eCode.code_middle
                If InStr(1, sLine, sFind, vbTextCompare) > 0 Then bContinue = True
            Case eCode.code_end
                If StrComp(Right(GetInfo(ASM_Code, sLine), Len(sFind)), sFind, vbTextCompare) = 0 Then bContinue = True
            Case eCode.code_all
                bContinue = True
        End Select
            
        If bContinue = True Then

            'adjust array
            lCount = lCount + 1
            ReDim Preserve atASM(1 To lCount)
        
            'save info
            With atASM(lCount)
                .lLine = i
                .sASM = GetInfo(ASM_Code, sLine)
                .sComment = GetInfo(ASM_Comment, sLine)
                .sOriginal = sOrigLine
            End With
        
        End If

10:
    Next i
    
    'return
    Filter = atASM

End Function

'init flexgrid
Private Sub InitFlexGrid()
Dim i As Long

    Set fg = ctlFG.GetFlexGrid
    
    With fg
    
        .GridLines = flexGridNone
        .AllowUserResizing = flexResizeNone
        .ScrollBars = flexScrollBarVertical
        .Appearance = flexFlat
    
        .Cols = 3
        .Rows = 1
    
        'headers
        .TextMatrix(0, 0) = "line"
        .TextMatrix(0, 1) = "ASM code"
        .TextMatrix(0, 2) = "comment"
    
        'size of columns
        .ColWidth(0) = .Width / 15 * 1
        .ColWidth(1) = .Width / 15 * 7
        .ColWidth(2) = .Width / 15 * 7
    
        'adjust headers
        For i = 0 To 2

            .ColAlignment(i) = flexAlignLeftCenter
            .col = i
            .Row = 0
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
            
        Next i

    End With
    
End Sub

'initialize filter (combobox)
Private Sub InitFilter()
    
    bLoaded = False
    
    'add filters
    With cmbFilter
    
        .Clear
    
        .AddItem Enum2String(LOCAL_)
        .ItemData(.NewIndex) = eFilterType.LOCAL_
        
        .AddItem Enum2String(INT_)
        .ItemData(.NewIndex) = eFilterType.INT_
        
        .AddItem Enum2String(RET_)
        .ItemData(.NewIndex) = eFilterType.RET_
        
        .AddItem Enum2String(Label_)
        .ItemData(.NewIndex) = eFilterType.Label_
        
        .AddItem Enum2String(EQU_)
        .ItemData(.NewIndex) = eFilterType.EQU_
        
        .AddItem Enum2String(DWORD_)
        .ItemData(.NewIndex) = eFilterType.DWORD_
        
        .AddItem Enum2String(BYTE_)
        .ItemData(.NewIndex) = eFilterType.BYTE_
        
        .AddItem Enum2String(BYTE_)
        .ItemData(.NewIndex) = eFilterType.WORD_
        
        .AddItem Enum2String(CALL_)
        .ItemData(.NewIndex) = eFilterType.CALL_
        
        .AddItem Enum2String(All_)
        .ItemData(.NewIndex) = eFilterType.All_
        
        .ListIndex = 0
        
    End With
    
    bLoaded = True
    
End Sub

'init name of modules (combobox)
Private Sub InitMod()
Dim i As Long, asMod() As String

    'init
    bLoaded = False
    cmbMod.Clear
    
    'get list of all modules
    asMod = EnumModuleNames()
    
    'check array
On Error Resume Next
    i = LBound(asMod)
    If Err.Number <> 0 Then Exit Sub
On Error GoTo 0

    'add names
    For i = LBound(asMod) To UBound(asMod)
        cmbMod.AddItem asMod(i)
    Next i

    'exit
    cmbMod.ListIndex = 0
    bLoaded = True

End Sub

'init name of procedures (combobox)
Private Sub InitProc()
Dim i As Long, asProc() As String
    
    'init
    bLoaded = False
    cmbProc.Clear
    
    'check combo and get list of all functions in module
    If cmbMod.ListIndex = -1 Then Exit Sub
    asProc = EnumFunctionNames(cmbMod.Text)
    
    'check array
On Error Resume Next
    i = LBound(asProc)
    If Err.Number <> 0 Then Exit Sub
On Error GoTo 0
    
    'add names
    For i = LBound(asProc) To UBound(asProc)
        cmbProc.AddItem asProc(i)
    Next i

    'exit
    cmbProc.ListIndex = 0
    bLoaded = True

End Sub

'convert enum to string
Private Function Enum2String(Filter As eFilterType) As String
    Enum2String = Choose(Filter, "Local variables (LOCAL)", "RET", "Labels", "Constants (EQU)", "Variables - DWORD", "Variables - BYTE", "Variables - WORD", "CALL", "Interrupts (INT)", "All ASM code")
End Function

'change line of code in a codewindow
' eType   - asm code or comment
' tInfAsm - info about asm line
Private Sub ChangeLineAndUpdateAsmInfo(eType As eGetAsmInfo, tInfAsm As tInfoASM)
Dim sNewLine As String
        
    With tInfAsm
        Select Case eType
            Case eGetAsmInfo.ASM_Code
                
                'create new line
                If Len(fg.Text) = 0 And Len(.sASM) = 0 Then
                    Exit Sub
                ElseIf Len(.sASM) <> 0 And Len(.sComment) <> 0 Then
                    sNewLine = Replace(.sOriginal, .sASM, fg.Text, , 1)
                ElseIf Len(.sASM) = 0 And Len(.sComment) = 0 Then
                    sNewLine = .sOriginal & " " & fg.Text
                ElseIf Len(.sASM) <> 0 And Len(.sComment) = 0 Then
                    sNewLine = Replace(.sOriginal, .sASM, fg.Text, , 1)
                ElseIf Len(.sASM) = 0 And Len(.sComment) <> 0 Then
                    sNewLine = Left(.sOriginal, Len(ASM_PREFIX)) & " " & fg.Text & " ;" & .sComment
                End If
                
                'update asm info
                .sASM = fg.Text
                
            Case eGetAsmInfo.ASM_Comment
                
                'create new line
                If Len(fg.Text) = 0 And Len(.sComment) = 0 Then
                    Exit Sub
                ElseIf Len(fg.Text) = 0 And Len(.sComment) <> 0 Then
                    sNewLine = Trim(Left(.sOriginal, Len(.sOriginal) - Len(.sComment) - 1))
                ElseIf Len(.sASM) <> 0 And Len(.sComment) <> 0 Then
                    sNewLine = Replace(.sOriginal, .sComment, fg.Text, , 1)
                ElseIf Len(.sASM) = 0 And Len(.sComment) = 0 Then
                    sNewLine = .sOriginal & " ;" & fg.Text
                ElseIf Len(.sASM) <> 0 And Len(.sComment) = 0 Then
                    sNewLine = .sOriginal & " ;" & fg.Text
                ElseIf Len(.sASM) = 0 And Len(.sComment) <> 0 Then
                    sNewLine = Replace(.sOriginal, .sComment, fg.Text, , 1)
                End If
                
                'update comment info
                .sComment = fg.Text
        
        End Select
        
        'replace line
        oThunVB.SetFunctionLine cmbMod.Text, cmbProc.Text, .lLine, sNewLine, True
        .sOriginal = sNewLine
        
    End With
    
End Sub
