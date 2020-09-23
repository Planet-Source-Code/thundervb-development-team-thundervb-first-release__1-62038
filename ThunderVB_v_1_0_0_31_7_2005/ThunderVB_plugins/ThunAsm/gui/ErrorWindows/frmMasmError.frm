VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmMasmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asm Error"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1305
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   7575
   End
   Begin VB.CommandButton cmd_can 
      Caption         =   "Cancel Compile"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdIng 
      Caption         =   "Ingore"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Retry"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txtMasmOut 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMasmError.frx":0000
      Top             =   120
      Width           =   7575
   End
   Begin CodeSenseCtl.CodeSense txtErr 
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "frmMasmError.frx":0014
      TabIndex        =   5
      Top             =   2880
      Width           =   7575
   End
End
Attribute VB_Name = "frmMasmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Form Created , intial version


Dim hf As Long

Private Sub cmd_can_Click()

    hf = 1
    Cancel_compile = True
    Me.Hide

End Sub

Private Sub cmdIng_Click()

    hf = 1
    Me.Hide

End Sub

Private Sub cmdRetry_Click()

    hf = 2
    Me.Hide

End Sub

Private Sub Form_Initialize()
    'init com controls peharps?? (for xp styles)
    'init sciedit
    Dim i As Long
    hf = 0
    Dim lang As language
    Dim glb As Globals
    Set glb = New Globals
    'txtErr.Lexer = asm

    
    'txtErr.MarginWidth(0) = 35
    'txtErr.MarginWidth(1) = 0
    'txtErr.MarginType(0) = SC_MARGIN_NUMBER
    txtErr.LineNumbering = True
    txtErr.LineNumberStart = 1
    txtErr.SetColor cmClrLineNumber, RGB(255, 255, 255)
    txtErr.SetColor cmClrLineNumberBk, RGB(120, 120, 120)
    'txtErr.StyleClearAll
     
    'WordList &cpuInstruction = *keywordlists[0];
    'WordList &mathInstruction = *keywordlists[1];
    'WordList &registers = *keywordlists[2];
    'WordList &directive = *keywordlists[3];
    'WordList &directiveOperand = *keywordlists[4];
    'WordList &extInstruction = *keywordlists[5];

  '  txtErr.KEYWORDS(0) = LCase("EAX EBX ECX EDX")
  '  txtErr.KEYWORDS(1) = LCase("AX BX CX DX")
  '  txtErr.KEYWORDS(2) = LCase("AH AL BH BL CH CL DH DL")
  '  txtErr.KEYWORDS(3) = LCase("CS DS ES FS GS SS")
  '  txtErr.KEYWORDS(4) = LCase("ESI EDI EBP EIP ESP")
  '  txtErr.KEYWORDS(5) = LCase("EFLAGS")
    Set lang = glb.GetLanguageDef("basic")
    
    lang.CaseSensitive = False
    lang.SingleLineComments = ";"
    lang.Keywords = "EAX" & vbLf & "EBX" & vbLf & "ECX" & vbLf & "EDX" & vbLf & "ESI" & vbLf & "EDI" & vbLf & "EBP" & vbLf & "EIP" & vbLf & "ESP" & vbLf & "EFLAGS"
    lang.Operators = "AX" & vbLf & "BX" & vbLf & "CX" & vbLf & "DX" & vbLf & "SI" & vbLf & "DI" & vbLf & _
                     "AH" & vbLf & "AL" & vbLf & "BH" & vbLf & "BL" & vbLf & "CH" & vbLf & "CL" & vbLf & "DH" & vbLf & "DL"
    lang.StringDelims = "'" & vbLf & """"
    
    lang.ScopeKeywords1 = ""
    lang.ScopeKeywords2 = ""
    'glb.UnregisterLanguage "_asm_x86_"
    On Error Resume Next
    glb.RegisterLanguage "_asm_x86_", lang
    
    txtErr.ColorSyntax = True
    txtErr.language = "_asm_x86_"
    
    Dim defCol As Long
    defCol = GetAsmWordColor("*default*")
    
    txtErr.SetColor cmClrText, defCol
        
    'standar
    'txtErr.StyleSetFore SCE_ASM_NUMBER, GetAsmWordColor("1234")
    txtErr.SetColor cmClrNumber, GetAsmWordColor("1234")
    'txtErr.StyleSetFore SCE_ASM_STRING, GetAsmWordColor("'this is a string'")
    txtErr.SetColor cmClrString, GetAsmWordColor("'this is a string'")
    'txtErr.StyleSetFore SCE_ASM_COMMENT, GetAsmWordColor(";")
    txtErr.SetColor cmClrComment, GetAsmWordColor(";")
    
    'remaped
    '
    'txtErr.StyleSetFore SCE_ASM_CPUINSTRUCTION, GetAsmWordColor("eax")
    'txtErr.StyleSetFore SCE_ASM_MATHINSTRUCTION, GetAsmWordColor("ax")
    'txtErr.StyleSetFore SCE_ASM_REGISTER, GetAsmWordColor("ah")
    'txtErr.StyleSetFore SCE_ASM_DIRECTIVE, GetAsmWordColor("cs")
    'txtErr.StyleSetFore SCE_ASM_DIRECTIVEOPERAND, GetAsmWordColor("esi")
    'txtErr.StyleSetFore SCE_ASM_EXTINSTRUCTION, GetAsmWordColor("EFLAGS")
    txtErr.SetColor cmClrKeyword, GetAsmWordColor("eax")
    txtErr.SetColor cmClrOperator, GetAsmWordColor("ax")

End Sub

Public Sub ShowError(masmErr As String, errText As String, Optional title As String = "Error on Asm code:")
    Dim t As Long, ln As String, st() As String, ErrLn As Long
          
    Form_Initialize
    
    hf = 0
    Me.caption = title
    Me.txtMasmOut = masmErr
    Me.txtErr.Text = errText
    Me.Show ' vbModal
          
    List1.Clear
    For t = 0 To UBound(Split(masmErr, vbNewLine))
        ln = Split(masmErr, vbNewLine)(t)
        st = Split(ln, ") : error ")
        If UBound(st) = 1 Then
        ErrLn = Mid$(st(0), InStrRev(st(0), "(") + 1)
        List1.AddItem "Line : " & ErrLn & " : " & st(1)
        List1.ItemData(List1.ListCount - 1) = ErrLn
        End If
    Next t
          
          
    Do
        Sleep 10
        DoEvents
    Loop While hf = 0
    
    Me.Hide
    
    If hf = 2 Then
        hf = 0
        retryAsm Me.txtErr.Text
    End If
    Unload Me
          
End Sub

Private Sub List1_Click()

    If List1.ListIndex >= 0 Then
    
        On Error Resume Next
        txtErr.SetCaretPos List1.ItemData(List1.ListIndex) - 1, 1
        'txtErr.Focus = True
        txtErr.SetFocus
    
    End If
    
End Sub



Public Function GetAsmWordColor(word As String) As Long
    'too slow for now
    Dim pls As PlugIn_List, i As Long
    pls = GetPlugInList()
    
    For i = 0 To pls.count - 1
        If pls.plugins(i).used = True And pls.plugins(i).Loaded = True Then
            If pls.plugins(i).interface.GetID = tvb_ThunIDE_p Then
                GetAsmWordColor = pls.plugins(i).interface.SendMessange(tvbm_User, tidem_GetAsmColor, , word)
                Exit Function
            End If
        End If
    Next i
    
    Select Case word
    
        Case "*default*"
            GetAsmWordColor = 0
        Case "1234"
            GetAsmWordColor = RGB(120, 120, 240)
        Case ";"
            GetAsmWordColor = RGB(30, 120, 10)
        Case """'this is a string'"""
            GetAsmWordColor = RGB(120, 120, 120)
        Case Else
            GetAsmWordColor = 0
            
    End Select
    
End Function

Private Sub Form_Load()
     
    Form_Initialize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    hf = 1
    Cancel = 1
    Me.Hide
    
End Sub

