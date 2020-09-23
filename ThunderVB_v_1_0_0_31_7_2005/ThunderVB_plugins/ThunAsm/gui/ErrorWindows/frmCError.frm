VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmCError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C Error"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin CodeSenseCtl.CodeSense txtErr 
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "frmCError.frx":0000
      TabIndex        =   5
      Top             =   2880
      Width           =   7575
   End
   Begin VB.TextBox txtclOut 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmCError.frx":0166
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Retry"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdIng 
      Caption         =   "Ingore"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmd_can 
      Caption         =   "Cancel Compile"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1305
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   7575
   End
End
Attribute VB_Name = "frmCError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Revision history:
'29/8/2004[dd/mm/yyyy] : Created by Raziel
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
          Dim i As Long
          
     hf = 0
     'txtErr.MarginWidth(0) = 35
     'txtErr.MarginWidth(1) = 0
     'txtErr.MarginType(0) = SC_MARGIN_NUMBER
     'txtErr.Lexer = CppNoCase
     'txtErr.StyleClearAll
      txtErr.LineNumbering = True
      txtErr.LineNumberStart = 1
      txtErr.SetColor cmClrLineNumber, RGB(255, 255, 255)
      txtErr.SetColor cmClrLineNumberBk, RGB(120, 120, 120)
    '"Primary keywords and identifiers",
    '"Secondary keywords and identifiers",
    '"Documentation comment keywords",
    '"Unused",
    '"Global classes and typedefs",
          
    'txtErr.Lexer = Cpp
          
    'txtErr.KEYWORDS(0) = LCase("and and_eq asm auto bitand bitor bool break case catch char " & _
                                     "class compl const const_cast continue default delete do double " & _
                                     "dynamic_cast else enum explicit export extern false float for " & _
                                     "friend goto if inline int long mutable namespace new not not_eq " & _
                                     "operator or or_eq private protected public register reinterpret_cast " & _
                                     "return short signed sizeof static static_cast struct switch " & _
                                     "template this throw true try typedef typeid typename union " & _
                                     "unsigned using virtual void volatile wchar_t while xor xor_eq")

          
    'For i = 0 To 31
    '    txtErr.StyleSetFore CInt(i), GetCWordColor("*default*")
    'Next i
     
    'standar
    'txtErr.StyleSetFore SCE_C_NUMBER, GetCWordColor("1234")
    txtErr.SetColor cmClrNumber, GetCWordColor("1234")
    'txtErr.StyleSetFore SCE_C_STRING, GetCWordColor(Add34("'this is a string'"))
    txtErr.SetColor cmClrString, GetCWordColor(Add34("this is a string"))
    'txtErr.StyleSetFore SCE_C_COMMENT, GetCWordColor("//")
    txtErr.SetColor cmClrComment, GetCWordColor("//")
    'remaped
    '
    'txtErr.StyleSetFore SCE_C_WORD, GetCWordColor("if")
    txtErr.SetColor cmClrKeyword, GetCWordColor("if")
    txtErr.language = "c/c++"
    txtErr.ColorSyntax = True
    
End Sub

Public Sub ShowError(clErr As String, errText As String, Optional title As String = "Error on C code:")
    Dim t As Long, ln As String, st() As String, ErrLn As Long
       
    Form_Initialize
    hf = 0
    Me.caption = title
    Me.txtclOut.Text = clErr
    List1.Clear
    For t = 0 To UBound(Split(clErr, vbNewLine))
        ln = Split(clErr, vbNewLine)(t)
        st = Split(ln, ") : error ")
        If UBound(st) = 1 Then
            ErrLn = Mid$(st(0), InStrRev(st(0), "(") + 1)
           List1.AddItem "Line : " & ErrLn & " : " & st(1)
           List1.ItemData(List1.ListCount - 1) = ErrLn
       End If
    Next t
       
    Me.txtErr.Text = errText
    Me.Show ' vbModal
    Do
        Sleep 10
        DoEvents
    Loop While hf = 0
    Me.Hide
    If hf = 2 Then
        hf = 0
        retryC Me.txtErr.Text
    End If
    Unload Me
       
End Sub

Private Sub Form_Load()
    Form_Initialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
      
    hf = 1
    Cancel = 1
    Me.Hide
      
End Sub

Private Sub List1_Click()

    If List1.ListIndex >= 0 Then
    
        On Error Resume Next
        txtErr.SetCaretPos List1.ItemData(List1.ListIndex) - 1, 1
        'txtErr.Focus = True
        txtErr.SetFocus
    
    End If
    
End Sub

Public Function GetCWordColor(word As String) As Long

    'too slow for now- prob is ok :)
    Dim pls As PlugIn_List, i As Long
    pls = GetPlugInList()
    
    For i = 0 To pls.count - 1
        If pls.plugins(i).used = True And pls.plugins(i).Loaded = True Then
            If pls.plugins(i).interface.GetID = tvb_ThunIDE_p Then
                GetCWordColor = pls.plugins(i).interface.SendMessange(tvbm_User, tidem_GetCColor, , word)
                Exit Function
            End If
        End If
    Next i
    
    Select Case word
    
        Case "*default*"
            GetCWordColor = 0
        Case "1234"
            GetCWordColor = RGB(120, 120, 240)
        Case "//"
            GetCWordColor = RGB(30, 120, 10)
        Case """'this is a string'"""
            GetCWordColor = RGB(120, 120, 120)
        Case Else
            GetCWordColor = 0
            
    End Select
    
End Function

