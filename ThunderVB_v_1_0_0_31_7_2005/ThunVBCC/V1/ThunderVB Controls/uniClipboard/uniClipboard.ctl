VERSION 5.00
Begin VB.UserControl uniClipboard 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "uniClipboard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1770
   End
End
Attribute VB_Name = "uniClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   --------------------
'   --- uniClipboard --- version 0.2.1
'   --------------------
'
'       made by Libor for ThunderVB
'
' Maybe you know that VBs clipboard does not handle Unicode strings.
' This UserControl does the job for you, you can use it to save/get Unicode strings to/from clipboard.
'
' Public methods - same as classic Clipboard. Only method GetText/SetText are unicode, other functions
' are only wrappers around classic VB Clipboard function.
'

Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Any, ByVal Src As Any, ByVal length As Long)

Private Const CF_UNICODETEXT As Long = 13

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Sub SetText(ByVal Str As String, Optional ByVal format As Integer = 0, Optional ByVal bUnicode As Boolean = True)
    '<EhHeader>
    On Error GoTo SetText_Err
    '</EhHeader>
Dim hMem As Long, pMem As Long
   
    If bUnicode = True Then
    
        'clear and open clipboard
        Me.Clear
        If OpenClipboard(UserControl.parent.hwnd) <> 0 Then
            
            'allocate memory for unicode string
            hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(Str) + 2)
            pMem = GlobalLock(hMem)
            
            'copy bytes to memory
            CopyMemory pMem, ByVal StrPtr(Str), LenB(Str) + 2
            GlobalUnlock hMem
            
            'place unicode to the clipboard (if failed then free alocated memory)
            If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then GlobalFree hMem
            CloseClipboard
        
        End If
        
    Else
    
        Clipboard.Clear
        Clipboard.SetText Str, format

    End If

    '<EhFooter>
    Exit Sub

SetText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniClipboard.SetText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub Clear()
    '<EhHeader>
    On Error GoTo Clear_Err
    '</EhHeader>
    Clipboard.Clear
    '<EhFooter>
    Exit Sub

Clear_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniClipboard.Clear " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Function GetText(Optional ByVal format As Integer = 0, Optional bUnicode As Boolean = True) As String
    '<EhHeader>
    On Error GoTo GetText_Err
    '</EhHeader>
Dim hMem As Long, i As Long, pString As Long

    If bUnicode = True Then

        'open clipboard and get pointer to data
        If OpenClipboard(UserControl.parent.hwnd) <> 0 Then
            hMem = GetClipboardData(CF_UNICODETEXT)
            If hMem <> 0 Then
                'get pointer to unicode string
                CopyMemory VarPtr(pString), ByVal hMem&, 4
                If pString <> 0 Then
                    'get length of string
                    i = lstrlen(pString)
                    If i <> 0 Then
                        GetText = Space(i)
                        'get unicode text
                        CopyMemory StrPtr(GetText), ByVal pString, LenB(GetText)
                    End If
                End If
            End If
            CloseClipboard
        End If
        
    Else
    
        GetText = Clipboard.GetText(format)
    
    End If
    
    '<EhFooter>
    Exit Function

GetText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniClipboard.GetText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function GetFormat(ByVal format As Integer) As Boolean
    '<EhHeader>
    On Error GoTo GetFormat_Err
    '</EhHeader>
    GetFormat = Clipboard.GetFormat(format)
    '<EhFooter>
    Exit Function

GetFormat_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniClipboard.GetFormat " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function GetData(Optional ByVal format As Integer = 0) As Variant
    '<EhHeader>
    On Error GoTo GetData_Err
    '</EhHeader>
    GetData = Clipboard.GetData(format)
    '<EhFooter>
    Exit Function

GetData_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniClipboard.GetData " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Sub SetData(ByVal Picture As IPictureDisp, Optional ByVal format As Integer = 0)
    '<EhHeader>
    On Error GoTo SetData_Err
    '</EhHeader>
    Call Clipboard.SetData(Picture, format)
    '<EhFooter>
    Exit Sub

SetData_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniClipboard.SetData " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
    lblCaption.Move 0, 0
    UserControl.Width = lblCaption.Width
    UserControl.Height = lblCaption.Height
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.uniClipboard.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
