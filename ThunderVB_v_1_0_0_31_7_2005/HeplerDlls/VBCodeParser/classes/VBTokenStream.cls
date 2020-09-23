VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBTokenStream"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Explicit

Private Tokens() As String, CurIndex As Long, Length As Long

Public Function CanRead() As Boolean
    
    CanRead = IsInRange(CurIndex)
    
End Function

Public Function IsInRange(index As Long) As Boolean
    
    IsInRange = (index < Length) And (index >= 0)
    
End Function

Public Sub FromTokenArray(tks() As String)
    
    CurIndex = 0
    Length = UBound(tks) + 1
    Tokens = tks
    
End Sub

Public Function GetNext() As String
    
    If CanRead Then
        GetNext = Tokens(CurIndex)
        CurIndex = CurIndex + 1
    End If
    
End Function

Public Function GetPrev() As String
    
    If IsInRange(CurIndex - 1) Then
        GetPrev = Tokens(CurIndex - 1)
        CurIndex = CurIndex - 1
    End If
    
End Function

Public Function PeakNext() As String
    
    If IsInRange(CurIndex) Then
        PeakNext = Tokens(CurIndex)
    End If
    
End Function

Public Function PeakPrev() As String
    
    If IsInRange(CurIndex - 1) Then
        PeakPrev = Tokens(CurIndex - 1)
    End If
    
End Function

Public Sub GoToNext()
    
    If IsInRange(CurIndex) Then
        CurIndex = CurIndex + 1
    End If
    
End Sub

Public Sub GoToPrev()
    
    If IsInRange(CurIndex) Then
        CurIndex = CurIndex - 1
    End If
    
End Sub


