Attribute VB_Name = "modClipboard"
Option Explicit

Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
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

Public Function SetText(ByVal sText As String) As Boolean

    If OpenClipboard(Form1.hwnd) <> 0 Then
        
        If EmptyClipboard = 0 Then
            CloseClipboard
            Exit Function
        End If
        
Dim hMem As Long, hPointer As Long
        hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(sText) + 2)
        hPointer = GlobalLock(hMem)
        CopyMemory hPointer, ByVal StrPtr(sText), LenB(sText) + 2
        
        If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then
            CloseClipboard
            GlobalFree hMem
            Exit Function
        End If
        
        CloseClipboard
        
        GlobalUnlock hMem
        
        SetText = True
        
    End If

End Function

Public Function Clear() As Boolean

    If OpenClipboard(Form1.hwnd) <> 0 Then
        EmptyClipboard
        CloseClipboard
        Clear = True
    End If

End Function

Public Function GetText() As String
Dim hMem As Long, i As Long, pString As Long

    If OpenClipboard(Form1.hwnd) <> 0 Then
        hMem = GetClipboardData(CF_UNICODETEXT)
        If hMem <> 0 Then
            CopyMemory VarPtr(pString), ByVal hMem&, 4
            If pString <> 0 Then
                i = lstrlen(pString)
                If i <> 0 Then
                    GetText = Space(i)
                    CopyMemory StrPtr(GetText), ByVal pString, LenB(GetText)
                End If
            End If
        End If
        CloseClipboard
    End If
    
End Function
