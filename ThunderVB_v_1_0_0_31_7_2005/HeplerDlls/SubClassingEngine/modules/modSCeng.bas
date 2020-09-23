Attribute VB_Name = "modSCeng"
Option Explicit

Private Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long


Public Type sc_entry
    Interface As ISubclass_Callbacks
    hwnd As Long
    mode As subclass_mode
    OldProc As Long
    Unicode As Boolean
    code As Long
    Used As Boolean
End Type

Public Type sc_list
    items() As sc_entry
    count As Long
End Type

Dim scw As sc_list

'subclassing engine

Public Function isMalloc() As IMalloc
Static im As IMalloc
   If (im Is Nothing) Then
      If Not (SHGetMalloc(im) = 0) Then
         ' Fatal error
         Err.Raise 7, , "Can't allocate memory"
      End If
   End If
   Set isMalloc = im
End Function


Public Sub SubClasshWnd(hwnd As Long, callbacks As ISubclass_Callbacks, mode As subclass_mode, Optional Unicode As Boolean = False)
Dim id As Long
    
    id = m_GetFreeID()
    With scw.items(id)
        .hwnd = hwnd
        Set .Interface = callbacks
        .mode = mode
        .Unicode = Unicode
        .Used = True
    End With
    
    m_SubClasshWnd scw.items(id), id
    
End Sub

Public Sub UnSubClasshWnd(hwnd As Long)
Dim id As Long
    
    id = m_GetIdFromhWnd(hwnd)
    If id = -1 Then Exit Sub
    m_DeSubClasshWnd scw.items(id)
    
End Sub

Private Function m_WindowProc(ByVal id As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tmp As sc_entry, lret As Long, cwp As Boolean, cap As Boolean
    
    'id = m_GetIdFromhWnd(hwnd)
    'If id = -1 Then Exit Function
    
    tmp = scw.items(id)
    
    If tmp.Interface Is Nothing Then
       scw.items(id).Used = False
       m_DeSubClasshWnd scw.items(id)
       Exit Function
    End If
    
    If tmp.mode = wproc_replace Then
        tmp.Interface.WindowProc hwnd, uMsg, wParam, lParam, lret, tmp.OldProc
    Else
        cwp = True
        cap = True
        tmp.Interface.BefWndProc hwnd, uMsg, wParam, lParam, lret, cwp, cap, tmp.OldProc
        
        If cwp Then
            lret = A_CallWindowProc(ByVal tmp.OldProc, ByVal hwnd, ByVal uMsg, ByVal wParam, ByVal lParam)
        End If
        
        If cap Then
            tmp.Interface.AftWndProc hwnd, uMsg, wParam, lParam, lret, cwp, tmp.OldProc
        End If
    End If
    
    m_WindowProc = lret
    
End Function

Private Function m_GetIdFromhWnd(ByVal hwnd As Long) As Long
Dim i As Long
    
    For i = 0 To scw.count - 1
        If (scw.items(i).hwnd = hwnd) Then
            If scw.items(i).Used = True Then
                m_GetIdFromhWnd = i
                Exit Function
            End If
        End If
    Next i
    m_GetIdFromhWnd = -1
    
End Function

Private Sub m_SubClasshWnd(entry As sc_entry, ByVal i As Long)
    
    With entry
        GenerateAsm i, AddressOf m_WindowProc, .code
        If .Unicode Then
            .OldProc = W_SetWindowLong(.hwnd, GWL_WNDPROC, .code)
        Else
            .OldProc = A_SetWindowLong(.hwnd, GWL_WNDPROC, .code)
        End If
    End With
    
End Sub

Private Sub m_DeSubClasshWnd(entry As sc_entry)

    entry.Used = False
    'ReDim entry.Code(0)
    isMalloc.Free entry.code
    entry.code = 0
    A_SetWindowLong entry.hwnd, GWL_WNDPROC, entry.OldProc
    
End Sub

Private Function m_GetFreeID() As Long
Dim i As Long

    For i = 0 To scw.count - 1
        If scw.items(i).Used = False Then
            m_GetFreeID = i
            Exit Function
        End If
    Next i
    
    ReDim Preserve scw.items(scw.count)
    m_GetFreeID = scw.count
    scw.count = scw.count + 1
    
End Function

Public Sub GenerateAsm(ByVal i As Long, ByVal procaddr As Long, ByRef CodeL As Long)
    Dim asz As Long
    Dim code() As Byte
    
    'CC                 'int 3 - for debuging
    'AppendToArray Code, asz, &HCC
    
    '58                 'pop eax
    AppendToArray code, asz, &H58
    
    '68 [7A 96 98 00]   'push value
    AppendToArray code, asz, &H68
    
    AppendToArray code, asz, i And 255
    AppendToArray code, asz, (i \ 256) And 255
    AppendToArray code, asz, (i \ 65536) And 255
    AppendToArray code, asz, (i \ 16777216) And 255
    
    '50                 'push eax
    AppendToArray code, asz, &H50
    
    'B8 [FF FF FF 0F]   'mov eax,value
    AppendToArray code, asz, &HB8
    
    AppendToArray code, asz, procaddr And 255
    AppendToArray code, asz, (procaddr \ 256) And 255
    AppendToArray code, asz, (procaddr \ 65536) And 255
    AppendToArray code, asz, (procaddr \ 16777216) And 255
    
    'FF E0             'jmp eax
    AppendToArray code, asz, &HFF
    AppendToArray code, asz, &HE0
    
    CodeL = isMalloc.Alloc(UBound(code) + 1)
    CopyMemory ByVal CodeL, code(0), UBound(code) + 1
    
End Sub

Public Function CountSubClassedWindows() As Long
Dim cnt As Long, i As Long

    For i = 0 To scw.count - 1
        If scw.items(i).Used Then
           cnt = cnt + 1
        End If
    Next i
    
    CountSubClassedWindows = cnt
    
End Function

Private Function AppendToArray(ByRef arr() As Byte, ByRef asz As Long, ByVal value As Byte)
    
    'If arr = 0 Then
    '    arr = isMalloc.Alloc(1)
    'Else
    '    arr = isMalloc.Realloc(ByVal arr, asz + 1)
    'End If
    ReDim Preserve arr(asz)
    arr(asz) = value
    asz = asz + 1
    'CopyMemory ByVal arr + asz, value, 1
    'asz = asz + 1
    
End Function
