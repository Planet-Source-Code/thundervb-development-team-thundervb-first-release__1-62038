Attribute VB_Name = "modSubClasser"
'Revision history:
'13/9/2004[dd/mm/yyyy] : Created by Raziel
'This is where all the subclassing init/update to curect window , message forwarding is done
'here is all the code that may crash something , and to keep that seperate the code that
'actualy does the message proccessing (witch is safe) is on a different module..
'This code acts as a abraction layer
'Works for MDI olny..
'
'1/10/2004[dd/mm/yyyy] : Edited by Raziel
'Works on SDI and MDI
'
'
'Notes
'hWnd In Window class is a hiden member .. to see it press f2
'and then right click , show hiden mebmers
'
'
'6/10/2004[dd/mm/yyyy] : Edited by Raziel
'Some heavy corection and fixes for mdi mode
'now , both MDIclient and child windows are hooked
'

'Wooow , even this is changed for the plugin system .. d@mit..
'All the functions are prefixed with sc , since from now on all functions
'that are on a speciacific grop are prefixed to make finding them easy..
Option Explicit


Public MainhWnd As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Const GWL_WNDPROC = (-4)

Type SubClassed
    
    hWnd As Long
    Previous As Long
    
End Type

Dim nCount As Long, sWind() As SubClassed, TimerId As Long

'if any new code windows are oppened then subclass them
Sub sc_CheckStatus()
Dim Temp As Long
    
    If DebugMode Then Exit Sub
    If VBI Is Nothing Then GoTo t2
    If VBI.MainWindow Is Nothing Then GoTo t2
    

    If MainhWnd <> 0 Then
    
        Temp = FindWindow("VbaWindow", vbNullString)
        If Temp = 0 Then GoTo t1
        sc_GetOldAddress Temp
        Exit Sub
t1:
        Temp = FindWindowEx(MainhWnd, 0, "MDIClient", vbNullString)
        If Temp = 0 Or (Temp = MainhWnd) Then GoTo nhp
        sc_GetOldAddress Temp
        sc_GetOldAddress GetParent(Temp)
nhp:
        Temp = GetCodeWindowHWnd(VBI)
        If Temp = 0 Then Exit Sub
        sc_GetOldAddress Temp
        
    Else
t2:
        Temp = FindWindow("VbaWindow", vbNullString)
        If Temp = 0 Then Exit Sub
        sc_GetOldAddress Temp
    End If

    
End Sub

Function sc_KillAllSubClasses() As Boolean
Dim i As Long, Error As Boolean
    LogMsg "Un initing Subclassed windows", "modSubClasser", "KillAllSubClasses"
           
    For i = 0 To nCount - 1
    
    If sWind(i).Previous = 0 Then
    
        ErrorBox "Subclassed window " & sWind(i).hWnd & ",id=" & i & vbNewLine & _
                 "Original wndproc is 0.This may cause a crash after addin unload", _
                 "modSubClasser", "CheckStatus"
                 
        Error = True
        
    Else
        
        LogMsg "Subclassed window " & i & " hWnd:" & sWind(i).hWnd & " RestoreTo : " & sWind(i).Previous, "modSubClasser", "KillAllSubClasses"
        
        SetWindowLong sWind(i).hWnd, GWL_WNDPROC, sWind(i).Previous
        
    End If
    
    Next i

End Function

Function sc_GetOldAddress(hWnd As Long) As Long
Dim i As Long
    'Exit Function
    For i = 0 To nCount - 1
    
    If sWind(i).hWnd = hWnd Then
        
        sc_GetOldAddress = sWind(i).Previous
        Exit Function
        
    End If
    
    Next i
    
    ReDim Preserve sWind(nCount)
    sWind(nCount).Previous = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    If sWind(nCount).Previous = 0 Then Exit Function
    sWind(nCount).hWnd = hWnd
    sc_GetOldAddress = sWind(nCount).Previous
    LogMsg "Subclassed " & sWind(nCount).hWnd & "(id:" & nCount & ")" & _
           "OldProc is " & sWind(nCount).Previous & " new is " & GetAddr(AddressOf WindowProc), "modSubClasser", "GetOldAddress"
    nCount = nCount + 1
    
End Function

Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If VBI Is Nothing Then GoTo def
    If VBI.ActiveWindow Is Nothing Then GoTo def
    If VBI.ActiveWindow.type <> 0 Then GoTo def
    Dim skipVB As Boolean, skipAft As Boolean
    
    WindowProcBef hWnd, uMsg, wParam, lParam, sc_GetOldAddress(hWnd), skipVB, skipAft
    
    If skipVB = False Then
        WindowProc = CallWindowProc(sc_GetOldAddress(hWnd), hWnd, uMsg, wParam, lParam)
    End If
    
    If skipAft = False Then
        WindowProcAft hWnd, uMsg, wParam, lParam, sc_GetOldAddress(hWnd), WindowProc
    End If
    
    Exit Function
    
def:
    WindowProc = CallWindowProc(sc_GetOldAddress(hWnd), hWnd, uMsg, wParam, lParam)
    
End Function

Function GetAddr(Temp As Long) As Long

    GetAddr = Temp

End Function

Public Sub sc_ApiTimer(ByVal aStart As Boolean)

    If DebugMode Then Exit Sub
'Exit Sub
    If aStart = False Then
        If TimerId Then
            KillTimer 0, TimerId
            TimerId = 0
        End If
    Else
        If TimerId Then
            KillTimer 0, TimerId
            TimerId = 0
        End If
        TimerId = SetTimer(0, 0, 200, AddressOf sc_CheckStatus)
    End If

End Sub

' return the handle of the current
' code window, 0 if none
Function GetCodeWindowHWnd(VBInstance As VBIDE.VBE) As Long
Dim hWnd As Long, caption As String
    
    If VBInstance.ActiveCodePane Is Nothing Then Exit Function
    caption = VBInstance.ActiveCodePane.Window.caption
    If VBInstance.DisplayModel = vbext_dm_MDI Then
        ' get the handle of the main window
        ' hWnd is a hidden, undocumented
        ' property
        hWnd = VBInstance.MainWindow.hWnd
        ' in MDI mode there is an
        ' intermediate window, of class\
        ' MDIClient
        hWnd = FindWindowEx(hWnd, 0, _
        "MDIClient", vbNullString)
        ' finally we can get the hWnd of
        ' the code window
        GetCodeWindowHWnd = _
        FindWindowEx(hWnd, 0, _
        "VbaWindow", caption)
    Else
        ' no intermediate window in SDI mode
        GetCodeWindowHWnd = _
        FindWindow("VbaWindow", caption)
    End If
End Function
