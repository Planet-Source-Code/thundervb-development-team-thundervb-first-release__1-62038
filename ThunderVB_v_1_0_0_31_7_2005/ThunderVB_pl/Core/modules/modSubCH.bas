Attribute VB_Name = "modSubCH"
'Revision history:
'13/9/2004[dd/mm/yyyy] : Created by Raziel
'This is where all the subclassed messanges proccesing is done...
'Any code that is inovled with the sublcass init / update/redirect
'is on the modSubClasser.bas , this code just proccess the messanges
'
'Mouse Wheel works on MDI enviroment
'
'1/10/2004[dd/mm/yyyy] : Edited by Raziel
'Works on SDI and MDI, mainly due to improvement on the subclasser
'
'
'6/10/2004[dd/mm/yyyy] : Edited by Raziel
'Base code for intelliAsm
'IntelliAsm form show/hide + tip changes
'
'10/10/2004[dd/mm/yyyy] : Edited by Raziel
'Many code fixes , intelliAsm is ready for use...
'
'22/10/2004[dd/mm/yyyy] : Edited by Raziel
'Many many fixes , code now implemetns and teh ctrl+i and ctrl+space
'(they do the same thing in our case)
'ToDo : we need a good asm definition file..
'

'Yep , ye found it , it is converted fot plyugin system too..

Option Explicit

Public Type sch_entry
    hWndFilter As ID_List
    MsgFilter As ID_List
    cb_int As ThunderVB_pl_sch_v1_0
End Type

Public Type sch_list
    item() As sch_entry
    count As Long
End Type

Public sccol As sch_list 'subclasses collection

'Called when the curent's window proc is called , before VB's one
Public Sub WindowProcBef(ByRef hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef PrevProc As Long, ByRef skipVB As Boolean, ByRef skipAft As Boolean)
    Dim i As Long, i2 As Long
    
    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveCodePane Is Nothing Then Exit Sub
    
    For i = 0 To sccol.count - 1
        With sccol.item(i)
            
            If .cb_int Is Nothing Then GoTo NextOne 'some basic and fast error checking..
        
            'if the filtering is enabled , check if the filter is ok ...
            If .hWndFilter.count > 0 Then
                For i2 = 0 To .hWndFilter.count - 1 'scan all hWnd's..
                    If .hWndFilter.Id(i2) = hWnd Then
                        GoTo Cont_hwnd 'if we found it , then good , we can go on ..
                    End If
                Next i2
                'nope , this must be filtered out..
                GoTo NextOne
Cont_hwnd:
            End If
            
            If .MsgFilter.count > 0 Then
                For i2 = 0 To .MsgFilter.count - 1 'scan all Msg's..
                    If .MsgFilter.Id(i2) = uMsg Then
                        GoTo Cont_msg 'if we found it , then good , we can go on ..
                    End If
                Next i2
                'nope , this must be filtered out..
                GoTo NextOne
Cont_msg:
            End If
            
            .cb_int.WindowProcBef hWnd, uMsg, wParam, lParam, PrevProc, skipVB, skipAft
            
        End With
NextOne:
    Next i
    
End Sub

'Called when the curent's window proc is called , After VB's one
Public Sub WindowProcAft(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, PrevProc As Long, ByRef RetValue As Long)
    Dim i As Long, i2 As Long
    
    If VBI Is Nothing Then Exit Sub
    If VBI.ActiveCodePane Is Nothing Then Exit Sub
    
    For i = 0 To sccol.count - 1
        With sccol.item(i)
            
            If .cb_int Is Nothing Then GoTo NextOne 'some basic and fast error checking..
            
            'if the filtering is enable , check if the filter is ok ...
            If .hWndFilter.count > 0 Then
                For i2 = 0 To .hWndFilter.count - 1 'scan all hWnd's..
                    If .hWndFilter.Id(i2) = hWnd Then
                        GoTo Cont_hwnd 'if we found it , then good , we can go on ..
                    End If
                Next i2
                'nope , this must be filtered out..
                GoTo NextOne
Cont_hwnd:
            End If
            
            If .MsgFilter.count > 0 Then
                For i2 = 0 To .MsgFilter.count - 1 'scan all Msg's..
                    If .MsgFilter.Id(i2) = uMsg Then
                        GoTo Cont_msg 'if we found it , then good , we can go on ..
                    End If
                Next i2
                'nope , this must be filtered out..
                GoTo NextOne
Cont_msg:
            End If
            
            .cb_int.WindowProcAft hWnd, uMsg, wParam, lParam, PrevProc, RetValue
            
        End With
NextOne:
    Next i
    
    
End Sub
