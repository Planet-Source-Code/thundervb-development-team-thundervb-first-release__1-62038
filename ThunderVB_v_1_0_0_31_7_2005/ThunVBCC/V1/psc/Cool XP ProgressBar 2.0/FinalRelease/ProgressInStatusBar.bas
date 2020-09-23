Attribute VB_Name = "ProgressInStatusBar"
'Simple Function to Show Progress In StatusBar
'Author: Mario Flores G
'E-mail: sistec_de_juarez@hotmail.com


'{{{8 May 2oo4... Little Add}}}}
'StatusBar.Panels(PanelNumber).Width = Progress.Width + 15
'Automatically Expand Selected StatusBar Panel Width ....
'Can Make the same for the Height but looks ugly when Progress is to Big..(My point of view)

'CD JUAREZ CHIHUAHUA MEXICO

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long



Const WM_USER    As Long = &H400
Const SB_GETRECT As Long = (WM_USER + 10)


Public Sub ShowProgressInStatusBar(ByRef Progress As Control, ByRef StatusBar As StatusBar, ByVal PanelNumber As Long)
    '<EhHeader>
    On Error GoTo ShowProgressInStatusBar_Err
    '</EhHeader>

    Dim TRC As RECT
    
        StatusBar.Panels(PanelNumber).Width = Progress.Width + 15
        SendMessageAny StatusBar.hwnd, SB_GETRECT, PanelNumber - 1, TRC
       
       'Set The ProgressBar Parent's Window = StatusBar
       'Center The ProgressBar in the Selected Panel (PanelNumber)
         
        With Progress
            SetParent .hwnd, StatusBar.hwnd
            .Move TRC.Left + ((TRC.Right - TRC.Left) / 2) - (.Width / 2), TRC.Top + ((TRC.Bottom - TRC.Top) / 2) - (.Height / 2), .Width, .Height
            .visible = True
            .Value = 0
        End With
        
    '<EhFooter>
    Exit Sub

ShowProgressInStatusBar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ProgressInStatusBar.ShowProgressInStatusBar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
