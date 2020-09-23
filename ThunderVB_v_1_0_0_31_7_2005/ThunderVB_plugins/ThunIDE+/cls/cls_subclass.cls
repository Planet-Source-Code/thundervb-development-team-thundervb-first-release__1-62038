VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_subclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetCaretPos Lib "user32" _
               (lpPoint As POINTAPI) As Long

Dim ptemp As POINTAPI
Dim ctrl_down As Boolean

Implements ThunderVB_pl_sch_v1_0

Private Sub ThunderVB_pl_sch_v1_0_WindowProcBef(hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, PrevProc As Long, skipVB As Boolean, skipAft As Boolean)
       On Error GoTo ErrH
       'Exit Sub
       If vbi Is Nothing Then Exit Sub
       If vbi.ActiveCodePane Is Nothing Then Exit Sub
          
       Select Case uMsg
                 
           Case WM_LBUTTONDOWN, WM_RBUTTONDOWN
               If bASM_IntelliSense Then
                   IntelliAsm.HideIntelliAsm
                   asmTT.HideToolTip
               End If
                  
          Case WM_KILLFOCUS
              If bASM_IntelliSense Then
                  IntelliAsm.HideIntelliAsm
                  If asmTT.Visible Then
                      asmTT.HideToolTip
                      wm_tthide = True
                  Else
                      wm_tthide = False
                  End If
              End If
                  
          Case 522 'Mouse wheel event
              If bASM_IntelliSense Then
                  If IntelliAsm.Visible = True Then
                      SendMessage IntelliAsm.iSe.memb_list.hWnd, uMsg, wParam, lParam
                      Exit Sub
                  End If
                  asmTT.HideToolTip
              End If
               Dim Top As Long
              If wParam < 0 Then
                  Top = vbi.ActiveCodePane.TopLine + 3
                  If Top > vbi.ActiveCodePane.codeModule.CountOfLines Then
                      Top = vbi.ActiveCodePane.codeModule.CountOfLines
                  End If
                  If TipVisible = False Then vbi.ActiveCodePane.TopLine = Top
                  If bASM_QuickWatch Then TipOffset = TipOffset + 1
               Else
                  Top = vbi.ActiveCodePane.TopLine - 3
                  If Top < 1 Then
                      Top = 1
                  End If
                  If TipVisible = False Then vbi.ActiveCodePane.TopLine = Top
                  If bASM_QuickWatch Then TipOffset = TipOffset - 1
              End If
              If bASM_QuickWatch Then TipReSetText
                  
           'Case -3 'WM_MOUSEMOVE
              
           '    If bASM_QuickWatch Then
           '       CheckToolTip
           '       TipReSetText
           '       RichWordOver hWnd, lParam And 65535, lParam \ 65536
           '    End If
              
          Case WM_KEYDOWN
              If bASM_IntelliSense Then
                  If ((wParam = VK_UP) Or (wParam = VK_DOWN) Or _
                       (wParam = vbKeyPageUp) Or (wParam = vbKeyPageDown)) And IntelliAsm.Visible = True Then
                          
                      IntelliAsmListSend uMsg, wParam, lParam
                      skipVB = True
                      skipAft = True
                          
                  ElseIf ((wParam = VK_ESCAPE)) And (IntelliAsm.Visible = True) Then
                      IntelliAsmHideAll
                      skipVB = True
                      skipAft = True
                  End If
              End If
                  
          Case WM_CHAR
              If bASM_IntelliSense Then
                  If IntelliAsm.Visible Then
                      If (InStrRev(" ,;]" & vbTab & vbCrLf, ChrW$(wParam)) > 0) And (Len(IntelliAsm.iSe.memb_list.Text) > 0) Then
                              IntelliAsm.iSe.list_DblClick
                      End If
                  End If
              End If
              
          Case WM_KEYUP
              If wParam = VK_CONTROL Then
                  ctrl_down = False
              End If
                  
      End Select
          
      If bASM_IntelliSense Then
          If GetAsyncKeyState(VK_CONTROL) = True Then
              If ((GetAsyncKeyState(VK_SPACE) <> 0) Or (GetAsyncKeyState(VK_I) <> 0)) = True Then
                  IntelliAsmChange hWnd, False
              End If
          End If
      End If
          
      Exit Sub
ErrH:
      ErrorBox Err.Description & " at line " & Erl & vbNewLine & "Params = " & _
       hWnd & " " & uMsg & " " & wParam & " " & lParam & " " & PrevProc & " " & skipVB & " " & skipAft _
       , "cls_subclass", "ThunderVB_pl_sch_v1_0_WindowProcBef"
End Sub

Private Sub ThunderVB_pl_sch_v1_0_WindowProcAft(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, PrevProc As Long, RetValue As Long)
   Dim temp As String
On Error GoTo ErrH:
       If vbi Is Nothing Then Exit Sub
       If vbi.ActiveCodePane Is Nothing Then Exit Sub
          
       Select Case uMsg
              
           Case WM_CHAR
           'If wParam = 3 Then MsgBox "copy"
               If bASM_IntelliSense Then
                   AutoExpandAsm_C
                   If wParam <> 27 Then
                       IntelliAsmChange hWnd, False
                   End If
              End If

          Case WM_KEYDOWN
           '    MsgBox wParam
           'If bASM_IntelliSense Then IntelliAsmChange hWnd, True
                  
           'Case WM_HOTKEY
           '    Select Case wParam
           '        Case hk_CopyHotKeyID
           '            'MsgBox "Ctrl+C"
           '            SendMessage hWnd, WM_KEYDOWN, &H11, &H1D0001
           '            SendMessage hWnd, WM_KEYDOWN, &H43, &H2E0001
           '            SendMessage hWnd, WM_CHAR, &H3, &H2E0001
           '            SendMessage hWnd, WM_KEYUP, &H43, &H2E0001
           '            SendMessage hWnd, WM_KEYUP, &H11, &H1D0001
           '            SendMessage hWnd, WM_COPY, 0, 0
           '            Dim rtf As RichTextBox
           '            Set rtf = frmImages.rtb
           '            rtf.TextRTF = ""
           '            rtf.Text = "'Colored by ThunderVB " & GetThunVBVer & vbNewLine & _
           '                                              Clipboard.GetText(vbCFText)
           '            Clipboard.SetText DoNonRealTimeColor(rtf), vbCFRTF
           '        Case hk_CutHotKeyID
           '            'MsgBox "Ctrl+X"
           '        Case hk_CtrlIHotKeyID
           '            'MsgBox "Ctrl+I"
           '        Case hk_CtrlSpaceHotKeyID
           '            'MsgBox "Ctrl+Space"
           '        Case Else
           '            MsgBox "Other hotkey ?? id=" & wParam
           '    End Select
                  
      End Select
      If Get_Misc(CopyTimeColoring) Then
          If GetAsyncKeyState(VK_CONTROL) <> 0 Then
              If ((GetAsyncKeyState(VK_C) <> 0) Or (GetAsyncKeyState(VK_X) <> 0)) = True Then
                   Dim rtf As RichTextBox
                  Set rtf = frmImages.rtb
                  rtf.TextRTF = ""
                  rtf.Text = "'Colored by " & GetThunIdeVersion & vbNewLine & _
                                                     Clipboard.GetText(vbCFText)
                  Clipboard.SetText DoNonRealTimeColor(rtf), vbCFRTF
              End If
          End If
      End If
      Exit Sub
ErrH:
      ErrorBox Err.Description & " at line " & Erl, "cls_subclass", "ThunderVB_pl_sch_v1_0_WindowProcAft"
End Sub
