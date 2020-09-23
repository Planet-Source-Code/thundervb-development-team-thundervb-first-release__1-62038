Attribute VB_Name = "modSubClass"
'modSubClass: Contains Sub class relate code

Option Explicit

'Used to call pWindowProc of the appropriate Control
Dim m_oTmpCtl As XTab

'Used to store a Object pointer
Dim m_lObjPtr As Long

Public Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '<EhHeader>
    On Error GoTo WndProc_Err
    '</EhHeader>
    
  On Error Resume Next

  m_lObjPtr = A_GetWindowLong(hwnd, GWL_USERDATA)

  CopyMemory m_oTmpCtl, m_lObjPtr, 4

  'Call the WindowProc function for the appropriate instance of our control
  WndProc = m_oTmpCtl.pWindowProc(hwnd, msg, wParam, lParam)

    
  'Destroy tmp control's interface copy (we just need the type defs)
  CopyMemory m_oTmpCtl, 0&, 4
    
  Set m_oTmpCtl = Nothing
    '<EhFooter>
    Exit Function

WndProc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.modSubClass.WndProc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


