VERSION 5.00
Begin VB.UserControl ScintillaEdit 
   BackColor       =   &H0080FFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "ScintillaEdit.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ScintillaEdit.ctx":0342
End
Attribute VB_Name = "ScintillaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' Scintilla Control Warper Made by drk||Raziel
' This is a VB6 User Control + WinApi
' The Control was made from the
' Documentation of the v1.6.1 Release
' For the latest version of Scintilla visit
' www.scintilla.org
' Please note that this control is
' available for download on www.pscode.com
' and that it has nothing to do
' with the author of Scintilla
' So don't BUG him for my bugz
' This is a part of ThunderVB project
' latest thunderVB  version can be found
' at http://thundervb.sourceforge.net/



' Using this code form PlanetSourceCode :
' [disclaimer of the code]
' Self SubClasser by
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' [end]
'==================================================================================================
'Subclasser declarations
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
'            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Other Decs
'======================

Dim SelVisible As Boolean
Const strDefText = "Scintilla Control Warper Made by drk||Raziel" & vbNewLine & _
                   "This is a VB6 User Control + WinApi" & vbNewLine & _
                   "The Control was made from the" & vbNewLine & _
                   "Documentation of the v1.6.1 Release" & vbNewLine & _
                   "For the latest version of Scintilla visit " & vbNewLine & _
                   "www.scintilla.org" & vbNewLine & _
                   "Please note that this control is" & vbNewLine & _
                   "available for download on www.pscode.com" & vbNewLine & _
                   "and that it has nothing to do" & vbNewLine & _
                   "with the author of Scintilla" & vbNewLine & _
                   "So don't BUG him for my bugz" & vbNewLine & _
                   "This cotnrol is developed as a part of the ThunderVB project"
           
Dim WndPointer As Long
Dim FunPointer As Long
Dim sci As Long
Dim lppw As Long
Dim ctr_style As Long 'control's style..
Dim bDoNotInit As Boolean

Public Event SciNotify(param As Long)

'=======================
'Subclasser :
'=======================

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
'Notes:
  'If you really know what you're doing, it's possible to change the values of the
  'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
  'values get passed to the default handler.. and optionaly, the 'after' callback
    '<EhHeader>
    On Error GoTo zSubclass_Proc_Err
    '</EhHeader>
  
  Select Case uMsg

        ' ======================================================
        ' Hide the TextBox when it loses focus (its LostFocus event it not fired
        ' when losing focus to a window outside the app).
    
        Case WM_NOTIFY
            If lParam <> 0 Then RaiseEvent SciNotify(lParam)

    
    End Select
    
    '<EhFooter>
    Exit Sub

zSubclass_Proc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zSubclass_Proc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    '<EhHeader>
    On Error GoTo Subclass_AddMsg_Err
    '</EhHeader>
  With sc_aSubData(zIdx(lng_hWnd))
    If When And 2 Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, 2, .nAddrSub)
    End If
    If When And 1 Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, 1, .nAddrSub)
    End If
  End With
    '<EhFooter>
    Exit Sub

Subclass_AddMsg_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Subclass_AddMsg " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
    '<EhHeader>
    On Error GoTo Subclass_DelMsg_Err
    '</EhHeader>
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
    '<EhFooter>
    Exit Sub

Subclass_DelMsg_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Subclass_DelMsg " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    '<EhHeader>
    On Error GoTo Subclass_InIDE_Err
    '</EhHeader>
  Debug.Assert zSetTrue(Subclass_InIDE)
    '<EhFooter>
    Exit Function

Subclass_InIDE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Subclass_InIDE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
    '<EhHeader>
    On Error GoTo Subclass_Start_Err
    '</EhHeader>
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                          'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = A_SetWindowLong(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call CopyMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                                 'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
    '<EhFooter>
    Exit Function

Subclass_Start_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Subclass_Start " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()

  Dim i As Long
  
  i = -1
  On Error Resume Next
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop

End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
On Error Resume Next
  With sc_aSubData(zIdx(lng_hWnd))
    Call A_SetWindowLong(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With

End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    '<EhHeader>
    On Error GoTo zAddMsg_Err
    '</EhHeader>
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = 2 Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
    '<EhFooter>
    Exit Sub

zAddMsg_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zAddMsg " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    '<EhHeader>
    On Error GoTo zAddrFunc_Err
    '</EhHeader>
  zAddrFunc = GetProcAddress(A_GetModuleHandle(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
    '<EhFooter>
    Exit Function

zAddrFunc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zAddrFunc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    '<EhHeader>
    On Error GoTo zDelMsg_Err
    '</EhHeader>
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
    '<EhFooter>
    Exit Sub

zDelMsg_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zDelMsg " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    '<EhHeader>
    On Error GoTo zIdx_Err
    '</EhHeader>
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
    '<EhFooter>
    Exit Function

zIdx_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zIdx " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    '<EhHeader>
    On Error GoTo zPatchRel_Err
    '</EhHeader>
  Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
    '<EhFooter>
    Exit Sub

zPatchRel_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zPatchRel " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    '<EhHeader>
    On Error GoTo zPatchVal_Err
    '</EhHeader>
  Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)
    '<EhFooter>
    Exit Sub

zPatchVal_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zPatchVal " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    '<EhHeader>
    On Error GoTo zSetTrue_Err
    '</EhHeader>
  zSetTrue = True
  bValue = True
    '<EhFooter>
    Exit Function

zSetTrue_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.zSetTrue " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'=================
'Rest Control Code
'=================

Public Function Send_SCI_message(ByVal msg As Long, ByVal par1 As Long, ByVal par2 As Long) As Long
    '<EhHeader>
    On Error GoTo Send_SCI_message_Err
    '</EhHeader>
    
    If (FunPointer = 0) Or (WndPointer = 0) Then
        Send_SCI_message = A_SendMessage(sci, msg, par1, par2)
    Else
        Send_SCI_message = CallBP(WndPointer, msg, par1, par2, FunPointer)
    End If
    
    '<EhFooter>
    Exit Function

Send_SCI_message_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Send_SCI_message " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function Send_SCI_messageStr(ByVal msg As Long, ByVal par1 As Long, ByVal par2 As String) As Long
    '<EhHeader>
    On Error GoTo Send_SCI_messageStr_Err
    '</EhHeader>
    
    If (FunPointer = 0) Or (WndPointer = 0) Then
        Send_SCI_messageStr = A_SendMessageStr(sci, msg, par1, par2)
    Else
        Dim Temp() As Byte
        sci_BstrToAnsi par2, Temp
        Send_SCI_messageStr = CallBP(WndPointer, msg, par1, VarPtr(Temp(0)), FunPointer)
    End If
    
    '<EhFooter>
    Exit Function

Send_SCI_messageStr_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Send_SCI_messageStr " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Private Sub UserControl_Initialize()
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>
    If A_LoadLibrary("SciLexer_tvb.dll") = 0 Then GoTo ErrExit
    If bDoNotInit = False Then
        ctr_style = WS_EX_CLIENTEDGE
    End If
    sci = A_CreateWindowEx(BorderStyle, "Scintilla", _
                           UserControl.Name, WS_CHILD Or WS_VISIBLE, 0, 0, 200, 200, _
                           UserControl.hwnd, 0, App.hInstance, 0)
    If sci = 0 Then GoTo ErrExit
    'DebugerRaise
    If CCode() Then
        FunPointer = Send_SCI_message(SCI_GETDIRECTFUNCTION, 0, 0)
        WndPointer = Send_SCI_message(SCI_GETDIRECTPOINTER, 0, 0)
    End If

    Subclass_Start UserControl.hwnd
    Subclass_AddMsg UserControl.hwnd, WM_NOTIFY
    If bDoNotInit = False Then
        Text = strDefText
        bDoNotInit = True
    End If
    '// now a wrapper to call Scintilla directly
    'sptr_t CallScintilla(unsigned int iMessage, uptr_t wParam, sptr_t lParam){
    '    return pSciMsg(pSciWndData, iMessage, wParam, lParam);
    '}
    Exit Sub
    
ErrExit:
    'Error Recovery
    
    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UserControl_Initialize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
    SetWindowPos sci, 0, 0, 0, UserControl.Width / 15, _
                 UserControl.Height / 15, 0
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error GoTo UserControl_Terminate_Err
    '</EhHeader>
    Subclass_StopAll
    DestroyWindow sci
    FunPointer = 0
    WndPointer = 0
    sci = 0
    '<EhFooter>
    Exit Sub

UserControl_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UserControl_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Property Get BorderStyle() As Long
    '<EhHeader>
    On Error GoTo BorderStyle_Err
    '</EhHeader>
    BorderStyle = ctr_style
    '<EhFooter>
    Exit Property

BorderStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BorderStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let BorderStyle(newS As Long)
    '<EhHeader>
    On Error GoTo BorderStyle_Err
    '</EhHeader>
    Dim oldS As Long
    If ctr_style <> newS Then
        Dim tbp As New PropertyBag, tbp2 As New PropertyBag, t As RECT
        A_SendMessage UserControl.hwnd, WM_SETREDRAW, False, 0
        oldS = ctr_style
        UserControl_WriteProperties tbp2
        ctr_style = newS
        UserControl_WriteProperties tbp
        UserControl_Terminate
        UserControl_Initialize
        If sci = 0 Then
            ctr_style = oldS
            UserControl_Initialize
            Set tbp = tbp2
        End If
        A_SendMessage UserControl.hwnd, WM_SETREDRAW, True, 0
        UserControl_Resize
        UserControl_ReadProperties tbp
        t.Top = 0: t.Left = 0: t.Right = -UserControl.Width / 15: t.Bottom = UserControl.Height / 15
        
    End If
    
    '<EhFooter>
    Exit Property

BorderStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BorderStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>

    'WriteToPropBag PropBag, HideSelection
    'WriteToPropBag PropBag, SelectionEnd
    'WriteToPropBag PropBag, SelectionStart
    'WriteToPropBag PropBag, Anchor
    'WriteToPropBag PropBag, CurrentPos
    'WriteToPropBag PropBag, UndoCollection
    'WriteToPropBag PropBag, ErrStatus
    'WriteToPropBag PropBag, EnableOverType
    'WriteToPropBag PropBag, TargetEnd
    'WriteToPropBag PropBag, TargetStart
    'WriteToPropBag PropBag, StyleBits
    'WriteToPropBag PropBag, ReadOlny
    'WriteToPropBag PropBag, Text
    'WriteToPropBag PropBag, SelectionMode
    
    If lppw = -1 Then
        'load defaults
        Text = strDefText
    End If
    
    lppw = 0
    WriteToPropBag PropBag, False
    WriteToPropBag PropBag, MOUSEDWELLTIME
    WriteToPropBag PropBag, MODEVENTMASK
    WriteToPropBag PropBag, EDGECOLOUR
    WriteToPropBag PropBag, EDGECOLUMN
    WriteToPropBag PropBag, edgeMode
    WriteToPropBag PropBag, ZOOM
    WriteToPropBag PropBag, wrapVisualFlagsLocation
    WriteToPropBag PropBag, LAYOUTCACHE
    WriteToPropBag PropBag, WRAPSTARTINDENT
    WriteToPropBag PropBag, wrapVisualFlags
    WriteToPropBag PropBag, wrapMode
    'WriteToPropBag PropBag, FOLDEXPANDED
    'WriteToPropBag PropBag, FOLDLEVEL
    'WriteToPropBag PropBag, DOCPOINTER
    WriteToPropBag PropBag, PRINTWRAPMODE
    WriteToPropBag PropBag, PRINTCOLOURMODE
    WriteToPropBag PropBag, PRINTMAGNIFICATION
    WriteToPropBag PropBag, HIGHLIGHTGUIDE
    WriteToPropBag PropBag, INDENTATIONGUIDES
    'WriteToPropBag PropBag, LINEINDENTATION
    WriteToPropBag PropBag, BACKSPACEUNINDENTS
    WriteToPropBag PropBag, tabIndents
    WriteToPropBag PropBag, Indent
    WriteToPropBag PropBag, useTabs
    WriteToPropBag PropBag, TabWidth
    'WriteToPropBag PropBag, Focus
    WriteToPropBag PropBag, codepage
    WriteToPropBag PropBag, TwoPhaseDraw
    WriteToPropBag PropBag, BufferEdDraw
    WriteToPropBag PropBag, UsePalette
    WriteToPropBag PropBag, MarginRight
    WriteToPropBag PropBag, MarginLeft
    'WriteToPropBag PropBag, MarginSensitive
    'WriteToPropBag PropBag, MarginMask
    'WriteToPropBag PropBag, MarginWidth
    'WriteToPropBag PropBag, MarginType
    WriteToPropBag PropBag, ControlCharSymbol
    WriteToPropBag PropBag, CaretWidth
    WriteToPropBag PropBag, CaretPeriod
    WriteToPropBag PropBag, CaretLineBack
    WriteToPropBag PropBag, CaretLineVisible
    WriteToPropBag PropBag, CaretFore
    'WriteToPropBag PropBag, LineState
    WriteToPropBag PropBag, ViewEOL
    WriteToPropBag PropBag, MouseDownCaptures
    WriteToPropBag PropBag, ViewWhiteSpace
    WriteToPropBag PropBag, EndAtLastLine
    WriteToPropBag PropBag, ScrollWidth
    WriteToPropBag PropBag, xOffset
    WriteToPropBag PropBag, VScrollBar
    WriteToPropBag PropBag, HScrollBar
    WriteToPropBag PropBag, HideSelection
    WriteToPropBag PropBag, SelectionEnd
    WriteToPropBag PropBag, SelectionStart
    WriteToPropBag PropBag, Anchor
    WriteToPropBag PropBag, CurrentPos
    WriteToPropBag PropBag, UndoCollection
    WriteToPropBag PropBag, ErrStatus
    WriteToPropBag PropBag, EnableOverType
    WriteToPropBag PropBag, TargetEnd
    WriteToPropBag PropBag, TargetStart
    WriteToPropBag PropBag, StyleBits
    WriteToPropBag PropBag, ReadOlny
    WriteToPropBag PropBag, Text
    WriteToPropBag PropBag, BorderStyle
    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UserControl_WriteProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
    lppw = 0
    'HideSelection = ReadfromPropBag(PropBag, False)
    'SelectionEnd = ReadfromPropBag(PropBag, 0)
    'SelectionStart = ReadfromPropBag(PropBag, 0)
    'Anchor = ReadfromPropBag(PropBag, 0)
    'CurrentPos = ReadfromPropBag(PropBag, 0)
    'UndoCollection = ReadfromPropBag(PropBag, True)
    'ErrStatus = ReadfromPropBag(PropBag, 0)
    'EnableOverType = ReadfromPropBag(PropBag, False)
    'TargetEnd = ReadfromPropBag(PropBag, 0)
    'TargetStart = ReadfromPropBag(PropBag, 0)
    'StyleBits = ReadfromPropBag(PropBag, 5)
    'ReadOlny = ReadfromPropBag(PropBag, False)
    'Text = ReadfromPropBag(PropBag, strDefText)
    'SelectionMode = ReadfromPropBag(PropBag, sci_SelectionMode.Sel_Stream)
    
    'If ReadfromPropBag(PropBag, True) Then lppw = -1: UserControl_WriteProperties PropBag
    
    lppw = 1
    MOUSEDWELLTIME = ReadfromPropBag(PropBag, 10000000)
    MODEVENTMASK = ReadfromPropBag(PropBag, 3959)
    EDGECOLOUR = ReadfromPropBag(PropBag, 12632256)
    EDGECOLUMN = ReadfromPropBag(PropBag, 0)
    edgeMode = ReadfromPropBag(PropBag, 0)
    ZOOM = ReadfromPropBag(PropBag, 0)
    wrapVisualFlagsLocation = ReadfromPropBag(PropBag, 0)
    LAYOUTCACHE = ReadfromPropBag(PropBag, 1)
    WRAPSTARTINDENT = ReadfromPropBag(PropBag, 0)
    wrapVisualFlags = ReadfromPropBag(PropBag, 0)
    wrapMode = ReadfromPropBag(PropBag, 0)
    'FOLDEXPANDED( = ReadfromPropBag(PropBag)
    'FOLDLEVEL = ReadfromPropBag(PropBag)
    'DOCPOINTER = ReadfromPropBag(PropBag)
    PRINTWRAPMODE = val(ReadfromPropBag(PropBag, 0))
    PRINTCOLOURMODE = ReadfromPropBag(PropBag, 1)
    PRINTMAGNIFICATION = ReadfromPropBag(PropBag, 0)
    HIGHLIGHTGUIDE = ReadfromPropBag(PropBag, 0)
    INDENTATIONGUIDES = ReadfromPropBag(PropBag, False)
    'LINEINDENTATION = ReadfromPropBag(PropBag)
    BACKSPACEUNINDENTS = ReadfromPropBag(PropBag, False)
    tabIndents = ReadfromPropBag(PropBag, True)
    Indent = ReadfromPropBag(PropBag, 0)
    useTabs = ReadfromPropBag(PropBag, True)
    TabWidth = ReadfromPropBag(PropBag, 8)
    'Focus = ReadfromPropBag(PropBag)
    codepage = ReadfromPropBag(PropBag, 0)
    TwoPhaseDraw = ReadfromPropBag(PropBag, True)
    BufferEdDraw = ReadfromPropBag(PropBag, True)
    UsePalette = ReadfromPropBag(PropBag, False)
    MarginRight = ReadfromPropBag(PropBag, 1)
    MarginLeft = ReadfromPropBag(PropBag, 1)
    'MarginSensitive = ReadfromPropBag(PropBag)
'    MarginMask = ReadfromPropBag(PropBag)
'    MarginWidth = ReadfromPropBag(PropBag)
'    MarginType = ReadfromPropBag(PropBag)
    ControlCharSymbol = ReadfromPropBag(PropBag, 0)
    CaretWidth = ReadfromPropBag(PropBag, 1)
    CaretPeriod = ReadfromPropBag(PropBag, 500)
    CaretLineBack = ReadfromPropBag(PropBag, 65535)
    CaretLineVisible = ReadfromPropBag(PropBag, False)
    CaretFore = ReadfromPropBag(PropBag, 0)
'    LineState = ReadfromPropBag(PropBag)
    ViewEOL = ReadfromPropBag(PropBag, False)
    MouseDownCaptures = ReadfromPropBag(PropBag, True)
    ViewWhiteSpace = ReadfromPropBag(PropBag, 0)
    EndAtLastLine = ReadfromPropBag(PropBag, True)
    ScrollWidth = ReadfromPropBag(PropBag, 200)
    xOffset = ReadfromPropBag(PropBag, 0)
    VScrollBar = ReadfromPropBag(PropBag, True)
    HScrollBar = ReadfromPropBag(PropBag, True)
    HideSelection = ReadfromPropBag(PropBag, False)
    SelectionEnd = ReadfromPropBag(PropBag, 0)
    SelectionStart = ReadfromPropBag(PropBag, 0)
    Anchor = ReadfromPropBag(PropBag, 0)
    CurrentPos = ReadfromPropBag(PropBag, 0)
    UndoCollection = ReadfromPropBag(PropBag, True)
    ErrStatus = ReadfromPropBag(PropBag, 0)
    EnableOverType = ReadfromPropBag(PropBag, False)
    TargetEnd = ReadfromPropBag(PropBag, 0)
    TargetStart = ReadfromPropBag(PropBag, 0)
    StyleBits = ReadfromPropBag(PropBag, 5)
    ReadOlny = ReadfromPropBag(PropBag, False)
    Text = ReadfromPropBag(PropBag, strDefText)
    BorderStyle = ReadfromPropBag(PropBag, 512)
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'SCI_GETTEXT(int length, char *text)
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo Text_Err
    '</EhHeader>
  Dim Temp As Long, tstr() As Byte

    Temp = Send_SCI_message(SCI_GETTEXT, 0, 0)
    If Temp < 2 Then Text = vbNullString: Exit Property
    ReDim tstr(Temp - 1)
    Send_SCI_message SCI_GETTEXT, Temp, VarPtr(tstr(0))
    ReDim Preserve tstr(Temp - 2)
    Text = StrConv(tstr, vbUnicode)
  
    '<EhFooter>
    Exit Property

Text_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Text " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETTEXT(<unused>, const char *text)
Public Property Let Text(newText As String)
    '<EhHeader>
    On Error GoTo Text_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_SETTEXT, 0, newText
  
    '<EhFooter>
    Exit Property

Text_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Text " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETSAVEPOINT
Public Function SetSavePoint()
    '<EhHeader>
    On Error GoTo SetSavePoint_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETSAVEPOINT, 0, 0
  
    '<EhFooter>
    Exit Function

SetSavePoint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SetSavePoint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETLINE(int line, char *text)
Public Function GetLine(numLine As Long) As String
    '<EhHeader>
    On Error GoTo GetLine_Err
    '</EhHeader>
    Dim tstr() As Byte, ln As Long
    
        ln = GetLineLen(numLine)
        If ln < 1 Then GetLine = vbNullString: Exit Function
        
        ReDim Preserve tstr(ln - 1)
        Send_SCI_message SCI_GETLINE, numLine, VarPtr(tstr(0))
        GetLine = StrConv(tstr, vbUnicode)
  
    '<EhFooter>
    Exit Function

GetLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


'SCI_REPLACESEL(<unused>, const char *text)
Public Sub SetSelectedText(tstr As String)
    '<EhHeader>
    On Error GoTo SetSelectedText_Err
    '</EhHeader>

    Send_SCI_messageStr SCI_REPLACESEL, 0, tstr

    '<EhFooter>
    Exit Sub

SetSelectedText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SetSelectedText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETREADONLY(bool readOnly)
Public Property Let ReadOlny(setRO As Boolean)
    '<EhHeader>
    On Error GoTo ReadOlny_Err
    '</EhHeader>
    
    Send_SCI_message SCI_SETREADONLY, setRO And 1, 0
    
    '<EhFooter>
    Exit Property

ReadOlny_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ReadOlny " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETREADONLY
Public Property Get ReadOlny() As Boolean
Attribute ReadOlny.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo ReadOlny_Err
    '</EhHeader>
    
    ReadOlny = Send_SCI_message(SCI_GETREADONLY, 0, 0)
    
    '<EhFooter>
    Exit Property

ReadOlny_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ReadOlny " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETTEXTRANGE(<unused>, TextRange *tr)
Public Function GetTextRange(FromChar As Long, ToChar As Long) As String
    '<EhHeader>
    On Error GoTo GetTextRange_Err
    '</EhHeader>
    Dim Temp As sci_TextRange
    
    If (ToChar - FromChar) < 1 Then Exit Function
    
    Temp.chrg.cpMin = FromChar
    Temp.chrg.cpMax = ToChar
    Temp.lpstrText = StrConv(Space(ToChar - FromChar + 1), vbFromUnicode)
    
    Send_SCI_message SCI_GETTEXTRANGE, 0, VarPtr(Temp)
    
    GetTextRange = Strings.Left$(StrConv(Temp.lpstrText, vbUnicode), ToChar - FromChar)
    
    '<EhFooter>
    Exit Function

GetTextRange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetTextRange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETSTYLEDTEXT(<unused>, TextRange *tr)
Public Function GetStyledTextRange(FromChar As Long, ToChar As Long) As Byte()
    '<EhHeader>
    On Error GoTo GetStyledTextRange_Err
    '</EhHeader>
    Dim Temp As sci_TextRange_Arr, tat() As Byte
    
    If (ToChar - FromChar) < 1 Then Exit Function
    
    Temp.chrg.cpMin = FromChar
    Temp.chrg.cpMax = ToChar
    ReDim tat((ToChar - FromChar) * 2 + 1)
    Temp.lpstrText = VarPtr(tat(0))
    
    Send_SCI_message SCI_GETSTYLEDTEXT, 0, VarPtr(Temp)
    
    ReDim Preserve tat((ToChar - FromChar) * 2 - 1)
    
    GetStyledTextRange = tat
    
    '<EhFooter>
    Exit Function

GetStyledTextRange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetStyledTextRange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_ALLOCATE(int bytes, <unused>)
Public Sub Allocate(bufSize As Long)
    '<EhHeader>
    On Error GoTo Allocate_Err
    '</EhHeader>
    
    Send_SCI_message SCI_ALLOCATE, bufSize, 0
    
    '<EhFooter>
    Exit Sub

Allocate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Allocate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_ADDTEXT(int length, const char *s)
Public Sub AddText(txt As String, Optional txtLen As Long = -1)
    '<EhHeader>
    On Error GoTo AddText_Err
    '</EhHeader>
    
    If txtLen = -1 Then txtLen = Len(txt)
    
    Send_SCI_messageStr SCI_ADDTEXT, txtLen, txt
    
    '<EhFooter>
    Exit Sub

AddText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AddText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_ADDSTYLEDTEXT(int length, cell *s)
Public Sub AddStyledText(StyledTxt() As Byte, Optional txtLen As Long = -1)
    '<EhHeader>
    On Error GoTo AddStyledText_Err
    '</EhHeader>
    
    If txtLen = -1 Then txtLen = UBound(StyledTxt) + 1
    
    Send_SCI_message SCI_ADDSTYLEDTEXT, txtLen, VarPtr(StyledTxt(0))
    
    '<EhFooter>
    Exit Sub

AddStyledText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AddStyledText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_APPENDTEXT(int length, const char *s)
Public Sub AppendText(txt As String, Optional txtLen As Long = -1)
    '<EhHeader>
    On Error GoTo AppendText_Err
    '</EhHeader>
    
    If txtLen = -1 Then txtLen = Len(txt)
    
    Send_SCI_messageStr SCI_APPENDTEXT, txtLen, txt
    
    '<EhFooter>
    Exit Sub

AppendText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AppendText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_INSERTTEXT(int pos, const char *text)
Public Sub InsertText(txt As String, Optional posLen As Long = -1)
    '<EhHeader>
    On Error GoTo InsertText_Err
    '</EhHeader>
    
    Send_SCI_messageStr SCI_INSERTTEXT, posLen, txt
    
    '<EhFooter>
    Exit Sub

InsertText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.InsertText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_CLEARALL
Public Sub ClearText()
    '<EhHeader>
    On Error GoTo ClearText_Err
    '</EhHeader>
    
    Send_SCI_message SCI_CLEARALL, 0, 0
    
    '<EhFooter>
    Exit Sub

ClearText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClearText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_CLEARDOCUMENTSTYLE
Public Sub ClearStyle()
    '<EhHeader>
    On Error GoTo ClearStyle_Err
    '</EhHeader>
    
    Send_SCI_message SCI_CLEARDOCUMENTSTYLE, 0, 0
    
    '<EhFooter>
    Exit Sub

ClearStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClearStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_GETCHARAT(int position)
Public Function GetCharAt(POS As Long) As String
    '<EhHeader>
    On Error GoTo GetCharAt_Err
    '</EhHeader>

    GetCharAt = Chr(Send_SCI_message(SCI_GETCHARAT, POS, 0))
    
    '<EhFooter>
    Exit Function

GetCharAt_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetCharAt " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETSTYLEAT(int position)
Public Function GetStyleAt(POS As Long) As Long
    '<EhHeader>
    On Error GoTo GetStyleAt_Err
    '</EhHeader>

    GetStyleAt = Send_SCI_message(SCI_GETSTYLEAT, POS, 0)
    
    '<EhFooter>
    Exit Function

GetStyleAt_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetStyleAt " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETSTYLEBITS(int bits)
Public Property Let StyleBits(Bits As Long)
    '<EhHeader>
    On Error GoTo StyleBits_Err
    '</EhHeader>

    Send_SCI_message SCI_SETSTYLEBITS, Bits, 0
    
    '<EhFooter>
    Exit Property

StyleBits_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleBits " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETSTYLEBITS
Public Property Get StyleBits() As Long
Attribute StyleBits.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo StyleBits_Err
    '</EhHeader>

    StyleBits = Send_SCI_message(SCI_GETSTYLEBITS, 0, 0)
    
    '<EhFooter>
    Exit Property

StyleBits_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleBits " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'Searching:
'

'SCI_FINDTEXT(int flags, TextToFind *ttf)
Public Function FindText(flags As sci_SearchFlags, cpMin As Long, cpMax As Long, TextToFind As String) As Long()
    '<EhHeader>
    On Error GoTo FindText_Err
    '</EhHeader>
Dim Temp As sci_TextToFind, tmp() As Byte, rez(1) As Long
    
    tmp = StrConv(TextToFind, vbFromUnicode)
    ReDim Preserve tmp(UBound(tmp) + 1)
    
    tmp(UBound(tmp)) = 0
    Temp.chrg.cpMin = cpMin
    Temp.chrg.cpMax = cpMax
    Temp.lpstrText = VarPtr(tmp(0))

    Send_SCI_message SCI_FINDTEXT, flags, VarPtr(Temp)
    rez(0) = Temp.chrgText.cpMin
    rez(1) = Temp.chrgText.cpMax
    FindText = rez
    
    '<EhFooter>
    Exit Function

FindText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FindText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SEARCHANCHOR
Public Sub SearchAnchor()
    '<EhHeader>
    On Error GoTo SearchAnchor_Err
    '</EhHeader>

    Send_SCI_message SCI_SEARCHANCHOR, 0, 0
    
    '<EhFooter>
    Exit Sub

SearchAnchor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SearchAnchor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SEARCHNEXT(int searchFlags, const char *text)
Public Function SearchNext(flags As sci_SearchFlags, TextToFind As String) As Long
    '<EhHeader>
    On Error GoTo SearchNext_Err
    '</EhHeader>
Dim tmp() As Byte
    
    tmp = StrConv(TextToFind, vbFromUnicode)
    ReDim Preserve tmp(UBound(tmp) + 1)
    tmp(UBound(tmp)) = 0
    SearchNext = Send_SCI_message(SCI_SEARCHNEXT, flags, VarPtr(tmp(0)))
    
    '<EhFooter>
    Exit Function

SearchNext_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SearchNext " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SEARCHPREV(int searchFlags, const char *text)
Public Function SearchPrev(flags As sci_SearchFlags, TextToFind As String) As Long
    '<EhHeader>
    On Error GoTo SearchPrev_Err
    '</EhHeader>
Dim tmp() As Byte
    
    tmp = StrConv(TextToFind, vbFromUnicode)
    ReDim Preserve tmp(UBound(tmp) + 1)
    tmp(UBound(tmp)) = 0
    SearchPrev = Send_SCI_message(SCI_SEARCHPREV, flags, VarPtr(tmp(0)))

    '<EhFooter>
    Exit Function

SearchPrev_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SearchPrev " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Search and replace using the target :
'

'SCI_GETTARGETSTART
Public Property Get TargetStart() As Long
Attribute TargetStart.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo TargetStart_Err
    '</EhHeader>

    TargetStart = Send_SCI_message(SCI_GETTARGETSTART, 0, 0)
    
    '<EhFooter>
    Exit Property

TargetStart_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TargetStart " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETTARGETSTART(int pos)
Public Property Let TargetStart(POS As Long)
    '<EhHeader>
    On Error GoTo TargetStart_Err
    '</EhHeader>
    
    Send_SCI_message SCI_SETTARGETSTART, POS, 0
    
    '<EhFooter>
    Exit Property

TargetStart_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TargetStart " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETTARGETEND
Public Property Get TargetEnd() As Long
Attribute TargetEnd.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo TargetEnd_Err
    '</EhHeader>

    TargetEnd = Send_SCI_message(SCI_GETTARGETEND, 0, 0)
    
    '<EhFooter>
    Exit Property

TargetEnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TargetEnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETTARGETEND(int pos)
Public Property Let TargetEnd(POS As Long)
    '<EhHeader>
    On Error GoTo TargetEnd_Err
    '</EhHeader>
    
    Send_SCI_message SCI_SETTARGETEND, POS, 0
    
    '<EhFooter>
    Exit Property

TargetEnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TargetEnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_TARGETFROMSELECTION
Public Sub TargetFromSelection()
    '<EhHeader>
    On Error GoTo TargetFromSelection_Err
    '</EhHeader>

    Send_SCI_message SCI_TARGETFROMSELECTION, 0, 0

    '<EhFooter>
    Exit Sub

TargetFromSelection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TargetFromSelection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'SCI_GETSEARCHFLAGS
Public Property Get SearchFlags() As sci_SearchFlags
    '<EhHeader>
    On Error GoTo SearchFlags_Err
    '</EhHeader>

    SearchFlags = Send_SCI_message(SCI_GETSEARCHFLAGS, 0, 0)
    
    '<EhFooter>
    Exit Property

SearchFlags_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SearchFlags " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETSEARCHFLAGS(int searchFlags)
Public Property Let SearchFlags(flags As sci_SearchFlags)
    '<EhHeader>
    On Error GoTo SearchFlags_Err
    '</EhHeader>
    
    Send_SCI_message SCI_SETSEARCHFLAGS, flags, 0
    
    '<EhFooter>
    Exit Property

SearchFlags_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SearchFlags " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


'SCI_SEARCHINTARGET(int length, const char *text)
Public Function SearchinTarget(strToFind As String) As Long
    '<EhHeader>
    On Error GoTo SearchinTarget_Err
    '</EhHeader>

    SearchinTarget = Send_SCI_messageStr(SCI_SEARCHINTARGET, Len(strToFind), strToFind)

    '<EhFooter>
    Exit Function

SearchinTarget_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SearchinTarget " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_REPLACETARGET(int length, const char *text)
Public Function ReplaceTarget(strWithText As String) As Long
    '<EhHeader>
    On Error GoTo ReplaceTarget_Err
    '</EhHeader>

    ReplaceTarget = Send_SCI_messageStr(SCI_REPLACETARGET, Len(strWithText), strWithText)

    '<EhFooter>
    Exit Function

ReplaceTarget_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ReplaceTarget " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_REPLACETARGETRE(int length, const char *text)
Public Function ReplaceTargetRegEx(strWithText As String) As Long
    '<EhHeader>
    On Error GoTo ReplaceTargetRegEx_Err
    '</EhHeader>

    ReplaceTargetRegEx = Send_SCI_messageStr(SCI_REPLACETARGETRE, Len(strWithText), strWithText)

    '<EhFooter>
    Exit Function

ReplaceTargetRegEx_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ReplaceTargetRegEx " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Overtype:
'

'SCI_GETOVERTYPE
Public Property Get EnableOverType() As Boolean
Attribute EnableOverType.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo EnableOverType_Err
    '</EhHeader>

    EnableOverType = Send_SCI_message(SCI_GETOVERTYPE, 0, 0)

    '<EhFooter>
    Exit Property

EnableOverType_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EnableOverType " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETOVERTYPE(bool overType)
Public Property Let EnableOverType(val As Boolean)
    '<EhHeader>
    On Error GoTo EnableOverType_Err
    '</EhHeader>

    Send_SCI_message SCI_SETOVERTYPE, val And 1, 0

    '<EhFooter>
    Exit Property

EnableOverType_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EnableOverType " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'Cut, copy and paste:
'

'Std clipboard funct
'Clip*
'SCI_CUT
Public Sub ClipCut()
    '<EhHeader>
    On Error GoTo ClipCut_Err
    '</EhHeader>

    Send_SCI_message SCI_CUT, 0, 0
    
    '<EhFooter>
    Exit Sub

ClipCut_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClipCut " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_COPY
Public Sub ClipCopy()
    '<EhHeader>
    On Error GoTo ClipCopy_Err
    '</EhHeader>

    Send_SCI_message SCI_COPY, 0, 0

    '<EhFooter>
    Exit Sub

ClipCopy_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClipCopy " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_PASTE
Public Sub ClipPase()
    '<EhHeader>
    On Error GoTo ClipPase_Err
    '</EhHeader>

    Send_SCI_message SCI_PASTE, 0, 0

    '<EhFooter>
    Exit Sub

ClipPase_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClipPase " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_CLEAR
Public Sub ClipClear()
    '<EhHeader>
    On Error GoTo ClipClear_Err
    '</EhHeader>

    Send_SCI_message SCI_CLEAR, 0, 0
    
    '<EhFooter>
    Exit Sub

ClipClear_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClipClear " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_CANPASTE
Public Sub ClipCanPaste()
    '<EhHeader>
    On Error GoTo ClipCanPaste_Err
    '</EhHeader>
    
        Send_SCI_message SCI_CANPASTE, 0, 0

    '<EhFooter>
    Exit Sub

ClipCanPaste_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClipCanPaste " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_COPYRANGE(int start, end)
Public Sub ClipCopyRange(FromChar As Long, ToChar As Long)
    '<EhHeader>
    On Error GoTo ClipCopyRange_Err
    '</EhHeader>
    
        Send_SCI_message SCI_COPYRANGE, FromChar, ToChar

    '<EhFooter>
    Exit Sub

ClipCopyRange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClipCopyRange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_COPYTEXT(int length, const char *text)
Public Sub ClipCopyText(strText As String)
    '<EhHeader>
    On Error GoTo ClipCopyText_Err
    '</EhHeader>
    
        Send_SCI_messageStr SCI_COPYTEXT, Len(strText), strText

    '<EhFooter>
    Exit Sub

ClipCopyText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ClipCopyText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Error handling:
'

'SCI_GETSTATUS
Public Property Get ErrStatus() As Long
Attribute ErrStatus.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo ErrStatus_Err
    '</EhHeader>

    ErrStatus = Send_SCI_message(SCI_GETSTATUS, 0, 0)

    '<EhFooter>
    Exit Property

ErrStatus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ErrStatus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETSTATUS(int status)
Public Property Let ErrStatus(newStat As Long)
    '<EhHeader>
    On Error GoTo ErrStatus_Err
    '</EhHeader>

    Send_SCI_message SCI_SETSTATUS, newStat, 0

    '<EhFooter>
    Exit Property

ErrStatus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ErrStatus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'From here , most of the code body  is generated using a tool
'That parsed the SCI help and created most of the code (i just did type fixing and
'ansi/unicode/pointer convertions ;) )
'
'

'SCI_UNDO
Public Function Undo() As Long
    '<EhHeader>
    On Error GoTo Undo_Err
    '</EhHeader>
  
  Undo = Send_SCI_message(SCI_UNDO, 0, 0)
  
    '<EhFooter>
    Exit Function

Undo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Undo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CANUNDO
Public Function CanUndo() As Boolean
    '<EhHeader>
    On Error GoTo CanUndo_Err
    '</EhHeader>
  
  CanUndo = Send_SCI_message(SCI_CANUNDO, 0, 0)
  
    '<EhFooter>
    Exit Function

CanUndo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CanUndo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_EMPTYUNDOBUFFER
Public Function EmptyUndoBuffer() As Long
    '<EhHeader>
    On Error GoTo EmptyUndoBuffer_Err
    '</EhHeader>
  
  EmptyUndoBuffer = Send_SCI_message(SCI_EMPTYUNDOBUFFER, 0, 0)
  
    '<EhFooter>
    Exit Function

EmptyUndoBuffer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EmptyUndoBuffer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_REDO
Public Function Redo() As Long
    '<EhHeader>
    On Error GoTo Redo_Err
    '</EhHeader>
  
  Redo = Send_SCI_message(SCI_REDO, 0, 0)
  
    '<EhFooter>
    Exit Function

Redo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Redo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CANREDO
Public Function CanRedo() As Boolean
    '<EhHeader>
    On Error GoTo CanRedo_Err
    '</EhHeader>
  
  CanRedo = Send_SCI_message(SCI_CANREDO, 0, 0)
  
    '<EhFooter>
    Exit Function

CanRedo_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CanRedo " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETUNDOCOLLECTION(bool collectUndo)
Public Property Let UndoCollection(CollectUndo As Boolean)
    '<EhHeader>
    On Error GoTo UndoCollection_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETUNDOCOLLECTION, CollectUndo And 1, 0
  
    '<EhFooter>
    Exit Property

UndoCollection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UndoCollection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETUNDOCOLLECTION
Public Property Get UndoCollection() As Boolean
Attribute UndoCollection.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo UndoCollection_Err
    '</EhHeader>
  
  UndoCollection = Send_SCI_message(SCI_GETUNDOCOLLECTION, 0, 0)
  
    '<EhFooter>
    Exit Property

UndoCollection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UndoCollection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_BEGINUNDOACTION
Public Function BeginUndoAction() As Long
    '<EhHeader>
    On Error GoTo BeginUndoAction_Err
    '</EhHeader>
  
  BeginUndoAction = Send_SCI_message(SCI_BEGINUNDOACTION, 0, 0)
  
    '<EhFooter>
    Exit Function

BeginUndoAction_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BeginUndoAction " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_ENDUNDOACTION
Public Function EndUndoAction() As Long
    '<EhHeader>
    On Error GoTo EndUndoAction_Err
    '</EhHeader>
  
  EndUndoAction = Send_SCI_message(SCI_ENDUNDOACTION, 0, 0)
  
    '<EhFooter>
    Exit Function

EndUndoAction_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EndUndoAction " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETTEXTLENGTH
Public Function TextLength() As Long
    '<EhHeader>
    On Error GoTo TextLength_Err
    '</EhHeader>
  
  TextLength = Send_SCI_message(SCI_GETTEXTLENGTH, 0, 0)
  
    '<EhFooter>
    Exit Function

TextLength_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TextLength " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETLENGTH
Public Function length() As Long
    '<EhHeader>
    On Error GoTo length_Err
    '</EhHeader>
  
  length = Send_SCI_message(SCI_GETLENGTH, 0, 0)
  
    '<EhFooter>
    Exit Function

length_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.length " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETLINECOUNT
Public Function LineCount() As Long
    '<EhHeader>
    On Error GoTo LineCount_Err
    '</EhHeader>
  
  LineCount = Send_SCI_message(SCI_GETLINECOUNT, 0, 0)
  
    '<EhFooter>
    Exit Function

LineCount_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LineCount " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETFIRSTVISIBLELINE
Public Function FirstVisibleLine() As Long
    '<EhHeader>
    On Error GoTo FirstVisibleLine_Err
    '</EhHeader>
  
  FirstVisibleLine = Send_SCI_message(SCI_GETFIRSTVISIBLELINE, 0, 0)
  
    '<EhFooter>
    Exit Function

FirstVisibleLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FirstVisibleLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_LINESONSCREEN
Public Function LinesOnScreen() As Long
    '<EhHeader>
    On Error GoTo LinesOnScreen_Err
    '</EhHeader>
  
  LinesOnScreen = Send_SCI_message(SCI_LINESONSCREEN, 0, 0)
  
    '<EhFooter>
    Exit Function

LinesOnScreen_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LinesOnScreen " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETMODIFY
Public Function Modify() As Boolean
    '<EhHeader>
    On Error GoTo Modify_Err
    '</EhHeader>
  
  Modify = Send_SCI_message(SCI_GETMODIFY, 0, 0)
  
    '<EhFooter>
    Exit Function

Modify_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Modify " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETSEL(int anchorPos, currentPos)
Public Sub Selection(anchorPos As Long, CurrentPos As Long)
    '<EhHeader>
    On Error GoTo Selection_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETSEL, anchorPos, CurrentPos
  
    '<EhFooter>
    Exit Sub

Selection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Selection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_GOTOPOS(int position)
Public Function GotoPos(position As Long) As Long
    '<EhHeader>
    On Error GoTo GotoPos_Err
    '</EhHeader>
  
  GotoPos = Send_SCI_message(SCI_GOTOPOS, position, 0)
  
    '<EhFooter>
    Exit Function

GotoPos_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GotoPos " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GOTOLINE(int line)
Public Function GotoLine(line As Long) As Long
    '<EhHeader>
    On Error GoTo GotoLine_Err
    '</EhHeader>
  
  GotoLine = Send_SCI_message(SCI_GOTOLINE, line, 0)
  
    '<EhFooter>
    Exit Function

GotoLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GotoLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETCURRENTPOS(int position)
Public Property Let CurrentPos(position As Long)
    '<EhHeader>
    On Error GoTo CurrentPos_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCURRENTPOS, position, 0
  
    '<EhFooter>
    Exit Property

CurrentPos_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CurrentPos " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCURRENTPOS
Public Property Get CurrentPos() As Long
Attribute CurrentPos.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo CurrentPos_Err
    '</EhHeader>
  
  CurrentPos = Send_SCI_message(SCI_GETCURRENTPOS, 0, 0)
  
    '<EhFooter>
    Exit Property

CurrentPos_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CurrentPos " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETANCHOR(int position)
Public Property Let Anchor(position As Long)
    '<EhHeader>
    On Error GoTo Anchor_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETANCHOR, position, 0
  
    '<EhFooter>
    Exit Property

Anchor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Anchor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETANCHOR
Public Property Get Anchor() As Long
Attribute Anchor.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo Anchor_Err
    '</EhHeader>
  
  Anchor = Send_SCI_message(SCI_GETANCHOR, 0, 0)
  
    '<EhFooter>
    Exit Property

Anchor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Anchor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETSELECTIONSTART(int position)
Public Property Let SelectionStart(position As Long)
    '<EhHeader>
    On Error GoTo SelectionStart_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETSELECTIONSTART, position, 0
  
    '<EhFooter>
    Exit Property

SelectionStart_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectionStart " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETSELECTIONSTART
Public Property Get SelectionStart() As Long
Attribute SelectionStart.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo SelectionStart_Err
    '</EhHeader>
  
  SelectionStart = Send_SCI_message(SCI_GETSELECTIONSTART, 0, 0)
  
    '<EhFooter>
    Exit Property

SelectionStart_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectionStart " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETSELECTIONEND(int position)
Public Property Let SelectionEnd(position As Long)
    '<EhHeader>
    On Error GoTo SelectionEnd_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETSELECTIONEND, position, 0
  
    '<EhFooter>
    Exit Property

SelectionEnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectionEnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETSELECTIONEND
Public Property Get SelectionEnd() As Long
Attribute SelectionEnd.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo SelectionEnd_Err
    '</EhHeader>
  
  SelectionEnd = Send_SCI_message(SCI_GETSELECTIONEND, 0, 0)
  
    '<EhFooter>
    Exit Property

SelectionEnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectionEnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SELECTALL
Public Sub SelectAll()
    '<EhHeader>
    On Error GoTo SelectAll_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SELECTALL, 0, 0
  
    '<EhFooter>
    Exit Sub

SelectAll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectAll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_LINEFROMPOSITION(int position)
Public Function LineFromPosition(position As Long) As Long
    '<EhHeader>
    On Error GoTo LineFromPosition_Err
    '</EhHeader>
  
  LineFromPosition = Send_SCI_message(SCI_LINEFROMPOSITION, position, 0)
  
    '<EhFooter>
    Exit Function

LineFromPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LineFromPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_POSITIONFROMLINE(int line)
Public Function PositionFromLine(line As Long) As Long
    '<EhHeader>
    On Error GoTo PositionFromLine_Err
    '</EhHeader>
  
  PositionFromLine = Send_SCI_message(SCI_POSITIONFROMLINE, line, 0)
  
    '<EhFooter>
    Exit Function

PositionFromLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PositionFromLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETLINEENDPOSITION(int line)
Public Function LineEndPosition(line As Long) As Long
    '<EhHeader>
    On Error GoTo LineEndPosition_Err
    '</EhHeader>
  
  LineEndPosition = Send_SCI_message(SCI_GETLINEENDPOSITION, line, 0)
  
    '<EhFooter>
    Exit Function

LineEndPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LineEndPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'heeh this is made by me not auto converted .. heheh lol
'SCI_LINELENGTH(int line).
Public Function GetLineLen(numLine As Long) As Long
    '<EhHeader>
    On Error GoTo GetLineLen_Err
    '</EhHeader>
    
    GetLineLen = Send_SCI_message(SCI_GETLINE, numLine, 0)
    
    '<EhFooter>
    Exit Function

GetLineLen_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetLineLen " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETCOLUMN(int position)
Public Function GetColumn(position As Long) As Long
    '<EhHeader>
    On Error GoTo GetColumn_Err
    '</EhHeader>
  
  GetColumn = Send_SCI_message(SCI_GETCOLUMN, position, 0)
  
    '<EhFooter>
    Exit Function

GetColumn_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetColumn " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_POSITIONFROMPOINT(int x, y)
Public Function PositionFromPoint(X As Long, Y As Long) As Long
    '<EhHeader>
    On Error GoTo PositionFromPoint_Err
    '</EhHeader>
  
  PositionFromPoint = Send_SCI_message(SCI_POSITIONFROMPOINT, X, Y)
  
    '<EhFooter>
    Exit Function

PositionFromPoint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PositionFromPoint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_POSITIONFROMPOINTCLOSE(int x, y)
Public Function PositionFromPointClose(X As Long, Y As Long) As Long
    '<EhHeader>
    On Error GoTo PositionFromPointClose_Err
    '</EhHeader>
  
  PositionFromPointClose = Send_SCI_message(SCI_POSITIONFROMPOINTCLOSE, X, Y)
  
    '<EhFooter>
    Exit Function

PositionFromPointClose_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PositionFromPointClose " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_POINTXFROMPOSITION(<unused>, position)
Public Function PointXFromPosition(position As Long) As Long
    '<EhHeader>
    On Error GoTo PointXFromPosition_Err
    '</EhHeader>
  
  PointXFromPosition = Send_SCI_message(SCI_POINTXFROMPOSITION, 0, position)
  
    '<EhFooter>
    Exit Function

PointXFromPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PointXFromPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_POINTYFROMPOSITION(<unused>, position)
Public Function PointYFromPosition(position As Long) As Long
    '<EhHeader>
    On Error GoTo PointYFromPosition_Err
    '</EhHeader>
  
  PointYFromPosition = Send_SCI_message(SCI_POINTYFROMPOSITION, 0, position)
  
    '<EhFooter>
    Exit Function

PointYFromPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PointYFromPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_HIDESELECTION(bool hide)
Public Property Let HideSelection(Hide As Boolean)
    '<EhHeader>
    On Error GoTo HideSelection_Err
    '</EhHeader>
  
   Send_SCI_message SCI_HIDESELECTION, Hide And 1, 0
   SelVisible = Hide
   
    '<EhFooter>
    Exit Property

HideSelection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HideSelection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'This is an enhasement ... it does not exist on the real control
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo HideSelection_Err
    '</EhHeader>
  
  HideSelection = SelVisible
  
    '<EhFooter>
    Exit Property

HideSelection_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HideSelection " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


'SCI_GETSELTEXT(<unused>, char *text)
Public Function GetSelText() As String
    '<EhHeader>
    On Error GoTo GetSelText_Err
    '</EhHeader>
    Dim Temp As Long, tstr() As Byte
    
    Temp = Send_SCI_message(SCI_GETSELTEXT, 0, 0)
    If Temp < 2 Then GetSelText = vbNullString: Exit Function
    ReDim tstr(Temp - 1)
    Send_SCI_message SCI_GETSELTEXT, 0, VarPtr(tstr(0))
    ReDim Preserve tstr(Temp - 2)
    GetSelText = StrConv(tstr, vbUnicode)
  
    '<EhFooter>
    Exit Function

GetSelText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetSelText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETCURLINE(int textLen, char *text)
Public Function GetCurLine() As String
    '<EhHeader>
    On Error GoTo GetCurLine_Err
    '</EhHeader>
    Dim Temp As Long, tstr() As Byte
    
    Temp = Send_SCI_message(SCI_GETCURLINE, 0, 0)
    If Temp < 2 Then GetCurLine = vbNullString: Exit Function
    ReDim tstr(Temp - 1)
    Send_SCI_message SCI_GETCURLINE, Temp, VarPtr(tstr(0))
    ReDim Preserve tstr(Temp - 2)
    GetCurLine = StrConv(tstr, vbUnicode)
  
    '<EhFooter>
    Exit Function

GetCurLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetCurLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SELECTIONISRECTANGLE
Public Function SelectionIsRectangle() As Boolean
    '<EhHeader>
    On Error GoTo SelectionIsRectangle_Err
    '</EhHeader>
  
  SelectionIsRectangle = Send_SCI_message(SCI_SELECTIONISRECTANGLE, 0, 0)
  
    '<EhFooter>
    Exit Function

SelectionIsRectangle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectionIsRectangle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'A note here :  the docs are wrong..
'This one takes one parameter..
'it should be SCI_SETSELECTIONMODE(int mode) instead of
'SCI_SETSELECTIONMODE
Public Property Let SelectionMode(Mode As sci_SelectionMode)
    '<EhHeader>
    On Error GoTo SelectionMode_Err
    '</EhHeader>

  Send_SCI_message SCI_SETSELECTIONMODE, Mode, 0

    '<EhFooter>
    Exit Property

SelectionMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectionMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETSELECTIONMODE
Public Property Get SelectionMode() As sci_SelectionMode
    '<EhHeader>
    On Error GoTo SelectionMode_Err
    '</EhHeader>
  
  SelectionMode = Send_SCI_message(SCI_GETSELECTIONMODE, 0, 0)
  
    '<EhFooter>
    Exit Property

SelectionMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelectionMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETLINESELSTARTPOSITION(int line)
Public Function GetLineSelStartPosition(line As Long) As Long
    '<EhHeader>
    On Error GoTo GetLineSelStartPosition_Err
    '</EhHeader>
  
  GetLineSelStartPosition = Send_SCI_message(SCI_GETLINESELSTARTPOSITION, line, 0)
  
    '<EhFooter>
    Exit Function

GetLineSelStartPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetLineSelStartPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETLINESELENDPOSITION(int line)
Public Function GetLineSelEndPosition(line As Long) As Long
    '<EhHeader>
    On Error GoTo GetLineSelEndPosition_Err
    '</EhHeader>
  
  GetLineSelEndPosition = Send_SCI_message(SCI_GETLINESELENDPOSITION, line, 0)
  
    '<EhFooter>
    Exit Function

GetLineSelEndPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetLineSelEndPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MOVECARETINSIDEVIEW
Public Function MoveCaretInsideView() As Long
    '<EhHeader>
    On Error GoTo MoveCaretInsideView_Err
    '</EhHeader>
  
  MoveCaretInsideView = Send_SCI_message(SCI_MOVECARETINSIDEVIEW, 0, 0)
  
    '<EhFooter>
    Exit Function

MoveCaretInsideView_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MoveCaretInsideView " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_WORDENDPOSITION(int position, bool onlyWordCharacters)
Public Function WordEndPosition(position As Long, onlyWordCharacters As Boolean) As Long
    '<EhHeader>
    On Error GoTo WordEndPosition_Err
    '</EhHeader>
  
  WordEndPosition = Send_SCI_message(SCI_WORDENDPOSITION, position, onlyWordCharacters And 1)
  
    '<EhFooter>
    Exit Function

WordEndPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WordEndPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_WORDSTARTPOSITION(int position, bool onlyWordCharacters)
Public Function WordStartPosition(position As Long, onlyWordCharacters As Boolean) As Long
    '<EhHeader>
    On Error GoTo WordStartPosition_Err
    '</EhHeader>
  
  WordStartPosition = Send_SCI_message(SCI_WORDSTARTPOSITION, position, onlyWordCharacters And 1)
  
    '<EhFooter>
    Exit Function

WordStartPosition_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WordStartPosition " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_POSITIONBEFORE(int position)
Public Function PositionBefore(position As Long) As Long
    '<EhHeader>
    On Error GoTo PositionBefore_Err
    '</EhHeader>
  
  PositionBefore = Send_SCI_message(SCI_POSITIONBEFORE, position, 0)
  
    '<EhFooter>
    Exit Function

PositionBefore_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PositionBefore " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_POSITIONAFTER(int position)
Public Function POSITIONAFTER(position As Long) As Long
    '<EhHeader>
    On Error GoTo POSITIONAFTER_Err
    '</EhHeader>
  
  POSITIONAFTER = Send_SCI_message(SCI_POSITIONAFTER, position, 0)
  
    '<EhFooter>
    Exit Function

POSITIONAFTER_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.POSITIONAFTER " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_TEXTWIDTH(int styleNumber, const char *text)
Public Function TextWidth(styleNumber As Long, Text As String) As Long
    '<EhHeader>
    On Error GoTo TextWidth_Err
    '</EhHeader>
  
  TextWidth = Send_SCI_messageStr(SCI_TEXTWIDTH, styleNumber, Text)
  
    '<EhFooter>
    Exit Function

TextWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TextWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_TEXTHEIGHT(int line)
Public Function TextHeight(line As Long) As Long
    '<EhHeader>
    On Error GoTo TextHeight_Err
    '</EhHeader>
  
  TextHeight = Send_SCI_message(SCI_TEXTHEIGHT, line, 0)
  
    '<EhFooter>
    Exit Function

TextHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TextHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CHOOSECARETX
Public Function ChooseCaretX() As Long
    '<EhHeader>
    On Error GoTo ChooseCaretX_Err
    '</EhHeader>
  
  ChooseCaretX = Send_SCI_message(SCI_CHOOSECARETX, 0, 0)
  
    '<EhFooter>
    Exit Function

ChooseCaretX_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ChooseCaretX " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function



'SCI_LINESCROLL(int column, line)
Public Function LineScroll(Column As Long, line As Long) As Long
    '<EhHeader>
    On Error GoTo LineScroll_Err
    '</EhHeader>
  
  LineScroll = Send_SCI_message(SCI_LINESCROLL, Column, line)
  
    '<EhFooter>
    Exit Function

LineScroll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LineScroll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SCROLLCARET
Public Function SCROLLCARET() As Long
    '<EhHeader>
    On Error GoTo SCROLLCARET_Err
    '</EhHeader>
  
  SCROLLCARET = Send_SCI_message(SCI_SCROLLCARET, 0, 0)
  
    '<EhFooter>
    Exit Function

SCROLLCARET_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SCROLLCARET " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETXCARETPOLICY(int caretPolicy, caretSlop)
Public Sub SetXCaretPolicy(caretPolicy As Long, caretSlop As Long)
    '<EhHeader>
    On Error GoTo SetXCaretPolicy_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETXCARETPOLICY, caretPolicy, caretSlop
  
    '<EhFooter>
    Exit Sub

SetXCaretPolicy_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SetXCaretPolicy " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETYCARETPOLICY(int caretPolicy, caretSlop)
Public Sub SetYCaretPolicy(caretPolicy As Long, caretSlop As Long)
    '<EhHeader>
    On Error GoTo SetYCaretPolicy_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETYCARETPOLICY, caretPolicy, caretSlop
  
    '<EhFooter>
    Exit Sub

SetYCaretPolicy_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SetYCaretPolicy " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

 
'SCI_SETVISIBLEPOLICY(int caretPolicy, caretSlop)
Public Sub SetVisiblePolicy(caretPolicy As Long, caretSlop As Long)
    '<EhHeader>
    On Error GoTo SetVisiblePolicy_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETVISIBLEPOLICY, caretPolicy, caretSlop
  
    '<EhFooter>
    Exit Sub

SetVisiblePolicy_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SetVisiblePolicy " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETHSCROLLBAR(bool visible)
Public Property Let HScrollBar(bvisible As Boolean)
    '<EhHeader>
    On Error GoTo HScrollBar_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETHSCROLLBAR, bvisible And 1, 0
  
    '<EhFooter>
    Exit Property

HScrollBar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HScrollBar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETHSCROLLBAR
Public Property Get HScrollBar() As Boolean
Attribute HScrollBar.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo HScrollBar_Err
    '</EhHeader>
  
  HScrollBar = Send_SCI_message(SCI_GETHSCROLLBAR, 0, 0)
  
    '<EhFooter>
    Exit Property

HScrollBar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HScrollBar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETVSCROLLBAR(bool visible)
Public Property Let VScrollBar(visible As Boolean)
    '<EhHeader>
    On Error GoTo VScrollBar_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETVSCROLLBAR, visible And 1, 0
  
    '<EhFooter>
    Exit Property

VScrollBar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.VScrollBar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETVSCROLLBAR
Public Property Get VScrollBar() As Boolean
Attribute VScrollBar.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo VScrollBar_Err
    '</EhHeader>
  
  VScrollBar = Send_SCI_message(SCI_GETVSCROLLBAR, 0, 0)
  
    '<EhFooter>
    Exit Property

VScrollBar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.VScrollBar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETXOFFSET
Public Property Get xOffset() As Long
Attribute xOffset.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo xOffset_Err
    '</EhHeader>
  
  xOffset = Send_SCI_message(SCI_GETXOFFSET, 0, 0)
  
    '<EhFooter>
    Exit Property

xOffset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.xOffset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETXOFFSET(int xOffset)
Public Property Let xOffset(xOffset As Long)
    '<EhHeader>
    On Error GoTo xOffset_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETXOFFSET, xOffset, 0
  
    '<EhFooter>
    Exit Property

xOffset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.xOffset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETSCROLLWIDTH(int pixelWidth)
Public Property Let ScrollWidth(pixelWidth As Long)
    '<EhHeader>
    On Error GoTo ScrollWidth_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETSCROLLWIDTH, pixelWidth, 0
  
    '<EhFooter>
    Exit Property

ScrollWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ScrollWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETSCROLLWIDTH
Public Property Get ScrollWidth() As Long
Attribute ScrollWidth.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo ScrollWidth_Err
    '</EhHeader>
  
  ScrollWidth = Send_SCI_message(SCI_GETSCROLLWIDTH, 0, 0)
  
    '<EhFooter>
    Exit Property

ScrollWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ScrollWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETENDATLASTLINE(bool endAtLastLine)
Public Property Let EndAtLastLine(bEndAtLastLine As Boolean)
    '<EhHeader>
    On Error GoTo EndAtLastLine_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETENDATLASTLINE, bEndAtLastLine And 1, 0
  
    '<EhFooter>
    Exit Property

EndAtLastLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EndAtLastLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETENDATLASTLINE
Public Property Get EndAtLastLine() As Boolean
Attribute EndAtLastLine.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo EndAtLastLine_Err
    '</EhHeader>
  
  EndAtLastLine = Send_SCI_message(SCI_GETENDATLASTLINE, 0, 0)
  
    '<EhFooter>
    Exit Property

EndAtLastLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EndAtLastLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETVIEWWS(int wsMode)
Public Property Let ViewWhiteSpace(wsMode As Long)
    '<EhHeader>
    On Error GoTo ViewWhiteSpace_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETVIEWWS, wsMode, 0
  
    '<EhFooter>
    Exit Property

ViewWhiteSpace_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ViewWhiteSpace " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETVIEWWS
Public Property Get ViewWhiteSpace() As Long
Attribute ViewWhiteSpace.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo ViewWhiteSpace_Err
    '</EhHeader>
  
  ViewWhiteSpace = Send_SCI_message(SCI_GETVIEWWS, 0, 0)
  
    '<EhFooter>
    Exit Property

ViewWhiteSpace_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ViewWhiteSpace " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETWHITESPACEFORE(bool useWhitespaceForeColour, colour)
Public Property Let WhiteSpaceFore(useWhitespaceForeColour As Boolean, colour As Long)
    '<EhHeader>
    On Error GoTo WhiteSpaceFore_Err
    '</EhHeader>

  MsgBox "add get"
  Send_SCI_message SCI_SETWHITESPACEFORE, useWhitespaceForeColour And 1, colour
  
    '<EhFooter>
    Exit Property

WhiteSpaceFore_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WhiteSpaceFore " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETWHITESPACEBACK(bool useWhitespaceBackColour, colour)
Public Property Let WhiteSpaceBack(useWhitespaceBackColour As Boolean, colour As Long)
    '<EhHeader>
    On Error GoTo WhiteSpaceBack_Err
    '</EhHeader>
  
  MsgBox "add get"
  Send_SCI_message SCI_SETWHITESPACEBACK, useWhitespaceBackColour And 1, colour
  
    '<EhFooter>
    Exit Property

WhiteSpaceBack_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WhiteSpaceBack " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCURSOR(int curType)
Public Property Let Cursor(curType As sci_CursorStyle)
    '<EhHeader>
    On Error GoTo Cursor_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCURSOR, curType, 0
  
    '<EhFooter>
    Exit Property

Cursor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Cursor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCURSOR
Public Property Get Cursor() As sci_CursorStyle
    '<EhHeader>
    On Error GoTo Cursor_Err
    '</EhHeader>
  
  Cursor = Send_SCI_message(SCI_GETCURSOR, 0, 0)
  
    '<EhFooter>
    Exit Property

Cursor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Cursor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETMOUSEDOWNCAPTURES(bool captures)
Public Property Let MouseDownCaptures(captures As Boolean)
    '<EhHeader>
    On Error GoTo MouseDownCaptures_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMOUSEDOWNCAPTURES, 1 And captures, 0
  
    '<EhFooter>
    Exit Property

MouseDownCaptures_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MouseDownCaptures " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMOUSEDOWNCAPTURES
Public Property Get MouseDownCaptures() As Boolean
Attribute MouseDownCaptures.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MouseDownCaptures_Err
    '</EhHeader>
  
  MouseDownCaptures = Send_SCI_message(SCI_GETMOUSEDOWNCAPTURES, 0, 0)
  
    '<EhFooter>
    Exit Property

MouseDownCaptures_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MouseDownCaptures " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETEOLMODE(int eolMode)
Public Property Let EOLMode(neweolMode As sci_EOLModes)
    '<EhHeader>
    On Error GoTo EOLMode_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETEOLMODE, EOLMode, 0
  
    '<EhFooter>
    Exit Property

EOLMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EOLMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETEOLMODE
Public Property Get EOLMode() As sci_EOLModes
    '<EhHeader>
    On Error GoTo EOLMode_Err
    '</EhHeader>
  
  EOLMode = Send_SCI_message(SCI_GETEOLMODE, 0, 0)
  
    '<EhFooter>
    Exit Property

EOLMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EOLMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_CONVERTEOLS(int eolMode)
Public Sub ConvertEOLs(neolMode As sci_EOLModes)
    '<EhHeader>
    On Error GoTo ConvertEOLs_Err
    '</EhHeader>
  
   Send_SCI_message SCI_CONVERTEOLS, EOLMode, 0
  
    '<EhFooter>
    Exit Sub

ConvertEOLs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ConvertEOLs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETVIEWEOL(bool visible)
Public Property Let ViewEOL(visible As Boolean)
    '<EhHeader>
    On Error GoTo ViewEOL_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETVIEWEOL, 1 And visible, 0
  
    '<EhFooter>
    Exit Property

ViewEOL_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ViewEOL " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETVIEWEOL
Public Property Get ViewEOL() As Boolean
Attribute ViewEOL.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo ViewEOL_Err
    '</EhHeader>
  
  ViewEOL = Send_SCI_message(SCI_GETVIEWEOL, 0, 0)
  
    '<EhFooter>
    Exit Property

ViewEOL_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ViewEOL " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETENDSTYLED
Public Property Get EndStyled() As Long
    '<EhHeader>
    On Error GoTo EndStyled_Err
    '</EhHeader>
  
  EndStyled = Send_SCI_message(SCI_GETENDSTYLED, 0, 0)
  
    '<EhFooter>
    Exit Property

EndStyled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EndStyled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_STARTSTYLING(int position, mask)
Public Sub StartStyling(position As Long, mask As Long)
    '<EhHeader>
    On Error GoTo StartStyling_Err
    '</EhHeader>
  
   Send_SCI_message SCI_STARTSTYLING, position, mask
  
    '<EhFooter>
    Exit Sub

StartStyling_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StartStyling " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETSTYLING(int length, style)
Public Sub SetStyling(length As Long, Style As Long)
    '<EhHeader>
    On Error GoTo SetStyling_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETSTYLING, length, Style
  
    '<EhFooter>
    Exit Sub

SetStyling_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SetStyling " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETSTYLINGEX(int length, const char *styles)
Public Sub SetStylengEx(nam As String)
    '<EhHeader>
    On Error GoTo SetStylengEx_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_SETSTYLINGEX, Len(nam), nam
  
    '<EhFooter>
    Exit Sub

SetStylengEx_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SetStylengEx " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETLINESTATE(int line, value )
Public Property Let LineState(line As Long, Value As Long)
    '<EhHeader>
    On Error GoTo LineState_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETLINESTATE, line, Value
  
    '<EhFooter>
    Exit Property

LineState_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LineState " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETLINESTATE(int line)
Public Property Get LineState(line As Long) As Long
Attribute LineState.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo LineState_Err
    '</EhHeader>
  
  LineState = Send_SCI_message(SCI_GETLINESTATE, line, 0)
  
    '<EhFooter>
    Exit Property

LineState_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LineState " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMAXLINESTATE
Public Function GetMaxLineState() As Long
    '<EhHeader>
    On Error GoTo GetMaxLineState_Err
    '</EhHeader>
  
  GetMaxLineState = Send_SCI_message(SCI_GETMAXLINESTATE, 0, 0)
  
    '<EhFooter>
    Exit Function

GetMaxLineState_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GetMaxLineState " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


'SCI_STYLERESETDEFAULT
Public Sub StyleResetDefault()
    '<EhHeader>
    On Error GoTo StyleResetDefault_Err
    '</EhHeader>
  
   Send_SCI_message SCI_STYLERESETDEFAULT, 0, 0
  
    '<EhFooter>
    Exit Sub

StyleResetDefault_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleResetDefault " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLECLEARALL
Public Sub StyleClearAll()
    '<EhHeader>
    On Error GoTo StyleClearAll_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLECLEARALL, 0, 0
  
    '<EhFooter>
    Exit Sub

StyleClearAll_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleClearAll " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'SCI_STYLESETFONT(int styleNumber, char *fontName)
Public Sub StyleSetFont(styleNumber As Long, fName As String)
    '<EhHeader>
    On Error GoTo StyleSetFont_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_STYLESETFONT, styleNumber, fName
  
    '<EhFooter>
    Exit Sub

StyleSetFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETSIZE(int styleNumber, sizeInPoints)
Public Sub STYLESETSIZE(styleNumber As Long, sizeInPoints As Long)
    '<EhHeader>
    On Error GoTo STYLESETSIZE_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETSIZE, styleNumber, sizeInPoints
  
    '<EhFooter>
    Exit Sub

STYLESETSIZE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.STYLESETSIZE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETBOLD(int styleNumber, bool bold)
Public Sub StyleSetBold(styleNumber As Long, bold As Boolean)
    '<EhHeader>
    On Error GoTo StyleSetBold_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETBOLD, styleNumber, 1 And bold
  
    '<EhFooter>
    Exit Sub

StyleSetBold_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetBold " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETITALIC(int styleNumber, bool italic)
Public Sub StyleSetItalic(styleNumber As Long, italic As Boolean)
    '<EhHeader>
    On Error GoTo StyleSetItalic_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETITALIC, styleNumber, 1 And italic
  
    '<EhFooter>
    Exit Sub

StyleSetItalic_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetItalic " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETUNDERLINE(int styleNumber, bool underline)
Public Sub StyleSetUnderline(styleNumber As Long, underline As Boolean)
    '<EhHeader>
    On Error GoTo StyleSetUnderline_Err
    '</EhHeader>
  
    Send_SCI_message SCI_STYLESETUNDERLINE, styleNumber, 1 And underline
  
    '<EhFooter>
    Exit Sub

StyleSetUnderline_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetUnderline " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETFORE(int styleNumber, colour)
Public Sub StyleSetFore(styleNumber As Long, colour As Long)
    '<EhHeader>
    On Error GoTo StyleSetFore_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETFORE, styleNumber, colour
  
    '<EhFooter>
    Exit Sub

StyleSetFore_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetFore " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETBACK(int styleNumber, colour)
Public Sub StyleSetBack(styleNumber As Long, colour As Long)
    '<EhHeader>
    On Error GoTo StyleSetBack_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETBACK, styleNumber, colour
  
    '<EhFooter>
    Exit Sub

StyleSetBack_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetBack " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETEOLFILLED(int styleNumber, bool eolFilled)
Public Sub STYLESETEOLFILLED(styleNumber As Long, eolFilled As Boolean)
    '<EhHeader>
    On Error GoTo STYLESETEOLFILLED_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETEOLFILLED, styleNumber, eolFilled And 1
  
    '<EhFooter>
    Exit Sub

STYLESETEOLFILLED_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.STYLESETEOLFILLED " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETCHARACTERSET(int styleNumber, charSet)
Public Sub StyleSetCharacterset(styleNumber As Long, charSet As Long)
    '<EhHeader>
    On Error GoTo StyleSetCharacterset_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETCHARACTERSET, styleNumber, charSet
  
    '<EhFooter>
    Exit Sub

StyleSetCharacterset_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetCharacterset " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETCASE(int styleNumber, caseMode)
Public Sub StyleSetCase(styleNumber As Long, caseMode As Long)
    '<EhHeader>
    On Error GoTo StyleSetCase_Err
    '</EhHeader>
     
    Send_SCI_message SCI_STYLESETCASE, styleNumber, caseMode
  
    '<EhFooter>
    Exit Sub

StyleSetCase_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetCase " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETVISIBLE(int styleNumber, bool visible)
Public Sub StyleSetVisible(styleNumber As Long, visible As Boolean)
    '<EhHeader>
    On Error GoTo StyleSetVisible_Err
    '</EhHeader>
  
   Send_SCI_message SCI_STYLESETVISIBLE, styleNumber, visible And 1
  
    '<EhFooter>
    Exit Sub

StyleSetVisible_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetVisible " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETCHANGEABLE(int styleNumber, bool changeable)
Public Sub StyleSetChangeable(styleNumber As Long, changeable As Boolean)
    '<EhHeader>
    On Error GoTo StyleSetChangeable_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETCHANGEABLE, styleNumber, 1 And changeable
  
    '<EhFooter>
    Exit Sub

StyleSetChangeable_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetChangeable " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_STYLESETHOTSPOT(int styleNumber, bool hotspot)
Public Sub StyleSetHotSpot(styleNumber As Long, hotspot As Boolean)
    '<EhHeader>
    On Error GoTo StyleSetHotSpot_Err
    '</EhHeader>
  
  Send_SCI_message SCI_STYLESETHOTSPOT, styleNumber, 1 And hotspot
  
    '<EhFooter>
    Exit Sub

StyleSetHotSpot_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.StyleSetHotSpot " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub



'SCI_SETSELFORE(bool useSelectionForeColour, colour)
Public Property Let SelcectedForeColor(useSelectionForeColour As Boolean, colour As Long)
    '<EhHeader>
    On Error GoTo SelcectedForeColor_Err
    '</EhHeader>

  MsgBox "Add Get"
  Send_SCI_message SCI_SETSELFORE, 1 And useSelectionForeColour, colour
  
    '<EhFooter>
    Exit Property

SelcectedForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelcectedForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETSELBACK(bool useSelectionBackColour, colour)
Public Property Let SelBack(useSelectionBackColour As Boolean, colour As Long)
    '<EhHeader>
    On Error GoTo SelBack_Err
    '</EhHeader>
  
  MsgBox "Add Get"
  Send_SCI_message SCI_SETSELBACK, 1 And useSelectionBackColour, colour
  
    '<EhFooter>
    Exit Property

SelBack_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SelBack " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCARETFORE(int colour)
Public Property Let CaretFore(colour As Long)
    '<EhHeader>
    On Error GoTo CaretFore_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCARETFORE, colour, 0
  
    '<EhFooter>
    Exit Property

CaretFore_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretFore " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCARETFORE
Public Property Get CaretFore() As Long
Attribute CaretFore.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo CaretFore_Err
    '</EhHeader>
  
  CaretFore = Send_SCI_message(SCI_GETCARETFORE, 0, 0)
  
    '<EhFooter>
    Exit Property

CaretFore_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretFore " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCARETLINEVISIBLE(bool show)
Public Property Let CaretLineVisible(show As Boolean)
    '<EhHeader>
    On Error GoTo CaretLineVisible_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCARETLINEVISIBLE, 1 And show, 0
  
    '<EhFooter>
    Exit Property

CaretLineVisible_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretLineVisible " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCARETLINEVISIBLE
Public Property Get CaretLineVisible() As Boolean
Attribute CaretLineVisible.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo CaretLineVisible_Err
    '</EhHeader>
  
  CaretLineVisible = Send_SCI_message(SCI_GETCARETLINEVISIBLE, 0, 0)
  
    '<EhFooter>
    Exit Property

CaretLineVisible_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretLineVisible " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property



'SCI_SETCARETLINEBACK(int colour)
Public Property Let CaretLineBack(colour As Long)
    '<EhHeader>
    On Error GoTo CaretLineBack_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCARETLINEBACK, colour, 0
  
    '<EhFooter>
    Exit Property

CaretLineBack_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretLineBack " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCARETLINEBACK
Public Property Get CaretLineBack() As Long
Attribute CaretLineBack.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo CaretLineBack_Err
    '</EhHeader>
  
  CaretLineBack = Send_SCI_message(SCI_GETCARETLINEBACK, 0, 0)
  
    '<EhFooter>
    Exit Property

CaretLineBack_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretLineBack " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCARETPERIOD(int milliseconds)
Public Property Let CaretPeriod(milliseconds As Long)
    '<EhHeader>
    On Error GoTo CaretPeriod_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCARETPERIOD, milliseconds, 0
  
    '<EhFooter>
    Exit Property

CaretPeriod_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretPeriod " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCARETPERIOD
Public Property Get CaretPeriod() As Long
Attribute CaretPeriod.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo CaretPeriod_Err
    '</EhHeader>
  
  CaretPeriod = Send_SCI_message(SCI_GETCARETPERIOD, 0, 0)
  
    '<EhFooter>
    Exit Property

CaretPeriod_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretPeriod " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCARETWIDTH(int pixels)
Public Property Let CaretWidth(pixels As Long)
    '<EhHeader>
    On Error GoTo CaretWidth_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCARETWIDTH, pixels, 0
  
    '<EhFooter>
    Exit Property

CaretWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCARETWIDTH
Public Property Get CaretWidth() As Long
Attribute CaretWidth.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo CaretWidth_Err
    '</EhHeader>
  
  CaretWidth = Send_SCI_message(SCI_GETCARETWIDTH, 0, 0)
  
    '<EhFooter>
    Exit Property

CaretWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CaretWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETHOTSPOTACTIVEFORE
Public Property Let HotSpotActiveFore(val As Long)
    '<EhHeader>
    On Error GoTo HotSpotActiveFore_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETHOTSPOTACTIVEFORE, val, 0
  
    '<EhFooter>
    Exit Property

HotSpotActiveFore_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HotSpotActiveFore " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETHOTSPOTACTIVEBACK
Public Property Let HOTSPOTACTIVEBACK(val As Long)
    '<EhHeader>
    On Error GoTo HOTSPOTACTIVEBACK_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETHOTSPOTACTIVEBACK, val, 0
  
    '<EhFooter>
    Exit Property

HOTSPOTACTIVEBACK_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HOTSPOTACTIVEBACK " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETHOTSPOTACTIVEUNDERLINE
Public Property Let HOTSPOTACTIVEUNDERLINE(val As Long)
    '<EhHeader>
    On Error GoTo HOTSPOTACTIVEUNDERLINE_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETHOTSPOTACTIVEUNDERLINE, val, 0
  
    '<EhFooter>
    Exit Property

HOTSPOTACTIVEUNDERLINE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HOTSPOTACTIVEUNDERLINE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETHOTSPOTSINGLELINE
Public Property Let HOTSPOTSINGLELINE(val As Long)
    '<EhHeader>
    On Error GoTo HOTSPOTSINGLELINE_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETHOTSPOTSINGLELINE, val, 0
  
    '<EhFooter>
    Exit Property

HOTSPOTSINGLELINE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HOTSPOTSINGLELINE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCONTROLCHARSYMBOL(int symbol)
Public Property Let ControlCharSymbol(symbol As Long)
    '<EhHeader>
    On Error GoTo ControlCharSymbol_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCONTROLCHARSYMBOL, symbol, 0
  
    '<EhFooter>
    Exit Property

ControlCharSymbol_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ControlCharSymbol " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCONTROLCHARSYMBOL
Public Property Get ControlCharSymbol() As Long
Attribute ControlCharSymbol.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo ControlCharSymbol_Err
    '</EhHeader>
  
  ControlCharSymbol = Send_SCI_message(SCI_GETCONTROLCHARSYMBOL, 0, 0)
  
    '<EhFooter>
    Exit Property

ControlCharSymbol_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ControlCharSymbol " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property



'SCI_SETMARGINTYPEN(int margin, type)
Public Property Let MarginType(margin As Long, mtype As Long)
    '<EhHeader>
    On Error GoTo MarginType_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMARGINTYPEN, margin, mtype
  
    '<EhFooter>
    Exit Property

MarginType_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginType " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMARGINTYPEN(int margin)
Public Property Get MarginType(margin As Long) As Long
Attribute MarginType.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MarginType_Err
    '</EhHeader>
  
  MarginType = Send_SCI_message(SCI_GETMARGINTYPEN, margin, 0)
  
    '<EhFooter>
    Exit Property

MarginType_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginType " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


'SCI_SETMARGINWIDTHN(int margin, pixelWidth)
Public Property Let MarginWidth(margin As Long, pixelWidth As Long)
    '<EhHeader>
    On Error GoTo MarginWidth_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMARGINWIDTHN, margin, pixelWidth
  
    '<EhFooter>
    Exit Property

MarginWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMARGINWIDTHN(int margin)
Public Property Get MarginWidth(margin As Long) As Long
Attribute MarginWidth.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MarginWidth_Err
    '</EhHeader>
  
  MarginWidth = Send_SCI_message(SCI_GETMARGINWIDTHN, margin, 0)
  
    '<EhFooter>
    Exit Property

MarginWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETMARGINMASKN(int margin, mask)
Public Property Let MarginMask(margin As Long, mask As Long)
    '<EhHeader>
    On Error GoTo MarginMask_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMARGINMASKN, margin, mask
  
    '<EhFooter>
    Exit Property

MarginMask_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginMask " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMARGINMASKN(int margin)
Public Property Get MarginMask(margin As Long) As Long
Attribute MarginMask.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MarginMask_Err
    '</EhHeader>
  
  MarginMask = Send_SCI_message(SCI_GETMARGINMASKN, margin, 0)
  
    '<EhFooter>
    Exit Property

MarginMask_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginMask " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETMARGINSENSITIVEN(int margin, bool sensitive)
Public Property Let MarginSensitive(margin As Long, sensitive As Boolean)
    '<EhHeader>
    On Error GoTo MarginSensitive_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMARGINSENSITIVEN, margin, 1 And sensitive
  
    '<EhFooter>
    Exit Property

MarginSensitive_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginSensitive " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMARGINSENSITIVEN(int margin)
Public Property Get MarginSensitive(margin As Long) As Boolean
Attribute MarginSensitive.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MarginSensitive_Err
    '</EhHeader>
  
  MarginSensitive = Send_SCI_message(SCI_GETMARGINSENSITIVEN, margin, 0)
  
    '<EhFooter>
    Exit Property

MarginSensitive_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginSensitive " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETMARGINLEFT(<unused>, pixels)
Public Property Let MarginLeft(pixels As Long)
    '<EhHeader>
    On Error GoTo MarginLeft_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMARGINLEFT, 0, pixels
  
    '<EhFooter>
    Exit Property

MarginLeft_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginLeft " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMARGINLEFT
Public Property Get MarginLeft() As Long
Attribute MarginLeft.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MarginLeft_Err
    '</EhHeader>
  
  MarginLeft = Send_SCI_message(SCI_GETMARGINLEFT, 0, 0)
  
    '<EhFooter>
    Exit Property

MarginLeft_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginLeft " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETMARGINRIGHT(<unused>, pixels)
Public Property Let MarginRight(pixels As Long)
    '<EhHeader>
    On Error GoTo MarginRight_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMARGINRIGHT, 0, pixels
  
    '<EhFooter>
    Exit Property

MarginRight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginRight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMARGINRIGHT
Public Property Get MarginRight() As Long
Attribute MarginRight.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MarginRight_Err
    '</EhHeader>
  
  MarginRight = Send_SCI_message(SCI_GETMARGINRIGHT, 0, 0)
  
    '<EhFooter>
    Exit Property

MarginRight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginRight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETFOLDMARGINCOLOUR(bool useSetting, colour)
Public Property Let FoldMarginColour(useSetting As Boolean, colour As Long)
    '<EhHeader>
    On Error GoTo FoldMarginColour_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETFOLDMARGINCOLOUR, 1 And useSetting, colour
  
    '<EhFooter>
    Exit Property

FoldMarginColour_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FoldMarginColour " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETFOLDMARGINHICOLOUR(bool useSetting, colour)
Public Property Let FoldMarginHiColour(useSetting As Boolean, colour As Long)
    '<EhHeader>
    On Error GoTo FoldMarginHiColour_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETFOLDMARGINHICOLOUR, 1 And useSetting, colour
  
    '<EhFooter>
    Exit Property

FoldMarginHiColour_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FoldMarginHiColour " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETMARGINTYPEN(int margin, iType)
Public Property Let MarginTypem(margin As Long, iType As Long)
    '<EhHeader>
    On Error GoTo MarginTypem_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMARGINTYPEN, margin, iType
  
    '<EhFooter>
    Exit Property

MarginTypem_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MarginTypem " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETUSEPALETTE(bool allowPaletteUse)
Public Property Let UsePalette(allowPaletteUse As Boolean)
    '<EhHeader>
    On Error GoTo UsePalette_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETUSEPALETTE, 1 And allowPaletteUse, 0
  
    '<EhFooter>
    Exit Property

UsePalette_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UsePalette " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETUSEPALETTE
Public Property Get UsePalette() As Boolean
Attribute UsePalette.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo UsePalette_Err
    '</EhHeader>
  
  UsePalette = Send_SCI_message(SCI_GETUSEPALETTE, 0, 0)
  
    '<EhFooter>
    Exit Property

UsePalette_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.UsePalette " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETBUFFEREDDRAW(bool isBuffered)
Public Property Let BufferEdDraw(isBuffered As Boolean)
    '<EhHeader>
    On Error GoTo BufferEdDraw_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETBUFFEREDDRAW, 1 And isBuffered, 0
  
    '<EhFooter>
    Exit Property

BufferEdDraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BufferEdDraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETBUFFEREDDRAW
Public Property Get BufferEdDraw() As Boolean
Attribute BufferEdDraw.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo BufferEdDraw_Err
    '</EhHeader>
  
  BufferEdDraw = Send_SCI_message(SCI_GETBUFFEREDDRAW, 0, 0)
  
    '<EhFooter>
    Exit Property

BufferEdDraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BufferEdDraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETTWOPHASEDRAW(bool twoPhase)
Public Property Let TwoPhaseDraw(twoPhase As Boolean)
    '<EhHeader>
    On Error GoTo TwoPhaseDraw_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETTWOPHASEDRAW, 1 And twoPhase, 0

    '<EhFooter>
    Exit Property

TwoPhaseDraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TwoPhaseDraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETTWOPHASEDRAW
Public Property Get TwoPhaseDraw() As Boolean
Attribute TwoPhaseDraw.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo TwoPhaseDraw_Err
    '</EhHeader>
  
  TwoPhaseDraw = Send_SCI_message(SCI_GETTWOPHASEDRAW, 0, 0)
  
    '<EhFooter>
    Exit Property

TwoPhaseDraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TwoPhaseDraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCODEPAGE(int codePage)
Public Property Let codepage(codepage As Long)
    '<EhHeader>
    On Error GoTo codepage_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCODEPAGE, codepage, 0
  
    '<EhFooter>
    Exit Property

codepage_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.codepage " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETCODEPAGE
Public Property Get codepage() As Long
Attribute codepage.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo codepage_Err
    '</EhHeader>
  
  codepage = Send_SCI_message(SCI_GETCODEPAGE, 0, 0)
  
    '<EhFooter>
    Exit Property

codepage_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.codepage " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETWORDCHARS(<unused>, const char *chars)
Public Property Let WORDCHARS(wChars As String)
    '<EhHeader>
    On Error GoTo WORDCHARS_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_SETWORDCHARS, 0, wChars
  
    '<EhFooter>
    Exit Property

WORDCHARS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WORDCHARS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETWHITESPACECHARS(<unused>, const char *chars)
Public Property Let WhiteSpaceChars(wsChars As String)
    '<EhHeader>
    On Error GoTo WhiteSpaceChars_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_SETWHITESPACECHARS, 0, wsChars
  
    '<EhFooter>
    Exit Property

WhiteSpaceChars_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WhiteSpaceChars " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETCHARSDEFAULT
Public Sub CharsDefault()
    '<EhHeader>
    On Error GoTo CharsDefault_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETCHARSDEFAULT, 0, 0
  
    '<EhFooter>
    Exit Sub

CharsDefault_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CharsDefault " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_GRABFOCUS
Public Sub GrabFocus()
    '<EhHeader>
    On Error GoTo GrabFocus_Err
    '</EhHeader>
  
    Send_SCI_message SCI_GRABFOCUS, 0, 0
  
    '<EhFooter>
    Exit Sub

GrabFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.GrabFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'SCI_SETFOCUS(bool focus)
Public Property Let Focus(Focus As Boolean)
    '<EhHeader>
    On Error GoTo Focus_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETFOCUS, 1 And Focus, 0
  
    '<EhFooter>
    Exit Property

Focus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Focus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETFOCUS
Public Property Get Focus() As Boolean
Attribute Focus.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo Focus_Err
    '</EhHeader>
  
  Focus = Send_SCI_message(SCI_GETFOCUS, 0, 0)
  
    '<EhFooter>
    Exit Property

Focus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Focus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


'SCI_BRACEHIGHLIGHT(int pos1, pos2)
Public Function BRACEHIGHLIGHT(pos1 As Long, pos2 As Long) As Long
    '<EhHeader>
    On Error GoTo BRACEHIGHLIGHT_Err
    '</EhHeader>
  
  BRACEHIGHLIGHT = Send_SCI_message(SCI_BRACEHIGHLIGHT, pos1, pos2)
  
    '<EhFooter>
    Exit Function

BRACEHIGHLIGHT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BRACEHIGHLIGHT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_BRACEBADLIGHT(int pos1)
Public Function BRACEBADLIGHT(pos1 As Long) As Long
    '<EhHeader>
    On Error GoTo BRACEBADLIGHT_Err
    '</EhHeader>
  
  BRACEBADLIGHT = Send_SCI_message(SCI_BRACEBADLIGHT, pos1, 0)
  
    '<EhFooter>
    Exit Function

BRACEBADLIGHT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BRACEBADLIGHT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_BRACEMATCH(int position, maxReStyle)
Public Function BRACEMATCH(position As Long, maxReStyle As Long) As Long
    '<EhHeader>
    On Error GoTo BRACEMATCH_Err
    '</EhHeader>
  
  BRACEMATCH = Send_SCI_message(SCI_BRACEMATCH, position, maxReStyle)
  
    '<EhFooter>
    Exit Function

BRACEMATCH_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BRACEMATCH " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function



'SCI_SETTABWIDTH(int widthInChars)
Public Property Let TabWidth(widthInChars As Long)
    '<EhHeader>
    On Error GoTo TabWidth_Err
    '</EhHeader>

  Send_SCI_message SCI_SETTABWIDTH, widthInChars, 0
  
    '<EhFooter>
    Exit Property

TabWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TabWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETTABWIDTH
Public Property Get TabWidth() As Long
Attribute TabWidth.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo TabWidth_Err
    '</EhHeader>
  
  TabWidth = Send_SCI_message(SCI_GETTABWIDTH, 0, 0)
  
    '<EhFooter>
    Exit Property

TabWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TabWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETUSETABS(bool useTabs)
Public Property Let useTabs(useTabs As Boolean)
    '<EhHeader>
    On Error GoTo useTabs_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETUSETABS, 1 And useTabs, 0
  
    '<EhFooter>
    Exit Property

useTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.useTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETUSETABS
Public Property Get useTabs() As Boolean
Attribute useTabs.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo useTabs_Err
    '</EhHeader>
  
  useTabs = Send_SCI_message(SCI_GETUSETABS, 0, 0)
  
    '<EhFooter>
    Exit Property

useTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.useTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETINDENT(int widthInChars)
Public Property Let Indent(widthInChars As Long)
    '<EhHeader>
    On Error GoTo Indent_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETINDENT, widthInChars, 0
  
    '<EhFooter>
    Exit Property

Indent_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Indent " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETINDENT
Public Property Get Indent() As Long
Attribute Indent.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo Indent_Err
    '</EhHeader>
  
  Indent = Send_SCI_message(SCI_GETINDENT, 0, 0)
  
    '<EhFooter>
    Exit Property

Indent_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Indent " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETTABINDENTS(bool tabIndents)
Public Property Let tabIndents(tabIndents As Boolean)
    '<EhHeader>
    On Error GoTo tabIndents_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETTABINDENTS, 1 And tabIndents, 0
  
    '<EhFooter>
    Exit Property

tabIndents_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.tabIndents " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETTABINDENTS
Public Property Get tabIndents() As Boolean
Attribute tabIndents.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo tabIndents_Err
    '</EhHeader>
  
  tabIndents = Send_SCI_message(SCI_GETTABINDENTS, 0, 0)
  
    '<EhFooter>
    Exit Property

tabIndents_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.tabIndents " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETBACKSPACEUNINDENTS(bool bsUnIndents)
Public Property Let BACKSPACEUNINDENTS(bsUnIndents As Boolean)
    '<EhHeader>
    On Error GoTo BACKSPACEUNINDENTS_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETBACKSPACEUNINDENTS, 1 And bsUnIndents, 0
  
    '<EhFooter>
    Exit Property

BACKSPACEUNINDENTS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BACKSPACEUNINDENTS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETBACKSPACEUNINDENTS
Public Property Get BACKSPACEUNINDENTS() As Boolean
Attribute BACKSPACEUNINDENTS.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo BACKSPACEUNINDENTS_Err
    '</EhHeader>
  
  BACKSPACEUNINDENTS = Send_SCI_message(SCI_GETBACKSPACEUNINDENTS, 0, 0)
  
    '<EhFooter>
    Exit Property

BACKSPACEUNINDENTS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.BACKSPACEUNINDENTS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETLINEINDENTATION(int line, indentation)
Public Property Let LINEINDENTATION(line As Long, indentation As Long)
    '<EhHeader>
    On Error GoTo LINEINDENTATION_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETLINEINDENTATION, line, indentation
  
    '<EhFooter>
    Exit Property

LINEINDENTATION_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LINEINDENTATION " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETLINEINDENTATION(int line)
Public Property Get LINEINDENTATION(line As Long) As Long
Attribute LINEINDENTATION.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo LINEINDENTATION_Err
    '</EhHeader>
  
  LINEINDENTATION = Send_SCI_message(SCI_GETLINEINDENTATION, line, 0)
  
    '<EhFooter>
    Exit Property

LINEINDENTATION_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LINEINDENTATION " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETLINEINDENTPOSITION(int line)
Public Property Get LINEINDENTPOSITION(line As Long) As Long
    '<EhHeader>
    On Error GoTo LINEINDENTPOSITION_Err
    '</EhHeader>
  
  LINEINDENTPOSITION = Send_SCI_message(SCI_GETLINEINDENTPOSITION, line, 0)
  
    '<EhFooter>
    Exit Property

LINEINDENTPOSITION_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LINEINDENTPOSITION " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETINDENTATIONGUIDES(bool view)
Public Property Let INDENTATIONGUIDES(View As Boolean)
    '<EhHeader>
    On Error GoTo INDENTATIONGUIDES_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETINDENTATIONGUIDES, 1 And View, 0
  
    '<EhFooter>
    Exit Property

INDENTATIONGUIDES_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.INDENTATIONGUIDES " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETINDENTATIONGUIDES
Public Property Get INDENTATIONGUIDES() As Boolean
Attribute INDENTATIONGUIDES.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo INDENTATIONGUIDES_Err
    '</EhHeader>
  
  INDENTATIONGUIDES = Send_SCI_message(SCI_GETINDENTATIONGUIDES, 0, 0)
  
    '<EhFooter>
    Exit Property

INDENTATIONGUIDES_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.INDENTATIONGUIDES " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETHIGHLIGHTGUIDE(int column)
Public Property Let HIGHLIGHTGUIDE(Column As Long)
    '<EhHeader>
    On Error GoTo HIGHLIGHTGUIDE_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETHIGHLIGHTGUIDE, Column, 0
  
    '<EhFooter>
    Exit Property

HIGHLIGHTGUIDE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HIGHLIGHTGUIDE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETHIGHLIGHTGUIDE
Public Property Get HIGHLIGHTGUIDE() As Long
Attribute HIGHLIGHTGUIDE.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo HIGHLIGHTGUIDE_Err
    '</EhHeader>
  
  HIGHLIGHTGUIDE = Send_SCI_message(SCI_GETHIGHLIGHTGUIDE, 0, 0)
  
    '<EhFooter>
    Exit Property

HIGHLIGHTGUIDE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HIGHLIGHTGUIDE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_MARKERDEFINE(int markerNumber, markerSymbols)
Public Function MARKERDEFINE(markerNumber As Long, markerSymbols As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERDEFINE_Err
    '</EhHeader>
  
  MARKERDEFINE = Send_SCI_message(SCI_MARKERDEFINE, markerNumber, markerSymbols)
  
    '<EhFooter>
    Exit Function

MARKERDEFINE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERDEFINE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERDEFINEPIXMAP(int markerNumber, const char *xpm)
Public Function MARKERDEFINEPIXMAP(markerNumber As Long, xpm As String) As Long
    '<EhHeader>
    On Error GoTo MARKERDEFINEPIXMAP_Err
    '</EhHeader>
  
  MARKERDEFINEPIXMAP = Send_SCI_messageStr(SCI_MARKERDEFINEPIXMAP, markerNumber, xpm)
  
    '<EhFooter>
    Exit Function

MARKERDEFINEPIXMAP_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERDEFINEPIXMAP " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERSETFORE(int markerNumber, colour)
Public Function MARKERSETFORE(markerNumber As Long, colour As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERSETFORE_Err
    '</EhHeader>
  
  MARKERSETFORE = Send_SCI_message(SCI_MARKERSETFORE, markerNumber, colour)
  
    '<EhFooter>
    Exit Function

MARKERSETFORE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERSETFORE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERSETBACK(int markerNumber, colour)
Public Function MARKERSETBACK(markerNumber As Long, colour As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERSETBACK_Err
    '</EhHeader>
  
  MARKERSETBACK = Send_SCI_message(SCI_MARKERSETBACK, markerNumber, colour)
  
    '<EhFooter>
    Exit Function

MARKERSETBACK_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERSETBACK " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERADD(int line, markerNumber)
Public Function MARKERADD(line As Long, markerNumber As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERADD_Err
    '</EhHeader>
  
  MARKERADD = Send_SCI_message(SCI_MARKERADD, line, markerNumber)
  
    '<EhFooter>
    Exit Function

MARKERADD_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERADD " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERDELETE(int line, markerNumber)
Public Function MARKERDELETE(line As Long, markerNumber As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERDELETE_Err
    '</EhHeader>
  
  MARKERDELETE = Send_SCI_message(SCI_MARKERDELETE, line, markerNumber)
  
    '<EhFooter>
    Exit Function

MARKERDELETE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERDELETE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERDELETEALL(int markerNumber)
Public Function MARKERDELETEALL(markerNumber As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERDELETEALL_Err
    '</EhHeader>
  
  MARKERDELETEALL = Send_SCI_message(SCI_MARKERDELETEALL, markerNumber, 0)
  
    '<EhFooter>
    Exit Function

MARKERDELETEALL_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERDELETEALL " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERGET(int line)
Public Function MARKERGET(line As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERGET_Err
    '</EhHeader>
  
  MARKERGET = Send_SCI_message(SCI_MARKERGET, line, 0)
  
    '<EhFooter>
    Exit Function

MARKERGET_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERGET " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERNEXT(int lineStart, markerMask)
Public Function MARKERNEXT(lineStart As Long, markerMask As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERNEXT_Err
    '</EhHeader>
  
  MARKERNEXT = Send_SCI_message(SCI_MARKERNEXT, lineStart, markerMask)
  
    '<EhFooter>
    Exit Function

MARKERNEXT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERNEXT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERPREVIOUS(int lineStart, markerMask)
Public Function MARKERPREVIOUS(lineStart As Long, markerMask As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERPREVIOUS_Err
    '</EhHeader>
  
  MARKERPREVIOUS = Send_SCI_message(SCI_MARKERPREVIOUS, lineStart, markerMask)
  
    '<EhFooter>
    Exit Function

MARKERPREVIOUS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERPREVIOUS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERLINEFROMHANDLE(int handle)
Public Function MARKERLINEFROMHANDLE(handle As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERLINEFROMHANDLE_Err
    '</EhHeader>
  
  MARKERLINEFROMHANDLE = Send_SCI_message(SCI_MARKERLINEFROMHANDLE, handle, 0)
  
    '<EhFooter>
    Exit Function

MARKERLINEFROMHANDLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERLINEFROMHANDLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_MARKERDELETEHANDLE(int handle)
Public Function MARKERDELETEHANDLE(handle As Long) As Long
    '<EhHeader>
    On Error GoTo MARKERDELETEHANDLE_Err
    '</EhHeader>
  
  MARKERDELETEHANDLE = Send_SCI_message(SCI_MARKERDELETEHANDLE, handle, 0)
  
    '<EhFooter>
    Exit Function

MARKERDELETEHANDLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MARKERDELETEHANDLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function




'SCI_INDICSETSTYLE(int indicatorNumber, indicatorStyle)
Public Function INDICSETSTYLE(indicatorNumber As Long, indicatorStyle As Long) As Long
    '<EhHeader>
    On Error GoTo INDICSETSTYLE_Err
    '</EhHeader>
  
  INDICSETSTYLE = Send_SCI_message(SCI_INDICSETSTYLE, indicatorNumber, indicatorStyle)
  
    '<EhFooter>
    Exit Function

INDICSETSTYLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.INDICSETSTYLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_INDICGETSTYLE(int indicatorNumber)
Public Function INDICGETSTYLE(indicatorNumber As Long) As Long
    '<EhHeader>
    On Error GoTo INDICGETSTYLE_Err
    '</EhHeader>
  
  INDICGETSTYLE = Send_SCI_message(SCI_INDICGETSTYLE, indicatorNumber, 0)
  
    '<EhFooter>
    Exit Function

INDICGETSTYLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.INDICGETSTYLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_INDICSETFORE(int indicatorNumber, colour)
Public Function INDICSETFORE(indicatorNumber As Long, colour) As Long
    '<EhHeader>
    On Error GoTo INDICSETFORE_Err
    '</EhHeader>
  
  INDICSETFORE = Send_SCI_message(SCI_INDICSETFORE, indicatorNumber, colour)
  
    '<EhFooter>
    Exit Function

INDICSETFORE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.INDICSETFORE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_INDICGETFORE(int indicatorNumber)
Public Function INDICGETFORE(indicatorNumber As Long) As Long
    '<EhHeader>
    On Error GoTo INDICGETFORE_Err
    '</EhHeader>
  
  INDICGETFORE = Send_SCI_message(SCI_INDICGETFORE, indicatorNumber, 0)
  
    '<EhFooter>
    Exit Function

INDICGETFORE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.INDICGETFORE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


'SCI_AUTOCSHOW(int lenEntered, const char *list)
Public Function AUTOCSHOW(lenEntered As Long, list As String) As Long
    '<EhHeader>
    On Error GoTo AUTOCSHOW_Err
    '</EhHeader>
  
  AUTOCSHOW = Send_SCI_messageStr(SCI_AUTOCSHOW, lenEntered, list)
  
    '<EhFooter>
    Exit Function

AUTOCSHOW_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSHOW " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCCANCEL
Public Function AUTOCCANCEL() As Long
    '<EhHeader>
    On Error GoTo AUTOCCANCEL_Err
    '</EhHeader>
  
  AUTOCCANCEL = Send_SCI_message(SCI_AUTOCCANCEL, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCCANCEL_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCCANCEL " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCACTIVE
Public Function AUTOCACTIVE() As Long
    '<EhHeader>
    On Error GoTo AUTOCACTIVE_Err
    '</EhHeader>
  
  AUTOCACTIVE = Send_SCI_message(SCI_AUTOCACTIVE, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCACTIVE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCACTIVE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCPOSSTART
Public Function AUTOCPOSSTART() As Long
    '<EhHeader>
    On Error GoTo AUTOCPOSSTART_Err
    '</EhHeader>
  
  AUTOCPOSSTART = Send_SCI_message(SCI_AUTOCPOSSTART, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCPOSSTART_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCPOSSTART " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCCOMPLETE
Public Function AUTOCCOMPLETE() As Long
    '<EhHeader>
    On Error GoTo AUTOCCOMPLETE_Err
    '</EhHeader>
  
  AUTOCCOMPLETE = Send_SCI_message(SCI_AUTOCCOMPLETE, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCCOMPLETE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCCOMPLETE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSTOPS(<unused>, const char *chars)
Public Function AUTOCSTOPS(chars As String) As Long
    '<EhHeader>
    On Error GoTo AUTOCSTOPS_Err
    '</EhHeader>
  
  AUTOCSTOPS = Send_SCI_messageStr(SCI_AUTOCSTOPS, 0, chars)
  
    '<EhFooter>
    Exit Function

AUTOCSTOPS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSTOPS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETSEPARATOR(char separator)
Public Function AUTOCSETSEPARATOR(SEPARATOR As Byte) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETSEPARATOR_Err
    '</EhHeader>
  
  AUTOCSETSEPARATOR = Send_SCI_message(SCI_AUTOCSETSEPARATOR, SEPARATOR, 0)
  
    '<EhFooter>
    Exit Function

AUTOCSETSEPARATOR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETSEPARATOR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETSEPARATOR
Public Function AUTOCGETSEPARATOR() As Long
    '<EhHeader>
    On Error GoTo AUTOCGETSEPARATOR_Err
    '</EhHeader>
  
  AUTOCGETSEPARATOR = Send_SCI_message(SCI_AUTOCGETSEPARATOR, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETSEPARATOR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETSEPARATOR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSELECT(<unused>, const char *select)
Public Function AUTOCSELECT(sselect As String) As Long
    '<EhHeader>
    On Error GoTo AUTOCSELECT_Err
    '</EhHeader>
  
  AUTOCSELECT = Send_SCI_messageStr(SCI_AUTOCSELECT, 0, sselect)
  
    '<EhFooter>
    Exit Function

AUTOCSELECT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSELECT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETCURRENT
Public Function AUTOCGETCURRENT() As Long
    '<EhHeader>
    On Error GoTo AUTOCGETCURRENT_Err
    '</EhHeader>
  
  AUTOCGETCURRENT = Send_SCI_message(SCI_AUTOCGETCURRENT, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETCURRENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETCURRENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETCANCELATSTART(bool cancel)
Public Function AUTOCSETCANCELATSTART(Cancel As Boolean) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETCANCELATSTART_Err
    '</EhHeader>
  
  AUTOCSETCANCELATSTART = Send_SCI_message(SCI_AUTOCSETCANCELATSTART, 1 And Cancel, 0)
  
    '<EhFooter>
    Exit Function

AUTOCSETCANCELATSTART_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETCANCELATSTART " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETCANCELATSTART
Public Function AUTOCGETCANCELATSTART() As Boolean
    '<EhHeader>
    On Error GoTo AUTOCGETCANCELATSTART_Err
    '</EhHeader>
  
  AUTOCGETCANCELATSTART = Send_SCI_message(SCI_AUTOCGETCANCELATSTART, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETCANCELATSTART_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETCANCELATSTART " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETFILLUPS(<unused>, const char *chars)
Public Function AUTOCSETFILLUPS(chars As String) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETFILLUPS_Err
    '</EhHeader>
  
  AUTOCSETFILLUPS = Send_SCI_messageStr(SCI_AUTOCSETFILLUPS, 0, chars)
  
    '<EhFooter>
    Exit Function

AUTOCSETFILLUPS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETFILLUPS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETCHOOSESINGLE(bool chooseSingle)
Public Function AUTOCSETCHOOSESINGLE(chooseSingle As Boolean) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETCHOOSESINGLE_Err
    '</EhHeader>
  
  AUTOCSETCHOOSESINGLE = Send_SCI_message(SCI_AUTOCSETCHOOSESINGLE, 1 And chooseSingle, 0)
  
    '<EhFooter>
    Exit Function

AUTOCSETCHOOSESINGLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETCHOOSESINGLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETCHOOSESINGLE
Public Function AUTOCGETCHOOSESINGLE() As Boolean
    '<EhHeader>
    On Error GoTo AUTOCGETCHOOSESINGLE_Err
    '</EhHeader>
  
  AUTOCGETCHOOSESINGLE = Send_SCI_message(SCI_AUTOCGETCHOOSESINGLE, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETCHOOSESINGLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETCHOOSESINGLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETIGNORECASE(bool ignoreCase)
Public Function AUTOCSETIGNORECASE(ignoreCase As Boolean) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETIGNORECASE_Err
    '</EhHeader>
  
  AUTOCSETIGNORECASE = Send_SCI_message(SCI_AUTOCSETIGNORECASE, 1 And ignoreCase, 0)
  
    '<EhFooter>
    Exit Function

AUTOCSETIGNORECASE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETIGNORECASE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETIGNORECASE
Public Function AUTOCGETIGNORECASE() As Long
    '<EhHeader>
    On Error GoTo AUTOCGETIGNORECASE_Err
    '</EhHeader>
  
  AUTOCGETIGNORECASE = Send_SCI_message(SCI_AUTOCGETIGNORECASE, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETIGNORECASE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETIGNORECASE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETAUTOHIDE(bool autoHide)
Public Function AUTOCSETAUTOHIDE(autoHide As Boolean) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETAUTOHIDE_Err
    '</EhHeader>
  
  AUTOCSETAUTOHIDE = Send_SCI_message(SCI_AUTOCSETAUTOHIDE, 1 And autoHide, 0)
  
    '<EhFooter>
    Exit Function

AUTOCSETAUTOHIDE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETAUTOHIDE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETAUTOHIDE
Public Function AUTOCGETAUTOHIDE() As Boolean
    '<EhHeader>
    On Error GoTo AUTOCGETAUTOHIDE_Err
    '</EhHeader>
  
  AUTOCGETAUTOHIDE = Send_SCI_message(SCI_AUTOCGETAUTOHIDE, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETAUTOHIDE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETAUTOHIDE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETDROPRESTOFWORD(bool dropRestOfWord)
Public Function AUTOCSETDROPRESTOFWORD(dropRestOfWord As Boolean) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETDROPRESTOFWORD_Err
    '</EhHeader>
  
  AUTOCSETDROPRESTOFWORD = Send_SCI_message(SCI_AUTOCSETDROPRESTOFWORD, 1 And dropRestOfWord, 0)
  
    '<EhFooter>
    Exit Function

AUTOCSETDROPRESTOFWORD_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETDROPRESTOFWORD " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETDROPRESTOFWORD
Public Function AUTOCGETDROPRESTOFWORD() As Boolean
    '<EhHeader>
    On Error GoTo AUTOCGETDROPRESTOFWORD_Err
    '</EhHeader>
  
  AUTOCGETDROPRESTOFWORD = Send_SCI_message(SCI_AUTOCGETDROPRESTOFWORD, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETDROPRESTOFWORD_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETDROPRESTOFWORD " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_REGISTERIMAGE
Public Function REGISTERIMAGE() As Long
    '<EhHeader>
    On Error GoTo REGISTERIMAGE_Err
    '</EhHeader>
  
  REGISTERIMAGE = Send_SCI_message(SCI_REGISTERIMAGE, 0, 0)
  
    '<EhFooter>
    Exit Function

REGISTERIMAGE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.REGISTERIMAGE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CLEARREGISTEREDIMAGES
Public Function CLEARREGISTEREDIMAGES() As Long
    '<EhHeader>
    On Error GoTo CLEARREGISTEREDIMAGES_Err
    '</EhHeader>
  
  CLEARREGISTEREDIMAGES = Send_SCI_message(SCI_CLEARREGISTEREDIMAGES, 0, 0)
  
    '<EhFooter>
    Exit Function

CLEARREGISTEREDIMAGES_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CLEARREGISTEREDIMAGES " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCSETTYPESEPARATOR(char separatorCharacter)
Public Function AUTOCSETTYPESEPARATOR(separatorCharacter As Byte) As Long
    '<EhHeader>
    On Error GoTo AUTOCSETTYPESEPARATOR_Err
    '</EhHeader>
  
  AUTOCSETTYPESEPARATOR = Send_SCI_message(SCI_AUTOCSETTYPESEPARATOR, separatorCharacter, 0)
  
    '<EhFooter>
    Exit Function

AUTOCSETTYPESEPARATOR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCSETTYPESEPARATOR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_AUTOCGETTYPESEPARATOR
Public Function AUTOCGETTYPESEPARATOR() As Long
    '<EhHeader>
    On Error GoTo AUTOCGETTYPESEPARATOR_Err
    '</EhHeader>
  
  AUTOCGETTYPESEPARATOR = Send_SCI_message(SCI_AUTOCGETTYPESEPARATOR, 0, 0)
  
    '<EhFooter>
    Exit Function

AUTOCGETTYPESEPARATOR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.AUTOCGETTYPESEPARATOR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_USERLISTSHOW(int listType, const char *list)
Public Function USERLISTSHOW(listType As Long, list As String) As Long
    '<EhHeader>
    On Error GoTo USERLISTSHOW_Err
    '</EhHeader>
  
  USERLISTSHOW = Send_SCI_messageStr(SCI_USERLISTSHOW, listType, list)
  
    '<EhFooter>
    Exit Function

USERLISTSHOW_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.USERLISTSHOW " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPSHOW(int posStart, const char *definition)
Public Function CALLTIPSHOW(posStart As Long, definition As String) As Long
    '<EhHeader>
    On Error GoTo CALLTIPSHOW_Err
    '</EhHeader>
  
  CALLTIPSHOW = Send_SCI_messageStr(SCI_CALLTIPSHOW, posStart, definition)
  
    '<EhFooter>
    Exit Function

CALLTIPSHOW_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPSHOW " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPCANCEL
Public Function CALLTIPCANCEL() As Long
    '<EhHeader>
    On Error GoTo CALLTIPCANCEL_Err
    '</EhHeader>
  
  CALLTIPCANCEL = Send_SCI_message(SCI_CALLTIPCANCEL, 0, 0)
  
    '<EhFooter>
    Exit Function

CALLTIPCANCEL_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPCANCEL " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPACTIVE
Public Function CALLTIPACTIVE() As Long
    '<EhHeader>
    On Error GoTo CALLTIPACTIVE_Err
    '</EhHeader>
  
  CALLTIPACTIVE = Send_SCI_message(SCI_CALLTIPACTIVE, 0, 0)
  
    '<EhFooter>
    Exit Function

CALLTIPACTIVE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPACTIVE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPPOSSTART
Public Function CALLTIPPOSSTART() As Long
    '<EhHeader>
    On Error GoTo CALLTIPPOSSTART_Err
    '</EhHeader>
  
  CALLTIPPOSSTART = Send_SCI_message(SCI_CALLTIPPOSSTART, 0, 0)
  
    '<EhFooter>
    Exit Function

CALLTIPPOSSTART_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPPOSSTART " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPSETHLT(int highlightStart, highlightEnd)
Public Function CALLTIPSETHLT(highlightStart As Long, highlightEnd As Long) As Long
    '<EhHeader>
    On Error GoTo CALLTIPSETHLT_Err
    '</EhHeader>
  
  CALLTIPSETHLT = Send_SCI_message(SCI_CALLTIPSETHLT, highlightStart, highlightEnd)
  
    '<EhFooter>
    Exit Function

CALLTIPSETHLT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPSETHLT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPSETBACK(int colour)
Public Function CALLTIPSETBACK(colour As Long) As Long
    '<EhHeader>
    On Error GoTo CALLTIPSETBACK_Err
    '</EhHeader>
  
  CALLTIPSETBACK = Send_SCI_message(SCI_CALLTIPSETBACK, colour, 0)
  
    '<EhFooter>
    Exit Function

CALLTIPSETBACK_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPSETBACK " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPSETFORE(int colour)
Public Function CALLTIPSETFORE(colour As Long) As Long
    '<EhHeader>
    On Error GoTo CALLTIPSETFORE_Err
    '</EhHeader>
  
  CALLTIPSETFORE = Send_SCI_message(SCI_CALLTIPSETFORE, colour, 0)
  
    '<EhFooter>
    Exit Function

CALLTIPSETFORE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPSETFORE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CALLTIPSETFOREHLT(int colour)
Public Function CALLTIPSETFOREHLT(colour As Long) As Long
    '<EhHeader>
    On Error GoTo CALLTIPSETFOREHLT_Err
    '</EhHeader>
  
  CALLTIPSETFOREHLT = Send_SCI_message(SCI_CALLTIPSETFOREHLT, colour, 0)
  
    '<EhFooter>
    Exit Function

CALLTIPSETFOREHLT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CALLTIPSETFOREHLT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


'SCI_ASSIGNCMDKEY(int keyDefinition, sciCommand)
Public Function ASSIGNCMDKEY(keyDefinition As Long, sciCommand As Long) As Long
    '<EhHeader>
    On Error GoTo ASSIGNCMDKEY_Err
    '</EhHeader>
  
  ASSIGNCMDKEY = Send_SCI_message(SCI_ASSIGNCMDKEY, keyDefinition, sciCommand)
  
    '<EhFooter>
    Exit Function

ASSIGNCMDKEY_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ASSIGNCMDKEY " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CLEARCMDKEY(int keyDefinition)
Public Function CLEARCMDKEY(keyDefinition As Long) As Long
    '<EhHeader>
    On Error GoTo CLEARCMDKEY_Err
    '</EhHeader>
  
  CLEARCMDKEY = Send_SCI_message(SCI_CLEARCMDKEY, keyDefinition, 0)
  
    '<EhFooter>
    Exit Function

CLEARCMDKEY_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CLEARCMDKEY " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_CLEARALLCMDKEYS
Public Function CLEARALLCMDKEYS() As Long
    '<EhHeader>
    On Error GoTo CLEARALLCMDKEYS_Err
    '</EhHeader>
  
  CLEARALLCMDKEYS = Send_SCI_message(SCI_CLEARALLCMDKEYS, 0, 0)
  
    '<EhFooter>
    Exit Function

CLEARALLCMDKEYS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CLEARALLCMDKEYS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_NULL
Public Function DoNull() As Long
    '<EhHeader>
    On Error GoTo DoNull_Err
    '</EhHeader>
  
  DoNull = Send_SCI_message(SCI_NULL, 0, 0)
  
    '<EhFooter>
    Exit Function

DoNull_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.DoNull " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_USEPOPUP(bool bEnablePopup)
Public Function USEPOPUP(bEnablePopup As Boolean) As Long
    '<EhHeader>
    On Error GoTo USEPOPUP_Err
    '</EhHeader>
  
  USEPOPUP = Send_SCI_message(SCI_USEPOPUP, 1 And bEnablePopup, 0)
  
    '<EhFooter>
    Exit Function

USEPOPUP_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.USEPOPUP " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_STARTRECORD
Public Function STARTRECORD() As Long
    '<EhHeader>
    On Error GoTo STARTRECORD_Err
    '</EhHeader>
  
  STARTRECORD = Send_SCI_message(SCI_STARTRECORD, 0, 0)
  
    '<EhFooter>
    Exit Function

STARTRECORD_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.STARTRECORD " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_STOPRECORD
Public Function STOPRECORD() As Long
    '<EhHeader>
    On Error GoTo STOPRECORD_Err
    '</EhHeader>
  
  STOPRECORD = Send_SCI_message(SCI_STOPRECORD, 0, 0)
  
    '<EhFooter>
    Exit Function

STOPRECORD_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.STOPRECORD " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_FORMATRANGE(bool bDraw, RangeToFormat *pfr)
Public Function FORMATRANGE(bDraw As Boolean, RangeToF As Long) As Long
    '<EhHeader>
    On Error GoTo FORMATRANGE_Err
    '</EhHeader>
  
  FORMATRANGE = Send_SCI_message(SCI_FORMATRANGE, 1 And bDraw, RangeToF)
  
    '<EhFooter>
    Exit Function

FORMATRANGE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FORMATRANGE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETPRINTMAGNIFICATION(int magnification)
Public Property Let PRINTMAGNIFICATION(magnification As Long)
    '<EhHeader>
    On Error GoTo PRINTMAGNIFICATION_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETPRINTMAGNIFICATION, magnification, 0
  
    '<EhFooter>
    Exit Property

PRINTMAGNIFICATION_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PRINTMAGNIFICATION " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETPRINTMAGNIFICATION
Public Property Get PRINTMAGNIFICATION() As Long
Attribute PRINTMAGNIFICATION.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo PRINTMAGNIFICATION_Err
    '</EhHeader>
  
  PRINTMAGNIFICATION = Send_SCI_message(SCI_GETPRINTMAGNIFICATION, 0, 0)
  
    '<EhFooter>
    Exit Property

PRINTMAGNIFICATION_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PRINTMAGNIFICATION " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETPRINTCOLOURMODE(int mode)
Public Property Let PRINTCOLOURMODE(Mode As Long)
    '<EhHeader>
    On Error GoTo PRINTCOLOURMODE_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETPRINTCOLOURMODE, Mode, 0
  
    '<EhFooter>
    Exit Property

PRINTCOLOURMODE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PRINTCOLOURMODE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETPRINTCOLOURMODE
Public Property Get PRINTCOLOURMODE() As Long
Attribute PRINTCOLOURMODE.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo PRINTCOLOURMODE_Err
    '</EhHeader>
  
  PRINTCOLOURMODE = Send_SCI_message(SCI_GETPRINTCOLOURMODE, 0, 0)
  
    '<EhFooter>
    Exit Property

PRINTCOLOURMODE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PRINTCOLOURMODE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETPRINTWRAPMODE
Public Property Get PRINTWRAPMODE() As Long
Attribute PRINTWRAPMODE.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo PRINTWRAPMODE_Err
    '</EhHeader>
  
  PRINTWRAPMODE = Send_SCI_message(SCI_GETPRINTWRAPMODE, 0, 0)
  
    '<EhFooter>
    Exit Property

PRINTWRAPMODE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PRINTWRAPMODE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETPRINTWRAPMODE(int wrapMode)
Public Property Let PRINTWRAPMODE(wrapMode As Long)
    '<EhHeader>
    On Error GoTo PRINTWRAPMODE_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETPRINTWRAPMODE, wrapMode, 0
  
    '<EhFooter>
    Exit Property

PRINTWRAPMODE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PRINTWRAPMODE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETDIRECTFUNCTION
Public Property Get DIRECTFUNCTION() As Long
    '<EhHeader>
    On Error GoTo DIRECTFUNCTION_Err
    '</EhHeader>
  
  DIRECTFUNCTION = Send_SCI_message(SCI_GETDIRECTFUNCTION, 0, 0)
  
    '<EhFooter>
    Exit Property

DIRECTFUNCTION_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.DIRECTFUNCTION " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETDIRECTPOINTER
Public Property Get DIRECTPOINTER() As Long
    '<EhHeader>
    On Error GoTo DIRECTPOINTER_Err
    '</EhHeader>
  
  DIRECTPOINTER = Send_SCI_message(SCI_GETDIRECTPOINTER, 0, 0)
  
    '<EhFooter>
    Exit Property

DIRECTPOINTER_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.DIRECTPOINTER " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETDOCPOINTER
Public Property Get DOCPOINTER() As Long
Attribute DOCPOINTER.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo DOCPOINTER_Err
    '</EhHeader>
  
  DOCPOINTER = Send_SCI_message(SCI_GETDOCPOINTER, 0, 0)
  
    '<EhFooter>
    Exit Property

DOCPOINTER_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.DOCPOINTER " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETDOCPOINTER(<unused>, document *pDoc)
Public Property Let DOCPOINTER(pDoc As Long)
    '<EhHeader>
    On Error GoTo DOCPOINTER_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETDOCPOINTER, 0, pDoc
  
    '<EhFooter>
    Exit Property

DOCPOINTER_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.DOCPOINTER " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_CREATEDOCUMENT
Public Function CREATEDOCUMENT() As Long
    '<EhHeader>
    On Error GoTo CREATEDOCUMENT_Err
    '</EhHeader>
  
  CREATEDOCUMENT = Send_SCI_message(SCI_CREATEDOCUMENT, 0, 0)
  
    '<EhFooter>
    Exit Function

CREATEDOCUMENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.CREATEDOCUMENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_ADDREFDOCUMENT(<unused>, document *pDoc)
Public Function ADDREFDOCUMENT(pDoc As Long) As Long
    '<EhHeader>
    On Error GoTo ADDREFDOCUMENT_Err
    '</EhHeader>
  
  ADDREFDOCUMENT = Send_SCI_message(SCI_ADDREFDOCUMENT, 0, pDoc)
  
    '<EhFooter>
    Exit Function

ADDREFDOCUMENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ADDREFDOCUMENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_RELEASEDOCUMENT(<unused>, document *pDoc)
Public Function RELEASEDOCUMENT(pDoc As Long) As Long
    '<EhHeader>
    On Error GoTo RELEASEDOCUMENT_Err
    '</EhHeader>
  
  RELEASEDOCUMENT = Send_SCI_message(SCI_RELEASEDOCUMENT, 0, pDoc)
  
    '<EhFooter>
    Exit Function

RELEASEDOCUMENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.RELEASEDOCUMENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_VISIBLEFROMDOCLINE(int docLine)
Public Function VISIBLEFROMDOCLINE(docLine As Long) As Long
    '<EhHeader>
    On Error GoTo VISIBLEFROMDOCLINE_Err
    '</EhHeader>
  
  VISIBLEFROMDOCLINE = Send_SCI_message(SCI_VISIBLEFROMDOCLINE, docLine, 0)
  
    '<EhFooter>
    Exit Function

VISIBLEFROMDOCLINE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.VISIBLEFROMDOCLINE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_DOCLINEFROMVISIBLE(int displayLine)
Public Function DOCLINEFROMVISIBLE(displayLine As Long) As Long
    '<EhHeader>
    On Error GoTo DOCLINEFROMVISIBLE_Err
    '</EhHeader>
  
  DOCLINEFROMVISIBLE = Send_SCI_message(SCI_DOCLINEFROMVISIBLE, displayLine, 0)
  
    '<EhFooter>
    Exit Function

DOCLINEFROMVISIBLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.DOCLINEFROMVISIBLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SHOWLINES(int lineStart, lineEnd)
Public Function SHOWLINES(lineStart As Long, lineEnd As Long) As Long
    '<EhHeader>
    On Error GoTo SHOWLINES_Err
    '</EhHeader>
  
  SHOWLINES = Send_SCI_message(SCI_SHOWLINES, lineStart, lineEnd)
  
    '<EhFooter>
    Exit Function

SHOWLINES_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.SHOWLINES " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_HIDELINES(int lineStart, lineEnd)
Public Function HIDELINES(lineStart As Long, lineEnd As Long) As Long
    '<EhHeader>
    On Error GoTo HIDELINES_Err
    '</EhHeader>
  
  HIDELINES = Send_SCI_message(SCI_HIDELINES, lineStart, lineEnd)
  
    '<EhFooter>
    Exit Function

HIDELINES_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.HIDELINES " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETLINEVISIBLE(int line)
Public Property Get LINEVISIBLE(line As Long) As Long
    '<EhHeader>
    On Error GoTo LINEVISIBLE_Err
    '</EhHeader>
  
  LINEVISIBLE = Send_SCI_message(SCI_GETLINEVISIBLE, line, 0)
  
    '<EhFooter>
    Exit Property

LINEVISIBLE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LINEVISIBLE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETFOLDLEVEL(int line, level)
Public Property Let FOLDLEVEL(line As Long, level As Long)
    '<EhHeader>
    On Error GoTo FOLDLEVEL_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETFOLDLEVEL, line, level
  
    '<EhFooter>
    Exit Property

FOLDLEVEL_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FOLDLEVEL " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETFOLDLEVEL(int line)
Public Property Get FOLDLEVEL(line As Long) As Long
Attribute FOLDLEVEL.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo FOLDLEVEL_Err
    '</EhHeader>
  
  FOLDLEVEL = Send_SCI_message(SCI_GETFOLDLEVEL, line, 0)
  
    '<EhFooter>
    Exit Property

FOLDLEVEL_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FOLDLEVEL " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETFOLDFLAGS(int flags)
Public Property Let FOLDFLAGS(flags As Long)
    '<EhHeader>
    On Error GoTo FOLDFLAGS_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETFOLDFLAGS, flags, 0
  
    '<EhFooter>
    Exit Property

FOLDFLAGS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FOLDFLAGS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property





'SCI_SETFOLDEXPANDED(int line, bool expanded)
Public Property Let FOLDEXPANDED(line As Long, expanded As Boolean)
    '<EhHeader>
    On Error GoTo FOLDEXPANDED_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETFOLDEXPANDED, line, 1 And expanded
  
    '<EhFooter>
    Exit Property

FOLDEXPANDED_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FOLDEXPANDED " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETFOLDEXPANDED(int line)
Public Property Get FOLDEXPANDED(line As Long) As Boolean
Attribute FOLDEXPANDED.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo FOLDEXPANDED_Err
    '</EhHeader>
  
  FOLDEXPANDED = Send_SCI_message(SCI_GETFOLDEXPANDED, line, 0)
  
    '<EhFooter>
    Exit Property

FOLDEXPANDED_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FOLDEXPANDED " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_TOGGLEFOLD(int line)
Public Function TOGGLEFOLD(line As Long) As Long
    '<EhHeader>
    On Error GoTo TOGGLEFOLD_Err
    '</EhHeader>
  
  TOGGLEFOLD = Send_SCI_message(SCI_TOGGLEFOLD, line, 0)
  
    '<EhFooter>
    Exit Function

TOGGLEFOLD_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.TOGGLEFOLD " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_ENSUREVISIBLE(int line)
Public Function EnsureVisible(line As Long) As Long
    '<EhHeader>
    On Error GoTo EnsureVisible_Err
    '</EhHeader>
  
  EnsureVisible = Send_SCI_message(SCI_ENSUREVISIBLE, line, 0)
  
    '<EhFooter>
    Exit Function

EnsureVisible_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EnsureVisible " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_ENSUREVISIBLEENFORCEPOLICY(int line)
Public Function ENSUREVISIBLEENFORCEPOLICY(line As Long) As Long
    '<EhHeader>
    On Error GoTo ENSUREVISIBLEENFORCEPOLICY_Err
    '</EhHeader>
  
  ENSUREVISIBLEENFORCEPOLICY = Send_SCI_message(SCI_ENSUREVISIBLEENFORCEPOLICY, line, 0)
  
    '<EhFooter>
    Exit Function

ENSUREVISIBLEENFORCEPOLICY_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ENSUREVISIBLEENFORCEPOLICY " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_GETLASTCHILD(int startLine, level)
Public Property Get LASTCHILD(startLine As Long, level As Long) As Long
    '<EhHeader>
    On Error GoTo LASTCHILD_Err
    '</EhHeader>
  
  LASTCHILD = Send_SCI_message(SCI_GETLASTCHILD, startLine, level)
  
    '<EhFooter>
    Exit Property

LASTCHILD_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LASTCHILD " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETFOLDPARENT(int startLine)
Public Property Get FOLDPARENT(startLine As Long) As Long
    '<EhHeader>
    On Error GoTo FOLDPARENT_Err
    '</EhHeader>
  
  FOLDPARENT = Send_SCI_message(SCI_GETFOLDPARENT, startLine, 0)
  
    '<EhFooter>
    Exit Property

FOLDPARENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.FOLDPARENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETWRAPMODE(int wrapMode)
Public Property Let wrapMode(wrapMode As Long)
    '<EhHeader>
    On Error GoTo wrapMode_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETWRAPMODE, wrapMode, 0
  
    '<EhFooter>
    Exit Property

wrapMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.wrapMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETWRAPMODE
Public Property Get wrapMode() As Long
Attribute wrapMode.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo wrapMode_Err
    '</EhHeader>
  
  wrapMode = Send_SCI_message(SCI_GETWRAPMODE, 0, 0)
  
    '<EhFooter>
    Exit Property

wrapMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.wrapMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETWRAPVISUALFLAGS(int wrapVisualFlags)
Public Property Let wrapVisualFlags(wrapVisualFlags As Long)
    '<EhHeader>
    On Error GoTo wrapVisualFlags_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETWRAPVISUALFLAGS, wrapVisualFlags, 0
  
    '<EhFooter>
    Exit Property

wrapVisualFlags_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.wrapVisualFlags " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETWRAPVISUALFLAGS
Public Property Get wrapVisualFlags() As Long
Attribute wrapVisualFlags.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo wrapVisualFlags_Err
    '</EhHeader>
  
  wrapVisualFlags = Send_SCI_message(SCI_GETWRAPVISUALFLAGS, 0, 0)
  
    '<EhFooter>
    Exit Property

wrapVisualFlags_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.wrapVisualFlags " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETWRAPSTARTINDENT(int indent)
Public Property Let WRAPSTARTINDENT(Indent As Long)
    '<EhHeader>
    On Error GoTo WRAPSTARTINDENT_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETWRAPSTARTINDENT, Indent, 0
  
    '<EhFooter>
    Exit Property

WRAPSTARTINDENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WRAPSTARTINDENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETWRAPSTARTINDENT
Public Property Get WRAPSTARTINDENT() As Long
Attribute WRAPSTARTINDENT.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo WRAPSTARTINDENT_Err
    '</EhHeader>
  
  WRAPSTARTINDENT = Send_SCI_message(SCI_GETWRAPSTARTINDENT, 0, 0)
  
    '<EhFooter>
    Exit Property

WRAPSTARTINDENT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WRAPSTARTINDENT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETLAYOUTCACHE(int cacheMode)
Public Property Let LAYOUTCACHE(cacheMode As Long)
    '<EhHeader>
    On Error GoTo LAYOUTCACHE_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETLAYOUTCACHE, cacheMode, 0
  
    '<EhFooter>
    Exit Property

LAYOUTCACHE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LAYOUTCACHE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETLAYOUTCACHE
Public Property Get LAYOUTCACHE() As Long
Attribute LAYOUTCACHE.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo LAYOUTCACHE_Err
    '</EhHeader>
  
  LAYOUTCACHE = Send_SCI_message(SCI_GETLAYOUTCACHE, 0, 0)
  
    '<EhFooter>
    Exit Property

LAYOUTCACHE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LAYOUTCACHE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


'SCI_LINESJOIN
Public Function LINESJOIN() As Long
    '<EhHeader>
    On Error GoTo LINESJOIN_Err
    '</EhHeader>
  
  LINESJOIN = Send_SCI_message(SCI_LINESJOIN, 0, 0)
  
    '<EhFooter>
    Exit Function

LINESJOIN_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LINESJOIN " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETWRAPVISUALFLAGSLOCATION(int wrapVisualFlagsLocation)
Public Property Let wrapVisualFlagsLocation(wrapVisualFlagsLocation As Long)
    '<EhHeader>
    On Error GoTo wrapVisualFlagsLocation_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETWRAPVISUALFLAGSLOCATION, wrapVisualFlagsLocation, 0
  
    '<EhFooter>
    Exit Property

wrapVisualFlagsLocation_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.wrapVisualFlagsLocation " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETWRAPVISUALFLAGSLOCATION
Public Property Get wrapVisualFlagsLocation() As Long
Attribute wrapVisualFlagsLocation.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo wrapVisualFlagsLocation_Err
    '</EhHeader>
  
  wrapVisualFlagsLocation = Send_SCI_message(SCI_GETWRAPVISUALFLAGSLOCATION, 0, 0)
  
    '<EhFooter>
    Exit Property

wrapVisualFlagsLocation_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.wrapVisualFlagsLocation " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_LINESSPLIT(int pixelWidth)
Public Function LINESSPLIT(pixelWidth As Long) As Long
    '<EhHeader>
    On Error GoTo LINESSPLIT_Err
    '</EhHeader>
  
  LINESSPLIT = Send_SCI_message(SCI_LINESSPLIT, pixelWidth, 0)
  
    '<EhFooter>
    Exit Function

LINESSPLIT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LINESSPLIT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_ZOOMIN
Public Function ZOOMIN() As Long
    '<EhHeader>
    On Error GoTo ZOOMIN_Err
    '</EhHeader>
  
  ZOOMIN = Send_SCI_message(SCI_ZOOMIN, 0, 0)
  
    '<EhFooter>
    Exit Function

ZOOMIN_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ZOOMIN " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_ZOOMOUT
Public Function ZOOMOUT() As Long
    '<EhHeader>
    On Error GoTo ZOOMOUT_Err
    '</EhHeader>
  
  ZOOMOUT = Send_SCI_message(SCI_ZOOMOUT, 0, 0)
  
    '<EhFooter>
    Exit Function

ZOOMOUT_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ZOOMOUT " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETZOOM(int zoomInPoints)
Public Property Let ZOOM(zoomInPoints As Long)
    '<EhHeader>
    On Error GoTo ZOOM_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETZOOM, zoomInPoints, 0
  
    '<EhFooter>
    Exit Property

ZOOM_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ZOOM " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETZOOM
Public Property Get ZOOM() As Long
Attribute ZOOM.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo ZOOM_Err
    '</EhHeader>
  
  ZOOM = Send_SCI_message(SCI_GETZOOM, 0, 0)
  
    '<EhFooter>
    Exit Property

ZOOM_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ZOOM " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property



'SCI_GETEDGEMODE
Public Property Get edgeMode() As Long
Attribute edgeMode.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo edgeMode_Err
    '</EhHeader>
  
  edgeMode = Send_SCI_message(SCI_GETEDGEMODE, 0, 0)
  
    '<EhFooter>
    Exit Property

edgeMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.edgeMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETEDGECOLUMN(int column)
Public Property Let EDGECOLUMN(Column As Long)
    '<EhHeader>
    On Error GoTo EDGECOLUMN_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETEDGECOLUMN, Column, 0
  
    '<EhFooter>
    Exit Property

EDGECOLUMN_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EDGECOLUMN " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETEDGECOLUMN
Public Property Get EDGECOLUMN() As Long
Attribute EDGECOLUMN.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo EDGECOLUMN_Err
    '</EhHeader>
  
  EDGECOLUMN = Send_SCI_message(SCI_GETEDGECOLUMN, 0, 0)
  
    '<EhFooter>
    Exit Property

EDGECOLUMN_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EDGECOLUMN " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETEDGECOLOUR(int colour)
Public Property Let EDGECOLOUR(colour As Long)
    '<EhHeader>
    On Error GoTo EDGECOLOUR_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETEDGECOLOUR, colour, 0
  
    '<EhFooter>
    Exit Property

EDGECOLOUR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EDGECOLOUR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETEDGECOLOUR
Public Property Get EDGECOLOUR() As Long
Attribute EDGECOLOUR.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo EDGECOLOUR_Err
    '</EhHeader>
  
  EDGECOLOUR = Send_SCI_message(SCI_GETEDGECOLOUR, 0, 0)
  
    '<EhFooter>
    Exit Property

EDGECOLOUR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.EDGECOLOUR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETEDGEMODE(int edgeMode)
Public Property Let edgeMode(edgeMode As Long)
    '<EhHeader>
    On Error GoTo edgeMode_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETEDGEMODE, edgeMode, 0
  
    '<EhFooter>
    Exit Property

edgeMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.edgeMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETLEXER(int lexer)
Public Property Let Lexer(llexer As sci_SCLex)
    '<EhHeader>
    On Error GoTo Lexer_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETLEXER, llexer, 0
  
    '<EhFooter>
    Exit Property

Lexer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Lexer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETLEXER
Public Property Get Lexer() As sci_SCLex
    '<EhHeader>
    On Error GoTo Lexer_Err
    '</EhHeader>
  
  Lexer = Send_SCI_message(SCI_GETLEXER, 0, 0)
  
    '<EhFooter>
    Exit Property

Lexer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.Lexer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property




'SCI_COLOURISE(int start, end)
Public Function COLOURISE(start As Long, ends As Long) As Long
    '<EhHeader>
    On Error GoTo COLOURISE_Err
    '</EhHeader>
  
  COLOURISE = Send_SCI_message(SCI_COLOURISE, start, ends)
  
    '<EhFooter>
    Exit Function

COLOURISE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.COLOURISE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'SCI_SETPROPERTY(const char *key, const char *value)
Public Property Let PROPERTY(key As String, Value As String)
    '<EhHeader>
    On Error GoTo PROPERTY_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_SETPROPERTY, StrPtr(StrConv(key, vbFromUnicode)), Value
  
    '<EhFooter>
    Exit Property

PROPERTY_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.PROPERTY " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETKEYWORDS(int keyWordSet, const char *keyWordList)
Public Property Let KEYWORDS(keyWordSet As Long, keyWordList As String)
    '<EhHeader>
    On Error GoTo KEYWORDS_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_SETKEYWORDS, keyWordSet, keyWordList
  
    '<EhFooter>
    Exit Property

KEYWORDS_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.KEYWORDS " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETLEXERLANGUAGE(<unused>, const char *name)
Public Property Let LEXERLANGUAGE(Name As String)
    '<EhHeader>
    On Error GoTo LEXERLANGUAGE_Err
    '</EhHeader>
  
  Send_SCI_messageStr SCI_SETLEXERLANGUAGE, 0, Name
  
    '<EhFooter>
    Exit Property

LEXERLANGUAGE_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LEXERLANGUAGE " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_LOADLEXERLIBRARY(<unused>, const char *path)
Public Function LOADLEXERLIBRARY(path As String) As Long
    '<EhHeader>
    On Error GoTo LOADLEXERLIBRARY_Err
    '</EhHeader>
  
  LOADLEXERLIBRARY = Send_SCI_messageStr(SCI_LOADLEXERLIBRARY, 0, path)
  
    '<EhFooter>
    Exit Function

LOADLEXERLIBRARY_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.LOADLEXERLIBRARY " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


'SCI_SETMODEVENTMASK(int eventMask)
Public Property Let MODEVENTMASK(eventMask As Long)
    '<EhHeader>
    On Error GoTo MODEVENTMASK_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMODEVENTMASK, eventMask, 0
  
    '<EhFooter>
    Exit Property

MODEVENTMASK_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MODEVENTMASK " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMODEVENTMASK
Public Property Get MODEVENTMASK() As Long
Attribute MODEVENTMASK.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MODEVENTMASK_Err
    '</EhHeader>
  
  MODEVENTMASK = Send_SCI_message(SCI_GETMODEVENTMASK, 0, 0)
  
    '<EhFooter>
    Exit Property

MODEVENTMASK_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MODEVENTMASK " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_SETMOUSEDWELLTIME
Public Property Let MOUSEDWELLTIME(tm As Long)
    '<EhHeader>
    On Error GoTo MOUSEDWELLTIME_Err
    '</EhHeader>
  
  Send_SCI_message SCI_SETMOUSEDWELLTIME, tm, 0

    '<EhFooter>
    Exit Property

MOUSEDWELLTIME_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MOUSEDWELLTIME " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'SCI_GETMOUSEDWELLTIME
Public Property Get MOUSEDWELLTIME() As Long
Attribute MOUSEDWELLTIME.VB_ProcData.VB_Invoke_Property = "MainSettings"
    '<EhHeader>
    On Error GoTo MOUSEDWELLTIME_Err
    '</EhHeader>
  
  MOUSEDWELLTIME = Send_SCI_message(SCI_GETMOUSEDWELLTIME, 0, 0)
  
    '<EhFooter>
    Exit Property

MOUSEDWELLTIME_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.MOUSEDWELLTIME " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property



Private Sub WriteToPropBag(toB As PropertyBag, Data As Variant)
    '<EhHeader>
    On Error GoTo WriteToPropBag_Err
    '</EhHeader>

    toB.WriteProperty CStr(lppw), Data
    lppw = lppw + 1

    '<EhFooter>
    Exit Sub

WriteToPropBag_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.WriteToPropBag " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Function ReadfromPropBag(toB As PropertyBag, Optional def As Variant = "none") As Variant
    '<EhHeader>
    On Error GoTo ReadfromPropBag_Err
    '</EhHeader>
    
    If def = "none" Then
        ReadfromPropBag = toB.ReadProperty(CStr(lppw), "")
    Else
        ReadfromPropBag = toB.ReadProperty(CStr(lppw), def)
    End If
    lppw = lppw + 1

    '<EhFooter>
    Exit Function

ReadfromPropBag_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.ScintillaEdit.ReadfromPropBag " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
