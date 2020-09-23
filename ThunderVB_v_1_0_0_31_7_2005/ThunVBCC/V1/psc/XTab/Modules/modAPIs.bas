Attribute VB_Name = "modAPIs"
'modAPI: Holds all the API declarations

Option Explicit


'===Types=============================================================================================================


Public Type OPENFILENAME      'for GetOpenFileName
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


'Bitmap type used to store Bitmap Data

'=====================================================================================================================
'===Constants=========================================================================================================

'Draw Text Constants
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000

'Subclass related constants
'Public Const GWL_WNDPROC = (-4)
'Public Const GWL_USERDATA = (-21)
'Public Const WM_LBUTTONDOWN As Long = &H201

'BitBlt Related Constants
Public Const SRCCOPY As Long = &HCC0020
Public Const SRCAND As Long = &H8800C6
Public Const SRCPAINT As Long = &HEE0086
Public Const SRCINVERT As Long = &H660046
Public Const WHITENESS As Long = &HFF0062

'DrawIcon Related Constants
Public Const DI_NORMAL As Long = &H3


'GetSystemMetrics Related Condtants
'Public Const SM_CXICON As Long = 11
'Public Const SM_CYICON As Long = 12
'Public Const SM_CXSMICON As Long = 49
'Public Const SM_CYSMICON As Long = 50
'=====================================================================================================================

'===Declarations======================================================================================================

'Drawing/Painting Declarations

Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long


'File Open Dialog Related Declarations
Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long

'Subclassing Related Declararions
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'Added 5th Oct 2004
'Following functions are used to preload and destroy Shell32.dll. See the Bug Fixes section for details
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long



'=====================================================================================================================



