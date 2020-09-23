VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
Begin VB.Form frmDllMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DllMain template"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin ThunVBCC_v1.UniLabel lblInfo 
      Height          =   225
      Left            =   120
      Top             =   480
      Width           =   600
      _ExtentX        =   79
      _ExtentY        =   26
      AutoSize        =   -1  'True
      CaptionB        =   "frmDllMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.isButton cmdClose 
      Height          =   300
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Icon            =   "frmDllMain.frx":0030
      Style           =   5
      Caption         =   "Close"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.isButton cmdToClipboard 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Icon            =   "frmDllMain.frx":004C
      Style           =   5
      Caption         =   "Clipboard"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.isButton cmdToCursor 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Icon            =   "frmDllMain.frx":0068
      Style           =   5
      Caption         =   "Cursor"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.HzxYCheckBox chbDllMain 
      Height          =   240
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0
      Caption         =   "Add DllMain function"
      Pic_UncheckedNormal=   "frmDllMain.frx":0084
      Pic_CheckedNormal=   "frmDllMain.frx":03D6
      Pic_MixedNormal =   "frmDllMain.frx":0728
      Pic_UncheckedDisabled=   "frmDllMain.frx":0A7A
      Pic_CheckedDisabled=   "frmDllMain.frx":0DCC
      Pic_MixedDisabled=   "frmDllMain.frx":111E
      Pic_UncheckedOver=   "frmDllMain.frx":1470
      Pic_CheckedOver =   "frmDllMain.frx":17C2
      Pic_MixedOver   =   "frmDllMain.frx":1B14
      Pic_UncheckedDown=   "frmDllMain.frx":1E66
      Pic_CheckedDown =   "frmDllMain.frx":21B8
      Pic_MixedDown   =   "frmDllMain.frx":250A
   End
   Begin ThunVBCC_v1.HzxYCheckBox chbConst 
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2434
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0
      Caption         =   "Add constants"
      Pic_UncheckedNormal=   "frmDllMain.frx":285C
      Pic_CheckedNormal=   "frmDllMain.frx":2BAE
      Pic_MixedNormal =   "frmDllMain.frx":2F00
      Pic_UncheckedDisabled=   "frmDllMain.frx":3252
      Pic_CheckedDisabled=   "frmDllMain.frx":35A4
      Pic_MixedDisabled=   "frmDllMain.frx":38F6
      Pic_UncheckedOver=   "frmDllMain.frx":3C48
      Pic_CheckedOver =   "frmDllMain.frx":3F9A
      Pic_MixedOver   =   "frmDllMain.frx":42EC
      Pic_UncheckedDown=   "frmDllMain.frx":463E
      Pic_CheckedDown =   "frmDllMain.frx":4990
      Pic_MixedDown   =   "frmDllMain.frx":4CE2
   End
End
Attribute VB_Name = "frmDllMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'13.09. 2004 - initial version - GUI, code
'23.04. 2005 - new gui - unicode controls, code revision

'constants
Private Const CONST_DLL As String = "Private Const DLL_PROCESS_ATTACH As Long = 1" & vbCrLf & _
                                    "Private Const DLL_PROCESS_DETACH As Long = 0" & vbCrLf & _
                                    "Private Const DLL_THREAD_ATTACH As Long = 2" & vbCrLf & _
                                    "Private Const DLL_THREAD_DETACH As Long = 3"

'DllMain
Private Const DLLMAIN_TEMPLATE As String = "Public Function DllMain(ByVal hInstDLL As Long, ByVal fdwReason As Long, ByVal lpvReserved As Long) As Long" & vbCrLf _
                                          & vbTab & "Select Case fdwReason" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_PROCESS_ATTACH" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_PROCESS_DETACH" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_THREAD_ATTACH" & vbCrLf _
                                          & vbTab & vbTab & "Case DLL_THREAD_DETACH" & vbCrLf _
                                          & vbTab & "End Select" & vbCrLf _
                                          & "End Function"

'dllmain name - default
Private Const DLLMAIN_NAME As String = "DllMain"

Private Const FORM_NAME As String = "frmDllMain"

'----------------------
'--- CONTROL EVENTS ---
'----------------------

Private Sub Form_Load()
    LogMsg "Loading " & Add34(Me.caption) & " window", FORM_NAME, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LogMsg "Unloading " & Add34(Me.caption) & " window", FORM_NAME, "Form_Unload"
End Sub

'close
Private Sub cmdClose_Click()
    Unload Me
End Sub

'paste code to clipboard
Private Sub cmdToClipboard_Click()
Dim sCode As String
    
    sCode = GetCode
    'replace DllMain with user defined
    If Len(frmIn.set_txtEntryPoint.Text) <> 0 Then sCode = Replace(sCode, DLLMAIN_NAME, frmIn.set_txtEntryPoint.Text)
    
    'save to clipboard
    Clipboard.Clear
    Clipboard.SetText sCode
    
End Sub

'paste template to cursor
Private Sub cmdToCursor_Click()
Dim sCode As String

    sCode = GetCode
  
    If Len(frmIn.set_txtEntryPoint.Text) <> 0 Then sCode = Replace(sCode, DLLMAIN_NAME, frmIn.set_txtEntryPoint.Text, , 1)

    'paste to cursor
    PutStringToCurCursor sCode
    
End Sub

'------------------------
'--- HELPER FUNCTIONS ---
'------------------------

'return template
Private Function GetCode() As String
    
    'add constants
    If chbConst.Value = 1 Then GetCode = CONST_DLL
    
    'add dllmain
    If chbDllMain.Value = 1 Then
        
        If Len(GetCode) <> 0 Then
            'add new line
            GetCode = GetCode & CrLf(2) & DLLMAIN_TEMPLATE
        Else
            GetCode = DLLMAIN_TEMPLATE
        End If
        
    End If

End Function
