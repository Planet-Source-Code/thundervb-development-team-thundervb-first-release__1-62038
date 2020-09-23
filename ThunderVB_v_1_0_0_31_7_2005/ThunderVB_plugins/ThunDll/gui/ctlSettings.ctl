VERSION 5.00
Begin VB.UserControl ctlSettings 
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   ScaleHeight     =   5670
   ScaleWidth      =   4455
   Begin VB.CheckBox chbExportSymbols 
      Caption         =   "* Export functions"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox txtBaseAddress 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "*"
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CheckBox chbLinkAsDLL 
      Caption         =   "* Create DLL"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtEntryPoint 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "*"
      Top             =   3480
      Width           =   1245
   End
   Begin VB.CheckBox chbDebugPreLoader 
      Caption         =   "* Debug ""Pre-Loader"""
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CheckBox chbFullLoading 
      Caption         =   "* Full loading"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox chbUsePreLoader 
      Caption         =   "* Use ""Pre-Loader"""
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cdmDLL_AddDllMain 
      Caption         =   "Add DllMain"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblDLL_1 
      AutoSize        =   -1  'True
      Caption         =   "* Base address  &&H"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   1365
   End
   Begin VB.Label lblDLL_2 
      AutoSize        =   -1  'True
      Caption         =   "* Entry-Point name"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   3480
      Width           =   1305
   End
End
Attribute VB_Name = "ctlSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cdmDLL_AddDllMain_Click()
    frmDllMain.Show vbModal, Me
End Sub

Public Function GetHwnd() As Long
    GetHwnd = UserControl.hWnd
End Function

Public Function GetEntryPointName() As String
    GetEntryPointName = txtEntryPoint.Text
End Function

Public Function GetMyObject(sName As String) As Control
    GetMyObject = UserControl.Controls(sName)
End Function

Public Function GetControlValue(sName As String) As String
Dim oControl As Control
    oControl = UserControl.Controls(sName)
    
    If oControl Is Nothing Then
        GetControlValue = vbNullString
    ElseIf TypeOf oControl Is TextBox Then
        GetControlValue = oControl.Text
    ElseIf TypeOf oControl Is CheckBox Or TypeOf oControl Is OptionButton Then
        GetControlValue = oControl.Value
    ElseIf TypeOf oControl Is Label Or TypeOf oControl Is CommandButton Then
        GetControlValue = oControl.Caption
    End If
    
End Function

