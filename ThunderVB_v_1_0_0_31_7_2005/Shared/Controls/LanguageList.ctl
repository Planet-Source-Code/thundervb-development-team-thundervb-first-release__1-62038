VERSION 5.00
Begin VB.UserControl LanguageList 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox cmbLang 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "LanguageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event LanguageChanged(newLang As tvb_Languages)

Private Sub cmbLang_Change()
    RaiseEvent LanguageChanged(cmbLang.ItemData(cmbLang.ListIndex))
End Sub

Private Sub cmbLang_Click()
    RaiseEvent LanguageChanged(cmbLang.ItemData(cmbLang.ListIndex))
End Sub

Private Sub UserControl_Initialize()
    Dim i As Long
    cmbLang.Clear
    For i = tvb_Languages.tvb_min To tvb_Languages.tvb_max
        cmbLang.AddItem Resource_LanguageIdToString(i)
        cmbLang.ItemData(cmbLang.ListCount - 1) = i
    Next i
    cmbLang.ListIndex = 0
    
End Sub
Public Sub SetLanguage(lang As tvb_Languages)

    On Error Resume Next
    
    
    If lang = Resource_GetCurLang() Then Exit Sub
    
    Dim i As Long
    For i = 0 To cmbLang.ListCount - 1
        If cmbLang.ItemData(i) = lang Then
            cmbLang.ListIndex = i
        End If
    Next i

    RaiseEvent LanguageChanged(cmbLang.ItemData(cmbLang.ListIndex))
    
End Sub
Private Sub UserControl_Resize()
    cmbLang.width = UserControl.width
    UserControl.height = cmbLang.height
End Sub
