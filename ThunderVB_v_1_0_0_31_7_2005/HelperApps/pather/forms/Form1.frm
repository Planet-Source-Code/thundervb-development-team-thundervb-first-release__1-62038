VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    'InstallDllTemplate
    
    'If Len(Command$) = 0 Then End
    Dim pars() As String, i As Long, t() As String
    'pars = Split(Command$, ";")
    'For i = 0 To UBound(pars)
    '    t = Split(pars(i), "?,")
    '    SaveSetting t(0), t(1), t(2), App.Path & t(3)
    'Next i
    SaveSetting "ThunderVB", "pl_ThunderAsm", "set_Paths_txtMasm", App.Path & "\Masm\ml.exe"
    'these are not used ..
    SaveSetting "ThunderVB", "pl_ThunderAsm", "set_Paths_txtLibFiles", "c:\"
    SaveSetting "ThunderVB", "pl_ThunderAsm", "set_Paths_txtIncFiles", "c:\"
    
    MsgBox "To use inmoduleC you must setup the VC++ compiler path" & vbNewLine & _
           "on thunAsm options.You will need Vc++ 2002/2003 or VC++ resource kit [witch is free]"
           
    'Shell "Register_dlls.bat"
    Shell "InstallThundervb_pl.exe"
    'InstallDllTemplate
    MsgBox "To add dll template to vb's new project list" & vbNewLine & _
           "copy it to vbdir\templates\projects\"
    End
End Sub
