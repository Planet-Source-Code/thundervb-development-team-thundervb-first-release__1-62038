Attribute VB_Name = "modPacker"
Option Explicit

Public oThunVB As ThunderVB_base
Public oMe As plugin

Public Const PLUGIN_NAME As String = "Packer"
Public Const MSG_TITLE As String = PLUGIN_NAME
Public Const APP_NAME As String = PLUGIN_NAME


Public Const PLUGIN_NAMEs As String = "Packer"
Public Const MSG_TITLEs As String = PLUGIN_NAMEs
Public Const APP_NAMEs As String = PLUGIN_NAMEs


Dim cph As Long, bSet As Boolean

'set path to the program
'- sDialogTitle - open dialog Title
'-txtTarget     - textbox where path will be stored
'-sAppName      - app name (eg. ml.exe or midl.exe)

Public Sub SetPath(sDialogTitle As String, txtTarget As TextBox, Optional sAppName As String = "")

    'set dialog title
    frmIn.cdSet.DialogTitle = sDialogTitle
    frmIn.cdSet.FileName = ""

    'set new init directory
    If Len(txtTarget.Text) <> 0 Then frmIn.cdSet.InitDir = Left(txtTarget.Text, InStrRev(txtTarget.Text, "\")) Else frmIn.cdSet.InitDir = App.Path & "\"

On Error Resume Next

    'select file
    frmIn.cdSet.ShowOpen
    'cancel was pressed
    If Err.Number = 32755 Then Exit Sub

On Error GoTo 0

    'check predefined app name
    If Len(sAppName) = 0 Then GoTo 10

    'check filename
    If StrComp(Right(frmIn.cdSet.FileName, Len(sAppName)), sAppName, vbTextCompare) <> 0 Then
        MsgBox "Select " & Add34(sAppName) & " file.", vbInformation, "Settings"
    Else
10:
        'store path to the textbox
        txtTarget.Text = frmIn.cdSet.FileName
    End If

End Sub

Public Sub Init_Hook()
Dim t As clsCpHook
    
    If bSet = False Then
    
        Set t = New clsCpHook
        cph = AddCPH(t)
        LogMsg "CPH added " & cph, PLUGIN_NAMEs, "modPacker", "Init_Hook"
        bSet = True
        
    End If
    
End Sub

Public Sub Unload_Hook()
    
    If bSet = True Then
        
        RemoveCPH cph
        LogMsg "CPH removed " & cph, PLUGIN_NAMEs, "modPacker", "Unload_Hook"
        cph = 0
        bSet = False
        
    End If
    
End Sub
