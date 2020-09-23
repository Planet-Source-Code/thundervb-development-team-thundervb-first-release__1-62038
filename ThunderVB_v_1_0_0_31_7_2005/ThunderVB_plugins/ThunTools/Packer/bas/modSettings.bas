Attribute VB_Name = "modSettings"
Option Explicit

'module for settings only

'*** Packer ***

Private bPacker_UsePacker As Boolean           'use packer
Private bPacker_ShowPackerOutPut As Boolean    'packer output
Private sPacker_CommandLine As String          'packer command line
Private sPacker_CmdLineDesc As String          'description of cmdline
Private sPacker_Path As String                 'path to packer

Public Enum PACKER_
    UsePacker
    CommandLine
    CmdLineDescription
    ShowPackerOutPut
    Path
End Enum

'--------------
'--- PACKER ---
'--------------

'get Packer settings
'parameter - ePacker - packer setting
'return -  "False"/"True" or string

Public Function Get_Packer(ByVal ePacker As PACKER_) As String

    Select Case ePacker
        Case PACKER_.UsePacker
            Get_Packer = bPacker_UsePacker
        Case PACKER_.ShowPackerOutPut
            Get_Packer = bPacker_ShowPackerOutPut
        Case PACKER_.CommandLine
            Get_Packer = sPacker_CommandLine
        Case PACKER_.CmdLineDescription
            Get_Packer = sPacker_CmdLineDesc
        Case PACKER_.Path
            Get_Packer = sPacker_Path
    End Select
    
End Function

'change Packer flags
'parameters - ePacker - flags
'           - sNewValue - new flag/path

Public Sub Let_Packer(ByVal ePacker As PACKER_, ByVal sNewValue As String)

    Select Case ePacker
        Case PACKER_.UsePacker
            bPacker_UsePacker = CBool(sNewValue)
        Case PACKER_.ShowPackerOutPut
            bPacker_ShowPackerOutPut = CBool(sNewValue)
        Case PACKER_.CommandLine
            sPacker_CommandLine = sNewValue
        Case PACKER_.CmdLineDescription
            sPacker_CmdLineDesc = sNewValue
        Case PACKER_.Path
            sPacker_Path = sNewValue
    End Select
    
End Sub

Public Function SaveSettingsToVariables(Optional bLoadForm As Boolean = False)
Dim sCmdLine As String, sCmdLineDesc As String, lPacker As Long
    
    If bLoadForm = True Then
        Load frmIn
        LoadSettings GLOBAL_, frmIn.pctSettings
        LoadSettings LOCAL_, frmIn.pctSettings
    End If
    
    GetSelPackerInfo sCmdLine, sCmdLineDesc
    
    With frmIn

        Let_Packer CmdLineDescription, sCmdLine
        Let_Packer CommandLine, sCmdLineDesc
        Let_Packer ShowPackerOutPut, .set_chbShowPackerOutPut.Value
        Let_Packer UsePacker, .set_chbUsePacker.Value
        Let_Packer Path, .set_txtPacker.Text

    End With
    
    If bLoadForm = True Then
        Unload frmIn
    End If
    
End Function

Public Function SaveOtherSettings(eScope As SET_SCOPE) As Boolean
    'other controls that must be saved
    'this function is called from SaveSettings (modSaveLoadSettings)
       
    If eScope = GLOBAL_ Then
    
        '--- global ---
    
    Else
    
        '--- local ---
    
    End If
    
    SaveOtherSettings = True
    
End Function

Public Function LoadOtherSettings(eScope As SET_SCOPE) As Boolean
    'other controls that must be inited
    'this function is called from LoadSettings (modSaveLoadSettings)
    
    If eScope = GLOBAL_ Then
    
        '--- global ---
    
    Else
    
        '--- local ---
        
        'one optionbutton must be checked
        With frmIn
            If .set_optPacker0.Value = False And .set_optPacker1.Value = False And .set_optPacker2.Value = False And .set_optPacker3.Value = False Then .set_optPacker0.Value = True
        End With
    
    End If
    
    LoadOtherSettings = True
    
End Function

Public Sub SetOtherDefaultSettings(eScope As SET_SCOPE)
    'other controls that must be set to default
    'this function is called from SetDefaultSetings (modSaveLoadSettings)
       
    If eScope = GLOBAL_ Then
    
        '--- global ---
    
    Else
    
        '--- local ---
            
    End If
       
End Sub

Private Function GetSelPackerInfo(Optional ByRef sCmdLine, Optional ByRef sCmdLineDesc)

    'one optionbutton must be checked
    With frmIn
        If .set_optPacker0.Value = False And .set_optPacker1.Value = False And .set_optPacker2.Value = False And .set_optPacker3.Value = False Then .set_optPacker0.Value = True
    End With

    With frmIn

        If .set_optPacker0.Value = True Then
            
            If IsMissing(sCmdLine) = False Then sCmdLine = ""
            If IsMissing(sCmdLineDesc) = False Then sCmdLineDesc = ""
        
        ElseIf .set_optPacker1.Value = True Then

            If IsMissing(sCmdLine) = False Then sCmdLine = .set_txtCmdLine1.Text
            If IsMissing(sCmdLineDesc) = False Then sCmdLineDesc = .set_txtDesc1.Text

        ElseIf .set_optPacker2.Value = True Then

            If IsMissing(sCmdLine) = False Then sCmdLine = .set_txtCmdLine2.Text
            If IsMissing(sCmdLineDesc) = False Then sCmdLineDesc = .set_txtDesc2.Text

        ElseIf .set_optPacker3.Value = True Then

            If IsMissing(sCmdLine) = False Then sCmdLine = .set_txtCmdLine3.Text
            If IsMissing(sCmdLineDesc) = False Then sCmdLineDesc = .set_txtDesc3.Text

        End If

    End With

End Function
