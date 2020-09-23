Attribute VB_Name = "modSettings"
Option Explicit

'module only for settings

Private Const MOD_NAME As String = "modSettings"

'*** General ***
'Tab General

Private bGeneral_PopUpExportWindow As Boolean
Private bGeneral_ListForAllModules As Boolean
Private bGeneral_HookCompiler As Boolean
Private bGeneral_SaveOBJ As Boolean
Private bGeneral_HideErrorDialogs As Boolean

Public Enum GENERAL_
    PopUpExportsWindow
    ListingsForAllModules
    HookCompiler
    SaveObjFiles
    HideErrorDialogs
End Enum

'*** PATHS ***
'Tab Paths

Private sPaths_Midl  As String        'path to midl.exe
Private sPaths_ML  As String          'path to ml.exe
Private sPaths_TextEditor As String   'path to text-editor
Private sPaths_INCFiles As String     'path to .INC files (directory)
Private sPaths_LIBFiles As String     'path to .LIB files (directory)
Private sPaths_CCompiler As String    'path to C Compiler

Public Enum PATHS_
    Midl
    ml
    TextEditor
    INCFiles_Directory
    LIBFiles_Directory
    CCompiler
End Enum

'*** Debug ***
'Tab Debug

Private bDebug_EnableOutPutToDebugLog As Boolean
Private bDebug_DelDebugLogBeforeCompiling As Boolean
Private bDebug_OutPutAssemblerMessToLog As Boolean
Private bDebug_OutPutMapFiles As Boolean

Private bDebug_DeleteLST As Boolean
Private bDebug_DeleteASM As Boolean

Public Enum DEBUG_
    EnableOutPutToDebugLog
    DeleteDebugLogBeforeCompiling
    OutPutAssemblerMessagesToLog
    OutPutMapFiles
    DeleteLST
    DeleteASM
End Enum

'*** Compile ***
'Tab Compile

Private bCompile_PauseBeforeAsm As Boolean
Private bCompile_PauseBeforeLink As Boolean
Private bCompile_ModifyCmdLine As Boolean
Private bCompile_SkipLinking As Boolean

Public Enum COMPILE_
    PauseBeforeAssembly
    PauseBeforeLinking
    ModifyCmdLine
    SkipLinking
End Enum

'*** ASM ***
'Tab ASM

Private bASM_FixASMListings As Boolean
Private bASM_CompileASMCode As Boolean

Public Enum ASM_
    FixASMListings
    CompileASMCode
End Enum

'*** C ***
'Tab C

Private bC_CompileCCode As Boolean

Public Enum C_
    CompileCCode
End Enum

'---------------
'--- GENERAL ---
'---------------

'return General settings
'parameter - eGeneral - General flag
'          - True/False - flag

Public Function Get_General(ByVal eGeneral As GENERAL_) As Boolean
         
10        Select Case eGeneral
              Case GENERAL_.HideErrorDialogs
20                Get_General = bGeneral_HideErrorDialogs
30            Case GENERAL_.HookCompiler
40                Get_General = bGeneral_HookCompiler
50            Case GENERAL_.ListingsForAllModules
60                Get_General = bGeneral_ListForAllModules
70            Case GENERAL_.PopUpExportsWindow
80                Get_General = bGeneral_PopUpExportWindow
90            Case GENERAL_.SaveObjFiles
100               Get_General = bGeneral_SaveOBJ
110       End Select

End Function

'change General flag
'parameters - eGeneral - general flag to change
'           - bNewValue - new flag

Public Sub Let_General(ByVal eGeneral As GENERAL_, bNewValue As Boolean)
         
10        Select Case eGeneral
              Case GENERAL_.HideErrorDialogs
20                bGeneral_HideErrorDialogs = bNewValue
30            Case GENERAL_.HookCompiler
40                bGeneral_HookCompiler = bNewValue
50            Case GENERAL_.ListingsForAllModules
60                bGeneral_ListForAllModules = bNewValue
70            Case GENERAL_.PopUpExportsWindow
80                bGeneral_PopUpExportWindow = bNewValue
90            Case GENERAL_.SaveObjFiles
100               bGeneral_SaveOBJ = bNewValue
110       End Select

End Sub

'-------------
'--- PATHS ---
'-------------

'return paths
'parameter - ePath - special path
'          - bWarning - when path is not specified, alert will appear
'return - ""     - path is not set
'       - string - path

Public Function Get_Paths(ByVal ePath As PATHS_, Optional bWarning As Boolean = False) As String
      Dim sText As String
         
10        Select Case ePath
              Case PATHS_.Midl
20                Get_Paths = sPaths_Midl
30                sText = "Path to MIDL (midl.exe) is not set." & vbCrLf & "Setting/Paths"
40            Case PATHS_.ml
50                Get_Paths = sPaths_ML
60                sText = "Path to ML (ml.exe) is not set." & vbCrLf & "Setting/Paths"
70            Case PATHS_.TextEditor
80                Get_Paths = sPaths_TextEditor
90                sText = "Path to your Text-Editor is not set." & vbCrLf & "Setting/Paths"
100           Case PATHS_.INCFiles_Directory
110               Get_Paths = sPaths_INCFiles
120               sText = "Path to .INC files is not set." & vbCrLf & "Setting/Paths"
130           Case PATHS_.LIBFiles_Directory
140               Get_Paths = sPaths_LIBFiles
150               sText = "Path to .LIB files is not set." & vbCrLf & "Setting/Paths"
160           Case PATHS_.CCompiler
170               Get_Paths = sPaths_CCompiler
180               sText = "Path to C compiler is not set." & vbCrLf & "Setting/Paths"
190       End Select

          'check path
200       If Len(Get_Paths) = 0 And bWarning = True Then
210           MsgBox sText, vbExclamation, "PATHS"
220       End If

End Function

'change Paths
'parameters - ePath - path
'           - sNewValue - new path

Public Sub Let_Paths(ByVal ePath As PATHS_, sNewValue As String)
        
10        Select Case ePath
              Case PATHS_.Midl
20                sPaths_Midl = sNewValue
30            Case PATHS_.ml
40                sPaths_ML = sNewValue
50            Case PATHS_.TextEditor
60                sPaths_TextEditor = sNewValue
70            Case PATHS_.CCompiler
80                sPaths_CCompiler = sNewValue
90            Case PATHS_.INCFiles_Directory
100               sPaths_INCFiles = sNewValue
110           Case PATHS_.LIBFiles_Directory
120               sPaths_LIBFiles = sNewValue
130       End Select
          
End Sub

'-------------
'--- DEBUG ---
'-------------

'get Debug flag
'parameter - eDebug_ - debug flag
'return - TRUE/FALSE

Public Function Get_Debug(ByVal eDebug As DEBUG_) As Boolean

10        Select Case eDebug
              Case DEBUG_.DeleteASM
20                Get_Debug = bDebug_DeleteASM
30            Case DEBUG_.DeleteDebugLogBeforeCompiling
40                Get_Debug = bDebug_DelDebugLogBeforeCompiling
50            Case DEBUG_.DeleteLST
60                Get_Debug = bDebug_DeleteLST
70            Case DEBUG_.EnableOutPutToDebugLog
80                Get_Debug = bDebug_EnableOutPutToDebugLog
90            Case DEBUG_.OutPutAssemblerMessagesToLog
100               Get_Debug = bDebug_OutPutAssemblerMessToLog
110           Case DEBUG_.OutPutMapFiles
120               Get_Debug = bDebug_OutPutMapFiles
130       End Select
              
End Function


'change Debug flag
'parameters - eDebug - flag
'           - bNewValue - new value

Public Sub Let_Debug(ByVal eDebug As DEBUG_, bNewValue As Boolean)

10        Select Case eDebug
              Case DEBUG_.DeleteASM
20                bDebug_DeleteASM = bNewValue
30            Case DEBUG_.DeleteDebugLogBeforeCompiling
40                bDebug_DelDebugLogBeforeCompiling = bNewValue
50            Case DEBUG_.DeleteLST
60                bDebug_DeleteLST = bNewValue
70            Case DEBUG_.EnableOutPutToDebugLog
80                bDebug_EnableOutPutToDebugLog = bNewValue
90            Case DEBUG_.OutPutAssemblerMessagesToLog
100               bDebug_OutPutAssemblerMessToLog = bNewValue
110           Case DEBUG_.OutPutMapFiles
120               bDebug_OutPutMapFiles = bNewValue
130       End Select
              
End Sub

'---------------
'--- COMPILE ---
'---------------

'get Compile flag
'parameter - eCompile - compile flag
'return - TRUE/FALSE

Public Function Get_Compile(ByVal eCompile As COMPILE_) As Boolean

10        Select Case eCompile
              Case COMPILE_.ModifyCmdLine
20                Get_Compile = bCompile_ModifyCmdLine
30            Case COMPILE_.PauseBeforeAssembly
40                Get_Compile = bCompile_PauseBeforeAsm
50            Case COMPILE_.PauseBeforeLinking
60                Get_Compile = bCompile_PauseBeforeLink
70            Case COMPILE_.SkipLinking
80                Get_Compile = bCompile_SkipLinking
90        End Select
          
End Function

'change Compile flag
'parameters - eCompile  - flag
'           - bNewValue - new flag

Public Sub Let_Compile(ByVal eCompile As COMPILE_, bNewValue As Boolean)

10        Select Case eCompile
              Case COMPILE_.ModifyCmdLine
20                bCompile_ModifyCmdLine = bNewValue
30            Case COMPILE_.PauseBeforeAssembly
40                bCompile_PauseBeforeAsm = bNewValue
50            Case COMPILE_.PauseBeforeLinking
60                bCompile_PauseBeforeLink = bNewValue
70            Case COMPILE_.SkipLinking
80                bCompile_SkipLinking = bNewValue
90        End Select
          
End Sub

'-----------
'--- ASM ---
'-----------

'get ASM settings
'parameter - eASM - asm setting
'return -  string/true/false

Public Function Get_ASM(ByVal eASM As ASM_) As Boolean

10        Select Case eASM
              Case ASM_.CompileASMCode
20                Get_ASM = bASM_CompileASMCode
30            Case ASM_.FixASMListings
40                Get_ASM = bASM_FixASMListings
50        End Select
          
End Function

'change ASM flags
'parameters - eASM - flags
'           - sNewValue - new setting

Public Sub Let_ASM(ByVal eASM As ASM_, bNewValue As Boolean)

10        Select Case eASM
              Case ASM_.CompileASMCode
20                bASM_CompileASMCode = bNewValue
30            Case ASM_.FixASMListings
40                bASM_FixASMListings = bNewValue
50        End Select
          
End Sub

'---------
'--- C ---
'---------

'get C settings
'parameter - eC - C setting
'return -  string/true/false

Public Function Get_C(ByVal eC As C_) As Boolean

10        Select Case eC
              Case C_.CompileCCode
20                Get_C = bC_CompileCCode
30        End Select
          
End Function

'change Code coloring flags
'parameters - eC - flags
'           - sNewValue - new value

Public Sub Let_C(ByVal eC As C_, bNewValue As Boolean)

10        Select Case eC
              Case C_.CompileCCode
20                bC_CompileCCode = bNewValue
30        End Select
          
End Sub

Public Function SaveSettingsToVariables(Optional bLoadForm As Boolean = False)
          
10        If bLoadForm = True Then
20            Load frmIn
30            LoadSettings GLOBAL_, frmIn.pctSettings
40            LoadSettings LOCAL_, frmIn.pctSettings
50        End If
          
60        With frmIn
          
              'Let_General HideErrorDialogs, .set_General_HideErrorDialogs
              'Let_General HookCompiler, .set_General_HookCompiler
              'Let_General ListingsForAllModules, .set_General_ListForAllMod
              'Let_General PopUpExportsWindow, .set_General_PopUpWindow
              'Let_General SaveObjFiles, .set_General_SaveOBJ
              
70            Let_Paths CCompiler, .set_Paths_txtCCompiler.Text
80            Let_Paths INCFiles_Directory, .set_Paths_txtIncFiles.Text
90            Let_Paths LIBFiles_Directory, .set_Paths_txtLibFiles.Text
100           Let_Paths ml, .set_Paths_txtMasm.Text
          
              'Let_Debug DeleteASM, .set_Debug_DeleteASM
              'Let_Debug DeleteDebugLogBeforeCompiling, .set_Debug_DelDebBefCom
              'Let_Debug DeleteLST, .set_Debug_DeleteLST
              'Let_Debug EnableOutPutToDebugLog, .set_Debug_OutDebLog
              'Let_Debug OutPutAssemblerMessagesToLog, .set_Debug_OutAsmToLog
              'Let_Debug OutPutMapFiles, .set_Debug_OutMapFiles
              
110           Let_Compile ModifyCmdLine, .set_Compile_ModifyCmdLine.Value
120           Let_Compile PauseBeforeAssembly, .set_Compile_PauseAsm.Value
130           Let_Compile PauseBeforeLinking, .set_Compile_PauseLink.Value
140           Let_Compile SkipLinking, .set_Compile_SkipLinking.Value
          
150           Let_ASM CompileASMCode, .set_ASM_CompileAsmCode.Value
160           Let_ASM FixASMListings, .set_ASM_FixAsmListings.Value
              
170           Let_C CompileCCode, .set_C_CompileCCode.Value
          
180       End With
          
190       If bLoadForm = True Then
200           Unload frmIn
210       End If
          
End Function

Public Function SaveOtherSettings(eScope As SET_SCOPE) As Boolean
          'other controls that must be saved
          'this function is called from SaveSettings (modSaveLoadSettings)
          
10        SaveOtherSettings = True
          
End Function

Public Function LoadOtherSettings(eScope As SET_SCOPE) As Boolean
          'other controls that must be inited
          'this function is called from LoadSettings (modSaveLoadSettings)
          
10        LoadOtherSettings = True
          
End Function

Public Sub SetOtherDefaultSettings(eScope As SET_SCOPE)
    'other controls that must be set to default
    'this function is called from SetDefaultSetings (modSaveLoadSettings)
       
End Sub
