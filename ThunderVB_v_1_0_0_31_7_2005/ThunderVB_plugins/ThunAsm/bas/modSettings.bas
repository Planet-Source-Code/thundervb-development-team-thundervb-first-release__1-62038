Attribute VB_Name = "modSettings"
Option Explicit

'module only for settings

'*** General ***
'Tab General

Private bGeneral_PopUpExportWindow As Boolean
Private bGeneral_ListForAllModules As Boolean
Private bGeneral_HookCompiler As Boolean
Private bGeneral_SaveOBJ As Boolean
Private bGeneral_HideErrorDialogs As Boolean
Private bGeneral_GenerateAsmCHeaders As Boolean

Public Enum GENERAL_
    PopUpExportsWindow
    ListingsForAllModules
    HookCompiler
    SaveObjFiles
    HideErrorDialogs
    GenerateAsmCHeaders
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
         
       Select Case eGeneral
           Case GENERAL_.HideErrorDialogs
               Get_General = bGeneral_HideErrorDialogs
           Case GENERAL_.HookCompiler
               Get_General = bGeneral_HookCompiler
           Case GENERAL_.ListingsForAllModules
               Get_General = bGeneral_ListForAllModules
           Case GENERAL_.PopUpExportsWindow
               Get_General = bGeneral_PopUpExportWindow
           Case GENERAL_.SaveObjFiles
              Get_General = bGeneral_SaveOBJ
            Case GENERAL_.GenerateAsmCHeaders
                Get_General = bGeneral_GenerateAsmCHeaders
      End Select

End Function

'change General flag
'parameters - eGeneral - general flag to change
'           - bNewValue - new flag

Public Sub Let_General(ByVal eGeneral As GENERAL_, bNewValue As Boolean)
         
       Select Case eGeneral
           Case GENERAL_.HideErrorDialogs
               bGeneral_HideErrorDialogs = bNewValue
           Case GENERAL_.HookCompiler
               bGeneral_HookCompiler = bNewValue
           Case GENERAL_.ListingsForAllModules
               bGeneral_ListForAllModules = bNewValue
           Case GENERAL_.PopUpExportsWindow
               bGeneral_PopUpExportWindow = bNewValue
           Case GENERAL_.SaveObjFiles
              bGeneral_SaveOBJ = bNewValue
            Case GENERAL_.GenerateAsmCHeaders
                bGeneral_GenerateAsmCHeaders = bNewValue
      End Select

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
         
       Select Case ePath
           Case PATHS_.Midl
               Get_Paths = sPaths_Midl
               sText = "Path to MIDL (midl.exe) is not set." & vbCrLf & "Setting/Paths"
           Case PATHS_.ml
               Get_Paths = sPaths_ML
               sText = "Path to ML (ml.exe) is not set." & vbCrLf & "Setting/Paths"
           Case PATHS_.TextEditor
               Get_Paths = sPaths_TextEditor
               sText = "Path to your Text-Editor is not set." & vbCrLf & "Setting/Paths"
          Case PATHS_.INCFiles_Directory
              Get_Paths = sPaths_INCFiles
              sText = "Path to .INC files is not set." & vbCrLf & "Setting/Paths"
          Case PATHS_.LIBFiles_Directory
              Get_Paths = sPaths_LIBFiles
              sText = "Path to .LIB files is not set." & vbCrLf & "Setting/Paths"
          Case PATHS_.CCompiler
              Get_Paths = sPaths_CCompiler
              sText = "Path to C compiler is not set." & vbCrLf & "Setting/Paths"
      End Select

       'check path
      If Len(Get_Paths) = 0 And bWarning = True Then
          MsgBoxX sText, vbExclamation, "PATHS"
      End If

End Function

'change Paths
'parameters - ePath - path
'           - sNewValue - new path

Public Sub Let_Paths(ByVal ePath As PATHS_, sNewValue As String)
        
       Select Case ePath
           Case PATHS_.Midl
               sPaths_Midl = sNewValue
           Case PATHS_.ml
               sPaths_ML = sNewValue
           Case PATHS_.TextEditor
               sPaths_TextEditor = sNewValue
           Case PATHS_.CCompiler
               sPaths_CCompiler = sNewValue
           Case PATHS_.INCFiles_Directory
              sPaths_INCFiles = sNewValue
          Case PATHS_.LIBFiles_Directory
              sPaths_LIBFiles = sNewValue
      End Select
          
End Sub

'-------------
'--- DEBUG ---
'-------------

'get Debug flag
'parameter - eDebug_ - debug flag
'return - TRUE/FALSE

Public Function Get_Debug(ByVal eDebug As DEBUG_) As Boolean

       Select Case eDebug
           Case DEBUG_.DeleteASM
               Get_Debug = bDebug_DeleteASM
           Case DEBUG_.DeleteDebugLogBeforeCompiling
               Get_Debug = bDebug_DelDebugLogBeforeCompiling
           Case DEBUG_.DeleteLST
               Get_Debug = bDebug_DeleteLST
           Case DEBUG_.EnableOutPutToDebugLog
               Get_Debug = bDebug_EnableOutPutToDebugLog
           Case DEBUG_.OutPutAssemblerMessagesToLog
              Get_Debug = bDebug_OutPutAssemblerMessToLog
          Case DEBUG_.OutPutMapFiles
              Get_Debug = bDebug_OutPutMapFiles
      End Select
              
End Function


'change Debug flag
'parameters - eDebug - flag
'           - bNewValue - new value

Public Sub Let_Debug(ByVal eDebug As DEBUG_, bNewValue As Boolean)

       Select Case eDebug
           Case DEBUG_.DeleteASM
               bDebug_DeleteASM = bNewValue
           Case DEBUG_.DeleteDebugLogBeforeCompiling
               bDebug_DelDebugLogBeforeCompiling = bNewValue
           Case DEBUG_.DeleteLST
               bDebug_DeleteLST = bNewValue
           Case DEBUG_.EnableOutPutToDebugLog
               bDebug_EnableOutPutToDebugLog = bNewValue
           Case DEBUG_.OutPutAssemblerMessagesToLog
              bDebug_OutPutAssemblerMessToLog = bNewValue
          Case DEBUG_.OutPutMapFiles
              bDebug_OutPutMapFiles = bNewValue
      End Select
              
End Sub

'---------------
'--- COMPILE ---
'---------------

'get Compile flag
'parameter - eCompile - compile flag
'return - TRUE/FALSE

Public Function Get_Compile(ByVal eCompile As COMPILE_) As Boolean

       Select Case eCompile
           Case COMPILE_.ModifyCmdLine
               Get_Compile = bCompile_ModifyCmdLine
           Case COMPILE_.PauseBeforeAssembly
               Get_Compile = bCompile_PauseBeforeAsm
           Case COMPILE_.PauseBeforeLinking
               Get_Compile = bCompile_PauseBeforeLink
           Case COMPILE_.SkipLinking
               Get_Compile = bCompile_SkipLinking
       End Select
          
End Function

'change Compile flag
'parameters - eCompile  - flag
'           - bNewValue - new flag

Public Sub Let_Compile(ByVal eCompile As COMPILE_, bNewValue As Boolean)

       Select Case eCompile
           Case COMPILE_.ModifyCmdLine
               bCompile_ModifyCmdLine = bNewValue
           Case COMPILE_.PauseBeforeAssembly
               bCompile_PauseBeforeAsm = bNewValue
           Case COMPILE_.PauseBeforeLinking
               bCompile_PauseBeforeLink = bNewValue
           Case COMPILE_.SkipLinking
               bCompile_SkipLinking = bNewValue
       End Select
          
End Sub

'-----------
'--- ASM ---
'-----------

'get ASM settings
'parameter - eASM - asm setting
'return -  string/true/false

Public Function Get_ASM(ByVal eASM As ASM_) As Boolean

       Select Case eASM
           Case ASM_.CompileASMCode
               Get_ASM = bASM_CompileASMCode
           Case ASM_.FixASMListings
               Get_ASM = bASM_FixASMListings
       End Select
          
End Function

'change ASM flags
'parameters - eASM - flags
'           - sNewValue - new setting

Public Sub Let_ASM(ByVal eASM As ASM_, bNewValue As Boolean)

       Select Case eASM
           Case ASM_.CompileASMCode
               bASM_CompileASMCode = bNewValue
           Case ASM_.FixASMListings
               bASM_FixASMListings = bNewValue
       End Select
          
End Sub

'---------
'--- C ---
'---------

'get C settings
'parameter - eC - C setting
'return -  string/true/false

Public Function Get_C(ByVal eC As C_) As Boolean

       Select Case eC
           Case C_.CompileCCode
               Get_C = bC_CompileCCode
       End Select
          
End Function

'change Code coloring flags
'parameters - eC - flags
'           - sNewValue - new value

Public Sub Let_C(ByVal eC As C_, bNewValue As Boolean)

       Select Case eC
           Case C_.CompileCCode
               bC_CompileCCode = bNewValue
       End Select
          
End Sub

Public Function SaveSettingsToVariables(Optional bLoadForm As Boolean = False)
          
       If bLoadForm = True Then
           Load frmIn
           LoadSettings GLOBAL_, frmIn.pctSettings
           LoadSettings LOCAL_, frmIn.pctSettings
       End If
          
       With frmIn
          
           'Let_General HideErrorDialogs, .set_General_HideErrorDialogs
           'Let_General HookCompiler, .set_General_HookCompiler
           'Let_General ListingsForAllModules, .set_General_ListForAllMod
           'Let_General PopUpExportsWindow, .set_General_PopUpWindow
           'Let_General SaveObjFiles, .set_General_SaveOBJ
              
           Let_Paths CCompiler, .set_Paths_txtCCompiler.Text
           Let_Paths INCFiles_Directory, .set_Paths_txtIncFiles.Text
           Let_Paths LIBFiles_Directory, .set_Paths_txtLibFiles.Text
           Let_Paths ml, .set_Paths_txtMasm.Text
          
           'Let_Debug DeleteASM, .set_Debug_DeleteASM
           'Let_Debug DeleteDebugLogBeforeCompiling, .set_Debug_DelDebBefCom
           'Let_Debug DeleteLST, .set_Debug_DeleteLST
           'Let_Debug EnableOutPutToDebugLog, .set_Debug_OutDebLog
           'Let_Debug OutPutAssemblerMessagesToLog, .set_Debug_OutAsmToLog
           'Let_Debug OutPutMapFiles, .set_Debug_OutMapFiles
              
          Let_Compile ModifyCmdLine, .set_Compile_ModifyCmdLine.Value
          Let_Compile PauseBeforeAssembly, .set_Compile_PauseAsm.Value
          Let_Compile PauseBeforeLinking, .set_Compile_PauseLink.Value
          Let_Compile SkipLinking, .set_Compile_SkipLinking.Value
          
          Let_ASM CompileASMCode, .set_ASM_CompileAsmCode.Value
          Let_ASM FixASMListings, .set_ASM_FixAsmListings.Value
              
          Let_C CompileCCode, .set_C_CompileCCode.Value
          
          Let_General GenerateAsmCHeaders, .set_Gen_chbGenAsmCHeaders.Value
          
      End With
          
      If bLoadForm = True Then
          Unload frmIn
      End If
          
End Function

Public Function SaveOtherSettings(eScope As SET_SCOPE) As Boolean
       'other controls that must be saved
       'this function is called from SaveSettings (modSaveLoadSettings)
          
       SaveOtherSettings = True
          
End Function

Public Function LoadOtherSettings(eScope As SET_SCOPE) As Boolean
       'other controls that must be inited
       'this function is called from LoadSettings (modSaveLoadSettings)
          
       LoadOtherSettings = True
          
End Function

Public Sub SetOtherDefaultSettings(eScope As SET_SCOPE)
    'other controls that must be set to default
    'this function is called from SetDefaultSetings (modSaveLoadSettings)
       
End Sub
