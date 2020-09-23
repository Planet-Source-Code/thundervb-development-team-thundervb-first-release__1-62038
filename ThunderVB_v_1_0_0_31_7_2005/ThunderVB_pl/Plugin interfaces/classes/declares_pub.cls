VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "declares_pub"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Enum pl_type

    'Type : for hook based
    ThunVB_DLLHook = 1 ' Dll hook
    ThunVB_CPHook = 2  ' CreateProcess hook
    
    'Type : for SubClassing Based
    ThunVB_SC_Light = 8 ' Sublcassing - less that 4 messanges
    ThunVB_SC_Med = 16 ' Sublcassing - less that 8 messanges
    ThunVB_SC_Heavy = 32 ' Sublcassing - more that 8 messanges
    
    'Type : for "simple" [no crazy shit around them like hooking] things
    ThunVB_plugin_CodeTool = 64  ' Code Enhasement->Code generator, Code checking ect
    ThunVB_plugin_MiscTool = 128 ' Tool->Misc things

End Enum

Public Enum pl_Speed_Req

    idle = 0  ' none   , eg asm function code prototype generator
    low = 1   ' low    , rare and not so power consuming
    Med = 2   ' Medium , well , consumes some power but it can be freerly enablen on all pc's..
    Hight = 4 ' Hight  , Consumes a considerable amount of power , but still usable on a p3 at 700 mhrz..
    Cazy = 8  ' too Hight  , Consumes much much power p4 > 1.5 giga to be enabled..
    
End Enum

Public Enum tvb_CP_CallOrder

    cpo_DoNotRun = 0
    cpo_First = 1
    cpo_1 = 2
    cpo_2 = 3
    cpo_3 = 4
    cpo_4 = 5
    cpo_Normal = 6
    cpo_6 = 7
    cpo_7 = 8
    cpo_8 = 9
    cpo_9 = 10
    cpo_last = 11
    
End Enum

Public Enum tvbMsg

    tvbm_ButtonClick = 0            'send when a button is clicked
    tvbm_ChangeLanguage_code = 1    'send when the language is changed
    tvbm_ChangeLanguage_gui = 2     'send when the language is changed , and all the plugins have got the tvbm_ChangeLanguage_Code message
    tvbm_User = 512
    
End Enum

Public Enum ThunVB_PluginId
    
    tvb_test = -1
    tvb_nonregisted = -2
    tvb_ThunderVB = 0
    tvb_ThunDll = 1
    tvb_ThunAsm = 2
    tvb_ThunIDE_p = 3
    
End Enum


Public Type PROCESS_INFORMATION

    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
    
End Type

Public Type SECURITY_ATTRIBUTES

    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
    
End Type

Public Type STARTUPINFO

    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
    
End Type

