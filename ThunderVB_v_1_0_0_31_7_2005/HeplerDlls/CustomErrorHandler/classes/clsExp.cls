VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EhModes
    Err_not_expected = 0
    Err_expected = 1
    Err_expected_AutoRestore = 2
End Enum

Public Static Sub SetexeptH()

    ErrorHanlder_vb.SetexeptH

End Sub


Public Static Sub FreeexeptH()

    KillHookList th

End Sub

Public Static Sub SetEhMode(newmode As EhModes)
    
    If IsAsmOn() And inited = True Then
        If EhModeI > UBound(ErrHmodes) Then
            ReDim Preserve ErrHmodes(EhModeI + 1)
        End If
    
        ErrHmodes(EhModeI) = ErrHmode
        ErrHmode = newmode
        
        EhModeI = EhModeI + 1
    End If
    
End Sub

Public Static Function GetEhMode() As EhModes

   GetEhMode = ErrHmode

End Function

Public Static Sub RestoreEhMode()

    If IsAsmOn() And inited = True Then
        EhModeI = EhModeI - 1
        If EhModeI < 0 Then
            ErrHmode = Err_not_expected
            EhModeI = 0
        Else
            ErrHmode = ErrHmodes(EhModeI)
        End If
    End If
    
End Sub


Public Static Sub SetLogger(new_Logger As ILogger)
    
    modLog.SetLogger new_Logger

End Sub
