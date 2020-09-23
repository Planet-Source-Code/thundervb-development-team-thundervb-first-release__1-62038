Attribute VB_Name = "modSharedPriv"
Option Explicit
'For things that must be on all the project and they must be
'a copy of them on each ..
'Like ide/exe detection

Public Function DebugMode() As Boolean

    DebugMode = App.LogMode = 0 'if on debug mode , return true , else false

End Function
