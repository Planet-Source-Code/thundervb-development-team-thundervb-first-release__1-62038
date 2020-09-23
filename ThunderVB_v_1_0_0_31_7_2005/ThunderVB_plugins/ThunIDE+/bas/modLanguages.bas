Attribute VB_Name = "modLanguages"
Option Explicit

'update all resouce indexes..
Public Sub lang_UpdateIds(lang As tvb_Languages)

       'pure nothing for now
          
End Sub

'Update all the gui for a new language;called after UpdateIds.
Public Sub lang_UpdateGui(lang As tvb_Languages)
              
       frmIn.lang_UpdateGui lang

End Sub
