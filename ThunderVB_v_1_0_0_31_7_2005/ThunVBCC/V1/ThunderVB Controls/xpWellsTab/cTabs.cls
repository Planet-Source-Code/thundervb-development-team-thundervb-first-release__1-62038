VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TabCaption           As String
Public TabLeft              As Long
Public TabTop               As Long
Public TabWidth             As Long
Public TabHeight            As Long
Public TabPicture           As StdPicture

Public Property Get TabIcon() As StdPicture
    '<EhHeader>
    On Error GoTo TabIcon_Err
    '</EhHeader>
    Set TabIcon = TabPicture
    '<EhFooter>
    Exit Property

TabIcon_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTabs.TabIcon " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set TabIcon(ByVal NewTabPicture As StdPicture)
    '<EhHeader>
    On Error GoTo TabIcon_Err
    '</EhHeader>
    Set TabPicture = NewTabPicture
    '<EhFooter>
    Exit Property

TabIcon_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTabs.TabIcon " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Private Sub Class_Terminate()
    '<EhHeader>
    On Error GoTo Class_Terminate_Err
    '</EhHeader>
    Set TabPicture = Nothing
    '<EhFooter>
    Exit Sub

Class_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.cTabs.Class_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
