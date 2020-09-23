VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThunderVB_pl_int_v1_0"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'PlugIn Interface

'Basic Interface...
Public Function InitSetLang(lang As tvb_Languages) 'called when teh dll is loaded onto memory
                                                   'no plugin releated things must be done here
                                                   'this is just to set the lagnuage of the decs ect..
End Function

Public Function GetName() As String         'Name of the addin

End Function

Public Function GetVersion() As String      'Version of the addin , string format

End Function

Public Function GetVersionNum() As Single() 'Version in x.x format , x is a float [so this can be extended to x.x.x.x]

End Function

Public Function GetDesciption() As String   'A short Desciption for this addin

End Function

Public Function GetDesciptionFull() As String 'A Desciption for this addin

End Function

Public Function GetType() As pl_type         'Addin Type

End Function

Public Function GetSpeed() As pl_Speed_Req   'How much power consuming this addin is

End Function

Public Function GetID() As ThunVB_PluginId     'Addin ID , this should be unique [email to the athors to get one]

End Function

Public Function GetWindowsVersion() As String 'windows version needed , for the user ..

End Function

'Loading/Unloading/GUI events

'On plugin load
Public Sub OnStartUp()

End Sub

'on plugin unload
Public Sub OnTermination()

End Sub

'Project related events..
Public Sub OnProjectAdded(project As VBProject)

End Sub

Public Sub OnProjectActivated(project As VBProject)

End Sub

Public Sub OnProjectRenamed(project As VBProject, OldName As String)

End Sub

Public Sub OnProjectRemoved(project As VBProject)

End Sub

'For GUI
'
'For The settings form...
'Load the gui , called when the settings window is opened
Public Sub OnGuiLoad()

End Sub

'called when teh apply/ok button is pressed
'can be recieved many times
Public Sub ApplySettings()

End Sub

'Called when the cancel button is pressed
Public Sub CancelSettings()

End Sub

'Called when the Settings form is closed..
'it is allways the last event sent..
Public Sub OnGuiUnLoad()

End Sub

'Show the config
'retuns : the hWnd of the config control
Public Function ShowConfig() As Long

End Function

'Hide the config
Public Sub HideConfig()

End Sub

'For the Credits control ..
'Retuns : the hWnd of the Credits control
Public Function ShowCredits() As Long

End Function

'Hide the credits control
Public Sub HideCredits()

End Sub

'function to send a notification to this addin
'For menu/button messanges and
'for plugin intercomunication..
'All messangers after tvbMsg.tvbm_User
'are not reserved..
Public Function SendMessange(msg As tvbMsg, Optional param1 As Long, Optional param2 As Long, Optional param3 As Variant, Optional param4 As Variant, Optional param5 As Variant, Optional param6 As Variant) As Variant

End Function


'sets the loging object..
Public Sub SetLogger(newlogger As ILogger)

End Sub
