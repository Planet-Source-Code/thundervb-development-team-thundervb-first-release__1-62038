Attribute VB_Name = "modThunIDE"
Option Explicit

Public oMe As plugin
Global rfile As New cResFile

Public Const PLUGIN_NAME As String = "ThunderIDE+"
Public Const MSG_TITLE As String = PLUGIN_NAME

Public Const PLUGIN_NAMEs As String = "ThunIDE"
Public Const MSG_TITLEs As String = PLUGIN_NAMEs

Public Const APP_NAME As String = PLUGIN_NAME
Public Const APP_NAMEs As String = PLUGIN_NAMEs

Public Function GetThunIdeVersion() As String
    GetThunIdeVersion = PLUGIN_NAME & " " & App.Major & "." & App.Minor & "." & App.Revision
End Function
