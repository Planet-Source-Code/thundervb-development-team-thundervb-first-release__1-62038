VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ITheme: This is the base interface for new themes
'        All the themes MUST implement this interface.
Option Explicit


' Initialization function, MUST be called before any other funciton
' is called Assignes the send control to a local variable
Public Sub SetControl(oXTab As XTab)

End Sub

' Erases the Previous Background and draws a new one
' (does not draw the tabs) , Called from Paint Event of the user
' control
Public Sub DrawBackground()

End Sub

' This is called from the DrawBackgroung function if the tab style
' is Property Page
Public Sub DrawBackgroundPropertyPages()

End Sub

' This is called from the DrawBackgroung function if the tab style
' is Tabbed Dialog
Public Sub DrawBackgroundTabbedDialog()

End Sub

' This is called when the user clicks a tab. Instead of drawing the
' complete user control we just draw the top strip
' (to improve performance and prevent flicker)
Public Sub DrawOnActiveTabChange()

End Sub

'Erases the previous tabs and draws each one by one
Public Sub DrawTabs()

End Sub

' This is called from the DrawTabs function if the tab style
' is Property Page
Public Sub DrawTabsPropertyPages()

End Sub

' This is called from the DrawTabs function if the tab style
' is Tabbed Dialog
Public Sub DrawTabsTabbedDialog()

End Sub

'Called when the user presses mouse on the clickable area of the user control
Public Sub MouseDownHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)

End Sub

'Called when the user releases mouse on the clickable area of the user control
Public Sub MouseUpHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)

End Sub

'Called as the user moves the mouse on the control
Public Sub MouseMoveHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)

End Sub

'Called when User control Get's or looses the focus
'This way we are able to prevent complete control redrawing
'only the focus rect is drawn or erased. Thus preventing flicker
Public Sub ShowHideFocus()

End Sub

'Called from ShowHideFocus for more speific processing
Public Sub ShowHideFocusPropertyPages()

End Sub

'Called from ShowHideFocus for more speific processing
Public Sub ShowHideFocusTabbedDialog()

End Sub

'Called when the user control's timer event occurs
Public Sub TimerEvent()

End Sub

'Function used to replace all the colors with the system colors
Public Sub ResetColorsToDefault()

End Sub

