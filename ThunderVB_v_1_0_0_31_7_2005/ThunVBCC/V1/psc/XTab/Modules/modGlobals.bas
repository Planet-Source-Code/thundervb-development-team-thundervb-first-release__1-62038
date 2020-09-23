Attribute VB_Name = "modGlobals"
'modGlobals: Hosts Global declarations

Option Explicit



'Made it public to passing and returning of this type. :(
Public Type TabInfo
  Caption As String                  'Caption for the tab
  ClickableRect As RECT              'Coordinates of the clickable rectangle
  ContainedControlsDetails As Collection
  Enabled As Boolean
  AccessKey As Integer
  TabPicture As StdPicture
End Type
