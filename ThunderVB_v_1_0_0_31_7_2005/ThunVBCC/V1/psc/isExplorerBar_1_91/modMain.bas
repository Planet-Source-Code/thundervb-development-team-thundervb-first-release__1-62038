Attribute VB_Name = "modMain"
'This Module is part of the Theme Explorer from VBAccelerator.com
'   I use It to allow themed controls in this App demo.
'   But If you use only a MANIFEST file, It will do the job.
'   (You know Steve Is TOO explicit in some things).
Option Explicit

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
   ' we need to call InitCommonControls before we
   ' can use XP visual styles.  Here I'm using
   ' InitCommonControlsEx, which is the extended
   ' version provided in v4.72 upwards (you need
   ' v6.00 or higher to get XP styles)
   On Error Resume Next
   ' this will fail if Comctl not available
   '  - unlikely now though!
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   
   ' now start the application
   On Error GoTo 0
   frmTest.Show
   
End Sub

