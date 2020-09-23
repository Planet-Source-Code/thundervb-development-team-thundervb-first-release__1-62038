VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#24.1#0"; "ThunVBCC_v1.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin ThunVBCC_v1.HzxYOption HzxYOption1 
      Height          =   225
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'we directly implement the interface
Implements ISubclass_Callbacks

Dim WithEvents subclasser_1 As SubclassEventImpl
Attribute subclasser_1.VB_VarHelpID = -1


Private Sub Form_Load()
    'init to sumbclass , using the implementation method
    'We are notified from teh callbacks
    SubClasshWnd Me.hwnd, Me, wproc_notify
    
    Set subclasser_1 = New SubclassEventImpl
    'Here we use the SubclassEventImpl witch is an implementaion of the
    'interface and converts the callbacks to events
    subclasser_1.SubClass Me.Frame1.hwnd, wproc_notify
    
    'this is also correct but the SubclassEventImpl has some propertys
    '(like isSubclassed)that work only if you init
    'the subclassing using the above way
    '
    'SubClasshWnd Me.Frame1.hwnd, subclasser_1, wproc_notify
    
End Sub


Private Sub subclasser_1_AftWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal CalledWProc As Boolean, ByVal pOriginalPRoc As Long)
    If uMsg = WM_LBUTTONUP Then MsgBox "Up After,Frame1"
End Sub

Private Sub Command1_Click()
    Dim t As Form1
    Set t = New Form1
    t.Show
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox "Up VB"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnSubClasshWnd Me.hwnd
    UnSubClasshWnd Me.Frame1.hwnd
End Sub

Private Sub ISubclass_Callbacks_AftWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal CalledWProc As Boolean, ByVal pOriginalPRoc As Long)
    If uMsg = WM_LBUTTONUP Then MsgBox "Up After"
End Sub


Private Sub ISubclass_Callbacks_BefWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, CallWProc As Boolean, CallAftProc As Boolean, ByVal pOriginalPRoc As Long)
    If uMsg = WM_LBUTTONUP Then MsgBox "Up Before"
End Sub


Private Sub ISubclass_Callbacks_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, ByVal pOriginalPRoc As Long)

End Sub


Private Sub subclasser_1_BefWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ReturnValue As Long, CallWProc As Boolean, CallAftProc As Boolean, ByVal pOriginalPRoc As Long)
    
    If uMsg = WM_LBUTTONUP Then MsgBox "Up Before,Frame1"
    If uMsg = WM_PAINT Then Debug.Print "wm_paint"
    If uMsg = WM_SETFOCUS Then Debug.Print "WM_SETFOCUS"
    If uMsg = WM_NCPAINT Then CallWProc = False

End Sub

Private Sub Timer1_Timer()
    Me.Caption = CountSubClassedWindows()
End Sub

