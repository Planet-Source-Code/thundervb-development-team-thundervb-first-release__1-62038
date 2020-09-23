VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cool XP ProgressBar 2.0"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   1560
      MouseIcon       =   "Form3.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Form3.frx":0152
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   252
      TabIndex        =   3
      Top             =   3360
      Width           =   3780
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "&Start"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Use Percent Text"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "&Exit"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   5640
      Width           =   1695
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   6
      ShowText        =   -1  'True
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   6
      ShowText        =   -1  'True
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   10
      Top             =   2160
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   6
      ShowText        =   -1  'True
   End
   Begin Project1.XP_ProgressBar XP 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Scrolling       =   6
      ShowText        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderColor     =   &H00C0C000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1420
      Left            =   25
      Top             =   4920
      Width           =   6390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Media Player ProgressBar Example"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PickColor"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   4440
      Width           =   405
   End
   Begin VB.Shape ShapeColor 
      BackColor       =   &H00C56A31&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2640
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   2640
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   6375
      Left            =   0
      Top             =   0
      Width           =   6450
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Dim MTime  As Long
Dim MDown  As Boolean

Private Sub Check2_Click()
XP(0).ShowText = IIf(Check2.Value = 0, False, True)
XP(1).ShowText = IIf(Check2.Value = 0, False, True)
XP(2).ShowText = IIf(Check2.Value = 0, False, True)
XP(3).ShowText = IIf(Check2.Value = 0, False, True)
End Sub

Private Sub cmdStart_Click()

If Timer1.Enabled = False Then
  Timer1.Enabled = True
  cmdStart.Caption = "&Stop"
Else
  Timer1.Enabled = False
  cmdStart.Caption = "&Start"

End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
MTime = 0
End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim R      As Integer
Dim G      As Integer
Dim B      As Integer
Dim PixCol As Long

PixCol = GetPixel(Picture1.hdc, X, Y)

'Convert to RGB
R = PixCol Mod 256
B = Int(PixCol / 65536)
G = (PixCol - (B * 65536) - R) / 256

If R < 0 Then R = 0
If G < 0 Then G = 0
If B < 0 Then B = 0


ShapeColor.BackColor = RGB(R, G, B)
XP(0).Color = ShapeColor.BackColor
XP(1).Color = ShapeColor.BackColor
XP(2).Color = ShapeColor.BackColor
XP(3).Color = ShapeColor.BackColor

MDown = True

End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MDown Then Picture1_MouseDown Button, Shift, X, Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MDown = False

End Sub
Private Sub Timer1_Timer()

MTime = MTime + 1

If MTime > XP(0).Max Then
    MTime = XP(0).Min
End If

XP(0).Value = MTime
XP(1).Value = MTime
XP(2).Value = MTime
XP(3).Value = MTime

LblValue = (100 * XP(0).Value) / XP(0).Max & " %"


End Sub
