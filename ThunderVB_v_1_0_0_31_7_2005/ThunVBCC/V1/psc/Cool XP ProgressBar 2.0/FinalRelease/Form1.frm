VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Mario Flores Cool Xp ProgressBar"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Use Metallic XP Scrolling"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Picture Style Scrolling"
      Height          =   495
      Left            =   6120
      TabIndex        =   24
      Top             =   7680
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":000C
      Left            =   240
      List            =   "Form1.frx":0025
      TabIndex        =   23
      Text            =   "HS_DIAGCROSS"
      Top             =   6960
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Custom Brush Scrolling"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use JavT Scrolling"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Media Player Style Scrolling"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   7680
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Pastel Color Scrolling"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Office XP Scrolling"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More"
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Search Style Scrolling"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Smooth Scrolling"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Use Standard Scrolling"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Value           =   -1  'True
      Width           =   1935
   End
   Begin StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   8400
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   4604
            Text            =   "Mario Flores G Cool Xp ProgressBar"
            TextSave        =   "Mario Flores G Cool Xp ProgressBar"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "9/14/2004"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   120
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Use Percent Text"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin Project1.XP_ProgressBar XP_ProgressBar1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MT Extra"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   12937777
      Value           =   1
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarx 
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
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
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarSilver 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
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
      Value           =   1
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarOlive 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
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
   End
   Begin Project1.XP_ProgressBar XP_ProgressBarBlue 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   480
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
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   3000
      MouseIcon       =   "Form1.frx":0084
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":01D6
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   3
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   4200
      TabIndex        =   12
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      Top             =   7200
      Width           =   405
   End
   Begin VB.Shape ShapeColor 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3720
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Blue XP"
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Olive XP"
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Example Color Silver XP"
      Height          =   195
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Picker Color Example"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   2400
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Dim MTime   As Long
Dim MDown   As Boolean
Dim LaBrush As BrushStyle

Const XPBlue_ProgressBar = &H2BD228
Const XPOlive_ProgressBar = &H4A86E4
Const XPSilver_ProgressBar = &H76AE83



Private Sub Check2_Click()
XP_ProgressBar1.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarBlue.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarOlive.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarSilver.ShowText = IIf(Check2.Value = 0, False, True)
XP_ProgressBarx.ShowText = IIf(Check2.Value = 0, False, True)
End Sub

Private Sub Combo1_Click()
LaBrush = Combo1.ListIndex
If Option1(7).Value = True Then Option1_Click 7
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub


Private Sub Command2_Click()
Form3.Show
End Sub



Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Form_Initialize()
'InitCommonControls
End Sub


Private Sub Form_Load()

Combo1.ListIndex = 4

XP_ProgressBar1.Max = 100
'XP_ProgressBar1.Min = 1

XP_ProgressBarBlue.Max = 100
XP_ProgressBarBlue.Min = 1

XP_ProgressBarOlive.Max = 100
XP_ProgressBarOlive.Min = 1

XP_ProgressBarSilver.Max = 100
XP_ProgressBarSilver.Min = 1

XP_ProgressBarx.Max = 100
XP_ProgressBarx.Min = 1



XP_ProgressBarBlue.Color = XPBlue_ProgressBar
XP_ProgressBarOlive.Color = XPOlive_ProgressBar
XP_ProgressBarSilver.Color = XPSilver_ProgressBar

XP_ProgressBar1.Color = vbHighlight
XP_ProgressBarx.Color = vbHighlight


ShapeColor.BackColor = vbHighlight

ShowProgressInStatusBar XP_ProgressBarx, StatusBar1, 3



End Sub



Private Sub Option1_Click(Index As Integer)

If Index = 2 Then
    Timer1.Interval = 20
Else                         'Time Interval yust to Show Search Style Demo.
    Timer1.Interval = 100
End If

If Index = 7 Then
XP_ProgressBar1.BrushStyle = LaBrush
XP_ProgressBarBlue.BrushStyle = LaBrush
XP_ProgressBarOlive.BrushStyle = LaBrush
XP_ProgressBarSilver.BrushStyle = LaBrush
XP_ProgressBarx.BrushStyle = LaBrush
End If

XP_ProgressBar1.Scrolling = Index
XP_ProgressBarBlue.Scrolling = Index
XP_ProgressBarOlive.Scrolling = Index
XP_ProgressBarSilver.Scrolling = Index
XP_ProgressBarx.Scrolling = Index

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
XP_ProgressBar1.Color = ShapeColor.BackColor
XP_ProgressBarx.Color = ShapeColor.BackColor

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

If MTime > XP_ProgressBar1.Max Then
    MTime = XP_ProgressBar1.Min
End If


XP_ProgressBar1.Value = MTime
XP_ProgressBarBlue.Value = MTime
XP_ProgressBarOlive.Value = MTime
XP_ProgressBarSilver.Value = MTime
XP_ProgressBarx.Value = MTime

LblValue = Round((100 * XP_ProgressBar1.Value) / XP_ProgressBar1.Max) & " %"


End Sub




Private Sub XP_ProgressBar2_GotFocus()

End Sub
