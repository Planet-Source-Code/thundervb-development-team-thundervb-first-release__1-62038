VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00E0F0F0&
   Caption         =   "Tab Test"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin TabTest.xpWellsTab xpWellsTab1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   5530
      Alignment       =   0
      TabHeight       =   25
      BackColor       =   14741744
      ForeColor       =   -2147483630
      ForeColorActive =   9982008
      ForeColorHot    =   16711680
      FrameColor      =   9915703
      MaskColor       =   16711935
      SelectedTab     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   4
      TabWidth1       =   70
      TabText1        =   "&Brian"
      TabPicture1     =   "frmTest.frx":0000
      TabWidth2       =   90
      TabText2        =   "&Shane"
      TabPicture2     =   "frmTest.frx":0352
      TabWidth3       =   60
      TabText3        =   "&Harry"
      TabPicture3     =   "frmTest.frx":06A4
      TabWidth4       =   50
      TabText4        =   "&Guy"
      TabPicture4     =   "frmTest.frx":09F6
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0F0F0&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   3
         Left            =   840
         ScaleHeight     =   1665
         ScaleWidth      =   3825
         TabIndex        =   4
         Top             =   1200
         Width           =   3855
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Guy`s Tab"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0F0F0&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   2
         Left            =   600
         ScaleHeight     =   1665
         ScaleWidth      =   3825
         TabIndex        =   3
         Top             =   960
         Width           =   3855
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Harry`s Tab"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0F0F0&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   360
         ScaleHeight     =   1665
         ScaleWidth      =   3825
         TabIndex        =   2
         Top             =   720
         Width           =   3855
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Shane`s Tab"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0F0F0&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   0
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   3825
         TabIndex        =   1
         Top             =   480
         Width           =   3855
         Begin VB.CommandButton Command1 
            Caption         =   "Go To Harry`s Tab"
            Height          =   495
            Left            =   3000
            TabIndex        =   5
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Brian`s Tab"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    xpWellsTab1.SelectedTab = 3
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long
    For i = 0 To 3
        pic(i).Visible = False
        pic(i).BorderStyle = 0
        pic(i).Left = 120
        pic(i).Top = 480
        pic(i).Height = 2535
        pic(i).Width = 4695
    Next i
    pic(0).Visible = True
End Sub

Private Sub xpWellsTab1_TabPressed(PreviousTab As Integer)
Dim i As Long
    For i = 0 To 3
        pic(i).Visible = False
    Next i
    pic(xpWellsTab1.SelectedTab - 1).Visible = True
End Sub
