VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
Begin VB.Form frmInternal 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Credits 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   5520
      ScaleHeight     =   5055
      ScaleWidth      =   5295
      TabIndex        =   5
      Top             =   0
      Width           =   5295
      Begin ThunVBCC_v1.UniLabel label4 
         Height          =   615
         Left            =   120
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1085
         Alignment       =   2
         CaptionB        =   "frmInternal.frx":0000
         CaptionLen      =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   20.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmInternal.frx":0034
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.PictureBox ConfigInternal 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   120
      ScaleHeight     =   5055
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin ThunVBCC_v1.UniLabel lblLabel3 
         Height          =   225
         Left            =   2760
         Top             =   840
         Width           =   1290
         _ExtentX        =   159
         _ExtentY        =   26
         AutoSize        =   -1  'True
         CaptionB        =   "frmInternal.frx":0213
         CaptionLen      =   19
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
      Begin ThunVBCC_v1.UniLabel lblLabel2 
         Height          =   225
         Left            =   0
         Top             =   840
         Width           =   1185
         _ExtentX        =   132
         _ExtentY        =   26
         AutoSize        =   -1  'True
         CaptionB        =   "frmInternal.frx":0259
         CaptionLen      =   16
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
      Begin ThunVBCC_v1.UniLabel lblLabel1 
         Height          =   225
         Left            =   1560
         Top             =   480
         Width           =   1980
         _ExtentX        =   238
         _ExtentY        =   26
         Alignment       =   2
         AutoSize        =   -1  'True
         CaptionB        =   "frmInternal.frx":0299
         CaptionLen      =   24
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
      Begin ThunVBCC_v1.UniLabel UniLabel1 
         Height          =   450
         Left            =   0
         Top             =   4440
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   794
         Alignment       =   2
         CaptionB        =   "frmInternal.frx":02E9
         CaptionLen      =   73
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
      Begin VB.ListBox lstLoaded 
         Height          =   3180
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ListBox lstAvail 
         Height          =   3180
         Left            =   2760
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "<<<"
         Height          =   615
         Left            =   2400
         TabIndex        =   2
         ToolTipText     =   "Load Selected Plugin"
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton cmdUnLoad 
         Caption         =   ">>>"
         Height          =   615
         Left            =   2400
         TabIndex        =   1
         ToolTipText     =   "Unload Selected Plugin"
         Top             =   2040
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmInternal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
    Dim i As Long
    On Error GoTo ext
    SetLoadScreenStatus lsS_Visible
    plugins.plugins(lstAvail.ItemData(lstAvail.ListIndex)).used = False 'free this slot
    i = LoadPlugin(plugins.plugins(lstAvail.ItemData(lstAvail.ListIndex)).dllfile)
    plugins.plugins(i).interface.OnGuiLoad
    SaveSetting APP_NAME, "plugins", plugins.plugins(i).dllfile & "_loaded", "true"
ext:
    RefreshLists
    SetLoadScreenStatus lsS_Hiden
End Sub

Public Sub lang_UpdateGui(lang As tvb_Languages)

        resfile.LoadFormFromResourceFile Me, "ThunderVB_pl", lang
        frmPlugIn.RefreshLists
    
End Sub


Private Sub cmdUnLoad_Click()
On Error GoTo ext
Dim i As Long
    SetLoadScreenStatus lsS_Visible
    i = lstLoaded.ItemData(lstLoaded.ListIndex)
    SaveSetting APP_NAME, "plugins", plugins.plugins(i).dllfile & "_loaded", "false"
    On Error Resume Next
    plugins.plugins(i).interface.HideCredits
    plugins.plugins(i).interface.HideConfig
    plugins.plugins(i).interface.OnGuiUnLoad
    On Error GoTo ext
    UnLoadPlugin lstLoaded.ItemData(lstLoaded.ListIndex)
    
ext:
    RefreshLists
    SetLoadScreenStatus lsS_Hiden
End Sub

Public Sub RefreshLists()
Dim i As Long
    lstLoaded.Clear
    lstAvail.Clear
    For i = 0 To plugins.count - 1
        If plugins.plugins(i).used = True Then
            If plugins.plugins(i).Loaded = True Then
                lstLoaded.AddItem plugins.plugins(i).Name
                lstLoaded.ItemData(lstLoaded.ListCount - 1) = i
            Else
                lstAvail.AddItem plugins.plugins(i).Name
                lstAvail.ItemData(lstAvail.ListCount - 1) = i
            End If
        End If
    Next i
    On Error Resume Next
    frmPlugIn.RefreshLists
End Sub

Private Sub Form_Load()
    Credits.Move 0, 0
    ConfigInternal.Move 0, 0
    
    RefreshLists

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub lstAvail_Click()
    lstLoaded.ListIndex = -1
End Sub

Private Sub lstLoaded_Click()
    Dim i As Long
    lstAvail.ListIndex = -1
    If lstLoaded.ListIndex > -1 Then
        i = lstLoaded.ItemData(lstLoaded.ListIndex)
    
        If plugins.plugins(i).dllfile = App.ProductName Then
            cmdLoad.Enabled = False
            cmdUnLoad.Enabled = False
        Else
            cmdLoad.Enabled = True
            cmdUnLoad.Enabled = True
        End If
    Else
        cmdLoad.Enabled = True
        cmdUnLoad.Enabled = True
    End If
    
End Sub


