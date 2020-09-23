VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
Begin VB.Form frmPlugIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure ThunderVB"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin ThunVBCC_v1.isButton isButton3 
      Height          =   300
      Left            =   240
      TabIndex        =   11
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Icon            =   "frmPlugIn.frx":0000
      Style           =   5
      Caption         =   "Report A Bug"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.XTab MainTab 
      Height          =   5775
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   10186
      TabCaption(0)   =   " Settings   "
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "curTab0"
      TabCaption(1)   =   "About      "
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "Info"
      TabCaption(2)   =   "Credits    "
      TabContCtrlCnt(2)=   1
      Tab(2)ContCtrlCap(1)=   "curTab2"
      ActiveTabHeight =   22
      InActiveTabHeight=   20
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   14215660
      ActiveTabBackEndColor=   14215660
      InActiveTabBackStartColor=   14215660
      InActiveTabBackEndColor=   12965594
      ActiveTabForeColor=   9982008
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.TextBox Info 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Text            =   "frmPlugIn.frx":001C
         Top             =   480
         Width           =   5055
      End
      Begin VB.PictureBox curTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5145
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   5145
         ScaleWidth      =   5025
         TabIndex        =   8
         Top             =   480
         WhatsThisHelpID =   1100
         Width           =   5025
         Begin ThunVBCC_v1.UniLabel Label2 
            Height          =   855
            Left            =   0
            Top             =   1920
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1508
            Alignment       =   2
            CaptionB        =   "frmPlugIn.frx":0035
            CaptionLen      =   22
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   21.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox curTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5175
         Index           =   0
         Left            =   120
         ScaleHeight     =   5175
         ScaleWidth      =   5055
         TabIndex        =   7
         Top             =   480
         WhatsThisHelpID =   700
         Width           =   5055
         Begin ThunVBCC_v1.UniLabel Label3 
            Height          =   735
            Left            =   0
            Top             =   2040
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   1296
            Alignment       =   2
            CaptionB        =   "frmPlugIn.frx":0081
            CaptionLen      =   22
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   21.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin ThunVBCC_v1.isButton isButton1 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Icon            =   "frmPlugIn.frx":00CD
      Style           =   4
      Caption         =   "GPF CRASH"
      IconAlign       =   1
      iNonThemeStyle  =   4
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.UniLabel Label4 
      Height          =   225
      Left            =   120
      Top             =   480
      Width           =   1110
      _ExtentX        =   132
      _ExtentY        =   26
      AutoSize        =   -1  'True
      CaptionB        =   "frmPlugIn.frx":00E9
      CaptionLen      =   14
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
   Begin ThunVBCC_v1.isButton cmdOk 
      Height          =   300
      Left            =   4200
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Icon            =   "frmPlugIn.frx":0125
      Style           =   5
      Caption         =   "Ok"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin ThunderVB_pl_v1_0.LanguageList cmbLang 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
   End
   Begin ThunVBCC_v1.isButton cmdCancel 
      Height          =   300
      Left            =   5400
      TabIndex        =   3
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Icon            =   "frmPlugIn.frx":0141
      Style           =   5
      Caption         =   "Cancel"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.isButton cmdApply 
      Height          =   300
      Left            =   6600
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Icon            =   "frmPlugIn.frx":015D
      Style           =   5
      Caption         =   "Apply"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ThunVBCC_v1.isButton isButton2 
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Icon            =   "frmPlugIn.frx":0179
      Style           =   4
      Caption         =   "VB UEXP"
      IconAlign       =   1
      iNonThemeStyle  =   4
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Public confhWnd As Long, oldpar As Long, oldConfhWnd As Long, oldCredhWnd As Long
Public credhWnd As Long, DoNotSendApp As Boolean


Public Sub RefreshLists()

Dim i As Long, li As Long
    li = List1.ListIndex
    
    List1.Clear
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            List1.AddItem plugins.plugins(i).Name
            List1.ItemData(List1.ListCount - 1) = i
        End If
    Next i
    
    If List1.ListCount > 0 Then List1.ListIndex = 0
    
End Sub

Public Sub cmbLang_LanguageChanged(newLang As tvb_Languages)
Dim i As Long
    Resource_SetCurLang newLang
    SaveSettingGlobal APP_NAMEs, "Language", (newLang)
    
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.InitSetLang newLang
            plugins.plugins(i).interface.SendMessange tvbm_ChangeLanguage_code, Resource_GetCurLang(), 0, Nothing
        End If
    Next i
    
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.SendMessange tvbm_ChangeLanguage_gui, Resource_GetCurLang(), 0, Nothing
        End If
    Next i
    
End Sub

Private Sub cmdApply_Click()
Dim i As Long
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.ApplySettings
        End If
    Next i
End Sub

Private Sub cmdCancel_Click()
Dim i As Long
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.CancelSettings
        End If
    Next i
    DoNotSendApp = True
    frmPlugin_Hide
End Sub

Private Sub cmdOk_Click()
Dim i As Long
    For i = 0 To plugins.count - 1
        If IsPluginLoaded(i) Then
            plugins.plugins(i).interface.ApplySettings
        End If
    Next i
    DoNotSendApp = True
    frmPlugin_Hide
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode <> 1 Then 'vbFormCode
        frmPlugin_Hide
    End If
    
End Sub



Private Sub isButton1_Click()
On Error GoTo er
t1
er:
End Sub
Private Sub t1()
t2
End Sub
Private Sub t2()
t3
End Sub
Public Sub t3()
    WinApiForVb.CopyMemory ByVal 0, 0, 100
End Sub


Private Sub isButton2_Click()
    Err.Raise 1000
End Sub

Private Sub isButton3_Click()

    RunExplorer "http://sourceforge.net/tracker/?atid=710267&group_id=128073&func=browse"
    
End Sub

Private Sub List1_Click()

Dim i As Long
    
    i = List1.ItemData(List1.ListIndex)
    
    With plugins.plugins(i)
           
        If oldpar = i Then
            Exit Sub
        Else
            If oldpar <> -1 Then
                If (Not (plugins.plugins(oldpar).interface Is Nothing)) Then
                    plugins.plugins(oldpar).interface.HideConfig
                    plugins.plugins(oldpar).interface.HideCredits
                    SetParent confhWnd, oldConfhWnd 'restore the window parent..
                    SetParent credhWnd, oldConfhWnd 'restore the window parent..
                End If
            End If
        End If
        
        oldpar = i
        
        confhWnd = .interface.ShowConfig
        
        If confhWnd <> 0 Then
            Label2.Visible = False
            Label3.Visible = False
        Else
            Label2.Visible = True
            Label3.Visible = True
        End If
        
        oldConfhWnd = SetParent(confhWnd, curTab(0).hWnd)
        
        credhWnd = .interface.ShowCredits
        oldCredhWnd = SetParent(credhWnd, curTab(2).hWnd)
        
        
    lang_UpdateTexts
    
    End With
    
    If Me.Visible Then SetFocusApi Me.hWnd     'set focus to this form.. (sometimes it flies around due to a vb6/winapi bug).. heh , a hiden form recieves the focus..
    
End Sub

Public Sub lang_UpdateGui(lang As tvb_Languages)

        'Resource_LoadFormFromResourceFile tvb_resfile, Me, "ThunderVB_pl", lang
        resfile.LoadFormFromResourceFile Me, "ThunderVB_pl", lang
        MainTab.TabCaption(0) = resfile.GetTextByIndex(lng_frmpl.MainTab.TabSettings)
        MainTab.TabCaption(1) = resfile.GetTextByIndex(lng_frmpl.MainTab.TabAbout)
        MainTab.TabCaption(2) = resfile.GetTextByIndex(lng_frmpl.MainTab.TabCredits)
        lang_UpdateTexts
        
End Sub

Private Sub lang_UpdateTexts()
Dim i As Long
        On Error GoTo err_handle
        i = List1.ItemData(List1.ListIndex)
        
        With plugins.plugins(i)
        
            Info.text = resfile.GetTextByIndex(lng_frmpl.Info.Name) & vbTab & ": " & .Name & vbNewLine & _
                        resfile.GetTextByIndex(lng_frmpl.Info.Version) & vbTab & ": " & .Version & vbNewLine & _
                        resfile.GetTextByIndex(lng_frmpl.Info.Speed) & vbTab & ": " & pl_Speed_req2String(.Speed) & vbNewLine & _
                        resfile.GetTextByIndex(lng_frmpl.Info.type) & vbTab & ": " & pl_type2String(.type) & vbNewLine & vbNewLine & _
                        resfile.GetTextByIndex(lng_frmpl.Info.WorksOnWin) & " :" & vbNewLine & _
                        .interface.GetWindowsVersion & ": " & vbNewLine & vbNewLine & _
                        resfile.GetTextByIndex(lng_frmpl.Info.ShortDesc) & " :" & vbNewLine & _
                        .Desciption & vbNewLine & vbNewLine & _
                        resfile.GetTextByIndex(lng_frmpl.Info.FullDesc) & " :" & vbNewLine & _
                        .DesciptionFull & vbNewLine
        End With

err_handle:
End Sub

Public Sub RunExplorer(url As String)
    Shell "explorer " & Add34(url)
End Sub
