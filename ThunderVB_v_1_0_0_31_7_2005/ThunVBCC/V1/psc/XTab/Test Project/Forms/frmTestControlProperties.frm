VERSION 5.00
Object = "*\A..\..\prjXTab.vbp"
Begin VB.Form frmTestControlProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                     :: XTabs :: by Neeraj Agrawal @ PSC ::"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmTestControlProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPreview 
      BackColor       =   &H00D8E9EC&
      Height          =   1965
      Left            =   30
      TabIndex        =   38
      Top             =   -120
      Width           =   6765
      Begin prjXTab.XTab XTab1 
         Height          =   1545
         Left            =   90
         TabIndex        =   0
         Top             =   270
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   2725
         TabCaption(0)   =   "Tab 0"
         TabContCtrlCnt(0)=   4
         Tab(0)ContCtrlCap(1)=   "Label4"
         Tab(0)ContCtrlCap(2)=   "Label3"
         Tab(0)ContCtrlCap(3)=   "Label2"
         Tab(0)ContCtrlCap(4)=   "Label1"
         TabCaption(1)   =   "Tab 1"
         TabCaption(2)   =   "Tab 2"
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-Try making the active tab smaller and inactive tab larger in height."
            Height          =   195
            Left            =   810
            TabIndex        =   42
            Top             =   1110
            Width           =   5820
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-Make sure the Xtab control has the focus and then use arrow keys."
            Height          =   195
            Left            =   810
            TabIndex        =   41
            Top             =   870
            Width           =   5835
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-Assign Acces keys using the && character in the tab caption."
            Height          =   195
            Left            =   810
            TabIndex        =   40
            Top             =   630
            Width           =   5145
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Try These:"
            Height          =   195
            Left            =   390
            TabIndex        =   39
            Top             =   360
            Width           =   930
         End
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "::Close::"
      Height          =   285
      Left            =   2453
      TabIndex        =   21
      Top             =   7650
      Width           =   2025
   End
   Begin VB.Frame fraGeneralProperties 
      Caption         =   " General Properties "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   30
      TabIndex        =   27
      Top             =   1890
      Width           =   6765
      Begin VB.TextBox txtTabCount 
         Height          =   285
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   1
         Top             =   270
         Width           =   600
      End
      Begin VB.ComboBox cboTabStyle 
         Height          =   315
         ItemData        =   "frmTestControlProperties.frx":000C
         Left            =   1710
         List            =   "frmTestControlProperties.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   1485
      End
      Begin VB.ComboBox cboTabTheme 
         Height          =   315
         ItemData        =   "frmTestControlProperties.frx":0037
         Left            =   5010
         List            =   "frmTestControlProperties.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   1365
      End
      Begin VB.TextBox txtYRadius 
         Height          =   285
         Left            =   5025
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1560
         Width           =   630
      End
      Begin VB.TextBox txtXRadius 
         Height          =   285
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1560
         Width           =   600
      End
      Begin VB.TextBox txtInActiveTabHeight 
         Height          =   285
         Left            =   5025
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1140
         Width           =   630
      End
      Begin VB.TextBox txtActiveTabHeight 
         Height          =   285
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1140
         Width           =   600
      End
      Begin VB.CheckBox chkShowFocusRect 
         Caption         =   "Show Focus Rect"
         Height          =   225
         Left            =   1710
         TabIndex        =   8
         Top             =   2220
         Width           =   1605
      End
      Begin VB.CommandButton cmdSelectMaskColor 
         Caption         =   "Select"
         Height          =   225
         Left            =   5700
         TabIndex        =   12
         Top             =   3390
         Width           =   645
      End
      Begin VB.CheckBox chkUseMaskColor 
         Caption         =   "Use Mask Color for Picture"
         Height          =   195
         Left            =   3450
         TabIndex        =   11
         Top             =   3420
         Width           =   2235
      End
      Begin VB.ComboBox cboPictureSize 
         Height          =   315
         ItemData        =   "frmTestControlProperties.frx":0097
         Left            =   1710
         List            =   "frmTestControlProperties.frx":00A1
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   1485
      End
      Begin VB.ComboBox cboPictureAlignment 
         Height          =   315
         ItemData        =   "frmTestControlProperties.frx":00B3
         Left            =   1710
         List            =   "frmTestControlProperties.frx":00C3
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3330
         Width           =   1485
      End
      Begin VB.Label lblTabCount 
         AutoSize        =   -1  'True
         Caption         =   "Tab Count:"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   315
         Width           =   810
      End
      Begin VB.Label lblTabStyle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Style:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblTabTheme 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Theme:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3450
         TabIndex        =   35
         Top             =   750
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Caption         =   "* X Radius and Y Radius apply only to  the Round Tabs Theme"
         ForeColor       =   &H80000010&
         Height          =   405
         Left            =   1710
         TabIndex        =   34
         Top             =   1890
         Width           =   4665
      End
      Begin VB.Label lblYRadius 
         AutoSize        =   -1  'True
         Caption         =   "Y Radius:"
         Height          =   195
         Left            =   3450
         TabIndex        =   33
         Top             =   1605
         Width           =   675
      End
      Begin VB.Label lblXRadius 
         AutoSize        =   -1  'True
         Caption         =   "X Radius:"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   1605
         Width           =   675
      End
      Begin VB.Label lblInactiveTabHeight 
         AutoSize        =   -1  'True
         Caption         =   "InActive Tab Height:"
         Height          =   195
         Left            =   3450
         TabIndex        =   31
         Top             =   1185
         Width           =   1485
      End
      Begin VB.Label lblActiveTabHeight 
         AutoSize        =   -1  'True
         Caption         =   "Active Tab Height:"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   1185
         Width           =   1335
      End
      Begin VB.Label lblPictureSize 
         AutoSize        =   -1  'True
         Caption         =   "Picture Size:"
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   2940
         Width           =   885
      End
      Begin VB.Label lblPictureAlignment 
         AutoSize        =   -1  'True
         Caption         =   "Picture Alignment:"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   3390
         Width           =   1305
      End
   End
   Begin VB.Frame fraTabSpecificProperties 
      Caption         =   " Tab Specific Properties "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   30
      TabIndex        =   22
      Top             =   5790
      Width           =   6765
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   285
         Left            =   5190
         TabIndex        =   20
         Top             =   1230
         Width           =   1245
      End
      Begin VB.CommandButton cmdRemovePic 
         Caption         =   "Remove"
         Height          =   225
         Left            =   2310
         TabIndex        =   19
         Top             =   1350
         Width           =   735
      End
      Begin VB.CheckBox chkTabEnabled 
         Caption         =   "Enabled"
         Height          =   195
         Left            =   3450
         TabIndex        =   16
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox txtTabCaption 
         Height          =   285
         Left            =   1710
         MaxLength       =   255
         TabIndex        =   17
         Top             =   690
         Width           =   4725
      End
      Begin VB.CommandButton cmdAssignPic 
         Caption         =   "Assign"
         Height          =   225
         Left            =   2310
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox picOuter 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1710
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1080
         Width           =   510
         Begin VB.Image imgPreview 
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   300
         Width           =   200
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   285
         Left            =   2970
         TabIndex        =   15
         Top             =   300
         Width           =   200
      End
      Begin VB.TextBox txtActiveTab 
         Height          =   285
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   13
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblTabCaption 
         AutoSize        =   -1  'True
         Caption         =   "Caption: "
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   735
         Width           =   660
      End
      Begin VB.Label lblPicture 
         AutoSize        =   -1  'True
         Caption         =   "Picture:"
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   1245
         Width           =   555
      End
      Begin VB.Label lblActiveTab 
         AutoSize        =   -1  'True
         Caption         =   "Active Tab:"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   345
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmTestControlProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bIgnoreChanges As Boolean      'used to prevent Setting of "m_bChanged" flag or property page
Private m_bChanged As Boolean


'=====Event Handlers==================================================================================================

Private Sub cboPictureAlignment_Click()
  If Not m_bIgnoreChanges Then
    XTab1.PictureAlign = cboPictureAlignment.ListIndex      'if the ignore flag is off then set the Pic Alignment
  End If
End Sub

Private Sub cboPictureSize_Click()
  If Not m_bIgnoreChanges Then
    XTab1.PictureSize = cboPictureSize.ListIndex      'if the ignore flag is off then set the Pic Size
  End If
End Sub

Private Sub cboTabStyle_Click()
  If Not m_bIgnoreChanges Then
    XTab1.TabStyle = cboTabStyle.ListIndex       'if the ignore flag is off then set the TabStyle
  End If
End Sub

Private Sub cboTabTheme_Click()
  If Not m_bIgnoreChanges Then
    XTab1.TabTheme = cboTabTheme.ListIndex 'if the ignore flag is off then set the Tab Theme
  End If
End Sub

Private Sub chkShowFocusRect_Click()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub chkTabEnabled_Click()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub chkUseMaskColor_Click()
  If Not m_bIgnoreChanges Then
    If chkUseMaskColor.Value = vbChecked Then
      XTab1.UseMaskColor = True
      cmdSelectMaskColor.Enabled = True
    Else
      XTab1.UseMaskColor = False
      cmdSelectMaskColor.Enabled = False
    End If
  End If
End Sub

Private Sub cmdApply_Click()
  If pValidatePropValues Then
    Call pSavePropValues
    m_bChanged = False
  Else
    m_bChanged = True
  End If
End Sub

Private Sub cmdAssignPic_Click()
  Dim sTmp As String
  
  sTmp = InputBox("Enter Path for the Image file.", "Image file")
  If sTmp <> "" Then
    
    'done to destroy the resource
    Set XTab1.TabPicture(XTab1.ActiveTab) = Nothing
    
    'done to destroy the resource
    Set imgPreview.Picture = Nothing
    
    Set imgPreview.Picture = LoadPicture(sTmp)  'load the whole picture from the file and store it into the tab's picture property
    'Set XTab1.TabPicture(XTab1.ActiveTab) = LoadPicture(sTmp)  'load the whole picture from the file and store it into the tab's picture property
    
    Set XTab1.TabPicture(XTab1.ActiveTab) = imgPreview.Picture
    
    'Set imgPreview.Picture = XTab1.TabPicture(XTab1.ActiveTab)
    
    Call pAlignImage
    
    Refresh    'donno but must refresh or else sometimes the picture goes away :(
    
  End If

End Sub

Private Sub cmdBack_Click()
  Unload Me
End Sub

Private Sub cmdNext_Click()
  Dim iTmp As Integer
  
  iTmp = XTab1.ActiveTab
  
  If iTmp < XTab1.TabCount - 1 Then
    
    If m_bChanged Then
      Call pSaveOnlyActiveTabControlProperties   'save existing props first
    End If
    
    XTab1.ActiveTab = iTmp + 1
    Call pGetAndFillPropValues
  End If
End Sub

Private Sub cmdPrevious_Click()
  
  If XTab1.ActiveTab > 0 Then
    
    If m_bChanged Then
      Call pSaveOnlyActiveTabControlProperties   'save existing props first
    End If
    
    XTab1.ActiveTab = XTab1.ActiveTab - 1
    Call pGetAndFillPropValues
  End If
End Sub

Private Sub cmdRemovePic_Click()
  Set XTab1.TabPicture(XTab1.ActiveTab) = Nothing
  Set imgPreview.Picture = Nothing
End Sub

Private Sub cmdSelectMaskColor_Click()
  On Error Resume Next
  
  'Swicth to Color's Property Page
  SendKeys "^+{Tab}"
End Sub


Private Sub Form_Load()
  Call pGetAndFillPropValues
End Sub

Private Sub txtActiveTab_Change()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub txtActiveTabHeight_Change()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub txtInActiveTabHeight_Change()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub txtTabCaption_Change()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub txtTabCount_Change()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub txtXRadius_Change()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

Private Sub txtYRadius_Change()
  If Not m_bIgnoreChanges Then m_bChanged = True     'if the ignore flag is off, indicate some property has m_bChanged
End Sub

'=====================================================================================================================


'====Private Functions================================================================================================
Private Sub pAlignImage()
    'do adjustments
    If imgPreview.Width < 32 Then
      imgPreview.Stretch = False
      imgPreview.Left = 16 - imgPreview.Width / 2
    Else
      imgPreview.Stretch = True
      imgPreview.Width = 32
      imgPreview.Left = 0
    End If
    
    If imgPreview.Height < 32 Then
      imgPreview.Stretch = False
      imgPreview.Top = 16 - imgPreview.Height / 2
    Else
      imgPreview.Stretch = True
      imgPreview.Height = 32
      imgPreview.Top = 0
    End If
End Sub


Private Function pValidatePropValues() As Boolean
  Dim oCtl As Control
  m_bIgnoreChanges = True
  
  For Each oCtl In Controls
    If TypeOf oCtl Is TextBox Then
      If oCtl.Name <> "txtTabCaption" Then
        oCtl.Text = Trim$(oCtl.Text)
        
        If IsNumeric(oCtl.Text) Then
          If CLng(oCtl.Text) < 0 Then
            oCtl.SetFocus
            MsgBox "Enter a value between 1 and 32767 in '" & Mid$(oCtl.Name, 4) & "' text box.", vbInformation
            m_bIgnoreChanges = False
            Exit Function
          ElseIf CLng(oCtl.Text) > 32767 Then
            oCtl.SetFocus
            MsgBox "Enter a value between 1 and 32767 in '" & Mid$(oCtl.Name, 4) & "' text box.", vbInformation
            m_bIgnoreChanges = False
            Exit Function
          End If
        Else
          oCtl.SetFocus
          MsgBox "Enter a value between 1 and 32767 in active text box.", vbInformation
          m_bIgnoreChanges = False
          Exit Function
        End If
      End If
    End If
  Next
  
  If CLng(txtActiveTab.Text) >= XTab1.TabCount Then
    txtActiveTab.SetFocus
    MsgBox "Active tab should be a number between 1 to TabCount -1", vbInformation, "Tab Count"
    m_bIgnoreChanges = False
    Exit Function
  End If
  
  
  m_bIgnoreChanges = False
  
  pValidatePropValues = True
End Function

Private Sub pSavePropValues()
  With XTab1
    '.bIgnoreRedraw = True    'was trying to prevent rerdraw as many properties are being m_bChanged...
                              'but can not access friend functions in property pages
                              'so dropped the idea. But left it in the place as there can
                              'be some other way of doing it.
    
    .TabCount = txtTabCount.Text
    
    Call pSaveOnlyActiveTabControlProperties
   
    '.TabStyle = cboTabStyle.ListIndex
    
    'update the theme only if its different than the current theme
    'If Not .TabTheme = cboTabTheme.ListIndex Then
    '  .TabTheme = cboTabTheme.ListIndex
    'End If
    
    .ActiveTabHeight = txtActiveTabHeight.Text
    .InActiveTabHeight = txtInActiveTabHeight.Text
    .XRadius = txtXRadius.Text
    .YRadius = txtYRadius.Text
    
    .ShowFocusRect = IIf(chkShowFocusRect.Value = vbChecked, True, False)
    
    '.bIgnoreRedraw = False
    '.pRefresh
  End With
End Sub
Private Sub pGetAndFillPropValues()
  m_bIgnoreChanges = True
  
  With XTab1
    txtTabCount.Text = .TabCount
    txtActiveTab.Text = .ActiveTab
    txtTabCaption.Text = .TabCaption(.ActiveTab)
    
    If .TabPicture(.ActiveTab) Is Nothing Then
      imgPreview.Picture = Nothing
      cmdRemovePic.Enabled = False
    Else
      imgPreview.Picture = .TabPicture(.ActiveTab)
      cmdRemovePic.Enabled = True
    End If
    
    Call pAlignImage
    
    chkTabEnabled.Value = IIf(.TabEnabled(.ActiveTab), vbChecked, vbUnchecked)
   
    cboTabStyle.ListIndex = .TabStyle
    cboTabTheme.ListIndex = .TabTheme
    
    txtActiveTabHeight.Text = .ActiveTabHeight
    txtInActiveTabHeight.Text = .InActiveTabHeight
    txtXRadius.Text = .XRadius
    txtYRadius.Text = .YRadius
    
    chkShowFocusRect.Value = IIf(.ShowFocusRect, vbChecked, vbUnchecked)
    If .UseMaskColor Then
      chkUseMaskColor.Value = vbChecked
      cmdSelectMaskColor.Enabled = True
    Else
      chkUseMaskColor.Value = vbUnchecked
      cmdSelectMaskColor.Enabled = False
    End If
    chkUseMaskColor.Value = IIf(.UseMaskColor, vbChecked, vbUnchecked)
    cboPictureSize.ListIndex = .PictureSize
    cboPictureAlignment.ListIndex = .PictureAlign
    
  End With
  m_bIgnoreChanges = False
End Sub

Private Sub pSaveOnlyActiveTabControlProperties()
    With XTab1
      '.bIgnoreRedraw = True    'was trying to prevent rerdraw as many properties are being m_bChanged...
                                'but can not access friend functions in property pages
                                'so dropped the idea. But left it in the place as there can
                                'be some other way of doing it.
      
      .TabCaption(.ActiveTab) = txtTabCaption.Text
      
      'Picture already set
      'If Not imgPreview.Picture Is Nothing Then
      '  Set .TabPicture(.ActiveTab) = LoadPicture(imgPreview.Tag, vbLPCustom, , 16, 16)
      'Else
      '  Set .TabPicture(.ActiveTab) = Nothing
      'End If
      
      .TabEnabled(.ActiveTab) = IIf(chkTabEnabled.Value = vbChecked, True, False)
      '.bIgnoreRedraw = False
      '.pRefresh
    End With
End Sub

'=====================================================================================================================

