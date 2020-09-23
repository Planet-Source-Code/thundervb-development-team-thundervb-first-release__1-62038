VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ISExplorerBar"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6780
   StartUpPosition =   1  'CenterOwner
   Begin ExplorerBarLiteTest.isExplorerBar isebAbout 
      Align           =   3  'Align Left
      Height          =   3435
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6059
      FontCharset     =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":09FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":0F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":18AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":1A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":1B5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
      Begin VB.Label lblDescription 
         Caption         =   "Comments, Suggestions and votes are wellcome."
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Tag             =   "Comments, Suggestions and votes are wellcome."
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "By Fred.cpp"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version ##.##"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "isExplorerBar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    isebAbout.SetImageList ImageList1
    isebAbout.AddSpecialGroup "About isExplorerBar", Me.Icon
    isebAbout.AddItem "Special Group", "e-mail", "E-mail to Fred.cpp", 7
    isebAbout.AddItem "Special Group", "homepage", "Visit Control HomePage", 3
    isebAbout.AddItem "Special Group", "About", "Show About Box", 1
    isebAbout.AddDetailsGroup "Details", "More Stuff", "If You like this project, please vote"
    lblVersion.Caption = "Version " & isebAbout.GetControlVersion
    isebAbout.Visible = True
End Sub

Private Sub isebAbout_ItemClick(sGroup As String, sItemKey As String)
    If sGroup = "Special Group" Then  'Special.
        Select Case sItemKey
            Case "e-mail" 'e-mail
                isebAbout.OpenLink "http://mx.geocities.com/fred_cpp/isexplorerbar_mail.htm"
            Case "homepage" 'Homepage
                isebAbout.OpenLink "http://mx.geocities.com/fred_cpp/isexplorerbar.htm"
            Case "About" ' you know...
                isebAbout.About
        End Select
    End If
End Sub

Private Sub isebAbout_ItemHover(sGroup As String, sItemKey As String)
    If sGroup = "Special Group" Then  'Special.
        Select Case sItemKey
            Case "e-mail" 'e-mail
                lblDescription.Caption = "If you have found a bug, or want to make any comment, please do It on the Planet-source-code.com page."
            Case "homepage" 'Homepage
                lblDescription.Caption = "If you want more Info about the control, visit It's home Page."
            Case "About" ' you know...
                lblDescription.Caption = "Show the About box generated by the control."
        End Select
    End If
End Sub

Private Sub isebAbout_ItemOut(sGroup As String, sItemKey As String)
    lblDescription.Caption = lblDescription.Tag
End Sub

