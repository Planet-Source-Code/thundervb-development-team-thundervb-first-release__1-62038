VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImgEdit 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   480
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Replace"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save As"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close && Save"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
   Begin ResEdit.LanguageList LanguageList1 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
   End
   Begin VB.Image picture1 
      Height          =   2295
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Language :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Data"
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmImgEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelLang As tvb_Languages
Dim entry As tvb_res_entry

Public Function ShowDialog(base As tvb_res_entry) As tvb_res_entry
    
    entry = base
    
    Text1.Text = entry.header.Id
    Dim temp As tvb_res_Data
    
    temp.Data = entry.Data
    temp.Length = entry.Length
    
    picture1.Picture = Resource_LoadImageFromResData(temp)
    
    SelLang = entry.header.language
    LanguageList1.SetLanguage entry.header.language
    
    Me.Show vbModal
    
    
    temp = Resource_SaveImageToResData(picture1)
    
    entry.Data = temp.Data
    entry.Length = temp.Length
    
    With entry.header
        .DataType = tvb_res_Image
        .Id = Text1.Text
        .language = SelLang
        .PackInfo = 0
        .PackMode = tvb_res_Stored
    End With
    
    ShowDialog = entry
    Unload Me
End Function

Private Sub Command1_Click()

    Me.Visible = False
    
End Sub

Private Sub Command2_Click()

    cd1.CancelError = False
    cd1.DialogTitle = "Save As"
    cd1.ShowSave
    SavePicture picture1.Picture, cd1.FileName
    
End Sub

Private Sub Command3_Click()

    cd1.CancelError = False
    cd1.DialogTitle = "Select an image.."
    cd1.ShowOpen
    picture1.Picture = LoadPicture(cd1.FileName)
    
End Sub

Private Sub LanguageList1_LanguageChanged(newLang As tvb_Languages)

    SelLang = newLang
    
End Sub

