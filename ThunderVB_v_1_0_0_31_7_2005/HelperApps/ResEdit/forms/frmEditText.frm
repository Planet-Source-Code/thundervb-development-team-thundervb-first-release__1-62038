VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#24.2#0"; "ThunVBCC_v1.ocx"
Begin VB.Form frmEditText 
   Caption         =   "Text Entry Info"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin ThunVBCC_v1.vbaRichEdit Text2 
      Height          =   2415
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4260
      Version         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ViewMode        =   0
      TextLimit       =   3276700
      AutoURLDetect   =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   480
      Width           =   4095
   End
   Begin ResEdit.LanguageList LanguageList1 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close && Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   840
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
   Begin VB.Label Label1 
      Caption         =   "Language :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmEditText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelLang As tvb_Languages

Public Function ShowDialog(base As tvb_res_entry) As tvb_res_entry
    
    Dim temp As tvb_res_entry
    temp = base
    
    Text1.Text = temp.header.Id
    Text2.Text = temp.Data
    
    SelLang = base.header.language
    LanguageList1.SetLanguage base.header.language
    
    Me.Show vbModal
    
    With temp.header
        .DataType = tvb_res_Text
        .Id = Text1.Text
        .language = SelLang
        .PackInfo = 0
        .PackMode = tvb_res_Stored
    End With
    
    temp.Data = Text2.Text
    temp.Length = UBound(temp.Data) + 1
    ShowDialog = temp
    
    Unload Me
End Function

Private Sub Command1_Click()
    Me.Visible = False
End Sub

Private Sub LanguageList1_LanguageChanged(newLang As tvb_Languages)

    SelLang = newLang
    
End Sub

