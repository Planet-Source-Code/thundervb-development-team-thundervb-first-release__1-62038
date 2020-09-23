VERSION 5.00
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
Begin VB.Form frmLoadscreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   Picture         =   "frmLoadscreen.frx":0000
   ScaleHeight     =   1455
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   8265
      TabIndex        =   0
      Top             =   1200
      Width           =   8295
      Begin ThunVBCC_v1.UniLabel lblstatus 
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         CaptionB        =   "frmLoadscreen.frx":1FB42
         CaptionLen      =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(c) 2004-2005 drkIIRaziel && Libor"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   1
         Top             =   0
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmLoadscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

