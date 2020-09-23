VERSION 5.00
Begin VB.Form frmThunDll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThunDll"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdStorno 
      Caption         =   "Storno"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame fraDll 
      Caption         =   "Settings"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.CheckBox chbExportSymbols 
         Caption         =   "* Export functions"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtBaseAddress 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Tag             =   "*"
         Top             =   3480
         Width           =   1245
      End
      Begin VB.CheckBox chbLinkAsDLL 
         Caption         =   "* Create DLL"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtEntryPoint 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Tag             =   "*"
         Top             =   3840
         Width           =   1245
      End
      Begin VB.CheckBox chbDebugPreLoader 
         Caption         =   "* Debug ""Pre-Loader"""
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   5400
         Width           =   1935
      End
      Begin VB.CheckBox chbFullLoading 
         Caption         =   "* Full loading"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CheckBox chbUsePreLoader 
         Caption         =   "* Use ""Pre-Loader"""
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton cdmDLL_AddDllMain 
         Caption         =   "Add DllMain"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label lblDLL_1 
         AutoSize        =   -1  'True
         Caption         =   "* Base address  &&H"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   3480
         Width           =   1365
      End
      Begin VB.Label lblDLL_2 
         AutoSize        =   -1  'True
         Caption         =   "* Entry-Point name"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmThunDll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cdmDLL_AddDllMain_Click()
    frmDllMain.Show
End Sub
