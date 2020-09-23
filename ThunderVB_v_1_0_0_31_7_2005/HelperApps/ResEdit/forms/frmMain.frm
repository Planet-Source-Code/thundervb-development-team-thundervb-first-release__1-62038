VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{972B81FA-2CBA-47A4-9D2B-259A900985D0}#25.1#0"; "ThunVBCC_v1_0.ocx"
Begin VB.Form frmMain 
   Caption         =   "ThunderVB resource Files Editor"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   5055
      Left            =   6360
      ScaleHeight     =   4995
      ScaleWidth      =   3675
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command11 
         Caption         =   "Replace with file"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Save as file"
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Load"
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin ThunVBCC_v1.vbaRichEdit Text1 
      Height          =   5055
      Left            =   6360
      TabIndex        =   9
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8916
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
      TextLimit       =   -1
      AutoURLDetect   =   0   'False
      TextOnly        =   -1  'True
      DisableNoScroll =   -1  'True
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9600
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Resource Files (*.gre)|*.gre"
   End
   Begin MSComctlLib.TreeView lstRes 
      Height          =   6375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11245
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command8 
      Caption         =   "New copy of selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add New Image entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add NewTextEntry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   5640
      Width           =   1815
   End
   Begin ThunVBCC_v1.vbaRichEdit Text2 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
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
      Text            =   "C:\develop\vb\vb-projects\ThunderVB_plugined\bin\tvb.gre"
      ViewMode        =   0
      ControlRightMargin=   100
      TextLimit       =   30000
      AutoURLDetect   =   0   'False
      TextOnly        =   -1  'True
      SingleLine      =   -1  'True
      ScrollBars      =   3
   End
   Begin VB.Image picture1 
      Height          =   5055
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSEH: ErrMsgBox
Option Explicit
Dim t As New cResFile
Dim nd() As Node
Dim ndc As Long

Public Enum States
    NoFileLoaded = 0
    FileLoaded = 1
    FileSaved = 2
    FileNotSaved = 3
End Enum

Public lState As States, Ri As Long

Private Sub Command1_Click()
    '<EhHeader>
    On Error GoTo Command1_Click_Err
    '</EhHeader>
    Dim temp As tvb_res_entry, i As Long
    
    SetState FileNotSaved
    i = GetSelectedID
        
    If i = -1 Then GoTo nop
    
    If t.Items(i).header.DataType = tvb_res_Image Then
        temp = t.Items(i)
    ElseIf t.Items(i).header.DataType = tvb_res_Text Then
        temp = t.Items(i)
    Else
        MsgBox "error"
    End If
    
nop:
    t.AddEntry frmEditText.ShowDialog(temp)
    RefreshList
    
    '<EhFooter>
    Exit Sub

Command1_Click_Err:
    MsgBox err.Description & vbCrLf & _
           "in ResEdit.frmMain.Command1_Click " & _
           "at line " & Erl

    '</EhFooter>
End Sub

Private Sub Command10_Click()

    cd1.CancelError = False
    cd1.DialogTitle = "Save As"
    cd1.ShowSave
    
    SaveFile_bin cd1.FileName, t.Items(Ri).Data
    
End Sub

Private Sub Command11_Click()
    
    cd1.CancelError = False
    cd1.DialogTitle = "Select an image.."
    cd1.ShowOpen
    
    Dim temp As tvb_res_entry
    
    temp = t.Items(Ri)
    temp.Data = LoadFile_bin(cd1.FileName)
    temp.Length = UBound(temp.Data)
    t.Items(Ri) = temp
    
    RefreshList
    
End Sub

Private Sub Command2_Click()
    '<EhHeader>
    On Error GoTo Command2_Click_Err
    '</EhHeader>
    Dim temp As tvb_res_entry
    
    SetState FileNotSaved
    t.AddEntry frmImgEdit.ShowDialog(temp)
    RefreshList
    '<EhFooter>
    Exit Sub

Command2_Click_Err:
    MsgBox err.Description & vbCrLf & _
           "in ResEdit.frmMain.Command2_Click " & _
           "at line " & Erl

    '</EhFooter>
End Sub

Private Sub Command3_Click()
    '<EhHeader>
    On Error GoTo Command3_Click_Err
    '</EhHeader>
Dim i As Long, temp2 As tvb_res_Data
    
    SetState FileNotSaved
    i = GetSelectedID
    If i = -1 Then Exit Sub
    
    If t.Items(i).header.DataType = tvb_res_Image Then
        Text1.Visible = False
        t.Items(i) = frmImgEdit.ShowDialog(t.Items(i))
        RefreshList
    ElseIf t.Items(i).header.DataType = tvb_res_Text Then
        t.Items(i) = frmEditText.ShowDialog(t.Items(i))
        RefreshList
    Else
        MsgBox "error"
    End If
    
    '<EhFooter>
    Exit Sub

Command3_Click_Err:
    MsgBox err.Description & vbCrLf & _
           "in ResEdit.frmMain.Command3_Click " & _
           "at line " & Erl

    '</EhFooter>
End Sub

Private Sub Command4_Click()
    '<EhHeader>
    On Error GoTo Command4_Click_Err
    '</EhHeader>
    Dim i As Long
    
    SetState FileNotSaved
    i = GetSelectedID
    If i = -1 Then Exit Sub
    
    t.RemoveEntry i
    RefreshList
    
    '<EhFooter>
    Exit Sub

Command4_Click_Err:
    MsgBox err.Description & vbCrLf & _
           "in ResEdit.frmMain.Command4_Click " & _
           "at line " & Erl

    '</EhFooter>
End Sub

Private Sub Command5_Click()
    '<EhHeader>
    On Error GoTo Command5_Click_Err
    '</EhHeader>

    cd1.CancelError = False
    cd1.DialogTitle = "Select a file.."
    cd1.ShowOpen
    Text2.Text = cd1.FileName
    Command7_Click
    
    '<EhFooter>
    Exit Sub

Command5_Click_Err:
    MsgBox err.Description & vbCrLf & _
           "in ResEdit.frmMain.Command5_Click " & _
           "at line " & Erl

    '</EhFooter>
End Sub

Private Sub Command6_Click()
    '<EhHeader>
    On Error GoTo Command6_Click_Err
    '</EhHeader>
    
    SetState FileSaved
    'Dim i As Long, temp As tvb_res_entry
    'For i = 0 To t.ItemCount - 1
    '    temp = t.Items(i)
    '    temp.header.PackMode = temp.header.PackMode Or tvb_res_Compressed
    '    t.Items(i) = temp
    'Next i
    
    t.SaveFileAs Text2.Text, ""
    
    RefreshList
    
    '<EhFooter>
    Exit Sub

Command6_Click_Err:
    MsgBox err.Description & vbCrLf & _
           "in ResEdit.frmMain.Command6_Click " & _
           "at line " & Erl

    '</EhFooter>
End Sub

Private Sub Command7_Click()

    
    't = Resource_LoadResourceFile(Text2.Text, "")
    Dim errstr As String
    If t.OpenFile(Text2.Text, "", , errstr) Then
        SetState FileLoaded
        RefreshList
    Else
        MsgBox "Error : " & errstr
    End If

End Sub

Private Sub Command8_Click()

    Dim i As Long
    
    SetState FileNotSaved
    i = GetSelectedID
    If i = -1 Then Exit Sub
    
    If t.Items(i).header.DataType = tvb_res_Text Then
        t.AddEntry frmEditText.ShowDialog(t.Items(i))
        RefreshList
    End If
    

End Sub

Private Sub RefreshList()

Dim i As Long
    
    lstRes.Nodes.Clear: ndc = 0

    For i = 0 To t.ItemCount - 1
        AddNode i
    Next i
    

End Sub
Private Sub AddNode(ind As Long)
    '<EhHeader>
    On Error GoTo AddNode_Err
    '</EhHeader>
    On Error GoTo err
            ReDim Preserve nd(ndc)
            Set nd(ndc) = lstRes.Nodes.Add(, , t.Items(ind).header.Id, t.Items(ind).header.Id & " (1)")
            lstRes.Nodes.Add nd(ndc), tvwChild, "s0" & ind, Resource_LanguageIdToString(t.Items(ind).header.language) & "; Size : " & FormatSize(t.Items(ind).Length)
            ndc = ndc + 1
    Exit Sub
err:
    Dim i As Long
    For i = 0 To ndc - 1
        If nd(i).Key = t.Items(ind).header.Id Then
            lstRes.Nodes.Add nd(i), tvwChild, "s0" & ind, Resource_LanguageIdToString(t.Items(ind).header.language) & "; Size : " & FormatSize(t.Items(ind).Length)
            nd(i).Text = nd(i).Key & " (" & nd(i).Children & ")"
        End If
    Next i
            
    '<EhFooter>
    Exit Sub

AddNode_Err:
    MsgBox err.Description & vbCrLf & _
           "in ResEdit.frmMain.AddNode " & _
           "at line " & Erl

    '</EhFooter>
End Sub

Private Sub RemoveNode(ind As Long)

End Sub

Private Sub Command9_Click()

    SetState FileLoaded
    t = Resource_NewFile("Test", "Test", "tvb_ResEdit")

End Sub

Private Sub Form_Load()

    SetState NoFileLoaded

End Sub

Private Sub lstRes_DblClick()

    Command3_Click
    
End Sub

Private Sub lstRes_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        Command4_Click
    End If

End Sub

Private Sub lstRes_NodeClick(ByVal Node As MSComctlLib.Node)

Dim i As Long, temp2 As tvb_res_Data

    If IsNumeric((Replace(Node.Key, "s0", ""))) Then
        i = Replace(Node.Key, "s0", "")
        If t.Items(i).header.DataType = tvb_res_Image Then
            Picture2.Visible = False
            Text1.Visible = False
            Set picture1.Picture = t.GetImageByIndex(i)
        ElseIf t.Items(i).header.DataType = tvb_res_Text Then
            Picture2.Visible = False
            TextEditable False
            Text1.Visible = True
            Text1.Text = trunString(t.GetTextByIndex(i), 2048)
        Else
            Picture2.Visible = True
            Ri = i
        End If
    Else
        Picture2.Visible = False
        Text1.Visible = True
        Text1.Text = ""
        Dim tmp As Node
        For Each tmp In lstRes.Nodes
            
            If tmp.Parent Is Nothing Then GoTo nextN
            If tmp.Parent = Node Then
                i = (Replace(tmp.Key, "s0", ""))
                TextEditable False
                If t.Items(i).header.DataType = tvb_res_Text Then
                    Text1.Text = Text1.Text & tmp.Text & ":" & """" & trunString(t.GetTextByIndex(i), 2048) & """" & vbNewLine
                ElseIf t.Items(i).header.DataType = tvb_res_Data Then
                    Text1.Text = Text1.Text & tmp.Text & ":" & """" & "Custom Data" & """" & vbNewLine
                ElseIf t.Items(i).header.DataType = tvb_res_Image Then
                    Text1.Text = Text1.Text & tmp.Text & ":" & """" & "Image Data" & """" & vbNewLine
                End If
            End If
nextN:
        Next
    End If

End Sub

Public Function GetSelectedID() As Long

    If lstRes.SelectedItem Is Nothing Then
        GetSelectedID = -1
    Else
    
        If IsNumeric((Replace(lstRes.SelectedItem.Key, "s0", ""))) Then
            GetSelectedID = (Replace(lstRes.SelectedItem.Key, "s0", ""))
        Else
            GetSelectedID = -1
        End If
    End If
    

End Function

Private Function FormatSize(size As Long) As String

    If size < 1024 Then
        FormatSize = size & " bytes"
    ElseIf size < 1024 ^ 2 Then
        FormatSize = Format$(size / 1024, "####.## kb")
    ElseIf size < 1024 ^ 3 Then
        FormatSize = Format$(size / (1024 ^ 2), "####.## mb")
    ElseIf size < 1024 ^ 4 Then
        FormatSize = Format$(size / (1024 ^ 3), "####.## gb")
    End If
    

End Function

Private Sub TextEditable(editable As Boolean)

    Text1.ReadOnly = Not editable
    If editable Then
        Text1.BackColor = &H8000000E
    Else
        Text1.BackColor = &H8000000F
    End If
    

End Sub

Public Sub SetState(State As States)


    Select Case State
        Case States.FileLoaded
            Command1.Enabled = True
            Command2.Enabled = True
            Command3.Enabled = True
            Command4.Enabled = True
            Command8.Enabled = True
            Command6.Enabled = True
            Command9.Enabled = True
            Text1.Enabled = True
            lstRes.Enabled = True
            'image1.Enabled = True
            
        Case States.FileNotSaved
            Command1.Enabled = True
            Command2.Enabled = True
            Command3.Enabled = True
            Command4.Enabled = True
            Command8.Enabled = True
            Command6.Enabled = True
            Command9.Enabled = True
            Text1.Enabled = True
            lstRes.Enabled = True
            'picimage1.Enabled = True
            
        Case States.FileSaved
            Command1.Enabled = True
            Command2.Enabled = True
            Command3.Enabled = True
            Command4.Enabled = True
            Command8.Enabled = True
            Command6.Enabled = False
            Command9.Enabled = True
            Text1.Enabled = True
            lstRes.Enabled = True
            'picimage1.Enabled = True
            
        Case States.NoFileLoaded
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command8.Enabled = False
            Command6.Enabled = False
            Command9.Enabled = True
            Text1.Enabled = False
            lstRes.Enabled = False
            'picimage1.Enabled = False
            
    End Select
    
    lState = State
    
End Sub

Public Function trunString(ByVal stri As String, ByVal maxlen As Long) As String
    
    If Len(stri) > maxlen Then
        trunString = Left$(stri, maxlen - 3) & "..."
    Else
        trunString = stri
    End If
    
End Function
