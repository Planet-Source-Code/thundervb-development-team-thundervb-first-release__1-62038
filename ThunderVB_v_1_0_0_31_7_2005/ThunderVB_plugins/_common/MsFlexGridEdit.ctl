VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl MsFlexGridEdit 
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   ScaleHeight     =   4080
   ScaleWidth      =   6225
   Begin VB.TextBox T1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid FG1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   30
      Cols            =   10
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FillStyle       =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "MsFlexGridEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Editable FlexGrid
'  designed for ThunderVB by Libor
'
'note : some parts of code are from source code MsFlexGridEdit (written by Peter Raddatz - rabbit@bluecrow.com)

'event - user is going to edit text in a cell
' FlexGrid        - object flexgrid
' CellText        - text in a selected cell
' DefaultEditText - default text in a textbox
' bEdit           - enable/disable editing
' eReason         - event that causes editing
Public Event BeforeTextEdit(ByRef FlexGrid As MSFlexGrid, ByRef CellText As String, ByRef DefaultEditText As String, ByRef bEdit As Boolean, ByVal eReason As eSTART_EDIT)

'event - user finished editing
Public Event AfterTextEdit(ByRef FlexGrid As MSFlexGrid, ByVal CellText As String, ByRef NewCellText As String, ByRef bContinue As Boolean, ByVal eReason As eSTOP_EDIT)
' FlexGrid    - object flexgrid
' CellText    - text in a selected cell
' NewCellText - next text that will be saved in a selected cell
' bContinue   - continue editing cell
' eReason     - event that causes editing

'we will need size of widht of a vertical-scrollbar
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE As Long = -16
Private Const WS_VSCROLL As Long = &H200000

Private Const SM_CXEDGE As Long = 45
Private Const SM_CXVSCROLL As Long = 2

Dim s As String, lScrollBarSize As Long, lBorderSize As Long
Dim sCellText As String, lRow As Long, lCol As Long
Dim sDefaultText As String

Public Enum eSTART_EDIT
    DOUBLE_CLICK          'double click on a cell
    KEY_DOWN_F2           'F2 was pressed
    KEY_DOWN_DELETE       'DELETE was pressed
    KEY_PRESS_RETURN      'RETURN was pressed
    KEY_PRESS_BACKSPACE   'F2 was pressed
    KEY_PRESS_CHAR        'another key was pressed
End Enum

Public Enum eSTOP_EDIT
    MOUSE_DOWN           'user clicks on another cell
    KEY_DOWN_LEFT        'LEFT arrow was pressed
    KEY_DOWN_RIGHT       'RIGHT arrow was pressed
    KEY_DOWN_DOWN        'DOWN arrow was pressed
    KEY_DOWN_UP          'UP arrow was pressed
    KEY_PRESS_RETURN_    'RETURN was pressed
End Enum

'edit text in cell
Private Sub FG1_DblClick()
    
    'check active cell
    If FG1.MouseCol = 0 Or FG1.MouseRow = 0 Or FG1.Rows = 1 Then Exit Sub
    Call Set_TextBox(DOUBLE_CLICK)
    
End Sub

Private Sub FG1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF2           'edit text in the cell
            Set_TextBox KEY_DOWN_F2
        Case vbKeyDelete       'delete text in the cell
            Set_TextBox KEY_DOWN_DELETE, ""
    End Select
    
End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case vbKeyReturn         'edit text in the cell
            Call Set_TextBox(KEY_PRESS_RETURN)
         Case vbKeyBack          'edit text in the cell and delete last char
            If Len(FG1.Text) > 0 Then Call Set_TextBox(KEY_PRESS_BACKSPACE, Left(FG1.Text, Len(FG1.Text) - 1))
        Case Else
            Call Set_TextBox(KEY_PRESS_CHAR, (KeyAscii)) 'edit text in the cell and add the char to the textbox
    End Select
    
End Sub

'this is needed if we click somewhere in FlexGrid where is no cell
Private Sub FG1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If T1.Visible = True Then
        Call RefreshCell(MOUSE_DOWN)
    End If
End Sub

Private Sub T1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape            'restore old text in the cell
            T1.Text = sDefaultText
            T1.SelStart = Len(T1)
        
        Case vbKeyLeft              'finish editing and move to the next cell
            If T1.SelStart = 0 Then
                If RefreshCell(KEY_DOWN_LEFT) = False Then Exit Sub
                If FG1.col > 1 Then
                    FG1.col = FG1.col - 1
                Else
                    If FG1.row > 1 Then
                        FG1.row = FG1.row - 1
                        FG1.col = FG1.Cols - 1
                    End If
                End If
            End If
            
        Case vbKeyUp                'finish editing and move to the next cell
            If RefreshCell(KEY_DOWN_UP) = False Then Exit Sub
            If FG1.row > 1 Then FG1.row = FG1.row - 1

        Case vbKeyRight             'finish editing and move to the next cell
            
            If T1.SelStart = Len(T1) Then
                If RefreshCell(KEY_DOWN_RIGHT) = False Then Exit Sub
                If FG1.col < FG1.Cols - 1 Then
                    FG1.col = FG1.col + 1
                Else
                    If T1.SelStart = Len(T1) And FG1.row < FG1.Rows - 1 Then
                        FG1.row = FG1.row + 1
                        FG1.col = 1
                    End If
                End If
            End If
        
        Case vbKeyDown          'finish editing and move to the next cell
            If RefreshCell(KEY_DOWN_DOWN) = False Then Exit Sub
            If FG1.row < FG1.Rows - 1 Then FG1.row = FG1.row + 1

    End Select
End Sub

Private Sub T1_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case vbKeyReturn         'finish editing
            KeyAscii = 0
            Call RefreshCell(KEY_PRESS_RETURN_)
    End Select
    
End Sub

'start editing
' sNewText - default text in the textbox
Private Sub Set_TextBox(eReason As eSTART_EDIT, Optional sNewText)
Dim bEdit As Boolean, lStyle As Long

    'init
    bEdit = True
    sCellText = FG1.Text
    sDefaultText = IIf(IsMissing(sNewText) = False, sNewText, FG1.Text)
    
    'save cell position
    lRow = FG1.row
    lCol = FG1.col
    
    'event
    RaiseEvent BeforeTextEdit(FG1, sCellText, sDefaultText, bEdit, eReason)
    'do not edit
    If bEdit = False Then Exit Sub
    
    'we will need info about vertical-scrollbar
    lStyle = GetWindowLong(FG1.hwnd, GWL_STYLE)
    'set new text
    FG1.Text = sCellText
    
    With T1
        
        'move textbox ever the cell
        .Top = FG1.Top + FG1.CellTop
        .Left = FG1.Left + FG1.CellLeft
        .Height = FG1.CellHeight
        
        'if v.scrollbar is visible and active column is the last column
        If (FG1.ScrollBars = flexScrollBarBoth Or FG1.ScrollBars = flexScrollBarVertical) And FG1.col = FG1.Cols - 1 And (lStyle And WS_VSCROLL) Then
Dim lColsWidth As Long, lTextBoxWidth As Long, lSpace As Long
            
            'width of all columns
            lColsWidth = FG1.ColPos(FG1.Cols - 1) + FG1.ColWidth(FG1.Cols - 1)
            'is there space between last column and scrollbar?
            lSpace = FG1.Width - (lColsWidth + lScrollBarSize * Screen.TwipsPerPixelX)
            
            'check space
            If lSpace > 0 Then GoTo 10
        
            If FG1.Appearance = flex3D Then
                'adjust textbox width (substract width of v.scrollbar and size of 3D border)
                .Width = FG1.CellWidth - Abs(lSpace) - lBorderSize * Screen.TwipsPerPixelX
            Else
                'adjust textbox width (substract width of v.scrollbar)
                .Width = FG1.CellWidth - Abs(lSpace)
            End If
                
        Else
10:
            .Width = FG1.CellWidth
        End If
           
        'set text
        .Text = sDefaultText
        'move textbox to the foreground
        .ZOrder (0)
        .Visible = True
        .SelStart = Len(.Text)
        .SetFocus
    
    End With
    
End Sub

'show tooltips
Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    With FG1
        'get text of a cell that is under mouse cursor
        s = .TextMatrix(.MouseRow, .MouseCol)
        'if width of cell is smaller than text
        If .ColWidth(.MouseCol) < UserControl.Parent.TextWidth(s) Then
            'set text of tooltip
            .ToolTipText = s
        Else
            'clear tooltip
            .ToolTipText = ""
        End If
    End With
    
End Sub

Private Sub FG1_Scroll()
    'init tooltip
    Call FG1_MouseMove(0, 0, 0, 0)
End Sub

'finish editing
Private Function RefreshCell(eReason As eSTOP_EDIT) As Boolean
Dim sNew As String, bContinue As Boolean, lRow1 As Long, lCol1 As Long
    'get new text
    sNew = T1.Text
    bContinue = False
    
    If lRow <> FG1.row Or lRow <> FG1.col Then
        lRow1 = FG1.row
        lCol1 = FG1.col
        FG1.row = lRow
        FG1.col = lCol
    End If
    
    'event
    RaiseEvent AfterTextEdit(FG1, sCellText, sNew, bContinue, eReason)
    
    If bContinue = True Then
        FG1.col = lCol
        FG1.row = lRow
        'T1.ZOrder (0)
        T1.SetFocus
        RefreshCell = False
        Exit Function
    End If
    
    If lRow1 <> 0 And lCol1 <> 0 Then
        FG1.col = lCol1
        FG1.row = lRow1
    End If
    
    'set new text and hide textbox
    FG1.TextMatrix(lRow, lCol) = sNew
    T1.Visible = False
    FG1.SetFocus
    RefreshCell = True
End Function

Public Function GetFlexGrid() As MSFlexGrid
    Set GetFlexGrid = FG1
End Function

'usercontrol events

Private Sub UserControl_Resize()
    FG1.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_Initialize()
'    get width of scrollbar and 3d border
    lScrollBarSize = GetSystemMetrics(SM_CXVSCROLL)
    lBorderSize = GetSystemMetrics(SM_CXEDGE) * 2
End Sub

