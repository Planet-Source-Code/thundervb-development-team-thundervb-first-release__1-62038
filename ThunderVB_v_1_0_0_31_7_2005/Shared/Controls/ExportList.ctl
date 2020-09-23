VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.UserControl ExportList 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   Begin MSComctlLib.ListView lvwExp 
      Height          =   2535
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuSelItemsOnly 
         Caption         =   "Selected items only"
      End
   End
End
Attribute VB_Name = "ExportList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Control created , intial version
'22/12/2004 - minor updates for new parsing system..

'22.01. 2005 - several new things were added (all changes are marked with string "PATCH LIBOR")
' new features - "Selected items only" - click on the list (right button on the mouse)
'                menu will appear. The purpose is to show only selected items.
'              - "Hidden members" - if name of module has suffix HIDDEN_MEMBER (look at the contants) then do not add any
'                function from this module to the list. If a name of function has suffix HIDDEN_MEMBER then do not add
'                add it to the list.
'11.02. 2005 - list control was replaced with listview

'PATCH LIBOR
Private Const SEPARATOR As String = "**@@split_@@**"   'splitter of functions names
Private Const HIDDEN_MEMBER As String = "__"           'if a function/module name has this suffix than do not add the function to a list

'PATCH LIBOR
Public Property Get FuncSeparator() As String
10        FuncSeparator = SEPARATOR
End Property

Private Sub lvwExp_ItemClick(ByVal Item As MSComctlLib.ListItem)
10        Item.Checked = Not Item.Checked
End Sub

'PATCH LIBOR
Private Sub lvwExp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If Button = vbRightButton Then PopupMenu mnuMain
End Sub

'PATCH LIBOR
Private Sub mnuSelItemsOnly_Click()
10        mnuSelItemsOnly.Checked = Not mnuSelItemsOnly.Checked
20        SelectedExports = SelectedExports
End Sub

Property Let SelectedExports(data As String)
      Dim temp() As String
          
10        temp = Split(data, SEPARATOR)
20        enumerate GetVBEObject, temp
          
End Property

Private Sub UserControl_Resize()
10        lvwExp.Move 0, 0, UserControl.Width / 15, UserControl.Height / 15
20        lvwExp.ColumnHeaders(1).Width = lvwExp.Width * 0.5
30        lvwExp.ColumnHeaders(2).Width = lvwExp.Width * 0.5 - 1
End Sub

Property Get SelectedExports() As String
      Dim data() As String, i As Long

10        ReDim Preserve data(1)
20        For i = 1 To lvwExp.ListItems.Count
30            If lvwExp.ListItems.Item(i).Checked = True Then
40                data(ArrUBound(data)) = Split(Trim(lvwExp.ListItems(i).Text), " ")(0)
50                ReDim Preserve data(ArrUBound(data) + 1)
60            End If
70        Next
          
80        ReDim Preserve data(ArrUBound(data) - 1)
90        SelectedExports = Join$(data, SEPARATOR)
          
End Property

Sub enumerate(vb As VBIDE.VBE, exports() As String)
      Dim modules() As String
      Dim fns() As String
      Dim i As Long
      Dim cObjC As Long
      Dim cObjM As Long

10        If vb Is Nothing Then Exit Sub
20        If vb.ActiveVBProject Is Nothing Then Exit Sub
30        If vb.ActiveVBProject.VBComponents Is Nothing Then Exit Sub
40        On Error GoTo clr
          
50        DoEvents
60        lvwExp.ListItems.Clear
          
          'enumerate the procedures in every module file within
          'the current project
70        modules = EnumModuleNames(vbext_ct_StdModule)
80        For cObjC = ArrLBound(modules) To ArrUBound(modules)
              
              'if a module should be hidden then do not add it to a list - PATCH LIBOR
90            If StrComp(Right(modules(cObjC), Len(HIDDEN_MEMBER)), HIDDEN_MEMBER, vbTextCompare) = 0 Then GoTo SkipModule
100           fns = EnumFunctionNames(modules(cObjC))
              
110           For cObjM = ArrLBound(fns) To ArrUBound(fns)
                  
                  'if a function should be hidden then do not add it to a list - PATCH LIBOR
120               If StrComp(Right(fns(cObjM), Len(HIDDEN_MEMBER)), HIDDEN_MEMBER, vbTextCompare) = 0 Then GoTo SkipFunction
                  
130               lvwExp.ListItems.Add , , "  " & fns(cObjM)
140               lvwExp.ListItems.Item(lvwExp.ListItems.Count).ListSubItems.Add , , modules(cObjC)
                  'check if the procedure is mardked to be exported.
                  'if so, tick the box next to it.
150               For i = ArrLBound(exports) To ArrUBound(exports)
160                   If exports(i) = fns(cObjM) Then
170                       lvwExp.ListItems(lvwExp.ListItems.Count).Checked = True
180                       Exit For  'PATCH LIBOR
190                   End If
200               Next i
                  
                  'if item is not selected and we want only selected items then remove this item
                  'PATCH - LIBOR
210               If mnuSelItemsOnly.Checked = True And lvwExp.ListItems(lvwExp.ListItems.Count).Checked = False Then
220                   lvwExp.ListItems.Remove lvwExp.ListItems.Count
230               End If
                  
SkipFunction:
240           Next cObjM
SkipModule:
250       Next cObjC

clr:
         
End Sub

Private Sub UserControl_Initialize()
10        With lvwExp
20            .View = lvwReport
30            .Checkboxes = True
40            .LabelEdit = lvwManual
50            .HideSelection = True
60            .FullRowSelect = True
70            .Sorted = False
              
80            .ColumnHeaders.Add , , "Name", .Width * 0.5
90            .ColumnHeaders.Add , , "Defined", .Width * 0.5 - 1
100       End With
End Sub
