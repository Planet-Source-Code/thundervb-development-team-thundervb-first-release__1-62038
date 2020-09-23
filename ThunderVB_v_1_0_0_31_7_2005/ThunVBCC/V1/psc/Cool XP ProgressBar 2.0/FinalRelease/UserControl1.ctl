VERSION 5.00
Begin VB.UserControl XP_ProgressBar 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ScaleHeight     =   990
   ScaleWidth      =   3000
   ToolboxBitmap   =   "UserControl1.ctx":0000
End
Attribute VB_Name = "XP_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------
'Mario Flores Cool Xp ProgressBar
'Emulating The Windows XP Progress Bar
'Open Source
'6 May 2004
'-----------------------------------------------------------
'Mario Flores Cool Xp ProgressBar 2.0
'MultiStyle ProgressBar
'Open Source
'September 12 2004
'-----------------------------------------------------------

'CD JUAREZ CHIHUAHUA MEXICO

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal fnStyle As Integer, ByVal COLORREF As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


'=====================================================
'TEXT FORMAT CONST
Const DT_SINGLELINE   As Long = &H20
Const DT_CALCRECT     As Long = &H400
'=====================================================

'=====================================================
'BORDER FIELD CONST
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'=====================================================

'=====================================================

'=====================================================

'=====================================================

'=====================================================

'=====================================================
'THE BRUSHSTYLE ENUM
Public Enum BrushStyle
 HS_HORIZONTAL = 0
 HS_VERTICAL = 1
 HS_FDIAGONAL = 2
 HS_BDIAGONAL = 3
 HS_CROSS = 4
 HS_DIAGCROSS = 5
 HS_SOLID = 6
End Enum
'=====================================================

'=====================================================
'THE COOL XP PROGRESSBAR 2.0 STYLES
Public Enum cScrolling
    ccScrollingStandard = 0
    ccScrollingSmooth = 1
    ccScrollingSearch = 2
    ccScrollingOfficeXP = 3
    ccScrollingPastel = 4
    ccScrollingJavT = 5
    ccScrollingMediaPlayer = 6
    ccScrollingCustomBrush = 7
    ccScrollingPicture = 8
    ccScrollingMetallic = 9
End Enum
'=====================================================

'=====================================================
'THE ORIENTATION ENUM
Public Enum cOrientation
    ccOrientationHorizontal = 0
    ccOrientationVertical = 1
End Enum
'=====================================================

'----------------------------------------------------
Private m_Color       As OLE_COLOR
Private m_hDC         As Long
Private m_hWnd        As Long        'PROPERTIES VARIABLES
Private m_Max         As Long
Private m_Min         As Long
Private m_Value       As Long
Private m_ShowText    As Boolean
Private m_Scrolling   As cScrolling
Private m_Orientation As cOrientation
Private m_Brush       As BrushStyle
Private m_Picture     As StdPicture
'----------------------------------------------------

'----------------------------------------------------
Private m_MemDC    As Boolean
Private m_ThDC     As Long
Private m_hBmp     As Long
Private m_hBmpOld  As Long
Private iFnt       As IFont
Private m_fnt      As IFont          'VARIABLES USED IN PROCESS
Private hFntOld    As Long
Private m_lWidth   As Long
Private m_lHeight  As Long
Private fPercent   As Double
Private tR         As RECT
Private TBR        As RECT
Private TSR        As RECT
Private AT         As RECT
Private lSegmentWidth   As Long
Private lSegmentSpacing As Long
'----------------------------------------------------



'==========================================================
'/---Draw ALL ProgressXP Bar  !!!!PUBLIC CALL!!!
'==========================================================

Public Sub DrawProgressBar()
    '<EhHeader>
    On Error GoTo DrawProgressBar_Err
    '</EhHeader>

            
            If m_Value > 100 Then m_Value = 100
            
            
            GetClientRect m_hWnd, tR               '//--- Reference = Control Client Area
              
            DrawFillRectangle tR, IIf(m_Scrolling = ccScrollingMediaPlayer, &H0, vbWhite), m_hDC '//--- Draw BackGround
            
            '//-- Draw ProgressBar Style
            
            '==========================================================
            '/---Draw METALLIC XP STYLE
            '==========================================================

            If m_Scrolling = ccScrollingMetallic Then
                   
                 DrawMetalProgressbar
                    

            '==========================================================
            '/---Draw OFFICE XP STYLE
            '==========================================================

            ElseIf m_Scrolling = ccScrollingOfficeXP Then
                   
                 DrawOfficeXPProgressbar
                    
            '==========================================================
            '/---Draw PASTEL XP STYLE
            '==========================================================

            ElseIf m_Scrolling = ccScrollingPastel Then
                 
                 DrawPastelProgressbar
                 
            '==========================================================
            '/---Draw JAVT XP STYLE
            '==========================================================

            ElseIf m_Scrolling = ccScrollingJavT Then
                 
                 DrawJavTProgressbar
             
            '==========================================================
            '/---Draw MEDIA PLAYER XP STYLE
            '==========================================================
 
            ElseIf m_Scrolling = ccScrollingMediaPlayer Then
            
                 DrawMediaProgressbar
            
            '==========================================================
            '/---Draw CUSTOM BRUSH XP WASH COLOR STYLE
            '==========================================================

            ElseIf m_Scrolling = ccScrollingCustomBrush Then
            
                 DrawCustomBrushProgressbar
             
            '==========================================================
            '/---Draw PICTURE STYLE
            '==========================================================

            ElseIf m_Scrolling = ccScrollingPicture Then
            
                 DrawPictureProgressbar
       
            Else
            
            '==========================================================
            '/---Draw WINDOWS XP STYLE
            '==========================================================

            
                CalcBarSize                            '//--- Calculate Progress and Percent Values
  
                PBarDraw                               '//--- Draw Scolling Bar (Inside Bar)
                  
                If m_Scrolling = 0 Then DrawDivisions  '//--- Draw SegmentSpacing (This Will Generate the Blocks Effect)
  
                pDrawBorder                            '//--- Draw The XP Look Border
            
            End If
            
            '==========================================================
            
            DrawTexto                                  '//--- Draw The Percent Text
            
            '==========================================================
            '/---Use the AntiFlicker DC
            '==========================================================

            If m_MemDC Then
                With UserControl
                    pDraw .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .ScaleLeft, .ScaleTop
                End With
            End If

    '<EhFooter>
    Exit Sub

DrawProgressBar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawProgressBar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'==========================================================
'/---OFFICE XP STYLE
'==========================================================
Private Sub DrawOfficeXPProgressbar()
    '<EhHeader>
    On Error GoTo DrawOfficeXPProgressbar_Err
    '</EhHeader>
        
        DrawRectangle tR, ShiftColorXP(m_Color, 100), m_hDC
             
        With TBR
          .Left = 1
          .Top = 1
          .Bottom = tR.Bottom - 1
          .Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 100)
        End With
             
        DrawFillRectangle TBR, ShiftColorXP(m_Color, 180), m_hDC
 
    '<EhFooter>
    Exit Sub

DrawOfficeXPProgressbar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawOfficeXPProgressbar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'==========================================================
'/---JAVT XP STYLE
'==========================================================
Private Sub DrawJavTProgressbar()
    '<EhHeader>
    On Error GoTo DrawJavTProgressbar_Err
    '</EhHeader>

       DrawRectangle tR, ShiftColorXP(m_Color, 10), m_hDC
       TBR.Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
       DrawGradient m_Color, ShiftColorXP(m_Color, 100), 2, 2, tR.Right - 2, tR.Bottom - 5, m_hDC ', True
       DrawGradient ShiftColorXP(m_Color, 250), m_Color, 3, 3, TBR.Right, tR.Bottom - 6, m_hDC  ', True
       DrawLine TBR.Right, 2, TBR.Right, tR.Bottom - 2, m_hDC, ShiftColorXP(m_Color, 25)
 
    '<EhFooter>
    Exit Sub

DrawJavTProgressbar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawJavTProgressbar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'==========================================================
'/---PICTURE STYLE
'==========================================================
Private Sub DrawPictureProgressbar()
    '<EhHeader>
    On Error GoTo DrawPictureProgressbar_Err
    '</EhHeader>

Dim Brush      As Long
Dim origBrush  As Long

       DrawEdge m_hDC, tR, 2, BF_RECT                       '//--- Draw ProgressBar Border
       
       If Nothing Is m_Picture Then Exit Sub                '//--- In Case No Picture is Choosen
              
       Brush = CreatePatternBrush(m_Picture.handle)         '//-- Use Pattern Picture Draw
       origBrush = SelectObject(m_hDC, Brush)
       TBR.Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
       
       PatBlt m_hDC, 2, 2, TBR.Right, tR.Bottom - 4, vbPatCopy
         
       SelectObject m_hDC, origBrush
       DeleteObject Brush
       
    '<EhFooter>
    Exit Sub

DrawPictureProgressbar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawPictureProgressbar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'==========================================================
'/---PASTEL XP STYLE
'==========================================================
Private Sub DrawPastelProgressbar()
    '<EhHeader>
    On Error GoTo DrawPastelProgressbar_Err
    '</EhHeader>
        DrawEdge m_hDC, tR, 6, BF_RECT
        DrawGradient ShiftColorXP(m_Color, 140), ShiftColorXP(m_Color, 200), 2, 2, tR.Left + (tR.Right - tR.Left - 4) * (m_Value / 100), tR.Bottom - 3, m_hDC, True
    '<EhFooter>
    Exit Sub

DrawPastelProgressbar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawPastelProgressbar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'==========================================================
'/---METALLIC XP STYLE
'==========================================================
Private Sub DrawMetalProgressbar()
    '<EhHeader>
    On Error GoTo DrawMetalProgressbar_Err
    '</EhHeader>
         TBR.Right = tR.Left + (tR.Right - tR.Left - 4) * (m_Value / 100)
         
         DrawGradient vbWhite, &HC0C0C0, 2, 2, tR.Right - 3, (tR.Bottom - 3) / 2, m_hDC
         DrawGradient BlendColor(&HC0C0C0, &H0, 255), &HC0C0C0, 2, (tR.Bottom - 3) / 2, tR.Right - 3, (tR.Bottom - 3) / 2, m_hDC
         DrawGradient ShiftColorXP(m_Color, 150), BlendColor(m_Color, &H0, 180), 2, 2, TBR.Right, (tR.Bottom - 3) / 2, m_hDC
         DrawGradient BlendColor(m_Color, &H0, 190), m_Color, 2, (tR.Bottom - 3) / 2, TBR.Right, (tR.Bottom - 3) / 2, m_hDC
        
         tR.Left = tR.Left + 3
         pDrawBorder
    
        
    '<EhFooter>
    Exit Sub

DrawMetalProgressbar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawMetalProgressbar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'==========================================================
'/---CUSTOM BRUSH XP STYLE
'==========================================================
Private Sub DrawCustomBrushProgressbar()
    '<EhHeader>
    On Error GoTo DrawCustomBrushProgressbar_Err
    '</EhHeader>
        
   Dim hBrush As Long
    
   DrawEdge m_hDC, tR, 9, BF_RECT
       
   With TBR
      .Left = 2
      .Top = 2
      .Bottom = tR.Bottom - 2
      .Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
   End With

   hBrush = CreateHatchBrush(m_Brush, GetLngColor(Color))
   SetBkColor m_hDC, ShiftColorXP(m_Color, 140)
   FillRect m_hDC, TBR, hBrush
   DeleteObject hBrush
                
    '<EhFooter>
    Exit Sub

DrawCustomBrushProgressbar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawCustomBrushProgressbar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'==========================================================
'/---MEDIA PROGRESS XP STYLE
'==========================================================
Private Sub DrawMediaProgressbar()
    '<EhHeader>
    On Error GoTo DrawMediaProgressbar_Err
    '</EhHeader>
        
        DrawRectangle tR, BlendColor(m_Color, &H0, 200), m_hDC
        DrawGradient &H0&, ShiftColorXP(GetLngColor(BlendColor(m_Color, &H0, 100)), 10), 2, 2, tR.Left + (tR.Right - tR.Left - 5) * (m_Value / 100), tR.Bottom - 2, m_hDC, True

    '<EhFooter>
    Exit Sub

DrawMediaProgressbar_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawMediaProgressbar " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'==========================================================
'/---Calculate Division Bars & Percent Values
'==========================================================

Private Sub CalcBarSize()
    '<EhHeader>
    On Error GoTo CalcBarSize_Err
    '</EhHeader>

      lSegmentWidth = IIf(m_Scrolling = 0, 6, 0) '/-- Windows Default
      lSegmentSpacing = 2                        '/-- Windows Default
            
      tR.Left = tR.Left + 3
   
      LSet TBR = tR

      fPercent = m_Value / 98
        
      If fPercent < 0# Then fPercent = 0#
   
      If m_Orientation = 0 Then
      
      '=======================================================================================
      '                                 Calc Horizontal ProgressBar
      '---------------------------------------------------------------------------------------
         
         TBR.Right = tR.Left + (tR.Right - tR.Left) * fPercent
         
         TBR.Right = TBR.Right - ((TBR.Right - TBR.Left) Mod (lSegmentWidth + lSegmentSpacing))
         
         If TBR.Right < tR.Left Then
            TBR.Right = tR.Left
         End If
                  
      Else
      
      '=======================================================================================
      '                                 Calc Vertical ProgressBar
      '---------------------------------------------------------------------------------------
         fPercent = 1# - fPercent
         TBR.Top = tR.Top + (tR.Bottom - tR.Top) * fPercent
         TBR.Top = TBR.Top - ((TBR.Top - TBR.Bottom) Mod (lSegmentWidth + lSegmentSpacing))
         If TBR.Top > tR.Bottom Then TBR.Top = tR.Bottom
    
         
      
      End If

    '<EhFooter>
    Exit Sub

CalcBarSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.CalcBarSize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'==========================================================
'/---Draw Division Bars
'==========================================================

Private Sub DrawDivisions()
    '<EhHeader>
    On Error GoTo DrawDivisions_Err
    '</EhHeader>
 Dim i As Long
 Dim hBr As Long
  
  hBr = CreateSolidBrush(vbWhite)
  
      LSet TSR = tR
      
       
      If m_Orientation = 0 Then
      
      
      '=======================================================================================
      '                                 Draw Horizontal ProgressBar
      '---------------------------------------------------------------------------------------
         For i = TBR.Left + lSegmentWidth To TBR.Right Step lSegmentWidth + lSegmentSpacing
            TSR.Left = i + 1
            TSR.Right = i + 1 + lSegmentSpacing
            FillRect m_hDC, TSR, hBr
         Next i
      '---------------------------------------------------------------------------------------
      
      Else
      
      '=======================================================================================
      '                                  Draw Vertical ProgressBar
      '---------------------------------------------------------------------------------------
         For i = TBR.Bottom To TBR.Top + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
            TSR.Top = i - 2
            TSR.Bottom = i - 2 + lSegmentSpacing
            FillRect m_hDC, TSR, hBr
         Next i
       '---------------------------------------------------------------------------------------
      
      End If
      
      DeleteObject hBr
     
    '<EhFooter>
    Exit Sub

DrawDivisions_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawDivisions " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'==========================================================
'/---Draw The ProgressXP Bar Border  ;)
'==========================================================

Private Sub pDrawBorder()
    '<EhHeader>
    On Error GoTo pDrawBorder_Err
    '</EhHeader>
Dim RTemp As RECT
 
 tR.Left = tR.Left - 3
 
 Let RTemp = tR
  
 
 DrawLine 2, 1, tR.Right - 2, 1, m_hDC, &HBEBEBE
 DrawLine 2, tR.Bottom - 2, tR.Right - 2, tR.Bottom - 2, m_hDC, &HEFEFEF
 DrawLine 1, 2, 1, tR.Bottom - 2, m_hDC, &HBEBEBE
 DrawLine 2, 2, 2, tR.Bottom - 2, m_hDC, &HEFEFEF
 DrawLine 2, 2, tR.Right - 2, 2, m_hDC, &HEFEFEF
 DrawLine tR.Right - 2, 2, tR.Right - 2, tR.Bottom - 2, m_hDC, &HEFEFEF
  
 DrawRectangle tR, GetLngColor(&H686868), m_hDC

 
 Call SetPixelV(m_hDC, 0, 0, GetLngColor(vbWhite))
 Call SetPixelV(m_hDC, 0, 1, GetLngColor(&HA6ABAC))
 Call SetPixelV(m_hDC, 0, 2, GetLngColor(&H7D7E7F))
 Call SetPixelV(m_hDC, 1, 0, GetLngColor(&HA7ABAC)) '//TOP RIGHT CORNER
 Call SetPixelV(m_hDC, 1, 1, GetLngColor(&H777777))
 Call SetPixelV(m_hDC, 2, 0, GetLngColor(&H7D7E7F))
 Call SetPixelV(m_hDC, 2, 2, GetLngColor(&HBEBEBE))
   
 Call SetPixelV(m_hDC, 0, tR.Bottom - 1, GetLngColor(vbWhite))
 Call SetPixelV(m_hDC, 1, tR.Bottom - 1, GetLngColor(&HA6ABAC))
 Call SetPixelV(m_hDC, 2, tR.Bottom - 1, GetLngColor(&H7D7E7F))
 Call SetPixelV(m_hDC, 0, tR.Bottom - 3, GetLngColor(&H7D7E7F)) '//BOTTOM RIGHT CORNER
 Call SetPixelV(m_hDC, 0, tR.Bottom - 2, GetLngColor(&HA7ABAC))
 Call SetPixelV(m_hDC, 1, tR.Bottom - 2, GetLngColor(&H777777))
 
 Call SetPixelV(m_hDC, tR.Right - 1, 0, GetLngColor(vbWhite))
 Call SetPixelV(m_hDC, tR.Right - 1, 1, GetLngColor(&HBEBEBE))
 Call SetPixelV(m_hDC, tR.Right - 1, 2, GetLngColor(&H7D7E7F)) '//TOP LEFT CORNER
 Call SetPixelV(m_hDC, tR.Right - 2, 2, GetLngColor(&HBEBEBE))
 Call SetPixelV(m_hDC, tR.Right - 2, 1, GetLngColor(&H686868))
 
 Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 1, GetLngColor(vbWhite))
 Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 2, GetLngColor(&HBEBEBE))
 Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 3, GetLngColor(&H7D7E7F))
 Call SetPixelV(m_hDC, tR.Right - 2, tR.Bottom - 2, GetLngColor(&H777777)) '//TOP RIGHT CORNER
 Call SetPixelV(m_hDC, tR.Right - 2, tR.Bottom - 1, GetLngColor(&HBEBEBE))
 Call SetPixelV(m_hDC, tR.Right - 3, tR.Bottom - 1, GetLngColor(&H7D7E7F))

 
    '<EhFooter>
    Exit Sub

pDrawBorder_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.pDrawBorder " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'==========================================================
'/---Draw The ProgressXP Bar ;)
'==========================================================

Private Sub PBarDraw()
    '<EhHeader>
    On Error GoTo PBarDraw_Err
    '</EhHeader>
Dim TempRect As RECT
Dim ITemp    As Long

If m_Orientation = 0 Then

    If TBR.Right <= 14 Then TBR.Right = 12
        
    TempRect.Left = 4
    TempRect.Right = IIf(TBR.Right + 4 > tR.Right, TBR.Right - 4, TBR.Right)
    TempRect.Top = 8
    TempRect.Bottom = tR.Bottom - 8

    '=======================================================================================
    '                                 Draw Horizontal ProgressBar
    '---------------------------------------------------------------------------------------
   
         
     If m_Scrolling = ccScrollingSearch Then
         GoSub HorizontalSearch
     Else
        DrawGradient ShiftColorXP(m_Color, 150), m_Color, 4, 3, TempRect.Right, 6, m_hDC
        DrawFillRectangle TempRect, m_Color, m_hDC
        DrawGradient m_Color, ShiftColorXP(m_Color, 150), 4, TempRect.Bottom - 2, TempRect.Right, 6, m_hDC
     End If
Else
    
    TempRect.Left = 9
    TempRect.Right = tR.Right - 8
    TempRect.Top = TBR.Top
    TempRect.Bottom = tR.Bottom
    
    '=======================================================================================
    '                                 Draw Vertical ProgressBar
    '---------------------------------------------------------------------------------------
   
    If m_Scrolling = ccScrollingSearch Then
         GoSub VerticalSearch
    Else
        DrawGradient ShiftColorXP(m_Color, 150), m_Color, 4, TBR.Top, 4, tR.Bottom, m_hDC, True
        DrawFillRectangle TempRect, m_Color, m_hDC
        DrawGradient m_Color, ShiftColorXP(m_Color, 150), tR.Right - 8, TBR.Top, 4, tR.Bottom, m_hDC, True
    End If
   
    '--------------------   <-------- Gradient Color From (- to +)
    '||||||||||||||||||||   <-------- Fill Color
    '--------------------   <-------- Gradient Color From (+ to -)

End If

Exit Sub

HorizontalSearch:
    
    
    For ITemp = 0 To 2
    
        With TempRect
          .Left = TBR.Right + ((lSegmentSpacing + 10) * (ITemp)) - (45 * ((100 - m_Value) / 100))
          .Right = .Left + 10
          .Top = 8
          .Bottom = tR.Bottom - 8
          DrawGradient ShiftColorXP(m_Color, 220 - (40 * ITemp)), ShiftColorXP(m_Color, 200 - (40 * ITemp)), .Left, 3, 9, tR.Bottom - 2, m_hDC, True
        End With
        
    Next ITemp

Return

VerticalSearch:
    
     
    For ITemp = 0 To 2
    
        With TempRect
          .Left = 8
          .Right = tR.Right - 8
          .Top = TBR.Top + ((lSegmentSpacing + 10) * ITemp)
          .Bottom = .Top + 10
          DrawGradient ShiftColorXP(m_Color, 220 - (40 * ITemp)), ShiftColorXP(m_Color, 200 - (40 * ITemp)), tR.Right - 2, .Top, 2, 9, m_hDC
        End With
        
    Next ITemp

Return

    '<EhFooter>
    Exit Sub

PBarDraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.PBarDraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'======================================================================
'DRAWS THE PERCENT TEXT ON PROGRESS BAR
Private Function DrawTexto()
    '<EhHeader>
    On Error GoTo DrawTexto_Err
    '</EhHeader>
Dim ThisText As String
Dim isAlpha  As Boolean

If (m_Scrolling = ccScrollingMediaPlayer Or m_Scrolling = ccScrollingMetallic) Then isAlpha = True


 If m_Scrolling = ccScrollingSearch Then
    ThisText = "Searching.."
 Else
    ThisText = Round(m_Value) & " %"
 End If

 If (m_ShowText) Then
           
      Set iFnt = Font                             '//--New Font
      hFntOld = SelectObject(m_hDC, iFnt.hFont)   '//--Use the New Font
      SetBkMode m_hDC, 1                          '//--Transparent Text
     
      '//--Use the Alpha Text Color Look if Progress is MediaPlayer Style, else Normal (Gray)
      SetTextColor m_hDC, GetLngColor(IIf(m_Scrolling = ccScrollingMediaPlayer, &HC0C0C0, vbBlack))
      
      CalculateAlphaTextRect ThisText             '//--Calculate The Text Rectangle
           
      '//-- If ProgressBar is already over the Text don't draw the old text, yust draw the Alpha Text
           'It saves some memory
      
      If ((tR.Right * (m_Value / 100)) <= AT.Right) Or Not isAlpha Then
            W_DrawText m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE
      End If
            
      SelectObject m_hDC, hFntOld  'Delete the Used Font
   
      '//--Use the Alpha Text Look if Progress is AlPhA Style
      If isAlpha Then DrawAlphaText ThisText
              
 End If


    '<EhFooter>
    Exit Function

DrawTexto_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawTexto " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
'======================================================================

'======================================================================
'ALPHA TEXT RECT FUNCTION
Private Sub CalculateAlphaTextRect(ByVal ThisText As String)
    '<EhHeader>
    On Error GoTo CalculateAlphaTextRect_Err
    '</EhHeader>

      '//--Calculates the Bounding Rects Of the Text using DT_CALCRECT
      W_DrawText m_hDC, ThisText, Len(ThisText), AT, DT_CALCRECT
      AT.Left = (tR.Right / 2) - ((AT.Right - AT.Left) / 2)
      AT.Top = (tR.Bottom / 2) - ((AT.Bottom - AT.Top) / 2)

    '<EhFooter>
    Exit Sub

CalculateAlphaTextRect_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.CalculateAlphaTextRect " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'ALPHA TEXT FUNCTION
Private Sub DrawAlphaText(ByVal ThisText As String)
    '<EhHeader>
    On Error GoTo DrawAlphaText_Err
    '</EhHeader>

 Set iFnt = Font                             '//--New Font
 hFntOld = SelectObject(m_hDC, iFnt.hFont)   '//--Use the New Font
 SetBkMode m_hDC, 1                          '//--Transparent Text
        
        
        '//-- This is When the Text is Drawn
            '//--Gives the Media Player Text Look (Changes Color When Progress is over the Text)
            
            If (tR.Right * (m_Value / 100)) >= AT.Left Then
                SetTextColor m_hDC, GetLngColor(IIf(m_Scrolling = ccScrollingMediaPlayer, ShiftColorXP(m_Color, 80), vbWhite))
                AT.Left = (tR.Right / 2) - ((AT.Right - AT.Left) / 2)
                AT.Right = (tR.Right * (m_Value / 100))
                W_DrawText m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE
                SelectObject m_hDC, hFntOld
            End If

    '<EhFooter>
    Exit Sub

DrawAlphaText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawAlphaText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long
    '<EhHeader>
    On Error GoTo GetLngColor_Err
    '</EhHeader>
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
    '<EhFooter>
    Exit Function

GetLngColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.GetLngColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
'======================================================================

'======================================================================
'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawRectangle(ByRef bRECT As RECT, ByVal Color As Long, ByVal hdc As Long)
    '<EhHeader>
    On Error GoTo DrawRectangle_Err
    '</EhHeader>

Dim hBrush As Long
    
    hBrush = CreateSolidBrush(Color)
    FrameRect hdc, bRECT, hBrush
    DeleteObject hBrush

    '<EhFooter>
    Exit Sub

DrawRectangle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawRectangle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'DRAWS A LINE WITH A DEFINED COLOR
Public Sub DrawLine( _
           ByVal X As Long, _
           ByVal Y As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal cHdc As Long, _
           ByVal Color As Long)
    '<EhHeader>
    On Error GoTo DrawLine_Err
    '</EhHeader>

    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim Outline As Long
    Dim POS     As POINTAPI

    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    
        MoveToEx cHdc, X, Y, POS
        LineTo cHdc, Width, Height
          
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1

    '<EhFooter>
    Exit Sub

DrawLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'BLENDS AN SPECIFIED COLOR TO GET XP COLOR LOOK
Private Function ShiftColorXP(ByVal MyColor As Long, ByVal Base As Long) As Long
    '<EhHeader>
    On Error GoTo ShiftColorXP_Err
    '</EhHeader>

    Dim R As Long, G As Long, B As Long, Delta As Long

    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    B = ((MyColor \ &H10000) Mod &H100)
    
    Delta = &HFF - Base

    B = Base + B * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF

    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255

    ShiftColorXP = R + 256& * G + 65536 * B

    '<EhFooter>
    Exit Function

ShiftColorXP_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.ShiftColorXP " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
'======================================================================

'======================================================================
'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
Public Sub DrawGradient(lEndColor As Long, lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hdc As Long, Optional bH As Boolean)
    '<EhHeader>
    On Error GoTo DrawGradient_Err
    '</EhHeader>
    On Error Resume Next
    
    ''Draw a Vertical Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    lEndColor = GetLngColor(lEndColor)
    lStartColor = GetLngColor(lStartColor)

    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    sR = (sR - eR) / IIf(bH, X2, Y2)
    sG = (sG - eG) / IIf(bH, X2, Y2)
    sB = (sB - eB) / IIf(bH, X2, Y2)
    
        
    For ni = 0 To IIf(bH, X2, Y2)
        
        If bH Then
            DrawLine X + ni, Y, X + ni, Y2, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Else
            DrawLine X, Y + ni, X2, Y + ni, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        End If
        
    Next ni
    '<EhFooter>
    Exit Sub

DrawGradient_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawGradient " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'BLENDS 2 COLORS WITH A PREDEFINED ALPHA VALUE
Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
    '<EhHeader>
    On Error GoTo BlendColor_Err
    '</EhHeader>
Dim lCFrom As Long
Dim lCTo As Long
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   
   lCFrom = GetLngColor(oColorFrom)
   lCTo = GetLngColor(oColorTo)
   
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
   
   BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
      )
      
    '<EhFooter>
    Exit Function

BlendColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.BlendColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
'======================================================================

'======================================================================
'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)
    '<EhHeader>
    On Error GoTo DrawFillRectangle_Err
    '</EhHeader>

Dim hBrush As Long
 
   hBrush = CreateSolidBrush(GetLngColor(Color))
   FillRect MyHdc, hRect, hBrush
   DeleteObject hBrush

    '<EhFooter>
    Exit Sub

DrawFillRectangle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.DrawFillRectangle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'CHECKS-CREATES CORRECT DIMENSIONS OF THE TEMP DC
Private Function ThDC(Width As Long, Height As Long) As Long
    '<EhHeader>
    On Error GoTo ThDC_Err
    '</EhHeader>
   If m_ThDC = 0 Then
      If (Width > 0) And (Height > 0) Then
         pCreate Width, Height
      End If
   Else
      If Width > m_lWidth Or Height > m_lHeight Then
         pCreate Width, Height
      End If
   End If
   ThDC = m_ThDC
    '<EhFooter>
    Exit Function

ThDC_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.ThDC " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function
'======================================================================

'======================================================================
'CREATES THE TEMP DC
Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
    '<EhHeader>
    On Error GoTo pCreate_Err
    '</EhHeader>
Dim lhDCC As Long
   pDestroy
   lhDCC = W_CreateDC("DISPLAY", "", "", ByVal 0&)
   If Not (lhDCC = 0) Then
      m_ThDC = CreateCompatibleDC(lhDCC)
      If Not (m_ThDC = 0) Then
         m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
         If Not (m_hBmp = 0) Then
            m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
            If Not (m_hBmpOld = 0) Then
               m_lWidth = Width
               m_lHeight = Height
               DeleteDC lhDCC
               Exit Sub
            End If
         End If
      End If
      DeleteDC lhDCC
      pDestroy
   End If
    '<EhFooter>
    Exit Sub

pCreate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.pCreate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'DRAWS THE TEMP DC
Public Sub pDraw( _
      ByVal hdc As Long, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
    '<EhHeader>
    On Error GoTo pDraw_Err
    '</EhHeader>
   If WidthSrc <= 0 Then WidthSrc = m_lWidth
   If HeightSrc <= 0 Then HeightSrc = m_lHeight
   BitBlt hdc, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, xSrc, ySrc, vbSrcCopy

    '<EhFooter>
    Exit Sub

pDraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.pDraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================

'======================================================================
'DESTROYS THE TEMP DC
Private Sub pDestroy()
    '<EhHeader>
    On Error GoTo pDestroy_Err
    '</EhHeader>
   If Not m_hBmpOld = 0 Then
      SelectObject m_ThDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If Not m_ThDC = 0 Then
      DeleteDC m_ThDC
      m_ThDC = 0
   End If
   m_lWidth = 0
   m_lHeight = 0
    '<EhFooter>
    Exit Sub

pDestroy_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.pDestroy " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'======================================================================



'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================
'USER CONTROL EVENTS
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================


Private Sub UserControl_Initialize()
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>

     Dim fnt As New StdFont
         Set Font = fnt

     With UserControl
        .BackColor = vbWhite
        .ScaleMode = vbPixels
     End With
     
     '----------------------------------------------------------
     'Default Values
     hdc = UserControl.hdc
     hwnd = UserControl.hwnd
     m_Max = 100
     m_Min = 0
     m_Value = 0
     m_Orientation = ccOrientationHorizontal
     m_Scrolling = ccScrollingStandard
     m_Color = GetLngColor(vbHighlight)
     DrawProgressBar
     '----------------------------------------------------------

    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.UserControl_Initialize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Paint()
    '<EhHeader>
    On Error GoTo UserControl_Paint_Err
    '</EhHeader>
 DrawProgressBar
    '<EhFooter>
    Exit Sub

UserControl_Paint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.UserControl_Paint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error GoTo UserControl_Resize_Err
    '</EhHeader>
hdc = UserControl.hdc
    '<EhFooter>
    Exit Sub

UserControl_Resize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.UserControl_Resize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error GoTo UserControl_Terminate_Err
    '</EhHeader>
 pDestroy 'Destroy Temp DC
    '<EhFooter>
    Exit Sub

UserControl_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.UserControl_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================
'USER CONTROL PROPERTIES
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================

Public Property Let BrushStyle(ByVal Style As BrushStyle)
   m_Brush = Style
   PropertyChanged "BrushStyle"
End Property

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/sets the color of the ProgressBar"
    '<EhHeader>
    On Error GoTo Color_Err
    '</EhHeader>
   Color = m_Color
    '<EhFooter>
    Exit Property

Color_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Color " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Color(ByVal lColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo Color_Err
    '</EhHeader>
   m_Color = GetLngColor(lColor)
   DrawProgressBar
    '<EhFooter>
    Exit Property

Color_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Color " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Font() As IFont
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
   Set Font = m_fnt
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set Font(ByRef fnt As IFont)
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
   Set m_fnt = fnt    'Defined By System but can change by user choice.(ADD Property!!)
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Font(ByRef fnt As IFont)
    '<EhHeader>
    On Error GoTo Font_Err
    '</EhHeader>
   Set m_fnt = fnt
    '<EhFooter>
    Exit Property

Font_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Font " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get hwnd() As Long
    '<EhHeader>
    On Error GoTo hwnd_Err
    '</EhHeader>
   hwnd = m_hWnd
    '<EhFooter>
    Exit Property

hwnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.hwnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let hwnd(ByVal chWnd As Long)
    '<EhHeader>
    On Error GoTo hwnd_Err
    '</EhHeader>
   m_hWnd = chWnd
    '<EhFooter>
    Exit Property

hwnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.hwnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get hdc() As Long
    '<EhHeader>
    On Error GoTo hdc_Err
    '</EhHeader>
   hdc = m_hDC
    '<EhFooter>
    Exit Property

hdc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.hdc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let hdc(ByVal cHdc As Long)
     '=============================================
   'AntiFlick...Cleaner HDC
    '<EhHeader>
    On Error GoTo hdc_Err
    '</EhHeader>
   m_hDC = ThDC(UserControl.ScaleWidth, UserControl.ScaleHeight)
   
   If m_hDC = 0 Then
      m_hDC = UserControl.hdc   'On Fail...Do it Normally
   Else
      m_MemDC = True
   End If
   '=============================================

    '<EhFooter>
    Exit Property

hdc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.hdc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Image() As StdPicture
    '<EhHeader>
    On Error GoTo Image_Err
    '</EhHeader>
    If Nothing Is m_Picture Then Exit Property
    Set Image = m_Picture
    '<EhFooter>
    Exit Property

Image_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Image " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set Image(ByVal handle As StdPicture)
    '<EhHeader>
    On Error GoTo Image_Err
    '</EhHeader>
   Set m_Picture = handle
   PropertyChanged "Image"
   DrawProgressBar
    '<EhFooter>
    Exit Property

Image_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Image " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Min() As Long
    '<EhHeader>
    On Error GoTo Min_Err
    '</EhHeader>
   Min = m_Min
    '<EhFooter>
    Exit Property

Min_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Min " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Min(ByVal cMin As Long)
    '<EhHeader>
    On Error GoTo Min_Err
    '</EhHeader>
   m_Min = cMin
   PropertyChanged "Min"
    '<EhFooter>
    Exit Property

Min_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Min " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Max() As Long
    '<EhHeader>
    On Error GoTo Max_Err
    '</EhHeader>
   Max = m_Max
    '<EhFooter>
    Exit Property

Max_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Max " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Max(ByVal cMax As Long)
    '<EhHeader>
    On Error GoTo Max_Err
    '</EhHeader>
   m_Max = cMax
   PropertyChanged "Max"
    '<EhFooter>
    Exit Property

Max_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Max " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Orientation() As cOrientation
    '<EhHeader>
    On Error GoTo Orientation_Err
    '</EhHeader>
   Orientation = m_Orientation
    '<EhFooter>
    Exit Property

Orientation_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Orientation " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Orientation(ByVal cOrientation As cOrientation)
    '<EhHeader>
    On Error GoTo Orientation_Err
    '</EhHeader>
   m_Orientation = cOrientation
   PropertyChanged "Orientation"
   DrawProgressBar
    '<EhFooter>
    Exit Property

Orientation_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Orientation " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Scrolling() As cScrolling
    '<EhHeader>
    On Error GoTo Scrolling_Err
    '</EhHeader>
   Scrolling = m_Scrolling
    '<EhFooter>
    Exit Property

Scrolling_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Scrolling " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Scrolling(ByVal lScrolling As cScrolling)
    '<EhHeader>
    On Error GoTo Scrolling_Err
    '</EhHeader>
   m_Scrolling = lScrolling
   PropertyChanged "Scrolling"
   DrawProgressBar
    '<EhFooter>
    Exit Property

Scrolling_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Scrolling " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ShowText() As Boolean
    '<EhHeader>
    On Error GoTo ShowText_Err
    '</EhHeader>
   ShowText = m_ShowText
    '<EhFooter>
    Exit Property

ShowText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.ShowText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ShowText(ByVal bShowText As Boolean)
    '<EhHeader>
    On Error GoTo ShowText_Err
    '</EhHeader>
   m_ShowText = bShowText
   PropertyChanged "ShowText"
   DrawProgressBar
    '<EhFooter>
    Exit Property

ShowText_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.ShowText " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Value() As Long
    '<EhHeader>
    On Error GoTo Value_Err
    '</EhHeader>
   Value = ((m_Value / 100) * m_Max) / IIf(m_Min > 0, m_Min, 1)
    '<EhFooter>
    Exit Property

Value_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Value " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Value(ByVal cValue As Long)
    '<EhHeader>
    On Error GoTo Value_Err
    '</EhHeader>
    m_Value = ((cValue * 100) / m_Max) + m_Min
    'PropertyChanged "Value"
    DrawProgressBar
    '<EhFooter>
    Exit Property

Value_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.Value " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'=======================================================================================================================
' USERCONTROL WRITE PROPERTIES
'=======================================================================================================================

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>
 Call PropBag.WriteProperty("Font", Font)
 Call PropBag.WriteProperty("BrushStyle", m_Brush, 4)
 Call PropBag.WriteProperty("Color", m_Color, vbHighlight)
 Call PropBag.WriteProperty("Image", m_Picture, Nothing)
 Call PropBag.WriteProperty("Max", m_Max, 100)
 Call PropBag.WriteProperty("Min", m_Min, 0)
 Call PropBag.WriteProperty("Orientation", m_Orientation, ccOrientationHorizontal)
 Call PropBag.WriteProperty("Scrolling", m_Scrolling, ccScrollingStandard)
 Call PropBag.WriteProperty("ShowText", m_ShowText, False)
 Call PropBag.WriteProperty("Value", m_Value, 0)
    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.UserControl_WriteProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
 End Sub

'=======================================================================================================================
' USERCONTROL READ PROPERTIES
'=======================================================================================================================

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
Set Font = PropBag.ReadProperty("Font")
m_Brush = PropBag.ReadProperty("BrushStyle", 4)
Color = PropBag.ReadProperty("Color", vbHighlight)
Set m_Picture = PropBag.ReadProperty("Image", Nothing)
Max = PropBag.ReadProperty("Max", 100)
Min = PropBag.ReadProperty("Min", 0)
Orientation = PropBag.ReadProperty("Orientation", ccOrientationHorizontal)
Scrolling = PropBag.ReadProperty("Scrolling", ccScrollingStandard)
ShowText = PropBag.ReadProperty("ShowText", False)
Value = PropBag.ReadProperty("Value", 0)
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XP_ProgressBar.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

