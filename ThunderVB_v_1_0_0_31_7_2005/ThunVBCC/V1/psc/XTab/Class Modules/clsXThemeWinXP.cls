VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xThemeWinXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'xThemeWinXP : Theme mimics stabdard Windows XP tabs. [ XP's HOT ;) ]
'              Can be customized
'
'              See ITheme for implemented function description

Option Explicit

Implements ITheme


'=====Constants=======================================================================================================
Private Const m_c_iFOCUS_RECT_AND_TEXT_DISTANCE = 2  'the distance between the text (caption) of the tab and the focus Rect
Private Const m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE As Integer = 7   'the distance between the text and the border in a Property Pages style tab
Private Const m_c_iPROP_PAGE_INACTIVE_TOP As Integer = 2      'the top for the property page (inactive property page)
'=====================================================================================================================


'=====Private Variables===============================================================================================

Private m_oXTab As XTab


'following property are cached (locally) for improved performance
Private m_lHwnd As Long
Private m_lHdc As Long
Private m_iActiveTab As Integer
Private m_iActiveTabHeight As Integer
Private m_iInActiveTabHeight As Integer
Private m_iTabCount As Integer
Private m_bIsFocused As Boolean
Private m_iScaleWidth As Integer
Private m_iScaleHeight As Integer
Private m_lOuterBorderColor As Long
Private m_lActiveTabForeColor As Long
Private m_lActiveTabBackStartColor As Long
Private m_lActiveTabBackEndColor As Long
Private m_lInActiveTabForeColor As Long
Private m_lInActiveTabBackStartColor As Long
Private m_lInActiveTabBackEndColor As Long
Private m_lDisabledTabForeColor As Long
Private m_lDisabledTabBackColor As Long
Private m_lHoverColorInverted As Long
Private m_iIconSize As Integer

Private m_bIsMouseOver As Boolean     'used to track the mouse

Private m_iMouseOverTabIndex As Integer   'stores the index of tab on which mouse is being moved
Private m_utRect As RECT      'used to prevent the trip to fetch the control properties
Private m_Pnt As POINTAPI


'=====================================================================================================================


'====Event Handlers===================================================================================================

Private Sub ITheme_DrawBackground()
    '<EhHeader>
    On Error GoTo ITheme_DrawBackground_Err
    '</EhHeader>
  Call pCacheControlProperties     'cache the control proerties for faster access
  Select Case m_oXTab.TabStyle
    Case xStyleTabbedDialog:
      Call ITheme_DrawBackgroundTabbedDialog
    Case xStylePropertyPages:
      Call ITheme_DrawBackgroundPropertyPages
  End Select
    '<EhFooter>
    Exit Sub

ITheme_DrawBackground_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_DrawBackground " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawBackgroundPropertyPages()
    '<EhHeader>
    On Error GoTo ITheme_DrawBackgroundPropertyPages_Err
    '</EhHeader>
  Call ITheme_DrawBackgroundTabbedDialog
    '<EhFooter>
    Exit Sub

ITheme_DrawBackgroundPropertyPages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_DrawBackgroundPropertyPages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawBackgroundTabbedDialog()
    '<EhHeader>
    On Error GoTo ITheme_DrawBackgroundTabbedDialog_Err
    '</EhHeader>
  Dim iTmp As Integer
  
  'get the larger of the active tab height and inactive tab height
  iTmp = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
  With m_oXTab
    
    .pCls   'clear control
    
    'fill background color based on tab's enabled property
    If .aryTabs(m_iActiveTab).Enabled Then
      .lBackColor = m_lActiveTabBackEndColor
    Else
      .lBackColor = m_lDisabledTabBackColor
    End If
        
    'draw inner shadow (left)
    .pLine 0, iTmp, 0, m_iScaleHeight - 1, m_lOuterBorderColor
    
    
    'draw inner shadow (right)
    .pLine m_iScaleWidth - 1, iTmp, m_iScaleWidth - 1, m_iScaleHeight - 1, m_lOuterBorderColor
    
    'draw inner shadow (bottom)
    .pLine 0, m_iScaleHeight - 1, m_iScaleWidth, m_iScaleHeight - 1, m_lOuterBorderColor
    '.pLine 1, m_iScaleHeight - 2, m_iScaleWidth - 1, m_iScaleHeight - 2, vb3DShadow
  End With
    '<EhFooter>
    Exit Sub

ITheme_DrawBackgroundTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_DrawBackgroundTabbedDialog " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'these functions were seperated acc to theme so as to allow painting of only the desired parts (and not the whole control)
Private Sub ITheme_DrawOnActiveTabChange()
    '<EhHeader>
    On Error GoTo ITheme_DrawOnActiveTabChange_Err
    '</EhHeader>
  Call ITheme_DrawTabs
    '<EhFooter>
    Exit Sub

ITheme_DrawOnActiveTabChange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_DrawOnActiveTabChange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawTabs()
    '<EhHeader>
    On Error GoTo ITheme_DrawTabs_Err
    '</EhHeader>
  Call pCacheControlProperties     'cache the control proerties for faster access
  With m_oXTab
    Select Case .TabStyle
      Case xStyleTabbedDialog:
        Call ITheme_DrawTabsTabbedDialog
      Case xStylePropertyPages:
        Call Itheme_DrawTabsPropertyPages
    End Select
  End With
    '<EhFooter>
    Exit Sub

ITheme_DrawTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_DrawTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Itheme_DrawTabsPropertyPages()
    '<EhHeader>
    On Error GoTo Itheme_DrawTabsPropertyPages_Err
    '</EhHeader>
  Dim iCnt As Integer
  Dim iTabWidth As Integer
  Dim utFontRect As RECT
  Dim sTmp As String
  Dim utTabInfo As TabInfo
  Dim iTmpHeight As Integer
  Dim iAdjustedIconSize As Integer
  Dim iTmpY As Integer
 
  'Set the Inactive tab's font as current font (since the TextWidth function
  'will use the current font's size)
  Set m_oXTab.oFont = m_oXTab.InActiveTabFont
  
  
  'store the larger height in tmp var
  iTmpHeight = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
  'initialize the clickable items
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oXTab.aryTabs(iCnt)
        
    sTmp = Replace$(utTabInfo.Caption, "&&", "&")
    
    If InStr(1, sTmp, "&") Then
      
      'if still there is one '&' in the string then reduce the width by one more character (since the '&' will be conveted into an underline when painted)
      sTmp = Mid$(sTmp, 1, Len(sTmp) - 1)
    End If
    
    
    If utTabInfo.TabPicture Is Nothing Then
      'get tab width acc to the text size and border
      iTabWidth = m_oXTab.pTextWidth(sTmp) + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2
    Else
      If iTmpHeight - 2 < m_iIconSize Then    '-2 for borders
        'here we adjust the size of the icon if it does not fit into current tab
        iAdjustedIconSize = iTmpHeight - 2
      Else
        iAdjustedIconSize = m_iIconSize
      End If

      'get tab width acc to the text size, border and Image
      iTabWidth = m_oXTab.pTextWidth(sTmp) + (m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2) + iAdjustedIconSize + 1
    End If
    
    
    'following adjustments are used in case of property pages only. We must shift
    'the left (+2) or (-2) to make it look like standard XP property pages
    With utTabInfo.ClickableRect
      If iCnt = 0 And iCnt <> m_iActiveTab Then
        .Left = m_c_iPROP_PAGE_INACTIVE_TOP
        .Right = .Left + iTabWidth - m_c_iPROP_PAGE_INACTIVE_TOP + 1
      Else
        If iCnt = 0 Then
          .Left = 0
        Else
          If iCnt = m_iActiveTab Or iCnt = m_iActiveTab + 1 Then
            .Left = m_oXTab.aryTabs(iCnt - 1).ClickableRect.Right
          
          Else
            '1 pixel distance between property pages (in XP)
            .Left = m_oXTab.aryTabs(iCnt - 1).ClickableRect.Right + 1
          End If
        End If
          
        .Right = .Left + iTabWidth
      End If
        
      If iCnt = m_iActiveTab Then
        If m_iActiveTabHeight > m_iInActiveTabHeight Then
          .Top = 0
        Else
          .Top = m_iInActiveTabHeight - m_iActiveTabHeight
        End If
        .Bottom = .Top + m_iActiveTabHeight
      Else
        If m_iInActiveTabHeight > m_iActiveTabHeight Then
          .Top = 0
          .Bottom = .Top + m_iInActiveTabHeight
        Else
          .Top = m_iActiveTabHeight - m_iInActiveTabHeight
          .Bottom = .Top + m_iInActiveTabHeight
        End If
        
      End If
        
    End With
    m_oXTab.aryTabs(iCnt) = utTabInfo   'assign the new tab info to the existing one
  Next
  
  'fill the tab strip with TabStripBackColor (customizable... so that tab's can easily blend with the background)
  m_oXTab.pLine 0, 0, m_iScaleWidth, iTmpHeight, m_oXTab.TabStripBackColor, True, True
  
  'Now Draw Each Tab
  For iCnt = 0 To m_iTabCount - 1
      
    utTabInfo = m_oXTab.aryTabs(iCnt)   'Get Local copy
    
    With utTabInfo.ClickableRect
    
      
      If iCnt = m_iActiveTab Then   'if we are drawing the active tab
        
        
        If utTabInfo.Enabled Then
          Call m_oXTab.pFillCurvedGradientR(utTabInfo.ClickableRect, m_lActiveTabBackStartColor, m_lActiveTabBackEndColor, 2, True, True)
        Else
          Call m_oXTab.pFillCurvedGradientR(utTabInfo.ClickableRect, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, True)
        End If
        
        'top-left corner
        m_oXTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lOuterBorderColor
        
        'top line
        m_oXTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lOuterBorderColor
        
        'top-right corner
        m_oXTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
      
        'right line
        m_oXTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom + 1, m_lOuterBorderColor
        
        'left line
        m_oXTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lOuterBorderColor
        
        Set m_oXTab.oFont = m_oXTab.ActiveTabFont 'set the font
        
        'set fore ground color
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      ElseIf iCnt = m_iActiveTab - 1 Then   'if we are drawing tab just b4 active tab, then
        
        If utTabInfo.Enabled Then
          'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
          If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, True)
          Else
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, False)
          End If
        Else
          'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
          If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, True)
          Else
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, False)
          End If
        End If
          
        'top-left corner
        m_oXTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lOuterBorderColor
          
        'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
        If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
          'top line
          m_oXTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lOuterBorderColor
            
          'top-right corner
          m_oXTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
          
          'right line
          m_oXTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom + 1, m_lOuterBorderColor
        Else
          'top line
          m_oXTab.pLine .Left + 2, .Top, .Right, .Top, m_lOuterBorderColor
        End If
          
    
        'bottom line
        m_oXTab.pLine .Left, .Bottom, .Right + 1, .Bottom, m_lOuterBorderColor
            
        'Left Line
        m_oXTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lOuterBorderColor
          
        Set m_oXTab.oFont = m_oXTab.InActiveTabFont 'set the font
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      ElseIf iCnt = m_iActiveTab + 1 Then   'if we are drawing tab just after the active tab, then
      
        If utTabInfo.Enabled Then
          'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
          If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, True)
          Else
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, False, True)
          End If
        Else
          'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
          If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, True, True)
          Else
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 2, False, True)
          End If
        End If
          
        'following adjustments are needed if the inactive tab's height is more than active tab's height (the tab's corner should be properly rounded)
        If m_iInActiveTabHeight + 2 > m_iActiveTabHeight Then
          'top-left corner
          m_oXTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lOuterBorderColor
          
          'top line
          m_oXTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lOuterBorderColor
            
          'Left Line
          m_oXTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lOuterBorderColor
        Else
          'top line
          m_oXTab.pLine .Left, .Top, .Right - 2, .Top, m_lOuterBorderColor
        End If
  
    
        'top-right corner
        m_oXTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
    
        'right line
        m_oXTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom, m_lOuterBorderColor
    
        'bottom line
        m_oXTab.pLine .Left, .Bottom, .Right + 1, .Bottom, m_lOuterBorderColor
          
        Set m_oXTab.oFont = m_oXTab.InActiveTabFont 'set the font
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
      
      Else  'other non active tab (must draw full curves on both the sides)
        
        If utTabInfo.Enabled Then
          Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 2, True, True)
        Else
          Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right - 1, .Bottom, m_lDisabledTabBackColor, m_lInActiveTabBackEndColor, 2, True, True)
        End If
          
        'top-left corner
        m_oXTab.pLine .Left, .Top + 2, .Left + 3, .Top - 1, m_lOuterBorderColor
            
        'top line
        m_oXTab.pLine .Left + 2, .Top, .Right - 2, .Top, m_lOuterBorderColor
    
        'top-right corner
        m_oXTab.pLine .Right - 2, .Top + 1, .Right, .Top + 3, m_lOuterBorderColor
    
        'right line
        m_oXTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom, m_lOuterBorderColor
    
        'bottom line
        m_oXTab.pLine .Left, .Bottom, .Right + 1, .Bottom, m_lOuterBorderColor
            
        'Left Line
        m_oXTab.pLine .Left, .Top + 2, .Left, .Bottom + 1, m_lOuterBorderColor
        
        Set m_oXTab.oFont = m_oXTab.InActiveTabFont 'set the font
        
        'set font color
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If

      End If
          
      'if its left most tab then adjust the bottom line
      If iCnt = 0 Then
        'bottom line
        m_oXTab.pLine .Left - 2, .Bottom, .Left, .Bottom, m_lOuterBorderColor
      End If
      
      'do the adjustments for the border
      utFontRect.Left = .Left + 2
      utFontRect.Top = .Top + 2
      utFontRect.Bottom = .Bottom
      utFontRect.Right = .Right - 2
      
      
      sTmp = utTabInfo.Caption
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        If iTmpHeight - 2 < m_iIconSize Then    '-2 for borders
          'here we adjust the size of the icon if it does not fit into current tab
          iAdjustedIconSize = iTmpHeight - 2
        Else
          iAdjustedIconSize = m_iIconSize
        End If
    
        iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
        
        Select Case m_oXTab.PictureAlign
          Case xAlignLeftEdge, xAlignLeftOfCaption:
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), utFontRect.Left + 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, utFontRect.Left + 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
            'shift the text to be drawn after the picture
            utFontRect.Left = (utFontRect.Left + iAdjustedIconSize + 6) - m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
            
            'call the API for the text drawing
            W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            
            
            'revert the changes so that the focus rectangle can be drawn for the whole tab's clickable area
            utFontRect.Left = (utFontRect.Left - iAdjustedIconSize - 6) + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
            
          Case xAlignRightEdge, xAlignRightOfCaption:
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), utFontRect.Right - iAdjustedIconSize - 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, utFontRect.Right - iAdjustedIconSize - 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
            'm_oXTab.pPaintPicture utTabInfo.TabPicture, utFontRect.Right - iAdjustedIconSize, iTmpY, iAdjustedIconSize, iAdjustedIconSize
            
            'shift the text to be drawn after the picture
            utFontRect.Right = (utFontRect.Right + 1) - iAdjustedIconSize - 6
            
            'call the API for the text drawing
            W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            
            'revert the changes so that the focus rectangle can be drawn for the whole tab's clickable area
            utFontRect.Right = (utFontRect.Right - 1) + iAdjustedIconSize + 6
            
        End Select
        
      Else
        
        'call the API for the text drawing
        W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        
      End If
      
      'call the API for the text drawing
      'DrawText m_lHDC, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
      
      If m_oXTab.bUserMode Then    'only if in the run mode
      
        If iCnt = m_iActiveTab And m_bIsFocused And m_oXTab.ShowFocusRect Then
          
          'draw focus rectangle
          Call DrawFocusRect(m_lHdc, utFontRect)
        End If
      End If
    End With
    
  Next
  
  'draw the line in the empty area after all the property pages heads are drawn
  m_oXTab.pLine m_oXTab.aryTabs(m_iTabCount - 1).ClickableRect.Right, m_oXTab.aryTabs(m_iTabCount - 1).ClickableRect.Bottom, m_iScaleWidth, m_oXTab.aryTabs(m_iTabCount - 1).ClickableRect.Bottom, m_lOuterBorderColor
  
  'force drawing with current hover color
  m_utRect = m_oXTab.aryTabs(m_iActiveTab).ClickableRect
  pDrawOverXored
    '<EhFooter>
    Exit Sub

Itheme_DrawTabsPropertyPages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.Itheme_DrawTabsPropertyPages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawTabsTabbedDialog()
    '<EhHeader>
    On Error GoTo ITheme_DrawTabsTabbedDialog_Err
    '</EhHeader>
  Dim iCnt As Integer
  Dim iTabWidth As Integer
  Dim utFontRect As RECT
  Dim sTmp As String
  Dim utTabInfo As TabInfo
  Dim iTmpW As Integer
  Dim iTmpH As Integer
  Dim iTmpX As Integer
  Dim iTmpY As Integer
  Dim iOrigLeft As Integer
  Dim iOrigRight As Integer
  Dim iAdjustedIconSize As Integer

  iTabWidth = m_iScaleWidth / m_iTabCount    'remember iTabWidth is an integer ... so the result is automatically rounded
  
  'initialize the clickable items
  For iCnt = 0 To m_iTabCount - 1
    
    utTabInfo = m_oXTab.aryTabs(iCnt)
    
    'no need to calculate the text size(like in property pages).... since this is a tabbed dialog style
    
    With utTabInfo.ClickableRect
      .Left = iCnt * iTabWidth
      .Right = .Left + iTabWidth
      
      If iCnt = m_iActiveTab Then
        If m_iActiveTabHeight > m_iInActiveTabHeight Then
          .Top = 0
        Else
          .Top = m_iInActiveTabHeight - m_iActiveTabHeight
        End If
        .Bottom = .Top + m_iActiveTabHeight
      Else
        If m_iInActiveTabHeight > m_iActiveTabHeight Then
          .Top = 0
        Else
          .Top = m_iActiveTabHeight - m_iInActiveTabHeight
        End If
        .Bottom = .Top + m_iInActiveTabHeight
      End If
      
    End With
    m_oXTab.aryTabs(iCnt) = utTabInfo
  Next
  
  'if the last tab is shorter or longer than the usual size.. then adjust it to perfect size
  utTabInfo.ClickableRect.Right = m_iScaleWidth - 1
  m_oXTab.aryTabs(iCnt - 1) = utTabInfo


  'added to prevent lines etc (we are filling the tab strip with the tab strip color)
  m_oXTab.pLine 0, 0, m_iScaleWidth, IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight), m_oXTab.TabStripBackColor, True, True
  
  'Now Draw Each Tab
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oXTab.aryTabs(iCnt)     'fetch local copy
    
    With utTabInfo.ClickableRect
      
      If iCnt = m_iActiveTab Then         'if we are drawing active tab then
      
        If utTabInfo.Enabled Then
          m_oXTab.pFillCurvedGradient .Left, .Top, .Right, .Bottom, m_lActiveTabBackStartColor, m_lActiveTabBackEndColor, 4, True, True
        Else
          m_oXTab.pFillCurvedGradient .Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 4, True, True
        End If

        'top-left corner
        m_oXTab.pLine .Left, .Top + 4, .Left + 4, .Top, m_lOuterBorderColor
        
        'top line
        m_oXTab.pLine .Left + 4, .Top, .Right - 4, .Top, m_lOuterBorderColor
        
        'top-right corner
        m_oXTab.pLine .Right - 4, .Top, .Right, .Top + 4, m_lOuterBorderColor
      
        'right line
        m_oXTab.pLine .Right, .Top + 4, .Right, .Bottom + 1, m_lOuterBorderColor
        
        If utTabInfo.Enabled Then
          m_oXTab.pLine .Left + 1, .Bottom + 1, .Right, .Bottom + 1, m_lActiveTabBackEndColor
        Else
          m_oXTab.pLine .Left + 1, .Bottom + 1, .Right, .Bottom + 1, m_lDisabledTabBackColor
        End If

        
        'left line
        m_oXTab.pLine .Left, .Top + 4, .Left, .Bottom + 1, m_lOuterBorderColor
        
        Set m_oXTab.oFont = m_oXTab.ActiveTabFont       'set the font
        
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      Else      'we are drawing inactive tab
      
        If utTabInfo.Enabled Then
          Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor, 4, True, True)
        Else
          Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor, 4, True, True)
        End If
        
        'top-left corner
        m_oXTab.pLine .Left, .Top + 4, .Left + 4, .Top, m_lOuterBorderColor
        
        'top line
        m_oXTab.pLine .Left + 4, .Top, .Right - 4, .Top, m_lOuterBorderColor
        
        'top-right corner
        m_oXTab.pLine .Right - 4, .Top, .Right, .Top + 4, m_lOuterBorderColor
        
        
        'right line
        m_oXTab.pLine .Right, .Top + 4, .Right, .Bottom + 1, m_lOuterBorderColor
        
        'bottom line
        m_oXTab.pLine .Left, .Bottom + 1, .Right + 1, .Bottom + 1, m_lOuterBorderColor
        
        'left line
        m_oXTab.pLine .Left, .Top + 4, .Left, .Bottom + 1, m_lOuterBorderColor
      
        Set m_oXTab.oFont = m_oXTab.InActiveTabFont   'set the font
        
        'set foreground color
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      End If

      'do the adjustments for the border
      utFontRect.Left = .Left + 3
      utFontRect.Top = .Top + 3
      utFontRect.Bottom = .Bottom
      utFontRect.Right = .Right - 2

      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        If utFontRect.Top + m_iIconSize > utFontRect.Bottom + 1 Then '+1 for minor adjustments
          'adjust if going out of current tab's bottom
          iAdjustedIconSize = (utFontRect.Bottom + 1) - utFontRect.Top
        Else
          iAdjustedIconSize = m_iIconSize
        End If
        
        iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
                  
        Select Case m_oXTab.PictureAlign
          Case xAlignLeftEdge:
            iTmpX = utFontRect.Left
            
            If iCnt = m_iActiveTab Then   'if active tab then give a comeup effect
              iTmpX = iTmpX + 1
              iTmpY = iTmpY - 1
              
              'make sure our adjustment dosen't make it out of the font area
              If iTmpY < utFontRect.Top Then iTmpY = utFontRect.Top
            End If
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
          Case xAlignRightEdge:
            iTmpX = utFontRect.Right - iAdjustedIconSize
            
            If iCnt = m_iActiveTab Then 'if active tab then give a comeup effect
              iTmpX = iTmpX - 1
              iTmpY = iTmpY - 1
              
              'make sure our adjustment dosen't make it out of the font area
              If iTmpY < utFontRect.Top Then iTmpY = utFontRect.Top
            End If
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
          Case xAlignLeftOfCaption:
            iOrigLeft = utFontRect.Left
          Case xAlignRightOfCaption:
            iOrigRight = utFontRect.Right
        End Select
        
      End If
      
      sTmp = utTabInfo.Caption
      
      'calculate the rect to draw the text, also modify the string to get ellipsis etc
      W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
      
      
      iTmpW = utFontRect.Right - utFontRect.Left + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE
      iTmpH = utFontRect.Bottom - utFontRect.Top + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2
      
      
      'do the adjustments to center the text (both vertically and horizontally)
      utFontRect.Left = (utFontRect.Left - (m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2)) + .Right / 2 - utFontRect.Right / 2
      utFontRect.Right = utFontRect.Left + iTmpW
      
      utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
      utFontRect.Bottom = utFontRect.Top + iTmpH
      
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        Select Case m_oXTab.PictureAlign
          Case xAlignLeftOfCaption:
            iTmpX = utFontRect.Left - iAdjustedIconSize - 1
            
            'make sure our adjustment dosen't make it out of the font area
            If iTmpX < iOrigLeft Then iTmpX = iOrigLeft
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
          Case xAlignRightOfCaption:
            iTmpX = utFontRect.Right + 1
            
            'make sure our adjustment dosen't make it out of the font area
            If iTmpX + iAdjustedIconSize > iOrigRight Then iTmpX = iOrigRight - iAdjustedIconSize
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If

          
        End Select
      End If
      
      'Now draw the text
      W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_SINGLELINE

      If m_oXTab.bUserMode Then    'only if in the run mode

        If iCnt = m_iActiveTab And m_bIsFocused And m_oXTab.ShowFocusRect Then
          'draw focus rectangle
          Call DrawFocusRect(m_lHdc, utFontRect)
        End If
      End If
    End With

  Next
  
  'store the larger tab height
  iCnt = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)

  'adjust the corners (whole tab control's corners)
  m_oXTab.pLine 0, iCnt + 1, 0, iCnt + 4, m_lOuterBorderColor
  m_oXTab.pLine m_iScaleWidth - 1, iCnt + 1, m_iScaleWidth - 1, iCnt + 4, m_lOuterBorderColor
  
  'force drawing with current hover color
  m_utRect = m_oXTab.aryTabs(m_iActiveTab).ClickableRect
  pDrawOverXored
    '<EhFooter>
    Exit Sub

ITheme_DrawTabsTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_DrawTabsTabbedDialog " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_MouseDownHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    '<EhHeader>
    On Error GoTo ITheme_MouseDownHanlder_Err
    '</EhHeader>
  Dim utTabInfo As TabInfo
  Dim iCnt As Integer
  Dim iX As Integer
  Dim iY As Integer
  iX = CInt(sngX)
  iY = CInt(sngY)

  'if the mouse is already over then stop the timer and reset the over flag
  If m_bIsMouseOver Then
    m_bIsMouseOver = False
    m_oXTab.pSetTimer (0)
  End If
  
  If iY > IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight) Then
      
      'if lower than the larger tab height then exit sub since anything lower than
      'active tab's height will not result in a tab switch
      Exit Sub
  End If

  'now go through each tab's rect to determine if the mouse was clicked within its boundaries
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oXTab.aryTabs(iCnt)
    If (iX >= utTabInfo.ClickableRect.Left And iX <= utTabInfo.ClickableRect.Right) And (iY >= utTabInfo.ClickableRect.Top And iY <= utTabInfo.ClickableRect.Bottom) And utTabInfo.Enabled Then
      
      'if its the active tab then no need to switch
      If m_iActiveTab <> iCnt Then
        m_oXTab.ActiveTab = iCnt
      End If
      
      Exit Sub  'our work is finished .... no need to itirate further
    End If
  Next
    '<EhFooter>
    Exit Sub

ITheme_MouseDownHanlder_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_MouseDownHanlder " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_MouseMoveHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    '<EhHeader>
    On Error GoTo ITheme_MouseMoveHanlder_Err
    '</EhHeader>
  Dim iCnt As Integer
  Dim iX As Integer
  Dim iY As Integer
  iX = CInt(sngX)
  iY = CInt(sngY)

  If m_bIsMouseOver Then      'no use goin in if already mouse is over
    Exit Sub
  End If
  
  If iY > IIf(m_iInActiveTabHeight > m_iActiveTabHeight, m_iInActiveTabHeight, m_iActiveTabHeight) Then
    'if lower than the larger tab height then exit sub since anything lower than
    'larger tab's height will not result in a tab switch
    
    Exit Sub
  End If
  
  
  Call pCacheControlProperties     'cache the control proerties for faster access

  For iCnt = 0 To m_iTabCount - 1
    If iX >= m_oXTab.aryTabs(iCnt).ClickableRect.Left And iX <= m_oXTab.aryTabs(iCnt).ClickableRect.Right And m_oXTab.aryTabs(iCnt).Enabled Then
        
      If iCnt = m_iActiveTab Then Exit Sub    'no need to draw for the active tab
      
      'store the index of the tab on which the mouse is over
      
      m_iMouseOverTabIndex = iCnt
      
      Select Case m_oXTab.TabStyle
        Case xStyleTabbedDialog:
          Call pDrawOnMouseOverTabbed
        Case xStylePropertyPages:
          Call pDrawOnMouseOverProperty
      End Select
      Exit Sub
    End If
  Next
  
    '<EhFooter>
    Exit Sub

ITheme_MouseMoveHanlder_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_MouseMoveHanlder " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_MouseUpHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  'Do Nothing
    '<EhHeader>
    On Error GoTo ITheme_MouseUpHanlder_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

ITheme_MouseUpHanlder_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_MouseUpHanlder " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'IMPORTANT: must be called before anything else is called... this function is called in the usercontrol
'the time the object is instantiated
Private Sub ITheme_SetControl(oXTab As XTab)
    '<EhHeader>
    On Error GoTo ITheme_SetControl_Err
    '</EhHeader>
  Set m_oXTab = oXTab
  m_oXTab.bAutoRedraw = False
    '<EhFooter>
    Exit Sub

ITheme_SetControl_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_SetControl " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub ITheme_ShowHideFocus()
    '<EhHeader>
    On Error GoTo ITheme_ShowHideFocus_Err
    '</EhHeader>
  Call pCacheControlProperties    'store in local variables
  
  Select Case m_oXTab.TabStyle
    Case xStylePropertyPages
      ITheme_ShowHideFocusPropertyPages
    Case xStyleTabbedDialog
      ITheme_ShowHideFocusTabbedDialog
  End Select
    '<EhFooter>
    Exit Sub

ITheme_ShowHideFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_ShowHideFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_ShowHideFocusPropertyPages()
    '<EhHeader>
    On Error GoTo ITheme_ShowHideFocusPropertyPages_Err
    '</EhHeader>
  Dim utFontRect As RECT
  Dim utTabInfo As TabInfo
  
  If Not m_oXTab.bUserMode Then       'only if in the run mode
    
    Exit Sub
    
  End If
  
  If Not m_oXTab.ShowFocusRect Then   'only if Show Focus Rect is true for the control
      
    Exit Sub
    
  End If
  
  utTabInfo = m_oXTab.aryTabs(m_iActiveTab)
  With utTabInfo.ClickableRect
    'do the adjustments for the border
    utFontRect.Left = .Left + 2
    utFontRect.Top = .Top + 2
    utFontRect.Bottom = .Bottom
    utFontRect.Right = .Right - 2
      
    If utTabInfo.Enabled Then             'done to allow proper drawing of focus rect
      m_oXTab.lForeColor = m_lActiveTabForeColor
    Else
      m_oXTab.lForeColor = m_lDisabledTabForeColor
    End If
    
    'show/hide the focus rectangle (drawn in XOR mode, so calling it again with same coords will erase it)
    Call DrawFocusRect(m_lHdc, utFontRect)
  End With
    '<EhFooter>
    Exit Sub

ITheme_ShowHideFocusPropertyPages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_ShowHideFocusPropertyPages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_ShowHideFocusTabbedDialog()
    '<EhHeader>
    On Error GoTo ITheme_ShowHideFocusTabbedDialog_Err
    '</EhHeader>
  Dim utFontRect As RECT
  Dim sTmp As String
  Dim utTabInfo As TabInfo
  Dim iTmpW As Integer
  Dim iTmpH As Integer
  
  If Not m_oXTab.bUserMode Then       'only if in the run mode
    
    Exit Sub
    
  End If
  
  If Not m_oXTab.ShowFocusRect Then   'only if Show Focus Rect is true for the control
      
    Exit Sub
    
  End If
  
  utTabInfo = m_oXTab.aryTabs(m_iActiveTab)
  With utTabInfo.ClickableRect
    'do the adjustments for the border
    utFontRect.Left = .Left + 3
    utFontRect.Top = .Top + 3
    utFontRect.Bottom = .Bottom
    utFontRect.Right = .Right - 2

    sTmp = utTabInfo.Caption
    
    Set m_oXTab.oFont = m_oXTab.ActiveTabFont
    
    'calculate the rect to draw the text, and get proper string (including ellipsis etc)
    W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
    
    
    iTmpW = utFontRect.Right - utFontRect.Left + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE
    iTmpH = utFontRect.Bottom - utFontRect.Top + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2
    
    
    'do the adjustments to center the text (both vertically and horizontally)
    utFontRect.Left = (utFontRect.Left - (m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2)) + .Right / 2 - utFontRect.Right / 2
    utFontRect.Right = utFontRect.Left + iTmpW
    
    utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
    utFontRect.Bottom = utFontRect.Top + iTmpH
    
    If utTabInfo.Enabled Then             'done to allow proper drawing of focus rect
      m_oXTab.lForeColor = m_lActiveTabForeColor
    Else
      m_oXTab.lForeColor = m_lDisabledTabForeColor
    End If
    
    'show/hide the focus rectangle (drawn in XOR mode, so calling it again with same coords will erase it)
    Call DrawFocusRect(m_lHdc, utFontRect)
  End With
    '<EhFooter>
    Exit Sub

ITheme_ShowHideFocusTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_ShowHideFocusTabbedDialog " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'timer function is used in this theme to track the mouse going out of the current tab's clickable regin
'if the mouse goes out then we must erase the hover effect line
Private Sub ITheme_TimerEvent()
    '<EhHeader>
    On Error GoTo ITheme_TimerEvent_Err
    '</EhHeader>
  
  'get cusror position
  Call GetCursorPos(m_Pnt)
  
  'convert coordinates
  Call ScreenToClient(m_lHwnd, m_Pnt)
  
  'check if the mouse is out of the clickable region
  If (m_Pnt.X < m_utRect.Left Or m_Pnt.X > m_utRect.Right) Or _
     (m_Pnt.Y < m_utRect.Top Or m_Pnt.Y > m_utRect.Bottom) Then
        
    m_oXTab.pSetTimer (0)   'disable the timer
        
    m_bIsMouseOver = False  'indicate mouse up
    
    Call pDrawOverXored     'draw the changes to clear the hover effect
  End If
    '<EhFooter>
    Exit Sub

ITheme_TimerEvent_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_TimerEvent " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'function will reset all the colors back to default colors
Private Sub ITheme_ResetColorsToDefault()
    '<EhHeader>
    On Error GoTo ITheme_ResetColorsToDefault_Err
    '</EhHeader>
  With m_oXTab
    
    .ActiveTabBackStartColor = &HFBFDFB
    .ActiveTabBackEndColor = &HFBFDFB
    .ActiveTabForeColor = vbButtonText
    
    .InActiveTabBackStartColor = &HFFFFFF
    .InActiveTabBackEndColor = &HEAF0F0
    .InActiveTabForeColor = vbButtonText
    
    On Error Resume Next
    .TabStripBackColor = m_oXTab.oAmbient.BackColor
    .DisabledTabBackColor = vbButtonFace
    .DisabledTabForeColor = &HA0A0A0
    
    '.TopLeftInnerBorderColor = vb3DHighlight
    .OuterBorderColor = &H9B9C91
    '.BottomRightInnerBorderColor = vb3DShadow
    
    .HoverColorInverted = &HC43800          'color drawn in XOR mode to achive the hover effect
  End With
    '<EhFooter>
    Exit Sub

ITheme_ResetColorsToDefault_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.ITheme_ResetColorsToDefault " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'=====================================================================================================================

'=====Private Functions===============================================================================================

'cache the properties from the control
'(this prevents trips to again and again fetch properties from the user control)
Private Sub pCacheControlProperties()
    '<EhHeader>
    On Error GoTo pCacheControlProperties_Err
    '</EhHeader>
  m_lHwnd = m_oXTab.lHwnd
  m_lHdc = m_oXTab.lHDC
  m_iActiveTab = m_oXTab.ActiveTab
  m_iActiveTabHeight = m_oXTab.ActiveTabHeight
  m_iInActiveTabHeight = m_oXTab.InActiveTabHeight
  m_iScaleWidth = m_oXTab.iScaleWidth
  m_iScaleHeight = m_oXTab.iScaleHeight
  m_iTabCount = m_oXTab.TabCount
  m_bIsFocused = m_oXTab.bIsFocused
  m_lOuterBorderColor = m_oXTab.OuterBorderColor
  
  m_lActiveTabForeColor = m_oXTab.ActiveTabForeColor
  m_lActiveTabBackStartColor = m_oXTab.ActiveTabBackStartColor
  m_lActiveTabBackEndColor = m_oXTab.ActiveTabBackEndColor
  
  m_lInActiveTabForeColor = m_oXTab.InActiveTabForeColor
  m_lInActiveTabBackStartColor = m_oXTab.InActiveTabBackStartColor
  m_lInActiveTabBackEndColor = m_oXTab.InActiveTabBackEndColor
  m_lHoverColorInverted = m_oXTab.HoverColorInverted
  m_lDisabledTabBackColor = m_oXTab.DisabledTabBackColor
  m_lDisabledTabForeColor = m_oXTab.DisabledTabForeColor
  
  'Get System's default size for a Icon.
  If m_oXTab.PictureSize = xSizeSmall Then
    m_iIconSize = GetSystemMetrics(SM_CXSMICON)
  Else
    m_iIconSize = GetSystemMetrics(SM_CXICON)
  End If
    '<EhFooter>
    Exit Sub

pCacheControlProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.pCacheControlProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


'function called when mouse hovers over the tab (tabbed)
Private Sub pDrawOnMouseOverTabbed()
    '<EhHeader>
    On Error GoTo pDrawOnMouseOverTabbed_Err
    '</EhHeader>
  m_bIsMouseOver = True
  
  m_utRect = m_oXTab.aryTabs(m_iMouseOverTabIndex).ClickableRect
  
  Call pDrawOverXored
  
  m_oXTab.pSetTimer (10)
    '<EhFooter>
    Exit Sub

pDrawOnMouseOverTabbed_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.pDrawOnMouseOverTabbed " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'used to draw the top of the tab Xored with Hover color
Private Sub pDrawOverXored()
    '<EhHeader>
    On Error GoTo pDrawOverXored_Err
    '</EhHeader>
  Select Case m_oXTab.TabStyle
    Case xStyleTabbedDialog:
      pDrawOverXoredTabbed
    Case xStylePropertyPages:
      pDrawOverXoredProperty
  End Select
    '<EhFooter>
    Exit Sub

pDrawOverXored_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.pDrawOverXored " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'used to draw the top of the tab Xored with Hover color (Tabbed Dialog)
Private Sub pDrawOverXoredTabbed()
    '<EhHeader>
    On Error GoTo pDrawOverXoredTabbed_Err
    '</EhHeader>
  m_oXTab.iDrawMode = vbXorPen
  With m_utRect
    m_oXTab.pFillCurvedSolid .Left, .Top, .Right, .Top + 3, m_lHoverColorInverted, 4, True, True
  End With
  m_oXTab.iDrawMode = vbCopyPen
    '<EhFooter>
    Exit Sub

pDrawOverXoredTabbed_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.pDrawOverXoredTabbed " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'used to draw the top of the tab Xored with Hover color (Property Pages)
Private Sub pDrawOverXoredProperty()
    '<EhHeader>
    On Error GoTo pDrawOverXoredProperty_Err
    '</EhHeader>
  m_oXTab.iDrawMode = vbXorPen
  With m_utRect
    If m_iMouseOverTabIndex = m_iActiveTab Then
      m_oXTab.pFillCurvedSolid .Left, .Top, .Right, .Top + 2, m_lHoverColorInverted, 2, True, True
    ElseIf m_iMouseOverTabIndex = m_iActiveTab - 1 Then
      m_oXTab.pFillCurvedSolid .Left, .Top, .Right, .Top + 2, m_lHoverColorInverted, 2, True, False
    ElseIf m_iMouseOverTabIndex = m_iActiveTab + 1 Then
      m_oXTab.pFillCurvedSolid .Left, .Top, .Right, .Top + 2, m_lHoverColorInverted, 2, False, True
    Else
      m_oXTab.pFillCurvedSolid .Left, .Top, .Right, .Top + 2, m_lHoverColorInverted, 2, True, True
    End If
  End With
  m_oXTab.iDrawMode = vbCopyPen
    '<EhFooter>
    Exit Sub

pDrawOverXoredProperty_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.pDrawOverXoredProperty " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'function called when mouse hovers over the tab (Property)
Private Sub pDrawOnMouseOverProperty()
    '<EhHeader>
    On Error GoTo pDrawOnMouseOverProperty_Err
    '</EhHeader>
  m_bIsMouseOver = True
  
  m_utRect = m_oXTab.aryTabs(m_iMouseOverTabIndex).ClickableRect
  
  Call pDrawOverXored
  
  'since the mouse is over the tab.. track the mouse going out of the clickable region
  m_oXTab.pSetTimer (10)
    '<EhFooter>
    Exit Sub

pDrawOnMouseOverProperty_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeWinXP.pDrawOnMouseOverProperty " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'=====================================================================================================================


