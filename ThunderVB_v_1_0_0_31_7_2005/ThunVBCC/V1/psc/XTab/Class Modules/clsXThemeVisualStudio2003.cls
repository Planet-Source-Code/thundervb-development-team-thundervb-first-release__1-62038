VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xThemeVisualStudio2003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'xThemeVisualStudio2003 :   Theme mimics the Visual Studio .Net 2003 tabs
'                           I found them cool so i implemented them
'
'                           See ITheme for implemented function description

Option Explicit

Implements ITheme

'====Constant Variables===============================================================================================
Private Const m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE As Integer = 7   'the distance between the text and the border in a Property Pages style tab
Private Const m_c_iPROP_PAGE_INACTIVE_TOP As Integer = 2      'the top for the property page (inactive property page)
'=====================================================================================================================

'====Private Variables================================================================================================
Private m_oXTab As XTab


'following property are cached (locally) for improved performance
Private m_lHdc As Long
Private m_iActiveTab As Integer
Private m_iActiveTabHeight As Integer
Private m_iInActiveTabHeight As Integer
Private m_iTabCount As Integer
Private m_bIsFocused As Boolean
Private m_iScaleWidth As Integer
Private m_iScaleHeight As Integer
Private m_lOuterBorderColor As Long
Private m_lBottomRightInnerBorderColor As Long
Private m_lActiveTabForeColor As Long
Private m_lActiveTabBackStartColor As Long
Private m_lActiveTabBackEndColor As Long
Private m_lInActiveTabForeColor As Long
Private m_lInActiveTabBackStartColor As Long
Private m_lInActiveTabBackEndColor As Long
Private m_lDisabledTabForeColor As Long
Private m_lDisabledTabBackColor As Long
Private m_iIconSize As Integer
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_DrawBackground " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawBackgroundPropertyPages()
    '<EhHeader>
    On Error GoTo ITheme_DrawBackgroundPropertyPages_Err
    '</EhHeader>
  Dim iTmp As Integer
  
  'get the larger of the active tab height and inactive tab height
  iTmp = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
  With m_oXTab
    
    .pCls   'clear the control
    
    
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
  End With
    '<EhFooter>
    Exit Sub

ITheme_DrawBackgroundPropertyPages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_DrawBackgroundPropertyPages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawBackgroundTabbedDialog()
    '<EhHeader>
    On Error GoTo ITheme_DrawBackgroundTabbedDialog_Err
    '</EhHeader>
  Call ITheme_DrawBackgroundPropertyPages
    '<EhFooter>
    Exit Sub

ITheme_DrawBackgroundTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_DrawBackgroundTabbedDialog " & _
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_DrawOnActiveTabChange " & _
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_DrawTabs " & _
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
  
  Dim iAdjustedIconSize As Integer
  Dim iTmpY As Integer
  Dim iTmpHeight As Integer

  'store the larger height in tmp var
  iTmpHeight = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
  
  'Set the active tab's font as current font (since the TextWidth function
  'will use the current font's size)
  Set m_oXTab.oFont = m_oXTab.ActiveTabFont
  
  'initialize the clickable items
  For iCnt = 0 To m_iTabCount - 1
    
    utTabInfo = m_oXTab.aryTabs(iCnt)     'get into local variable
        
    sTmp = Replace$(utTabInfo.Caption, "&&", "&")
    
    If InStr(1, sTmp, "&") Then
      
      'if still there is one '&' in the string then reduce the width by one more character (since the '&' will be conveted into an underline when painted)
      sTmp = Mid$(sTmp, 1, Len(sTmp) - 1)
    End If
    
    
    If utTabInfo.TabPicture Is Nothing Then
      'get tab width acc to the text size and border
      iTabWidth = m_oXTab.pTextWidth(sTmp) + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2
    Else
      If iTmpHeight - 2 < m_iIconSize Then    '-6 for borders
        'here we adjust the size of the icon if it does not fit into current tab
        iAdjustedIconSize = iTmpHeight - 2
      Else
        iAdjustedIconSize = m_iIconSize
      End If

      'get tab width acc to the text size, border and Image
      iTabWidth = m_oXTab.pTextWidth(sTmp) + (m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2) + iAdjustedIconSize + 1
    End If
    
    
    'get tab width acc to the text size and border
    'iTabWidth = m_oXTab.pTextWidth(sTmp) + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2
    
    
    
    'following adjustments are used in case of property pages only. We must shift
    'the left (+2) or (-2) to make it look like standard property pages
    With utTabInfo.ClickableRect
      If iCnt = 0 And iCnt <> m_iActiveTab Then
        .Left = m_c_iPROP_PAGE_INACTIVE_TOP
        .Right = .Left + iTabWidth - m_c_iPROP_PAGE_INACTIVE_TOP + 1
      Else
        If iCnt = 0 Then
          .Left = 0
        Else
          'If iCnt = m_iActiveTab Or iCnt = m_iActiveTab + 1 Then
          .Left = m_oXTab.aryTabs(iCnt - 1).ClickableRect.Right
          
          ' Else
          '1 pixel distance between property pages (in XP)
          '   .Left = m_oXTab.aryTabs(iCnt - 1).ClickableRect.Right + 1
          ' End If
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
    
    m_oXTab.aryTabs(iCnt) = utTabInfo       'assign the new tab info to the existing one
    
  Next
  

  'fill the tab strip with TabStripBackColor (customizable... so that tab's can easily blend with the background)
  m_oXTab.pLine 0, 0, m_iScaleWidth, IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight), m_oXTab.TabStripBackColor, True, True
  
  'Now Draw Each Tab
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oXTab.aryTabs(iCnt)     'fetch local copy
    
    With utTabInfo.ClickableRect
      
      If iCnt = m_iActiveTab Then         'if we are drawing the active tab
      
        If utTabInfo.Enabled Then
          m_oXTab.pFillCurvedGradient .Left, .Top, .Right, .Bottom, m_lActiveTabBackStartColor, m_lActiveTabBackEndColor
        Else
          m_oXTab.pFillCurvedGradient .Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor
        End If

        'top line
        m_oXTab.pLine .Left, .Top, .Right, .Top, m_lOuterBorderColor
        
      
        'right line
        m_oXTab.pLine .Right, .Top, .Right, .Bottom + 2, m_lBottomRightInnerBorderColor
        
        If utTabInfo.Enabled Then
          'bottom line  (actually we must erase the previously drawn background (since this is the active tab)
          m_oXTab.pLine .Left, .Bottom + 1, .Right, .Bottom + 1, m_lActiveTabBackEndColor
        Else
          'bottom line  (actually we must erase the previously drawn background (since this is the active tab)
          m_oXTab.pLine .Left, .Bottom + 1, .Right, .Bottom + 1, m_lDisabledTabBackColor
        End If

        
        'left line
        m_oXTab.pLine .Left, .Top, .Left, .Bottom + 2, m_lOuterBorderColor
        
        Set m_oXTab.oFont = m_oXTab.ActiveTabFont       'set the active tab font as current font
        
        'set fore color
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      Else      'its an inactive tab
      
        If utTabInfo.Enabled Then
          If iCnt = m_iActiveTab + 1 Then   'if we are drawing tab just after active tab, then
          
            Call m_oXTab.pFillCurvedGradient(.Left + 1, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor)
            
          Else  'we are drawing tab just b4 active tab, then
          
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lInActiveTabBackStartColor, m_lInActiveTabBackEndColor)
            
          End If
        Else
          If iCnt = m_iActiveTab + 1 Then 'if we are drawing tab just after active tab, then
          
            Call m_oXTab.pFillCurvedGradient(.Left + 1, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
            
          Else    'we are drawing tab just b4 active tab, then
          
            Call m_oXTab.pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
            
          End If
        End If
        
        
        'following is special in case of Visual studio .Net 2003 tabs , a simple line seperates two inactive tabs
        If iCnt <> m_iActiveTab - 1 Then
          'right line
          m_oXTab.pLine .Right - 1, .Top + 2, .Right - 1, .Bottom - 2, m_lInActiveTabForeColor
        End If
        
        'bottom line
        m_oXTab.pLine .Left, .Bottom + 1, .Right + 1, .Bottom + 1, m_lOuterBorderColor
        
      
        Set m_oXTab.oFont = m_oXTab.InActiveTabFont   'set the font
        
        'set fore color
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      End If

      'do the adjustments for the border
      utFontRect.Left = .Left + 2
      utFontRect.Top = .Top + 2
      utFontRect.Bottom = .Bottom
      utFontRect.Right = .Right - 1
      
      
      sTmp = utTabInfo.Caption
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        If iTmpHeight - 6 < m_iIconSize Then    '-6 for borders
          'here we adjust the size of the icon if it does not fit into current tab
          iAdjustedIconSize = iTmpHeight - 6
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
      'w_DrawText m_lHDC, sTmp, -1, utFontRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
      
      If m_oXTab.bUserMode Then    'only if in the run mode
      
        If iCnt = m_iActiveTab And m_bIsFocused And m_oXTab.ShowFocusRect Then
          'draw focus rect
          Call DrawFocusRect(m_lHdc, utFontRect)
        End If
      End If
      
    End With

    

  Next
  
  'store the larger tab height
  iCnt = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)

  'adjust the corners
  m_oXTab.pLine 0, iCnt + 1, 0, iCnt + 4, m_lOuterBorderColor
  m_oXTab.pLine m_iScaleWidth - 1, iCnt + 1, m_iScaleWidth - 1, iCnt + 4, m_lOuterBorderColor
  
  'draw the line in the empty area after all the property pages heads are drawn
  m_oXTab.pLine m_oXTab.aryTabs(m_iTabCount - 1).ClickableRect.Right, m_oXTab.aryTabs(m_iTabCount - 1).ClickableRect.Bottom + 1, m_iScaleWidth, m_oXTab.aryTabs(m_iTabCount - 1).ClickableRect.Bottom + 1, m_lOuterBorderColor
  
    '<EhFooter>
    Exit Sub

Itheme_DrawTabsPropertyPages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.Itheme_DrawTabsPropertyPages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawTabsTabbedDialog()
    '<EhHeader>
    On Error GoTo ITheme_DrawTabsTabbedDialog_Err
    '</EhHeader>
  Call Itheme_DrawTabsPropertyPages 'call the same function (VS tabs and proeprty pages are same)
    '<EhFooter>
    Exit Sub

ITheme_DrawTabsTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_DrawTabsTabbedDialog " & _
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_MouseDownHanlder " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_MouseMoveHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  'do Nothing
    '<EhHeader>
    On Error GoTo ITheme_MouseMoveHanlder_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

ITheme_MouseMoveHanlder_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_MouseMoveHanlder " & _
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_MouseUpHanlder " & _
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_SetControl " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub ITheme_ShowHideFocus()
    '<EhHeader>
    On Error GoTo ITheme_ShowHideFocus_Err
    '</EhHeader>
  Call pCacheControlProperties
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_ShowHideFocus " & _
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
    utFontRect.Right = .Right - 1
      
         
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_ShowHideFocusPropertyPages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_ShowHideFocusTabbedDialog()
    '<EhHeader>
    On Error GoTo ITheme_ShowHideFocusTabbedDialog_Err
    '</EhHeader>
  ITheme_ShowHideFocusPropertyPages
    '<EhFooter>
    Exit Sub

ITheme_ShowHideFocusTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_ShowHideFocusTabbedDialog " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub ITheme_TimerEvent()
  'not used
    '<EhHeader>
    On Error GoTo ITheme_TimerEvent_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

ITheme_TimerEvent_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_TimerEvent " & _
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
    
    .ActiveTabBackStartColor = vbButtonFace
    .ActiveTabBackEndColor = vbButtonFace
    .ActiveTabForeColor = vbButtonText
    
    .InActiveTabBackStartColor = vb3DLight
    .InActiveTabBackEndColor = vb3DLight
    .InActiveTabForeColor = &H80000011  'vbDisabledText
    
    .TabStripBackColor = vb3DLight
    .DisabledTabBackColor = vbButtonFace
    .DisabledTabForeColor = vb3DDKShadow
    
    'Not Reset for this theme
    '.TopLeftInnerBorderColor
    
    .OuterBorderColor = vb3DHighlight
    .BottomRightInnerBorderColor = vb3DShadow
    
    'Not Reset for this theme
    ' .HoverColorInverted
  End With
    '<EhFooter>
    Exit Sub

ITheme_ResetColorsToDefault_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.ITheme_ResetColorsToDefault " & _
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
  m_lHdc = m_oXTab.lHDC
  m_iActiveTab = m_oXTab.ActiveTab
  m_iActiveTabHeight = m_oXTab.ActiveTabHeight
  m_iInActiveTabHeight = m_oXTab.InActiveTabHeight
  m_iScaleWidth = m_oXTab.iScaleWidth
  m_iScaleHeight = m_oXTab.iScaleHeight
  m_iTabCount = m_oXTab.TabCount
  m_bIsFocused = m_oXTab.bIsFocused
  m_lOuterBorderColor = m_oXTab.OuterBorderColor
  m_lBottomRightInnerBorderColor = m_oXTab.BottomRightInnerBorderColor
  
  m_lActiveTabForeColor = m_oXTab.ActiveTabForeColor
  m_lActiveTabBackStartColor = m_oXTab.ActiveTabBackStartColor
  m_lActiveTabBackEndColor = m_oXTab.ActiveTabBackEndColor
  
  m_lInActiveTabForeColor = m_oXTab.InActiveTabForeColor
  m_lInActiveTabBackStartColor = m_oXTab.InActiveTabBackStartColor
  m_lInActiveTabBackEndColor = m_oXTab.InActiveTabBackEndColor
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
           "in ThunVBCC_v1_0.xThemeVisualStudio2003.pCacheControlProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'=====================================================================================================================

