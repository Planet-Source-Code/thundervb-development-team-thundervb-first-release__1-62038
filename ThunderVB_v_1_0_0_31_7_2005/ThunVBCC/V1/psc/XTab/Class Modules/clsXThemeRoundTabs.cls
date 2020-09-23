VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xThemeRoundTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'xThemeRoundTabs : Rounded Tabs example. This just show's a basic idea.
'                  We can do wonders with this and a little APIs.
'
'                  See ITheme for implemented function description

Option Explicit

Implements ITheme

'== Constants ========================================================================================================
Private Const m_c_iFOCUS_RECT_AND_TEXT_DISTANCE = 2  'the distance between the text (caption) of the tab and the focus Rect
Private Const m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE As Integer = 7   'the distance between the text and the border in a Property Pages style tab
Private Const m_c_iPROP_PAGE_INACTIVE_TOP As Integer = 2      'the top for the property page (inactive property page)
'=====================================================================================================================

'== Private Variables ================================================================================================

Private m_oXTab As XTab
Private m_bIsBackgroundPaintDelayed As Boolean      'See ITheme_DrawBackground() to get a description of this flag


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
Private m_lActiveTabForeColor As Long
Private m_lActiveTabBackStartColor As Long
Private m_lInActiveTabForeColor As Long
Private m_lInActiveTabBackStartColor As Long
Private m_lDisabledTabForeColor As Long
Private m_lDisabledTabBackColor As Long
Private m_iXRadius As Integer
Private m_iYRadius As Integer
Private m_iIconSize As Integer
'=====================================================================================================================



'==== Implemented Functions ==========================================================================================
Private Sub ITheme_DrawBackground()
  ' The below like is added to allow skipping of background paint.
  ' We want to draw the tabs and background one after another for
  ' this theme. We set the flag to true and then call this function
  ' again from the draw tabs() routine
    '<EhHeader>
    On Error GoTo ITheme_DrawBackground_Err
    '</EhHeader>
  If Not m_bIsBackgroundPaintDelayed Then
    m_bIsBackgroundPaintDelayed = True
  Else
    Call pCacheControlProperties     'cache the control proerties for faster access
    Select Case m_oXTab.TabStyle
      Case xStyleTabbedDialog:
        Call ITheme_DrawBackgroundTabbedDialog
      Case xStylePropertyPages:
        Call ITheme_DrawBackgroundPropertyPages
    End Select
  End If
    '<EhFooter>
    Exit Sub

ITheme_DrawBackground_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_DrawBackground " & _
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_DrawBackgroundPropertyPages " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_DrawBackgroundTabbedDialog()
    '<EhHeader>
    On Error GoTo ITheme_DrawBackgroundTabbedDialog_Err
    '</EhHeader>
  Dim iTmp As Integer
  Dim utRect As RECT
  
  'get the larger of the active tab height and inactive tab height
  iTmp = IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight)
  
  With m_oXTab
    
    'fill background color based on tab's enabled property
    If .aryTabs(m_iActiveTab).Enabled Then
      .lFillColor = m_lActiveTabBackStartColor
    Else
      .lFillColor = m_lDisabledTabBackColor
    End If
        
    .iFillStyle = vbFSSolid   'must set to transparent (since we are setting it to Solid in the DrawTabs() )
   
    .pLine 0, iTmp + 1, m_iScaleWidth - 1, m_iScaleHeight - 1, m_lOuterBorderColor, True, False
    
    'erase the bottom line for active tab
    utRect = .aryTabs(m_iActiveTab).ClickableRect
    
    If .aryTabs(m_iActiveTab).Enabled Then
      .pLine utRect.Left + 1, utRect.Bottom - 1, utRect.Right - 1, utRect.Bottom - 1, m_lActiveTabBackStartColor, False, False
    Else
      .pLine utRect.Left + 1, utRect.Bottom - 1, utRect.Right - 1, utRect.Bottom - 1, m_lDisabledTabBackColor, False, False
    End If
    
    

  End With
    '<EhFooter>
    Exit Sub

ITheme_DrawBackgroundTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_DrawBackgroundTabbedDialog " & _
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_DrawOnActiveTabChange " & _
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
  
  m_bIsBackgroundPaintDelayed = True
  
  'draw bg again to give a nice effect
  ITheme_DrawBackground
    '<EhFooter>
    Exit Sub

ITheme_DrawTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_DrawTabs " & _
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
  Dim iTmpW As Integer
  Dim iTmpH As Integer
  Dim iAdjustedIconSize As Integer
  Dim iTmpX As Integer
  Dim iTmpY As Integer
  Dim iTmpHeight As Integer
  Dim iOrigLeft As Integer
  Dim iOrigRight As Integer

 
  'Set the active tab's font as current font (since the TextWidth function
  'will use the current font's size)
  Set m_oXTab.oFont = m_oXTab.ActiveTabFont
  
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
      iTabWidth = m_oXTab.pTextWidth(sTmp) + (m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2) + iAdjustedIconSize + 4
    End If
    
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
          .Left = m_oXTab.aryTabs(iCnt - 1).ClickableRect.Right + 1
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
      .Bottom = .Bottom + 2
    End With
    m_oXTab.aryTabs(iCnt) = utTabInfo
  Next
  
  'fill the tab strip with TabStripBackColor (customizable... so that tab's can easily blend with the background)
  m_oXTab.pLine 0, 0, m_iScaleWidth, IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight), m_oXTab.TabStripBackColor, True, True
  
  
  'set the fill style to Solid
  m_oXTab.iFillStyle = vbFSSolid

        
  'Now Draw Each Tab
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oXTab.aryTabs(iCnt)     'fetch local copy
    
    m_oXTab.lForeColor = m_lOuterBorderColor
    
    With utTabInfo.ClickableRect
      
      If iCnt = m_iActiveTab Then     'if we are drawing active tab
        
        
        'we'll use solid colors for background (since we are calling RoundRect API)
        If utTabInfo.Enabled Then
          m_oXTab.lFillColor = m_lActiveTabBackStartColor
        Else
          m_oXTab.lFillColor = m_lDisabledTabBackColor
        End If
        
        'draw round effect
        Call RoundRect(m_lHdc, .Left, .Top, .Right, m_iScaleHeight - 1, m_iXRadius, m_iYRadius)

        Set m_oXTab.oFont = m_oXTab.ActiveTabFont       'set the font
        
        'set foreground color
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      Else
        
        'we'll use solid colors for background (since we are calling RoundRect API)
        If utTabInfo.Enabled Then
          m_oXTab.lFillColor = m_lInActiveTabBackStartColor
        Else
          m_oXTab.lFillColor = m_lDisabledTabBackColor
        End If
        
        'draw rounded rectangle
        Call RoundRect(m_lHdc, .Left, .Top, .Right, m_iScaleHeight - 1, m_iXRadius, m_iYRadius)
     
        Set m_oXTab.oFont = m_oXTab.InActiveTabFont   'set the font
        
        'set foreground color
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      End If
      
      'do the adjustments for the border
      utFontRect.Left = .Left
      utFontRect.Top = .Top
      utFontRect.Bottom = .Bottom
      utFontRect.Right = .Right - 1

      
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        If utFontRect.Top + m_iIconSize > utFontRect.Bottom + 1 Then '+1 for minor adjustments
          'adjust if going out of current tab's bottom
          iAdjustedIconSize = (utFontRect.Bottom - 4) - utFontRect.Top
        Else
          iAdjustedIconSize = m_iIconSize
        End If
        
        iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
                  
        Select Case m_oXTab.PictureAlign
          Case xAlignLeftEdge, xAlignLeftOfCaption:
            
            iTmpX = utFontRect.Left + 4 'move the icon a little far from left edge
            
          
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If
            
            iOrigLeft = iTmpX + iAdjustedIconSize + 2
            iOrigRight = iOrigLeft + (utFontRect.Right - utFontRect.Left)
            utFontRect.Left = iOrigLeft
             
            
          Case xAlignRightEdge, xAlignRightOfCaption:
            iOrigLeft = utFontRect.Left
            iOrigRight = utFontRect.Right
            'utFontRect.Left = 5
            utFontRect.Right = utFontRect.Right - iAdjustedIconSize - 2
            
        End Select
        
      End If
      
      
      sTmp = utTabInfo.Caption
      
      'calculate the rect to draw the text, also modify the string to get ellipsis etc
      W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
      
      
      iTmpW = utFontRect.Right - utFontRect.Left + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE
      iTmpH = utFontRect.Bottom - utFontRect.Top + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2
      
      
      'do the adjustments to center the text (both vertically and horizontally)
      'utFontRect.Left = (utFontRect.Left - (m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2)) + .Right / 2 - utFontRect.Right / 2
      utFontRect.Left = (utFontRect.Left + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE)
      If utTabInfo.TabPicture Is Nothing Then
        utFontRect.Right = utFontRect.Left + iTmpW
      Else
        utFontRect.Right = utFontRect.Left + iTmpW - m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
      End If
      
      utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
      utFontRect.Bottom = utFontRect.Top + iTmpH
      
      If Not utTabInfo.TabPicture Is Nothing Then
        
        
        Select Case m_oXTab.PictureAlign
          Case xAlignLeftEdge, xAlignLeftOfCaption:
            
            utFontRect.Left = iOrigLeft
          
            'Now draw the text
            W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_SINGLELINE
            
          Case xAlignRightEdge, xAlignRightOfCaption:
   
        
            iTmpW = utFontRect.Right
            utFontRect.Right = iTmpW + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
            
            'Now draw the text
            W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_SINGLELINE
            
            iTmpX = utFontRect.Right + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE - 2
            
            'utFontRect.Right = iTmpW
            
            'make sure our adjustment dosen't make it out of the font area
            If iTmpX + iAdjustedIconSize > iOrigRight Then iTmpX = iOrigRight - iAdjustedIconSize
            
            If utTabInfo.TabPicture.Type = vbPicTypeBitmap And m_oXTab.UseMaskColor Then
              
              Call DrawImage(m_lHdc, utTabInfo.TabPicture.handle, g_pGetRGBFromOLE(m_oXTab.PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            Else
              
              Call m_oXTab.pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
              
            End If

          
        End Select
      Else
            'Now draw the text
            W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_SINGLELINE
      End If
      


      If m_oXTab.bUserMode Then    'only if in the run mode

        If iCnt = m_iActiveTab And m_bIsFocused And m_oXTab.ShowFocusRect Then
          'show the focus rectangle
          Call DrawFocusRect(m_lHdc, utFontRect)
        End If
      End If
    End With
  Next
  
 
    '<EhFooter>
    Exit Sub

Itheme_DrawTabsPropertyPages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.Itheme_DrawTabsPropertyPages " & _
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
      
      .Bottom = .Bottom + 2
    End With
    
    m_oXTab.aryTabs(iCnt) = utTabInfo
  Next
  
  'if the last tab is shorter or longer than the usual size.. then adjust it to perfect size
  utTabInfo.ClickableRect.Right = m_iScaleWidth
  m_oXTab.aryTabs(iCnt - 1) = utTabInfo

  'added to prevent lines etc (we are filling the tab strip with the tab strip color)
  m_oXTab.pLine 0, 0, m_iScaleWidth, IIf(m_iActiveTabHeight > m_iInActiveTabHeight, m_iActiveTabHeight, m_iInActiveTabHeight), m_oXTab.TabStripBackColor, True, True
  
  m_oXTab.iFillStyle = vbFSSolid
  m_oXTab.lForeColor = m_lOuterBorderColor
        
  'Now Draw Each Tab
  For iCnt = 0 To m_iTabCount - 1
    utTabInfo = m_oXTab.aryTabs(iCnt)     'fetch local copy
    
    m_oXTab.lForeColor = m_lOuterBorderColor
    
    With utTabInfo.ClickableRect
      
      If iCnt = m_iActiveTab Then     'if we are drawing the active tab then
        
        
        'we'll use solid colors for background (since we are calling RoundRect API)
        If utTabInfo.Enabled Then
          m_oXTab.lFillColor = m_lActiveTabBackStartColor
        Else
          m_oXTab.lFillColor = m_lDisabledTabBackColor
        End If
        
        'draw rounded rectangle
        Call RoundRect(m_lHdc, .Left, .Top, .Right, m_iScaleHeight - 1, m_iXRadius, m_iYRadius)

        Set m_oXTab.oFont = m_oXTab.ActiveTabFont       'set the font
        
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      Else                        'we are drawing inactive tab
      
        'we'll use solid colors for background (since we are calling RoundRect API)
        If utTabInfo.Enabled Then
          m_oXTab.lFillColor = m_lInActiveTabBackStartColor
        Else
          m_oXTab.lFillColor = m_lDisabledTabBackColor
        End If
        
        'draw rounded rectangle
        Call RoundRect(m_lHdc, .Left, .Top, .Right, m_iScaleHeight - 1, m_iXRadius, m_iYRadius)
     
        Set m_oXTab.oFont = m_oXTab.InActiveTabFont   'set the font
        
        If utTabInfo.Enabled Then
          m_oXTab.lForeColor = m_lInActiveTabForeColor
        Else
          m_oXTab.lForeColor = m_lDisabledTabForeColor
        End If
        
      End If

      'do the adjustments for the border
      utFontRect.Left = .Left
      utFontRect.Top = .Top
      utFontRect.Bottom = .Bottom
      utFontRect.Right = .Right - 1

    If Not utTabInfo.TabPicture Is Nothing Then
        
        If utFontRect.Top + m_iIconSize > utFontRect.Bottom + 1 Then '+1 for minor adjustments
          'adjust if going out of current tab's bottom
          iAdjustedIconSize = (utFontRect.Bottom - 4) - utFontRect.Top
        Else
          iAdjustedIconSize = m_iIconSize
        End If
        
        iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
                  
        Select Case m_oXTab.PictureAlign
          Case xAlignLeftEdge:
            iTmpX = utFontRect.Left + 4
            
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
            iTmpX = utFontRect.Right - iAdjustedIconSize - 4
            
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
  
    '<EhFooter>
    Exit Sub

ITheme_DrawTabsTabbedDialog_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_DrawTabsTabbedDialog " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_MouseDownHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    '<EhHeader>
    On Error GoTo ITheme_MouseDownHanlder_Err
    '</EhHeader>
  Dim iCnt As Integer
  Dim iX As Integer
  Dim iY As Integer
  Dim utTabInfo As TabInfo
  
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
    
    If iX >= utTabInfo.ClickableRect.Left And iX <= utTabInfo.ClickableRect.Right And iY >= utTabInfo.ClickableRect.Top And iY <= utTabInfo.ClickableRect.Bottom And utTabInfo.Enabled Then
      
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_MouseDownHanlder " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub ITheme_MouseMoveHanlder(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
  'not used
    '<EhHeader>
    On Error GoTo ITheme_MouseMoveHanlder_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

ITheme_MouseMoveHanlder_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_MouseMoveHanlder " & _
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_MouseUpHanlder " & _
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_SetControl " & _
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_ShowHideFocus " & _
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
  Dim iTmpW As Integer
  Dim iTmpH As Integer
  Dim sTmp As String
  Dim iAdjustedIconSize As Integer
  Dim iOrigLeft As Integer
  
  
  If Not m_oXTab.bUserMode Then       'only if in the run mode
    
    Exit Sub
    
  End If
  
  If Not m_oXTab.ShowFocusRect Then   'only if Show Focus Rect is true for the control
      
    Exit Sub
    
  End If
  
  utTabInfo = m_oXTab.aryTabs(m_iActiveTab)
  With utTabInfo.ClickableRect
    'do the adjustments for the border
    utFontRect.Left = .Left
    utFontRect.Top = .Top
    utFontRect.Bottom = .Bottom
    utFontRect.Right = .Right - 1

    sTmp = utTabInfo.Caption
    
    If Not utTabInfo.TabPicture Is Nothing Then
      
      If utFontRect.Top + m_iIconSize > utFontRect.Bottom + 1 Then '+1 for minor adjustments
        'adjust if going out of current tab's bottom
        iAdjustedIconSize = (utFontRect.Bottom + 1) - utFontRect.Top
      Else
        iAdjustedIconSize = m_iIconSize
      End If
      
      Select Case m_oXTab.PictureAlign
        Case xAlignLeftEdge, xAlignLeftOfCaption:
          iOrigLeft = utFontRect.Left
          utFontRect.Left = utFontRect.Left + iAdjustedIconSize - 1
      End Select
      
    End If
          
    
    'set the active tab font as current font... used to get proper values when calculating text size
    Set m_oXTab.oFont = m_oXTab.ActiveTabFont
    
    'calculate the rect to draw the text, and get proper string (including ellipsis etc)
    W_DrawText m_lHdc, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
    
    
    iTmpW = utFontRect.Right - utFontRect.Left + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE
    iTmpH = utFontRect.Bottom - utFontRect.Top + m_c_iFOCUS_RECT_AND_TEXT_DISTANCE / 2
    
    
    'do the adjustments to center the text (both vertically and horizontally)
    utFontRect.Left = (utFontRect.Left + m_c_iPROP_PAGE_BORDER_AND_TEXT_DISTANCE)
    utFontRect.Right = utFontRect.Left + iTmpW
    
    utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
    utFontRect.Bottom = utFontRect.Top + iTmpH

    If utTabInfo.Enabled Then             'done to allow proper drawing of focus rect
      m_oXTab.lForeColor = m_lActiveTabForeColor
    Else
      m_oXTab.lForeColor = m_lDisabledTabForeColor
    End If
    
    Select Case m_oXTab.PictureAlign
      Case xAlignRightEdge, xAlignRightOfCaption:
        'utFontRect.Right = utFontRect.Right - 2
    End Select
    
    'show/hide the focus rectangle (drawn in XOR mode, so calling it again with same coords will erase it)
    Call DrawFocusRect(m_lHdc, utFontRect)
  End With
    '<EhFooter>
    Exit Sub

ITheme_ShowHideFocusPropertyPages_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_ShowHideFocusPropertyPages " & _
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
    utFontRect.Left = .Left
    utFontRect.Top = .Top
    utFontRect.Bottom = .Bottom
    utFontRect.Right = .Right - 1

    sTmp = utTabInfo.Caption
    
    'set the active tab font as current font... used to get proper values when calculating text size
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_ShowHideFocusTabbedDialog " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub ITheme_TimerEvent()
  'Do Nothing (will never be called for this theme)
    '<EhHeader>
    On Error GoTo ITheme_TimerEvent_Err
    '</EhHeader>
    '<EhFooter>
    Exit Sub

ITheme_TimerEvent_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_TimerEvent " & _
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
    
    .ActiveTabBackStartColor = &HF8F8F8
    
    'Not used in this Theme
    '.ActiveTabBackEndColor
    
    .ActiveTabForeColor = &HA76D50
    
    .InActiveTabBackStartColor = &HE5E5E5
    
    'Not used in this Theme
    '.InActiveTabBackEndColor
    .InActiveTabForeColor = &H909090
    
    .TabStripBackColor = vbButtonFace
    .DisabledTabBackColor = vbButtonFace
    .DisabledTabForeColor = vb3DDKShadow
    
    '.TopLeftInnerBorderColor = vb3DHighlight
    .OuterBorderColor = &H909090
    '.BottomRightInnerBorderColor = vb3DShadow
    
    'Not Reset for this theme
    ' .HoverColorInverted
  End With
    '<EhFooter>
    Exit Sub

ITheme_ResetColorsToDefault_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.xThemeRoundTabs.ITheme_ResetColorsToDefault " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'=====================================================================================================================


'==== Private Functions ==============================================================================================

'cache the properties from the control (this prevents trips to again and again fetch properties from the user control)
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
  
  m_lActiveTabForeColor = m_oXTab.ActiveTabForeColor
  m_lActiveTabBackStartColor = m_oXTab.ActiveTabBackStartColor
  
  
  m_lInActiveTabForeColor = m_oXTab.InActiveTabForeColor
  m_lInActiveTabBackStartColor = m_oXTab.InActiveTabBackStartColor
  
  m_lDisabledTabBackColor = m_oXTab.DisabledTabBackColor
  m_lDisabledTabForeColor = m_oXTab.DisabledTabForeColor
  m_iXRadius = m_oXTab.XRadius
  m_iYRadius = m_oXTab.YRadius
  
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
           "in ThunVBCC_v1_0.xThemeRoundTabs.pCacheControlProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'=====================================================================================================================

