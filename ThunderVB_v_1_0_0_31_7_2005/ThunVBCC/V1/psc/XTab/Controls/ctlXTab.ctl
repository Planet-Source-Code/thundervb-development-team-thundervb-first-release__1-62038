VERSION 5.00
Begin VB.UserControl XTab 
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2130
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "ctlXTab.ctx":0000
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   142
   ToolboxBitmap   =   "ctlXTab.ctx":002D
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   450
      Top             =   900
   End
End
Attribute VB_Name = "XTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'                                   ||Jai Maata Di||
'
'Project:   XTab
'
'Version:   1.1.1
'
'Original
'Release :  On 1st Oct 2004
'
'
'By:        Neeraj Agrawal
'           (project dedicated to my parents... i love them a lot :) )
'
'Mail:      nja91@yahoo.com
'           neeraj_agrawal_ind@rediffmail.com
'
'
'Revisions:
'
' Release 1.1.1 : On 9th Jan 2005
'         1)Added a method called "CopyTabImagesFromImageList()" which can be called at
'           runtime to copy images from an Image List or compatible control into the tabs
'           (will find a better way to do so in future release(s)), Based
'           on a request by: Mr. Mel Reyes
'
'       --> Bugs Fixed:
'         1)Bug Reported by: Mr. Bob Churchill
'           When Multiple forms containing XTab are unloaded **sometimes** we get
'           a "Client Side Not Available" error.
'           For more details about the error please contact me.
'           Solution: This happens sometimes because the control gets a "GotFocus" event
'           When the form is unloaded. This is strange behaviour and the Ambient object
'           becomes unavailable. To solve this I had to check for Ambient Object before usin it.
'
'           Thanks Bob for reporting the bug
'
' Release 1.1.0 : On 27th Oct 2004
'         1) Added Icon Support for all the themes. Includes Mask Color and other props.
'         2) Removed repeated themes.
'         3) Modified Property Page for better UI.
'
'  Release 1.0.1 : On 05 Oct 2004
'       --> Bugs Fixed:
'         1)Bug Reported by: Maurice
'           Creating an exe with the control in XP and using a Manifest file
'           causes a General Protection Fault "after" exiting the application.
'           This is not a problem with the control, but with the ComCtl32.dll
'           version of XP. This problem occurs even if u have no code at all in
'           the user cotrol. For more details about the error please contact me.
'           Solution found from a article By: Steve McMahon <steve@vbaccelerator.com>
'
'           Thanks Maurice for reporting the bug
'           And Thanks Steave for your wonderful article.
'

'
'
'Major
'Features:  Possible MS Tabbed Dialog Replacement, NO Dependencies,
'           Many Themes to choose from, Flicker free and Optimized
'           Tab Switching, Can have Access keys for Individual tabs,
'           New Events like BeforeTabSwitch,TabSwitch,AfterCompleteInit
'
'Comments:  Thanx for downloading this control. Please vote for me
'           if you like the control.
'
'           I have tried to optimize the code to prevent flickers,
'           memory leaks, slow drawing etc. But as we all know,
'           no matter how good we optimize...there's always
'           room for better optimization. So if you find any
'           bugs/errors/functionality glitches please inform me.
'
'           ------------------------------------------------------
'           Even if you have some complaints you are welcome to
'           discuss them with me. It will be my pleasure to hear
'           from you.
'           -----------------------------------------------------
'
'       (1) This Control has NO external dependencies (Neither Reference
'           nor controls) except the standard controls and dependencies.
'
'       (2) When i was developing the control i tried a lot to make
'           the tabs behave like Microsoft Tabbed Dialogs do at
'           Design time. i.e. being able to switch tabs the same
'           way we switch them at runtime. And not having to right-
'           click the control and select "Edit". I tried many ways
'           but faild to do so. When i was trying a thought came
'           to my mind, since VB is showing us the tab in design time,
'           it is surely instantiated. So I can't trap
'           any events. But since there is a working instance of
'           Control on the form, its events are also available,
'           and VB is freezing the events and not allowing the
'           events to come to control. So i thought maybe subclass
'           is the solution to this problem. And you know what..
'           a lil API Declaration Loader help and it worked... :)
'
'           So now you can see these tabs behave same as Ms Tabs.
'           But as you may be knowing there is a lot of problem
'           with subclass when we want to debug. So i'v added a
'           Conditional compilation directive for the subclass.
'           To prevent subclassing (that is when u are debugging)
'           set the debug flag to 1, To enable subclassing set it
'           to 0 again.
'
'       (3) As i planned to develop a tab control which can
'           mimic Standard Tabs, XP Tabs,Visual Studio Tabs etc,
'           and add some more feature like gradiants etc. I knew
'           the code was going to be big. So i wanted some way to
'           make the drawnig related code independent of the
'           functionality. Only possible solution i saw was to
'           use some Object Oriented ways.. i.e. classes ;) .
'           So finally i implemented all the theme specific code
'           into Classes. All the themes implement from the common
'           Interface "ITheme", this was done to allow consistency
'           among themes.
''            --Problem i faced was to allow functions such as
'           Line, TextHeight,TextWidth etc to be accisible to themes.
'           But didn 't figured out any straight way to do it. So i had
'           to write Wrapper Functions and Properties with Friend
'           Access. If you find any better way of doing this please
'           let me know.
'
'       (4) Earlier i didn't wanted to include tab picture property.
'           As we can't use an Image list with UserControls. And I didn't
'           Wanted to use API way for maintaining and accessing an Image list
'           And also because images stored (even 16x16, 32x32) does slow app
'           the loading time since the data comes from a stream. But ultimately
'           added since many friends among us will surely demand an iconed version
'           of the control.
'
'       (5) To understand the code, breakpoint on major events and procedures
'           in the selected theme.
'
'       (6) Error Handlers need to be implemented.
'
'
'IMPORTANT:  Not All Properties apply to all the themes. Some properties have
'            no effect in some theme, but affect other themes. To set
'            properties in a better way use the property page.
'
'
'Yet To Implement:
'            There are many things i want to implement. Some examples are
'            Tab Orientation, Multiple rows, etc.
'            I'd tried to given user an opp to implement all these and many
'            more by using xThemeOwnerDrawn. But this theme is not complete
'            and not tested also.
'            Some or all of these features will be included in future
'            releases, based on feedback i get.
'            --> SOME FUNKY THEMES ARE ALSO EXPECTED IN FUTURE [ maybe from your side :) ] <--
'
'=====================================================================================================================


Option Explicit


'===Declarations======================================================================================================


Public Type ControlDetails          'Stores Contained Control details
  ControlID As String               '--> Control's id created using Control.Name and Control.Index (index used to allow name conflicts for control arrays)
  TabStop As Boolean                '-->Original tab stop for the control. Since we set the tabstop to false, we must store the original value
End Type

Public Enum Style                     'Style for the tab
  xStyleTabbedDialog                  '-->Tabbed Dialog
  xStylePropertyPages                 '-->Property Pages
End Enum

Public Enum Theme
  xThemeWin9x                         '--> Windows 9x,2000,ME style
  xThemeWinXP                         '--> Windows XP style
  xThemeVisualStudio2003              '--> Visual Studio .Net 2003 style
  xThemeRoundTabs                     '--> Rounded Tabs
  xThemeOwnerDrawn                    '--> Owner drawn. Nothing on our side. We send events to client which he
                                      '    can interpret our events and draw accordingly
End Enum

Public Enum PictureAlign
  
  xAlignLeftEdge                      '-->Left edge of the Tab
  xAlignRightEdge                     '-->Right Edge of the Tab
  xAlignLeftOfCaption                 '-->Left of the caption
  xAlignRightOfCaption                '-->Right of the caption
  
End Enum

Public Enum PictureSize               ' Determines Picture size on tabs
  xSizeSmall
  xSizeLarge
End Enum
'=====================================================================================================================


'===Constants=========================================================================================================

'default property values
Private Const m_def_iTabCount As Integer = 3
Private Const m_def_iActiveTab As Integer = 0
Private Const m_def_iLastActiveTab As Integer = m_def_iActiveTab
Private Const m_def_sCaption As String = "Tab"              'Default caption that is appended to form default name "Tab 0", "Tab 1" etc
Private Const m_def_bTabEnabled As Boolean = True
Private Const m_def_iStyle As Integer = xStyleTabbedDialog
Private Const m_def_iActiveTabHeight As Integer = 20        'Default Height for Active tab
Private Const m_def_iInActiveTabHeight As Integer = 18      'Default Height for Inactive tab
Private Const m_def_iTheme As Integer = xThemeWin9x
Private Const m_def_bShowFocusRect As Boolean = True
Private Const m_def_lHoverColorInvered As Long = &HC43800
Private Const m_def_lActiveTabBackStartColor As Long = vbButtonFace
Private Const m_def_lActiveTabBackEndColor As Long = vbButtonFace
Private Const m_def_iXRadius As Integer = 10   'radius used in the RoundTabs Theme to draw the rounded tab
Private Const m_def_iYRadius As Integer = 10  'radius used in the RoundTabs Theme to draw the rounded tab

Private Const m_def_lInActiveTabBackStartColor As Long = vbButtonFace
Private Const m_def_lInActiveTabBackEndColor As Long = vbButtonFace

Private Const m_def_lActiveTabForeColor As Long = vbButtonText
Private Const m_def_lInActiveTabForeColor As Long = vbButtonText
Private Const m_def_lOuterBorderColor As Long = vb3DDKShadow
Private Const m_def_lTopLeftInnerBorderColor As Long = vb3DHighlight
Private Const m_def_lBottomRightInnerBorderColor As Long = vb3DShadow
Private Const m_def_lDisabledTabBackColor As Long = vb3DShadow
Private Const m_def_lDisabledTabForeColor As Long = vb3DHighlight
'Private Const m_def_lTabStripBackColor as Long =  vbButtonFace    'commented coz we'll use ambient.back color for this

Private Const m_def_iPictureAlign As Integer = xAlignLeftEdge
Private Const m_def_iPictureSize As Integer = xSizeSmall
Private Const m_def_lPictureMaskColor As Long = &HC0C0C0                 'Standard gray color as Transparany color
Private Const m_def_bUseMaskColor As Boolean = True

'=====================================================================================================================

'===Property Variables================================================================================================
Private m_iTabCount As Integer            'The Number of tabs
Private m_iActiveTab As Integer           'Stores the Active Tab index
Private m_iLastActiveTab As Integer       'Stores the Last Active Tab Index
Private m_iActiveTabHeight As Integer     'Active Tab Height
Private m_iInActiveTabHeight As Integer   'InActive Tab Height
Private m_enmStyle As Style               'style for the tab
Private m_enmTheme As Theme               'theme for the tabs
Private m_bShowFocusRect As Boolean       'focus rectangle
Private m_lActiveTabBackStartColor As OLE_COLOR   'Active tab's background color
Private m_lActiveTabBackEndColor As OLE_COLOR
Private m_lInActiveTabBackStartColor As OLE_COLOR 'InActive tab's background color
Private m_lInActiveTabBackEndColor As OLE_COLOR 'InActive tab's background color
Private m_lActiveTabForeColor As OLE_COLOR
Private m_lInActiveTabForeColor As OLE_COLOR
Private m_oActiveTabFont As StdFont
Private m_oInActiveTabFont As StdFont
Private m_lOuterBorderColor As OLE_COLOR
Private m_lTopLeftInnerBorderColor As OLE_COLOR
Private m_lBottomRightInnerBorderColor As OLE_COLOR
Private m_lTabStripBackColor As OLE_COLOR
Private m_lDisabledTabBackColor As OLE_COLOR
Private m_lDisabledTabForeColor As OLE_COLOR
Private m_lHoverColorInverted As OLE_COLOR        'the color used to draw on Mouse Over (Inverted)
Private m_iXRadius As Integer
Private m_iYRadius As Integer
Private m_enmPictureAlign As PictureAlign
Private m_enmPictureSize As PictureSize
Private m_bUseMaskColor As Boolean
'=====================================================================================================================


'===Private Vairbles==================================================================================================
Private m_aryTabs() As TabInfo              'array of tabs
Private m_bAreControlsAdded As Boolean
Private m_bIsFocused As Boolean             'determines if the control is focused
Private m_oTheme As ITheme
Private m_bIsRecursive As Boolean
Private m_bCancelFlag As Boolean            'Used to pass as a cancel flag with the events to the container.

''Commented folllowing line bcoz i wanted to set this from property page,
''but property page does not allow friend member access :( thats why
''commented the code
'Private m_bIgnoreRedraw As Boolean          'Prevent redraw in case we are changing multiple properties (like in property page)

Private m_lOrigWndProc As Long              'stores the original address of the WndProc (used to call CallWindowProc() for default processing)
Private m_lhWndOrigSubClassed As Long       'Handle (hWnd) of the subclassed window.

'Added 5th Oct 2004
Private m_lhModShell32 As Long              'Will store the handle to the Shell 32.dll (we are preloading Shell32.dll in Initialize event and
                                            'unloading it in the Terminate Event to prevent App crashes on XP. For details c the BugFixes section
                                            
'=====================================================================================================================

'=Public Events=======================================================================================================


' Note that bCancel is passed by Reference in below event. This event is called just before a
' tab is being switched, we can prevent tab switch by making bCancel as true
Public Event BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
Attribute BeforeTabSwitch.VB_Description = "This is a new event, this event occurs just before the tab is being switched. User can prevent switching by setting bCancel flag."
                                                                                  
' if we Set bCancel in the BeforeTabSwitch following event will not occur.
Public Event TabSwitch(ByVal iLastActiveTab As Integer)
Attribute TabSwitch.VB_Description = "Occurs After the tabs are switched."

Public Event Click()
Attribute Click.VB_Description = "Click Event."
Public Event DblClick()
Attribute DblClick.VB_Description = "Double Click Event."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "KeyDown Event."
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Key Up Event."
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "KeyPress Event."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Mouse Down Event."
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Mouse Move Event."
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


'IMPORTANT EVENT :  used to solve a bug with original ssTab....
'used to tell container when the tab is completely initialised.
Public Event AfterCompleteInit()
Attribute AfterCompleteInit.VB_Description = "This event occurs when the control is completely loaded and all the contained controls have also been loaded."
Attribute AfterCompleteInit.VB_MemberFlags = "200"
                                      

'following are related to owner draw theme

' sent to inform container that the background needs to be drawn
Public Event DrawBackground(ByVal lHwnd As Long, ByVal lHDC As Long)
Attribute DrawBackground.VB_Description = "Event Occures when owner draw theme is selected."

' sent to inform container that the Tabs needs to be drawn
' (client must take care not to clear existing background drawn in
' DrawBackground event)
Public Event DrawTabs(ByVal lHwnd As Long, ByVal lHDC As Long)
Attribute DrawTabs.VB_Description = "Event Occures when owner draw theme is selected."

' sent to inform container that the Tabs needs to be drawn
' as a result of active tab change (due to mouse/keyboard or through code)
' (client must take care not to clear existing background drawn in
' DrawBackground event)
' This event provides an optimized way redraw only the tabs. Instead of
' redrawing complete control
Public Event DrawOnActiveTabChange(ByVal lHwnd As Long, ByVal lHDC As Long)
Attribute DrawOnActiveTabChange.VB_Description = "Event Occures when owner draw theme is selected."

' sent to inform container that the Focus rectangle needs to be drawn or erased
' This event provides an optimized way redraw only the focus rect. Instead of
' redrawing complete control. Normally focus rect is drawn in XOR mode, so that
' it can be cleared in a second call
Public Event ShowHideFocus(ByVal lHwnd As Long, ByVal lHDC As Long, ByVal bIsFocused As Boolean)
Attribute ShowHideFocus.VB_Description = "Event Occures when owner draw theme is selected."

'=====================================================================================================================


'===Public Properties=================================================================================================

Public Property Get LastActiveTab() As Integer        'Note: Read Only Property: returns the last active tab index
Attribute LastActiveTab.VB_Description = "Readonly property. Returns Last active tab's index."
Attribute LastActiveTab.VB_ProcData.VB_Invoke_Property = ";Position"
    '<EhHeader>
    On Error GoTo LastActiveTab_Err
    '</EhHeader>
  LastActiveTab = m_iLastActiveTab
    '<EhFooter>
    Exit Property

LastActiveTab_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.LastActiveTab " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get handle() As Long
Attribute handle.VB_Description = "Handle for the control."
    '<EhHeader>
    On Error GoTo handle_Err
    '</EhHeader>
  handle = UserControl.hwnd
    '<EhFooter>
    Exit Property

handle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.handle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property
Public Property Get DC() As Long
Attribute DC.VB_Description = "Returns the Device Contect for the control."
    '<EhHeader>
    On Error GoTo DC_Err
    '</EhHeader>
  DC = UserControl.hdc
    '<EhFooter>
    Exit Property

DC_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.DC " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get Enabled() As Boolean        'dont confuse with TabEnable Poperty that's for individual tab
Attribute Enabled.VB_Description = "Does this need an explanation ??"
                                                ' and this is for the whole control
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
  Enabled = UserControl.Enabled
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.Enabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let Enabled(bNewValue As Boolean) 'dont confuse with TabEnable Poperty that's for individual tab
    '<EhHeader>
    On Error GoTo Enabled_Err
    '</EhHeader>
  UserControl.Enabled() = bNewValue               ' and this is for the whole control
  
  Refresh
  
  PropertyChanged "Enabled"
  
    '<EhFooter>
    Exit Property

Enabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.Enabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TabCount() As Integer             'the number of tabs
Attribute TabCount.VB_Description = "The Total number of tabs."
Attribute TabCount.VB_ProcData.VB_Invoke_Property = ";Behavior"
    '<EhHeader>
    On Error GoTo TabCount_Err
    '</EhHeader>
  TabCount = m_iTabCount
    '<EhFooter>
    Exit Property

TabCount_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabCount " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TabCount(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo TabCount_Err
    '</EhHeader>
  If iNewValue < 1 Then
    Err.Raise 380 ' invalid property value
    Exit Property    'if out of range
  End If
  
  m_iTabCount = iNewValue
  
  Call pHandleTabCount    'handle the change in tabcount (i.e. resize/initialize the array of tabs)
  Refresh

  PropertyChanged ("TabCount")
    '<EhFooter>
    Exit Property

TabCount_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabCount " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ActiveTab() As Integer        'active tab index
Attribute ActiveTab.VB_Description = "The Active tab whose controls are visible."
Attribute ActiveTab.VB_ProcData.VB_Invoke_Property = ";Position"
Attribute ActiveTab.VB_MemberFlags = "200"
    '<EhHeader>
    On Error GoTo ActiveTab_Err
    '</EhHeader>
  ActiveTab = m_iActiveTab
    '<EhFooter>
    Exit Property

ActiveTab_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTab " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


Public Property Let ActiveTab(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo ActiveTab_Err
    '</EhHeader>
  If iNewValue < 0 Or iNewValue >= m_iTabCount Then
    Err.Raise 380 ' invalid property value
    Exit Property    'if out of range
  End If
  
  'if already we are on the same tab (this is important or else all
  ' the contained controls for active tab will be moved to -75000 and so...
  If iNewValue = m_iActiveTab Then Exit Property
  
  m_bCancelFlag = False
  
  'raise event and confirm that the user want to allow the tab switch
  RaiseEvent BeforeTabSwitch(iNewValue, m_bCancelFlag)
  
  
  'if user set the cancel flag in the BeforeTabSwitch event
  If m_bCancelFlag Then Exit Property
  
  
  m_iLastActiveTab = m_iActiveTab       'store current tab in last active tab
  
  m_iActiveTab = iNewValue
  
  Call pHandleContainedControls         'Show/Hide Controls for active tab
  
  Call pDrawOnActiveTabChange           'now draw the tabs with changed state
 
  PropertyChanged "ActiveTab"
  RaiseEvent TabSwitch(m_iLastActiveTab)
    '<EhFooter>
    Exit Property

ActiveTab_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTab " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TabStyle() As Style           'Style for the tab
Attribute TabStyle.VB_Description = "Style for the tab."
Attribute TabStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo TabStyle_Err
    '</EhHeader>
  TabStyle = m_enmStyle
    '<EhFooter>
    Exit Property

TabStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TabStyle(enmNewStyle As Style)
    '<EhHeader>
    On Error GoTo TabStyle_Err
    '</EhHeader>
  m_enmStyle = enmNewStyle
  
  Refresh
  
  PropertyChanged "TabStyle"
    '<EhFooter>
    Exit Property

TabStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TabTheme() As Theme
Attribute TabTheme.VB_Description = "Theme for the tab."
Attribute TabTheme.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo TabTheme_Err
    '</EhHeader>
  TabTheme = m_enmTheme
    '<EhFooter>
    Exit Property

TabTheme_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabTheme " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TabTheme(enmNewTheme As Theme)
    '<EhHeader>
    On Error GoTo TabTheme_Err
    '</EhHeader>
  m_enmTheme = enmNewTheme
  
  pSetThemeObject (m_enmTheme)
  
  Refresh   'redraw
  
  PropertyChanged ("TabTheme")
    '<EhFooter>
    Exit Property

TabTheme_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabTheme " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ActiveTabHeight() As Integer
Attribute ActiveTabHeight.VB_Description = "Height for the Active Tab."
Attribute ActiveTabHeight.VB_ProcData.VB_Invoke_Property = ";Scale"
    '<EhHeader>
    On Error GoTo ActiveTabHeight_Err
    '</EhHeader>
  ActiveTabHeight = m_iActiveTabHeight
    '<EhFooter>
    Exit Property

ActiveTabHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ActiveTabHeight(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo ActiveTabHeight_Err
    '</EhHeader>
  m_iActiveTabHeight = iNewValue
  
  Refresh
  
  PropertyChanged "ActiveTabHeight"
    '<EhFooter>
    Exit Property

ActiveTabHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get InActiveTabHeight() As Integer
Attribute InActiveTabHeight.VB_Description = "Height for the InActiveTab."
Attribute InActiveTabHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo InActiveTabHeight_Err
    '</EhHeader>
  InActiveTabHeight = m_iInActiveTabHeight
    '<EhFooter>
    Exit Property

InActiveTabHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let InActiveTabHeight(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo InActiveTabHeight_Err
    '</EhHeader>
  m_iInActiveTabHeight = iNewValue
  
  Refresh
  
  
  PropertyChanged "InActiveTabHeight"
    '<EhFooter>
    Exit Property

InActiveTabHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Determines wheher focus rectangle is displayed."
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo ShowFocusRect_Err
    '</EhHeader>
  ShowFocusRect = m_bShowFocusRect
    '<EhFooter>
    Exit Property

ShowFocusRect_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ShowFocusRect " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ShowFocusRect(bNewValue As Boolean)
    '<EhHeader>
    On Error GoTo ShowFocusRect_Err
    '</EhHeader>
  m_bShowFocusRect = bNewValue
  PropertyChanged (ShowFocusRect)
    '<EhFooter>
    Exit Property

ShowFocusRect_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ShowFocusRect " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ActiveTabBackStartColor() As OLE_COLOR
Attribute ActiveTabBackStartColor.VB_Description = "Start  color for the Active Tab's Background Gradient"
Attribute ActiveTabBackStartColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo ActiveTabBackStartColor_Err
    '</EhHeader>
  ActiveTabBackStartColor = m_lActiveTabBackStartColor
    '<EhFooter>
    Exit Property

ActiveTabBackStartColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabBackStartColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ActiveTabBackStartColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ActiveTabBackStartColor_Err
    '</EhHeader>
  m_lActiveTabBackStartColor = lNewValue
  
  Refresh
  
  PropertyChanged ("ActiveTabBackStartColor")
    '<EhFooter>
    Exit Property

ActiveTabBackStartColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabBackStartColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ActiveTabBackEndColor() As OLE_COLOR
Attribute ActiveTabBackEndColor.VB_Description = "End color for the Active Tab's Background Gradient"
Attribute ActiveTabBackEndColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo ActiveTabBackEndColor_Err
    '</EhHeader>
  ActiveTabBackEndColor = m_lActiveTabBackEndColor
    '<EhFooter>
    Exit Property

ActiveTabBackEndColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabBackEndColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ActiveTabBackEndColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ActiveTabBackEndColor_Err
    '</EhHeader>
  m_lActiveTabBackEndColor = lNewValue
  
  Refresh
  
  PropertyChanged ("ActiveTabBackEndColor")
    '<EhFooter>
    Exit Property

ActiveTabBackEndColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabBackEndColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get InActiveTabBackStartColor() As OLE_COLOR
Attribute InActiveTabBackStartColor.VB_Description = "Start color for the InActive Tab's Gradient."
Attribute InActiveTabBackStartColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo InActiveTabBackStartColor_Err
    '</EhHeader>
  InActiveTabBackStartColor = m_lInActiveTabBackStartColor
    '<EhFooter>
    Exit Property

InActiveTabBackStartColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabBackStartColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let InActiveTabBackStartColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo InActiveTabBackStartColor_Err
    '</EhHeader>
  m_lInActiveTabBackStartColor = lNewValue
  
  Refresh
  
  PropertyChanged ("InActiveTabBackStartColor")
    '<EhFooter>
    Exit Property

InActiveTabBackStartColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabBackStartColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get InActiveTabBackEndColor() As OLE_COLOR
Attribute InActiveTabBackEndColor.VB_Description = "End color for the InActive Tab's Gradient."
Attribute InActiveTabBackEndColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo InActiveTabBackEndColor_Err
    '</EhHeader>
  InActiveTabBackEndColor = m_lInActiveTabBackEndColor
    '<EhFooter>
    Exit Property

InActiveTabBackEndColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabBackEndColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let InActiveTabBackEndColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo InActiveTabBackEndColor_Err
    '</EhHeader>
  m_lInActiveTabBackEndColor = lNewValue
  
  Refresh
  
  PropertyChanged ("InActiveTabBackEndColor")
    '<EhFooter>
    Exit Property

InActiveTabBackEndColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabBackEndColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get ActiveTabForeColor() As OLE_COLOR
Attribute ActiveTabForeColor.VB_Description = "Foreground Text  Color  for Active Tab."
Attribute ActiveTabForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo ActiveTabForeColor_Err
    '</EhHeader>
  ActiveTabForeColor = m_lActiveTabForeColor
    '<EhFooter>
    Exit Property

ActiveTabForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let ActiveTabForeColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo ActiveTabForeColor_Err
    '</EhHeader>
  m_lActiveTabForeColor = lNewValue
  
  Refresh
  
  PropertyChanged ("ActiveTabForeColor")
    '<EhFooter>
    Exit Property

ActiveTabForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


Public Property Get InActiveTabForeColor() As OLE_COLOR
Attribute InActiveTabForeColor.VB_Description = "Foreground Text color for Inactive Tab."
Attribute InActiveTabForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo InActiveTabForeColor_Err
    '</EhHeader>
  InActiveTabForeColor = m_lInActiveTabForeColor
    '<EhFooter>
    Exit Property

InActiveTabForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let InActiveTabForeColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo InActiveTabForeColor_Err
    '</EhHeader>
  m_lInActiveTabForeColor = lNewValue
  
  Refresh
  
  PropertyChanged ("InActiveTabForeColor")
    '<EhFooter>
    Exit Property

InActiveTabForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


Public Property Get ActiveTabFont() As StdFont
Attribute ActiveTabFont.VB_Description = "Font used for the Active Tab."
Attribute ActiveTabFont.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo ActiveTabFont_Err
    '</EhHeader>
  Set ActiveTabFont = m_oActiveTabFont
    '<EhFooter>
    Exit Property

ActiveTabFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set ActiveTabFont(oNewFnt As StdFont)
    '<EhHeader>
    On Error GoTo ActiveTabFont_Err
    '</EhHeader>
  Set m_oActiveTabFont = oNewFnt
  
  Refresh
  
  PropertyChanged ("ActiveTabFont")
    '<EhFooter>
    Exit Property

ActiveTabFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ActiveTabFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get InActiveTabFont() As StdFont
Attribute InActiveTabFont.VB_Description = "Font used for Inactive Tab."
Attribute InActiveTabFont.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo InActiveTabFont_Err
    '</EhHeader>
  Set InActiveTabFont = m_oInActiveTabFont
    '<EhFooter>
    Exit Property

InActiveTabFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set InActiveTabFont(oNewFnt As StdFont)
    '<EhHeader>
    On Error GoTo InActiveTabFont_Err
    '</EhHeader>
  Set m_oInActiveTabFont = oNewFnt
  
  Refresh
  
  PropertyChanged ("InActiveTabFont")
    '<EhFooter>
    Exit Property

InActiveTabFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.InActiveTabFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get OuterBorderColor() As OLE_COLOR
Attribute OuterBorderColor.VB_Description = "The color used to draw outer border."
Attribute OuterBorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo OuterBorderColor_Err
    '</EhHeader>
  OuterBorderColor = m_lOuterBorderColor
    '<EhFooter>
    Exit Property

OuterBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.OuterBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let OuterBorderColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo OuterBorderColor_Err
    '</EhHeader>
  m_lOuterBorderColor = lNewValue
  Refresh
  PropertyChanged ("OuterBorderColor")
    '<EhFooter>
    Exit Property

OuterBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.OuterBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TopLeftInnerBorderColor() As OLE_COLOR
Attribute TopLeftInnerBorderColor.VB_Description = "Color used to draw the Left and Top Inner Borders."
Attribute TopLeftInnerBorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo TopLeftInnerBorderColor_Err
    '</EhHeader>
  TopLeftInnerBorderColor = m_lTopLeftInnerBorderColor
    '<EhFooter>
    Exit Property

TopLeftInnerBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TopLeftInnerBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TopLeftInnerBorderColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo TopLeftInnerBorderColor_Err
    '</EhHeader>
  m_lTopLeftInnerBorderColor = lNewValue
  Refresh
  PropertyChanged ("TopLeftInnerBorderColor")
    '<EhFooter>
    Exit Property

TopLeftInnerBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TopLeftInnerBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get BottomRightInnerBorderColor() As OLE_COLOR
Attribute BottomRightInnerBorderColor.VB_Description = "This is the shadow color for Right and Bottom lines."
    '<EhHeader>
    On Error GoTo BottomRightInnerBorderColor_Err
    '</EhHeader>
  BottomRightInnerBorderColor = m_lBottomRightInnerBorderColor
    '<EhFooter>
    Exit Property

BottomRightInnerBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.BottomRightInnerBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let BottomRightInnerBorderColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo BottomRightInnerBorderColor_Err
    '</EhHeader>
  m_lBottomRightInnerBorderColor = lNewValue
  Refresh
  PropertyChanged ("BottomRightInnerBorderColor")
    '<EhFooter>
    Exit Property

BottomRightInnerBorderColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.BottomRightInnerBorderColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TabStripBackColor() As OLE_COLOR
Attribute TabStripBackColor.VB_Description = "Background color of the tabstrip."
Attribute TabStripBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo TabStripBackColor_Err
    '</EhHeader>
  TabStripBackColor = m_lTabStripBackColor
    '<EhFooter>
    Exit Property

TabStripBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabStripBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TabStripBackColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo TabStripBackColor_Err
    '</EhHeader>
  m_lTabStripBackColor = lNewValue
  
  Refresh
  
  PropertyChanged ("TabStripBackColor")
    '<EhFooter>
    Exit Property

TabStripBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabStripBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisabledTabBackColor() As OLE_COLOR
Attribute DisabledTabBackColor.VB_Description = "Back color of the disabled Tab."
Attribute DisabledTabBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo DisabledTabBackColor_Err
    '</EhHeader>
  DisabledTabBackColor = m_lDisabledTabBackColor
    '<EhFooter>
    Exit Property

DisabledTabBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.DisabledTabBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisabledTabBackColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo DisabledTabBackColor_Err
    '</EhHeader>
  m_lDisabledTabBackColor = lNewValue
  
  Refresh
  
  PropertyChanged ("DisabledTabBackColor")
    '<EhFooter>
    Exit Property

DisabledTabBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.DisabledTabBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get DisabledTabForeColor() As OLE_COLOR
Attribute DisabledTabForeColor.VB_Description = "Foreground Text color for this control."
Attribute DisabledTabForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo DisabledTabForeColor_Err
    '</EhHeader>
  DisabledTabForeColor = m_lDisabledTabForeColor
    '<EhFooter>
    Exit Property

DisabledTabForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.DisabledTabForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let DisabledTabForeColor(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo DisabledTabForeColor_Err
    '</EhHeader>
  m_lDisabledTabForeColor = lNewValue
  
  Refresh
  
  PropertyChanged ("DisabledTabForeColor")
    '<EhFooter>
    Exit Property

DisabledTabForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.DisabledTabForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TabCaption(iTabIndex As Integer) As String
Attribute TabCaption.VB_Description = "Get/Set Active Tab's caption."
Attribute TabCaption.VB_ProcData.VB_Invoke_Property = ";Text"
    '<EhHeader>
    On Error GoTo TabCaption_Err
    '</EhHeader>
  If iTabIndex > -1 And iTabIndex < m_iTabCount Then
    TabCaption = m_aryTabs(iTabIndex).Caption
  End If
    '<EhFooter>
    Exit Property

TabCaption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabCaption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TabCaption(iTabIndex As Integer, sTabCaption As String)
    '<EhHeader>
    On Error GoTo TabCaption_Err
    '</EhHeader>
  Dim sTmp As String
  Dim lPos As Long
  
  If iTabIndex > -1 And iTabIndex < m_iTabCount Then
  
  
    ' first get the existing caption's access key and remove it from the
    ' current "AccessKeys" property for the control
    sTmp = Replace$(m_aryTabs(iTabIndex).Caption, "&&", Chr$(1))
    
    lPos = InStrRev(sTmp, "&")
    If m_aryTabs(iTabIndex).AccessKey <> 0 Then
      'remove from AccessKey
      AccessKeys = Replace$(AccessKeys, UCase$(Mid$(sTmp, lPos + 1, 1)), "")
    End If
    
    m_aryTabs(iTabIndex).Caption = sTabCaption
    
    'Now get the new caption's access key and append it to the "AccessKeys" property
    sTmp = Replace$(m_aryTabs(iTabIndex).Caption, "&&", Chr$(1))
    
    lPos = InStrRev(sTmp, "&")
    If lPos Then
      'note: we are using Ucase$... since we store all the access keys in uppercase only
      m_aryTabs(iTabIndex).AccessKey = Asc(UCase$(Mid$(sTmp, lPos + 1, 1)))
      AccessKeys = AccessKeys & Chr$(m_aryTabs(iTabIndex).AccessKey)
    Else
      m_aryTabs(iTabIndex).AccessKey = 0    'reset the access key
    End If
   
    PropertyChanged "TabCaption"
    
    Refresh
  Else
    
    Err.Raise 9 ' Subscript Out of Range
    Exit Property    'if out of range
  End If
    '<EhFooter>
    Exit Property

TabCaption_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabCaption " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get TabPicture(iTabIndex As Integer) As StdPicture
Attribute TabPicture.VB_Description = "Picture for the active tab. Only used by the supported themes."
Attribute TabPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo TabPicture_Err
    '</EhHeader>
  If iTabIndex > -1 And iTabIndex < m_iTabCount Then
    Set TabPicture = m_aryTabs(iTabIndex).TabPicture
  Else
    Err.Raise 9 ' Subscript Out of Range
    Exit Property    'if out of range
  End If
    '<EhFooter>
    Exit Property

TabPicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabPicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Set TabPicture(iTabIndex As Integer, oTabPicture As StdPicture)
    '<EhHeader>
    On Error GoTo TabPicture_Err
    '</EhHeader>
  If iTabIndex > -1 And iTabIndex < m_iTabCount Then
    Set m_aryTabs(iTabIndex).TabPicture = oTabPicture
      
    Refresh
      
    PropertyChanged ("TabPicture")
  Else
    Err.Raise 9 ' Subscript Out of Range
    Exit Property    'if out of range
  End If
    '<EhFooter>
    Exit Property

TabPicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabPicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


Public Property Get TabEnabled(iTabIndex As Integer) As Boolean
Attribute TabEnabled.VB_Description = "Individual tab's enabled property."
Attribute TabEnabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    '<EhHeader>
    On Error GoTo TabEnabled_Err
    '</EhHeader>
  If iTabIndex > -1 And iTabIndex < m_iTabCount Then
    TabEnabled = m_aryTabs(iTabIndex).Enabled
  End If
    '<EhFooter>
    Exit Property

TabEnabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabEnabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let TabEnabled(iTabIndex As Integer, bNewValue As Boolean)
    '<EhHeader>
    On Error GoTo TabEnabled_Err
    '</EhHeader>
  If iTabIndex > -1 And iTabIndex < m_iTabCount Then
    m_aryTabs(iTabIndex).Enabled = bNewValue
    PropertyChanged "TabEnabled"
    
    Refresh
    
  Else
    Err.Raise 9 ' Subscript Out of Range
    Exit Property    'if out of range
  End If
    '<EhFooter>
    Exit Property

TabEnabled_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.TabEnabled " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get HoverColorInverted() As OLE_COLOR
Attribute HoverColorInverted.VB_Description = "This is the XORED color which will be used while when mouse hovers in the supported themes."
Attribute HoverColorInverted.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo HoverColorInverted_Err
    '</EhHeader>
  HoverColorInverted = m_lHoverColorInverted
    '<EhFooter>
    Exit Property

HoverColorInverted_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.HoverColorInverted " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let HoverColorInverted(lNewValue As OLE_COLOR)
    '<EhHeader>
    On Error GoTo HoverColorInverted_Err
    '</EhHeader>
  m_lHoverColorInverted = lNewValue
  
  Refresh
  
  PropertyChanged ("HoverColorInverted")
    '<EhFooter>
    Exit Property

HoverColorInverted_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.HoverColorInverted " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


Public Property Get XRadius() As Integer
Attribute XRadius.VB_Description = "Horizontal radius of the rounded rectangle for the ""Rounded Tabs"" Theme."
Attribute XRadius.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo XRadius_Err
    '</EhHeader>
  XRadius = m_iXRadius
    '<EhFooter>
    Exit Property

XRadius_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.XRadius " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let XRadius(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo XRadius_Err
    '</EhHeader>
  m_iXRadius = iNewValue
  Refresh
  PropertyChanged "XRadius"
    '<EhFooter>
    Exit Property

XRadius_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.XRadius " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get YRadius() As Integer
Attribute YRadius.VB_Description = "Vertical radius of the rounded rectangle for the ""Rounded Tabs"" Theme."
Attribute YRadius.VB_ProcData.VB_Invoke_Property = ";Appearance"
    '<EhHeader>
    On Error GoTo YRadius_Err
    '</EhHeader>
  YRadius = m_iYRadius
    '<EhFooter>
    Exit Property

YRadius_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.YRadius " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let YRadius(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo YRadius_Err
    '</EhHeader>
  m_iYRadius = iNewValue
  Refresh
  PropertyChanged "YRadius"
    '<EhFooter>
    Exit Property

YRadius_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.YRadius " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get PictureAlign() As PictureAlign
    '<EhHeader>
    On Error GoTo PictureAlign_Err
    '</EhHeader>
  PictureAlign = m_enmPictureAlign
    '<EhFooter>
    Exit Property

PictureAlign_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.PictureAlign " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let PictureAlign(iNewValue As PictureAlign)
    '<EhHeader>
    On Error GoTo PictureAlign_Err
    '</EhHeader>
  m_enmPictureAlign = iNewValue
    
  Refresh   'reflect changes
    
  PropertyChanged "PictureAlign"
    '<EhFooter>
    Exit Property

PictureAlign_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.PictureAlign " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get PictureSize() As PictureSize
    '<EhHeader>
    On Error GoTo PictureSize_Err
    '</EhHeader>
  PictureSize = m_enmPictureSize
    '<EhFooter>
    Exit Property

PictureSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.PictureSize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let PictureSize(iNewSize As PictureSize)
    '<EhHeader>
    On Error GoTo PictureSize_Err
    '</EhHeader>
  m_enmPictureSize = iNewSize
  
  Refresh   'redraw changes
  
  PropertyChanged ("PictureSize") 'mark for serialization
    '<EhFooter>
    Exit Property

PictureSize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.PictureSize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get PictureMaskColor() As OLE_COLOR
    '<EhHeader>
    On Error GoTo PictureMaskColor_Err
    '</EhHeader>
  PictureMaskColor = UserControl.MaskColor
    '<EhFooter>
    Exit Property

PictureMaskColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.PictureMaskColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let PictureMaskColor(lNewColor As OLE_COLOR)
    '<EhHeader>
    On Error GoTo PictureMaskColor_Err
    '</EhHeader>
  UserControl.MaskColor = lNewColor
  
  Refresh
  
  PropertyChanged "PictureMaskColor"
    '<EhFooter>
    Exit Property

PictureMaskColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.PictureMaskColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Get UseMaskColor() As Boolean
    '<EhHeader>
    On Error GoTo UseMaskColor_Err
    '</EhHeader>
  UseMaskColor = m_bUseMaskColor
    '<EhFooter>
    Exit Property

UseMaskColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UseMaskColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Public Property Let UseMaskColor(bNewValue As Boolean)
    '<EhHeader>
    On Error GoTo UseMaskColor_Err
    '</EhHeader>
  m_bUseMaskColor = bNewValue
  
  Refresh
  
  PropertyChanged "UseMaskColor"
    '<EhFooter>
    Exit Property

UseMaskColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UseMaskColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property


'=====================================================================================================================


'===Public Procedues==================================================================================================

'Copy From a standard Image List to current tabs images

Public Sub CopyTabImagesFromImageList(ByRef oIml As Object)
    '<EhHeader>
    On Error GoTo CopyTabImagesFromImageList_Err
    '</EhHeader>
  Dim iTmp As Integer
  
  On Error GoTo Finally 'if the number of images is less than number of tabs error may occur
  
  For iTmp = 0 To UBound(m_aryTabs)
    Set m_aryTabs(iTmp).TabPicture = Nothing  'Free Existing Picture
    Set m_aryTabs(iTmp).TabPicture = oIml.ListImages(iTmp + 1).Picture
  Next
  
Finally:
  Refresh
  'Do nothing as this is possibly because the no. of images is less
    '<EhFooter>
    Exit Sub

CopyTabImagesFromImageList_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.CopyTabImagesFromImageList " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Reset all the colors to the default colors
Public Sub ResetAllColors()
    '<EhHeader>
    On Error GoTo ResetAllColors_Err
    '</EhHeader>
  
  
  m_bIsRecursive = True   'Prevent Redrawing untill all the properties are set (in the ResetColorsToDefault function)
  
  'Call the Theme specific function for resetting the colors
  Call m_oTheme.ResetColorsToDefault
  
  m_bIsRecursive = False 'Prevent Redrawing untill all the properties are set (in the ResetColorsToDefault function)
  
  Refresh       'Now Force Redraw
  
    '<EhFooter>
    Exit Sub

ResetAllColors_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.ResetAllColors " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'=====================================================================================================================


'===Friend Properties========(These properties are not accessible outside current project)============================
Friend Property Get lHwnd() As Long
    '<EhHeader>
    On Error GoTo lHwnd_Err
    '</EhHeader>
  lHwnd = hwnd
    '<EhFooter>
    Exit Property

lHwnd_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.lHwnd " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get lHDC() As Long
    '<EhHeader>
    On Error GoTo lHDC_Err
    '</EhHeader>
  lHDC = hdc
    '<EhFooter>
    Exit Property

lHDC_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.lHDC " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let lBackColor(lNewValue As Long)
    '<EhHeader>
    On Error GoTo lBackColor_Err
    '</EhHeader>
  BackColor = lNewValue
    '<EhFooter>
    Exit Property

lBackColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.lBackColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let lForeColor(lNewValue As Long)
    '<EhHeader>
    On Error GoTo lForeColor_Err
    '</EhHeader>
  ForeColor = lNewValue
    '<EhFooter>
    Exit Property

lForeColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.lForeColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let lFillColor(lNewValue As Long)
    '<EhHeader>
    On Error GoTo lFillColor_Err
    '</EhHeader>
  FillColor = lNewValue
    '<EhFooter>
    Exit Property

lFillColor_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.lFillColor " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get iFillStyle() As Integer
    '<EhHeader>
    On Error GoTo iFillStyle_Err
    '</EhHeader>
  iFillStyle = FillStyle
    '<EhFooter>
    Exit Property

iFillStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.iFillStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let iFillStyle(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo iFillStyle_Err
    '</EhHeader>
  FillStyle = iNewValue
    '<EhFooter>
    Exit Property

iFillStyle_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.iFillStyle " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get iDrawMode() As Integer
    '<EhHeader>
    On Error GoTo iDrawMode_Err
    '</EhHeader>
  iDrawMode = DrawMode
    '<EhFooter>
    Exit Property

iDrawMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.iDrawMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let iDrawMode(iNewValue As Integer)
    '<EhHeader>
    On Error GoTo iDrawMode_Err
    '</EhHeader>
  DrawMode = iNewValue
    '<EhFooter>
    Exit Property

iDrawMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.iDrawMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Set oFont(oNewFont As StdFont)
    '<EhHeader>
    On Error GoTo oFont_Err
    '</EhHeader>
  Set Font = oNewFont
    '<EhFooter>
    Exit Property

oFont_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.oFont " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get iScaleWidth() As Integer
    '<EhHeader>
    On Error GoTo iScaleWidth_Err
    '</EhHeader>
  iScaleWidth = ScaleWidth
    '<EhFooter>
    Exit Property

iScaleWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.iScaleWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get iScaleHeight() As Integer
    '<EhHeader>
    On Error GoTo iScaleHeight_Err
    '</EhHeader>
  iScaleHeight = ScaleHeight
    '<EhFooter>
    Exit Property

iScaleHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.iScaleHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'Friend Property Get bIgnoreRedraw() As Boolean
'  bIgnoreRedraw = m_bIgnoreRedraw
'End Property
'
'Friend Property Let bIgnoreRedraw(iNewValue As Boolean)
'  m_bIgnoreRedraw = iNewValue
'End Property

Friend Property Get aryTabs(iIndex As Integer) As TabInfo
    '<EhHeader>
    On Error GoTo aryTabs_Err
    '</EhHeader>
  aryTabs = m_aryTabs(iIndex)
    '<EhFooter>
    Exit Property

aryTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.aryTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let aryTabs(iIndex As Integer, utNewValue As TabInfo)
    '<EhHeader>
    On Error GoTo aryTabs_Err
    '</EhHeader>
  m_aryTabs(iIndex) = utNewValue
    '<EhFooter>
    Exit Property

aryTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.aryTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get bIsFocused() As Boolean
    '<EhHeader>
    On Error GoTo bIsFocused_Err
    '</EhHeader>
  bIsFocused = m_bIsFocused
    '<EhFooter>
    Exit Property

bIsFocused_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.bIsFocused " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let bIsFocused(bNewValue As Boolean)
    '<EhHeader>
    On Error GoTo bIsFocused_Err
    '</EhHeader>
  m_bIsFocused = bNewValue
    '<EhFooter>
    Exit Property

bIsFocused_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.bIsFocused " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get bUserMode() As Boolean
    '<EhHeader>
    On Error GoTo bUserMode_Err
    '</EhHeader>
  On Error Resume Next            'Used to prevent an error which occurs when a
                                  'form with this control gets unloaded
                                  'This is strange but the control gets a "GotFocus" event
                                  'sometimes when the container form is unloaded
  bUserMode = Ambient.UserMode
    '<EhFooter>
    Exit Property

bUserMode_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.bUserMode " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get bAutoRedraw() As Boolean
    '<EhHeader>
    On Error GoTo bAutoRedraw_Err
    '</EhHeader>
  bAutoRedraw = AutoRedraw
    '<EhFooter>
    Exit Property

bAutoRedraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.bAutoRedraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Let bAutoRedraw(bNewValue As Boolean)
    '<EhHeader>
    On Error GoTo bAutoRedraw_Err
    '</EhHeader>
  AutoRedraw = bNewValue
    '<EhFooter>
    Exit Property

bAutoRedraw_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.bAutoRedraw " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

Friend Property Get oAmbient() As Object
    '<EhHeader>
    On Error GoTo oAmbient_Err
    '</EhHeader>
  Set oAmbient = Ambient
    '<EhFooter>
    Exit Property

oAmbient_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.oAmbient " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Property

'=====================================================================================================================

'===Friend Procedures============(These Procedures are not accessible outside current project)========================

' Function wraps local Line Function . This is done to allow access
' from outside classes
Friend Sub pLine(ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long, ByVal lColor As Long, Optional ByVal bBox As Boolean, Optional ByVal bFill As Boolean)
    '<EhHeader>
    On Error GoTo pLine_Err
    '</EhHeader>
  If bBox And bFill Then
    Line (lX1, lY1)-(lX2, lY2), lColor, BF
  ElseIf bBox Then
    Line (lX1, lY1)-(lX2, lY2), lColor, B
  Else
    Line (lX1, lY1)-(lX2, lY2), lColor
  End If
    '<EhFooter>
    Exit Sub

pLine_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pLine " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Friend Function pTextWidth(sText As String) As Integer
    '<EhHeader>
    On Error GoTo pTextWidth_Err
    '</EhHeader>
  pTextWidth = TextWidth(sText)
    '<EhFooter>
    Exit Function

pTextWidth_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pTextWidth " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Friend Function pTextHeight(sText As String) As Integer
    '<EhHeader>
    On Error GoTo pTextHeight_Err
    '</EhHeader>
  pTextHeight = TextHeight(sText)
    '<EhFooter>
    Exit Function

pTextHeight_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pTextHeight " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

' Wrapper function for Form's PaintPicture Method
Friend Sub pPaintPicture(oStdPic As StdPicture, ByVal iX1 As Integer, ByVal iY1 As Integer, Optional ByVal iWidth As Integer = -1, Optional ByVal iHeight As Integer = -1)
    '<EhHeader>
    On Error GoTo pPaintPicture_Err
    '</EhHeader>
  If iWidth = -1 And iHeight = -1 Then
    Call PaintPicture(oStdPic, iX1, iY1)
  ElseIf iWidth = -1 Then
    Call PaintPicture(oStdPic, iX1, iY1, , iHeight)
  ElseIf iHeight = -1 Then
    Call PaintPicture(oStdPic, iX1, iY1, iWidth)
  Else
    Call PaintPicture(oStdPic, iX1, iY1, iWidth, iHeight)
  End If
    '<EhFooter>
    Exit Sub

pPaintPicture_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pPaintPicture " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' called from xThemeOwnerDrawn to raise local event
Friend Function pRaiseDrawBackground()
    '<EhHeader>
    On Error GoTo pRaiseDrawBackground_Err
    '</EhHeader>
  RaiseEvent DrawBackground(hwnd, hdc)
    '<EhFooter>
    Exit Function

pRaiseDrawBackground_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pRaiseDrawBackground " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

' called from xThemeOwnerDrawn to raise local event
Friend Function pRaiseDrawTabs()
    '<EhHeader>
    On Error GoTo pRaiseDrawTabs_Err
    '</EhHeader>
  RaiseEvent DrawTabs(hwnd, hdc)
    '<EhFooter>
    Exit Function

pRaiseDrawTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pRaiseDrawTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

' called from xThemeOwnerDrawn to raise local event
Friend Sub pRaiseDrawOnActiveTabChange()
    '<EhHeader>
    On Error GoTo pRaiseDrawOnActiveTabChange_Err
    '</EhHeader>
  RaiseEvent DrawOnActiveTabChange(hwnd, hdc)
    '<EhFooter>
    Exit Sub

pRaiseDrawOnActiveTabChange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pRaiseDrawOnActiveTabChange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' called from xThemeOwnerDrawn to raise local event
Friend Sub pRaiseShowHideFocus()
    '<EhHeader>
    On Error GoTo pRaiseShowHideFocus_Err
    '</EhHeader>
  RaiseEvent ShowHideFocus(hwnd, hdc, m_bIsFocused)
    '<EhFooter>
    Exit Sub

pRaiseShowHideFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pRaiseShowHideFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Friend Sub pCls()
    '<EhHeader>
    On Error GoTo pCls_Err
    '</EhHeader>
  Cls
    '<EhFooter>
    Exit Sub

pCls_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pCls " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Friend Sub pRefresh()
    '<EhHeader>
    On Error GoTo pRefresh_Err
    '</EhHeader>
  Refresh
    '<EhFooter>
    Exit Sub

pRefresh_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pRefresh " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' function used to set the timer to ON/OFF
' based on the parameter value
Friend Sub pSetTimer(ByVal iInterval As Integer)
    '<EhHeader>
    On Error GoTo pSetTimer_Err
    '</EhHeader>
  If iInterval = 0 Then
    tmrMain.Enabled = False
  Else
    tmrMain.Interval = iInterval
    tmrMain.Enabled = True
  End If
    '<EhFooter>
    Exit Sub

pSetTimer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pSetTimer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


' Kinda over-riden function for pFillCurveSolid, performs same job,
' but takes integers instead of Rect as parameter
Friend Sub pFillCurvedSolid(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal lColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '<EhHeader>
    On Error GoTo pFillCurvedSolid_Err
    '</EhHeader>
  Dim utRect As RECT
  
  utRect.Left = lLeft
  utRect.Top = lTop
  utRect.Right = lRight
  utRect.Bottom = lBottom
  
  Call pFillCurvedSolidR(utRect, lColor, iCurveValue, bCurveLeft, bCurveRight)
    '<EhFooter>
    Exit Sub

pFillCurvedSolid_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pFillCurvedSolid " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' function used to Fill a rectangular area by Solid color.
' This function can draw using the curve value to generate a rounded rect kinda effect
Friend Sub pFillCurvedSolidR(utRect As RECT, ByVal lColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '<EhHeader>
    On Error GoTo pFillCurvedSolidR_Err
    '</EhHeader>

  Dim intCnt As Integer
    
  If iCurveValue = -1 Then
    For intCnt = utRect.Top To utRect.Bottom
      Line (utRect.Left, intCnt)-(utRect.Right, intCnt), lColor
    Next
  Else
    If bCurveLeft And bCurveRight Then
      For intCnt = utRect.Top To utRect.Bottom
        Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right - iCurveValue, intCnt), lColor
            
        If iCurveValue > 0 Then
          iCurveValue = iCurveValue - 1
        End If
      Next
    ElseIf bCurveLeft Then
      For intCnt = utRect.Top To utRect.Bottom
        Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right, intCnt), lColor
            
        If iCurveValue > 0 Then
          iCurveValue = iCurveValue - 1
        End If
      Next
    Else    'curve right
      For intCnt = utRect.Top To utRect.Bottom
        Line (utRect.Left, intCnt)-(utRect.Right - iCurveValue, intCnt), lColor
            
        If iCurveValue > 0 Then
          iCurveValue = iCurveValue - 1
        End If
      Next
    End If
  End If
    '<EhFooter>
    Exit Sub

pFillCurvedSolidR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pFillCurvedSolidR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Kinda over-riden function for pFillCurvedGradientR, performs same job,
' but takes integers instead of Rect as parameter
Friend Sub pFillCurvedGradient(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '<EhHeader>
    On Error GoTo pFillCurvedGradient_Err
    '</EhHeader>
  Dim utRect As RECT
  
  utRect.Left = lLeft
  utRect.Top = lTop
  utRect.Right = lRight
  utRect.Bottom = lBottom
  
  Call pFillCurvedGradientR(utRect, lStartColor, lEndColor, iCurveValue, bCurveLeft, bCurveRight)
    '<EhFooter>
    Exit Sub

pFillCurvedGradient_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pFillCurvedGradient " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' function used to Fill a rectangular area by Gradient
' This function can draw using the curve value to generate a rounded rect kinda effect
Friend Sub pFillCurvedGradientR(utRect As RECT, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '<EhHeader>
    On Error GoTo pFillCurvedGradientR_Err
    '</EhHeader>
    
  Dim sngRedInc As Single, sngGreenInc As Single, sngBlueInc As Single
  Dim sngRed As Single, sngGreen As Single, sngBlue As Single
    
  lStartColor = g_pGetRGBFromOLE(lStartColor)
  lEndColor = g_pGetRGBFromOLE(lEndColor)
    
  Dim intCnt As Integer
  sngRedInc = (pGetRValue(lEndColor) - pGetRValue(lStartColor)) / (utRect.Bottom - utRect.Top)
  sngGreenInc = (pGetGValue(lEndColor) - pGetGValue(lStartColor)) / (utRect.Bottom - utRect.Top)
  sngBlueInc = (pGetBValue(lEndColor) - pGetBValue(lStartColor)) / (utRect.Bottom - utRect.Top)
    
  sngRed = pGetRValue(lStartColor)
  sngGreen = pGetGValue(lStartColor)
  sngBlue = pGetBValue(lStartColor)
    
  If iCurveValue = -1 Then
    For intCnt = utRect.Top To utRect.Bottom
      Line (utRect.Left, intCnt)-(utRect.Right, intCnt), RGB(sngRed, sngGreen, sngBlue)
      sngRed = sngRed + sngRedInc
      sngGreen = sngGreen + sngGreenInc
      sngBlue = sngBlue + sngBlueInc
    Next
  Else
    If bCurveLeft And bCurveRight Then
      For intCnt = utRect.Top To utRect.Bottom
        Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right - iCurveValue, intCnt), RGB(sngRed, sngGreen, sngBlue)
  
        sngRed = sngRed + sngRedInc
        sngGreen = sngGreen + sngGreenInc
        sngBlue = sngBlue + sngBlueInc
            
        If iCurveValue > 0 Then
          iCurveValue = iCurveValue - 1
        End If
      Next
    ElseIf bCurveLeft Then
      For intCnt = utRect.Top To utRect.Bottom
        Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right, intCnt), RGB(sngRed, sngGreen, sngBlue)
  
        sngRed = sngRed + sngRedInc
        sngGreen = sngGreen + sngGreenInc
        sngBlue = sngBlue + sngBlueInc
            
        If iCurveValue > 0 Then
          iCurveValue = iCurveValue - 1
        End If
      Next
    Else    'curve right
      For intCnt = utRect.Top To utRect.Bottom
        Line (utRect.Left, intCnt)-(utRect.Right - iCurveValue, intCnt), RGB(sngRed, sngGreen, sngBlue)
  
        sngRed = sngRed + sngRedInc
        sngGreen = sngGreen + sngGreenInc
        sngBlue = sngBlue + sngBlueInc
            
        If iCurveValue > 0 Then
          iCurveValue = iCurveValue - 1
        End If
      Next
    End If
  End If
    '<EhFooter>
    Exit Sub

pFillCurvedGradientR_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pFillCurvedGradientR " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


' Local WndProc. This function is called from the main WndProc in the modAPI.
' Its for subclassing.
Friend Function pWindowProc(ByVal lHwnd As Long, ByVal lMsg As Long, ByVal lwParam As Long, ByVal llParam As Long) As Long
    '<EhHeader>
    On Error GoTo pWindowProc_Err
    '</EhHeader>
  If lMsg = WM_LBUTTONDOWN Then
      Call UserControl_MouseDown(vbLeftButton, 0, llParam And &HFFFF&, llParam \ &H10000 And &HFFFF&)
  End If
  pWindowProc = A_CallWindowProc(m_lOrigWndProc, lHwnd, lMsg, lwParam, llParam)
    '<EhFooter>
    Exit Function

pWindowProc_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pWindowProc " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'=====================================================================================================================


'=====Event Handlers==================================================================================================

Private Sub UserControl_Initialize()
  'Added 5th Oct 2004
  'Preload Shell32.dll , For details c the BugFixes section
    '<EhHeader>
    On Error GoTo UserControl_Initialize_Err
    '</EhHeader>
  m_lhModShell32 = LoadLibrary("Shell32.dll")
    '<EhFooter>
    Exit Sub

UserControl_Initialize_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_Initialize " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error GoTo UserControl_Terminate_Err
    '</EhHeader>
  Call pDestroyResources    'call to free up system resources by deleting pictures etc
  
  Call pEndSubClass         'end subclassing
  
  'Added 5th Oct 2004
  'Free Shell32.dll , For details c the BugFixes section
  FreeLibrary (m_lhModShell32)
    '<EhFooter>
    Exit Sub

UserControl_Terminate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_Terminate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_InitProperties()
    '<EhHeader>
    On Error GoTo UserControl_InitProperties_Err
    '</EhHeader>
  Dim iCnt As Integer
  
  'initialize all properties
  
  m_iTabCount = m_def_iTabCount
  m_iActiveTab = m_def_iActiveTab
  m_iLastActiveTab = m_def_iLastActiveTab
  
  m_lActiveTabBackStartColor = m_def_lActiveTabBackStartColor
  m_lActiveTabBackEndColor = m_def_lActiveTabBackEndColor
  
  m_lInActiveTabBackStartColor = m_def_lInActiveTabBackStartColor
  m_lInActiveTabBackEndColor = m_def_lInActiveTabBackEndColor
  
  m_lActiveTabForeColor = m_def_lActiveTabForeColor
  m_lInActiveTabForeColor = m_def_lInActiveTabForeColor
  m_iActiveTabHeight = m_def_iActiveTabHeight
  m_iInActiveTabHeight = m_def_iInActiveTabHeight
  m_bShowFocusRect = m_def_bShowFocusRect
  Set m_oActiveTabFont = Ambient.Font
  m_oActiveTabFont.bold = True    'by default the active tab's font is bold
  Set m_oInActiveTabFont = Ambient.Font
  m_lOuterBorderColor = m_def_lOuterBorderColor
  m_lTopLeftInnerBorderColor = m_def_lTopLeftInnerBorderColor
  m_lBottomRightInnerBorderColor = m_def_lBottomRightInnerBorderColor
  m_lTabStripBackColor = Ambient.BackColor
  m_lDisabledTabBackColor = m_def_lDisabledTabBackColor
  m_lDisabledTabForeColor = m_def_lDisabledTabForeColor
  m_lHoverColorInverted = m_def_lHoverColorInvered
  m_enmPictureAlign = m_def_iPictureAlign
  m_enmPictureSize = m_def_iPictureSize
  m_bUseMaskColor = m_def_bUseMaskColor
  UserControl.MaskColor = m_def_lPictureMaskColor
  
  
  
  m_iXRadius = m_def_iXRadius
  m_iYRadius = m_def_iYRadius
  
  
  m_enmStyle = m_def_iStyle
  
  ReDim m_aryTabs(m_iTabCount - 1)            'redim the tabs array
  
  'initialize the tabs
  For iCnt = 0 To m_iTabCount - 1
    m_aryTabs(iCnt).Caption = m_def_sCaption & " " & iCnt
    Set m_aryTabs(iCnt).ContainedControlsDetails = New Collection
    m_aryTabs(iCnt).Enabled = m_def_bTabEnabled
  Next
  
  
  pSetThemeObject (m_enmTheme)    'set initial theme object to the m_enmTheme
  
  'if its not user mode then call the code to start subclassing
  If Not Ambient.UserMode Then
    pStartSubClass
  End If
  
    '<EhFooter>
    Exit Sub

UserControl_InitProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_InitProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_ReadProperties_Err
    '</EhHeader>
  Dim iCnt As Integer
  Dim iCnt2 As Integer
  Dim iCCCount As Integer
  Dim oControlDetails As ControlDetails
  
  
  'Read previously saved the property values
  
  m_iActiveTab = PropBag.ReadProperty("ActiveTab", m_def_iActiveTab)
  m_iActiveTabHeight = PropBag.ReadProperty("ActiveTabHeight", m_def_iActiveTabHeight)
  m_iInActiveTabHeight = PropBag.ReadProperty("InActiveTabHeight", m_def_iInActiveTabHeight)
  m_enmStyle = PropBag.ReadProperty("TabStyle", m_def_iStyle)
  
  
  m_enmTheme = PropBag.ReadProperty("TabTheme", m_def_iTheme)
  
  pSetThemeObject (m_enmTheme)   'set initial theme object to the m_enmTheme
  
  m_bShowFocusRect = PropBag.ReadProperty("ShowFocusRect", m_def_bShowFocusRect)
  
  m_lActiveTabBackStartColor = PropBag.ReadProperty("ActiveTabBackStartColor", m_def_lActiveTabBackStartColor)
  m_lActiveTabBackEndColor = PropBag.ReadProperty("ActiveTabBackEndColor", m_def_lActiveTabBackEndColor)
  
  m_lInActiveTabBackStartColor = PropBag.ReadProperty("InActiveTabBackStartColor", m_def_lInActiveTabBackStartColor)
  m_lInActiveTabBackEndColor = PropBag.ReadProperty("InActiveTabBackEndColor", m_def_lInActiveTabBackEndColor)
  
  m_lActiveTabForeColor = PropBag.ReadProperty("ActiveTabForeColor", m_def_lActiveTabForeColor)
  m_lInActiveTabForeColor = PropBag.ReadProperty("InActiveTabForeColor", m_def_lInActiveTabForeColor)
  Set m_oActiveTabFont = PropBag.ReadProperty("ActiveTabFont", Ambient.Font)
  Set m_oInActiveTabFont = PropBag.ReadProperty("InActiveTabFont", Ambient.Font)
  m_lOuterBorderColor = PropBag.ReadProperty("OuterBorderColor", m_def_lOuterBorderColor)
  m_lTopLeftInnerBorderColor = PropBag.ReadProperty("TopLeftInnerBorderColor", m_def_lTopLeftInnerBorderColor)
  m_lBottomRightInnerBorderColor = PropBag.ReadProperty("BottomRightInnerBorderColor", m_def_lBottomRightInnerBorderColor)
  m_lTabStripBackColor = PropBag.ReadProperty("TabStripBackColor", Ambient.BackColor)
  m_lDisabledTabBackColor = PropBag.ReadProperty("DisabledTabBackColor", m_def_lDisabledTabBackColor)
  m_lDisabledTabForeColor = PropBag.ReadProperty("DisabledTabForeColor", m_def_lDisabledTabForeColor)
  m_lHoverColorInverted = PropBag.ReadProperty("HoverColorInverted", m_def_lHoverColorInvered)
  m_iXRadius = PropBag.ReadProperty("XRadius", m_def_iXRadius)
  m_iYRadius = PropBag.ReadProperty("YRadius", m_def_iYRadius)
  m_enmPictureAlign = PropBag.ReadProperty("PictureAlign", m_def_iPictureAlign)
  m_enmPictureSize = PropBag.ReadProperty("PictureSize", m_def_iPictureSize)
  
  m_bUseMaskColor = PropBag.ReadProperty("UseMaskColor", m_def_bUseMaskColor)
  UserControl.MaskColor = PropBag.ReadProperty("PictureMaskColor", m_def_lPictureMaskColor)
  
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  
  
  
  m_iTabCount = PropBag.ReadProperty("TabCount", m_def_iTabCount)
  
  ReDim m_aryTabs(m_iTabCount - 1)            'redim the tabs array
  
  For iCnt = 0 To m_iTabCount - 1
      
      
    m_aryTabs(iCnt).Caption = PropBag.ReadProperty("TabCaption(" & iCnt & ")", m_def_sCaption & iCnt + 1)
      
      
    m_aryTabs(iCnt).Enabled = PropBag.ReadProperty("TabEnabled(" & iCnt & ")", m_def_bTabEnabled)
    m_aryTabs(iCnt).AccessKey = PropBag.ReadProperty("TabAccessKey(" & iCnt & ")", 0)
    Set m_aryTabs(iCnt).TabPicture = PropBag.ReadProperty("TabPicture(" & iCnt & ")", Nothing)
      
    iCCCount = PropBag.ReadProperty("TabContCtrlCnt(" & iCnt & ")", 0)
      
    Set m_aryTabs(iCnt).ContainedControlsDetails = New Collection 'Instantiate
      
    'read all the control ids
    For iCnt2 = 1 To iCCCount
      oControlDetails.ControlID = PropBag.ReadProperty("Tab(" & iCnt & ")ContCtrlCap(" & iCnt2 & ")", "")
      'oControlDetails.TabStop = True
      Call m_aryTabs(iCnt).ContainedControlsDetails.Add(oControlDetails, oControlDetails.ControlID)
    Next
  Next
  
  Call pAssignAccessKeys        'since new caption data is read. Now extract and set the acess keys for control
  
  
  
  'if its not user mode then call the code to start subclassing
  If Not Ambient.UserMode Then
    pStartSubClass
  End If
  
    '<EhFooter>
    Exit Sub

UserControl_ReadProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_ReadProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error GoTo UserControl_WriteProperties_Err
    '</EhHeader>
  Dim iCnt As Integer
  Dim iCnt2 As Integer
  
  
  
  ' save the existing controls to the collection, i.e. if controls are placed on the
  ' tab and this event is called, the new controls must be added too. Following is the
  ' call for it
  Call pAddRemainingControlsToActiveTab
  
  
  Call PropBag.WriteProperty("TabCount", m_iTabCount, m_def_iTabCount)
  
  For iCnt = 0 To m_iTabCount - 1
      
    Call PropBag.WriteProperty("TabCaption(" & iCnt & ")", m_aryTabs(iCnt).Caption, m_def_sCaption & iCnt + 1)
    Call PropBag.WriteProperty("TabEnabled(" & iCnt & ")", m_aryTabs(iCnt).Enabled, True)
    Call PropBag.WriteProperty("TabAccessKey(" & iCnt & ")", m_aryTabs(iCnt).AccessKey, 0)
    Call PropBag.WriteProperty("TabPicture(" & iCnt & ")", m_aryTabs(iCnt).TabPicture, Nothing)
      
    If Not m_aryTabs(iCnt).ContainedControlsDetails Is Nothing Then
      Call PropBag.WriteProperty("TabContCtrlCnt(" & iCnt & ")", m_aryTabs(iCnt).ContainedControlsDetails.Count, 0)
        
      For iCnt2 = 1 To m_aryTabs(iCnt).ContainedControlsDetails.Count
        Call PropBag.WriteProperty("Tab(" & iCnt & ")ContCtrlCap(" & iCnt2 & ")", m_aryTabs(iCnt).ContainedControlsDetails(iCnt2).ControlID)
      Next
        
    End If
  Next
  
  
  Call PropBag.WriteProperty("ActiveTab", m_iActiveTab, m_def_iActiveTab)
  Call PropBag.WriteProperty("ActiveTabHeight", m_iActiveTabHeight, m_def_iActiveTabHeight)
  Call PropBag.WriteProperty("InActiveTabHeight", m_iInActiveTabHeight, m_def_iInActiveTabHeight)
  Call PropBag.WriteProperty("TabStyle", m_enmStyle, m_def_iStyle)
  Call PropBag.WriteProperty("TabTheme", m_enmTheme, m_def_iTheme)
  Call PropBag.WriteProperty("ShowFocusRect", m_bShowFocusRect, m_def_bShowFocusRect)
  
  Call PropBag.WriteProperty("ActiveTabBackStartColor", m_lActiveTabBackStartColor, m_def_lActiveTabBackStartColor)
  Call PropBag.WriteProperty("ActiveTabBackEndColor", m_lActiveTabBackEndColor, m_def_lActiveTabBackEndColor)
  
  Call PropBag.WriteProperty("InActiveTabBackStartColor", m_lInActiveTabBackStartColor, m_def_lInActiveTabBackStartColor)
  Call PropBag.WriteProperty("InActiveTabBackEndColor", m_lInActiveTabBackEndColor, m_def_lInActiveTabBackEndColor)
  
  Call PropBag.WriteProperty("ActiveTabForeColor", m_lActiveTabForeColor, m_def_lActiveTabForeColor)
  Call PropBag.WriteProperty("InActiveTabForeColor", m_lInActiveTabForeColor, m_def_lInActiveTabForeColor)
  Call PropBag.WriteProperty("ActiveTabFont", m_oActiveTabFont, Ambient.Font)
  Call PropBag.WriteProperty("InActiveTabFont", m_oInActiveTabFont, Ambient.Font)
  Call PropBag.WriteProperty("OuterBorderColor", m_lOuterBorderColor, m_def_lOuterBorderColor)
  Call PropBag.WriteProperty("TopLeftInnerBorderColor", m_lTopLeftInnerBorderColor, m_def_lTopLeftInnerBorderColor)
  Call PropBag.WriteProperty("BottomRightInnerBorderColor", m_lBottomRightInnerBorderColor, m_def_lBottomRightInnerBorderColor)
  Call PropBag.WriteProperty("TabStripBackColor", m_lTabStripBackColor, Ambient.BackColor)
  Call PropBag.WriteProperty("DisabledTabBackColor", m_lDisabledTabBackColor, m_def_lDisabledTabBackColor)
  Call PropBag.WriteProperty("DisabledTabForeColor", m_lDisabledTabForeColor, m_def_lDisabledTabForeColor)
  Call PropBag.WriteProperty("HoverColorInverted", m_lHoverColorInverted, m_def_lHoverColorInvered)
  Call PropBag.WriteProperty("XRadius", m_iXRadius, m_def_iXRadius)
  Call PropBag.WriteProperty("YRadius", m_iYRadius, m_def_iYRadius)
  Call PropBag.WriteProperty("PictureAlign", m_enmPictureAlign, m_def_iPictureAlign)
  Call PropBag.WriteProperty("PictureSize", m_enmPictureSize, m_def_iPictureSize)
  Call PropBag.WriteProperty("UseMaskColor", m_bUseMaskColor, m_def_bUseMaskColor)
  Call PropBag.WriteProperty("PictureMaskColor", UserControl.MaskColor, m_def_lPictureMaskColor)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)

    '<EhFooter>
    Exit Sub

UserControl_WriteProperties_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_WriteProperties " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub tmrMain_Timer()
    '<EhHeader>
    On Error GoTo tmrMain_Timer_Err
    '</EhHeader>
  m_oTheme.TimerEvent
    '<EhFooter>
    Exit Sub

tmrMain_Timer_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.tmrMain_Timer " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_Click()
    '<EhHeader>
    On Error GoTo UserControl_Click_Err
    '</EhHeader>
  RaiseEvent Click
    '<EhFooter>
    Exit Sub

UserControl_Click_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_Click " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_DblClick()
    '<EhHeader>
    On Error GoTo UserControl_DblClick_Err
    '</EhHeader>
  RaiseEvent DblClick
    '<EhFooter>
    Exit Sub

UserControl_DblClick_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_DblClick " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyDown_Err
    '</EhHeader>
  
  
  If m_bIsRecursive Then Exit Sub   'if its a recursive call then exit the sub
  
  ' raise event, note: Byref arguments user can change there value to
  ' control how tabs behave on key down
  RaiseEvent KeyDown(KeyCode, Shift)
  
  m_bIsRecursive = True
  AutoRedraw = True
  
  'c if proper key pressed
  If (KeyCode = vbKeyPageDown And ((Shift And vbCtrlMask) > 0)) Or (KeyCode = vbKeyRight) Then
    'if right key or the Ctrl+PageDown Key is pressed
  
    If m_iActiveTab < m_iTabCount - 1 Then    'if standing on some middle tab
      If m_aryTabs(m_iActiveTab + 1).Enabled Then

        ActiveTab = m_iActiveTab + 1 'then increment tab by 1
      End If
    Else                                      'we are standing on the last tab
      If m_aryTabs(0).Enabled Then
        ActiveTab = 0                 'then set it to 0
      End If
    End If
      
  ElseIf KeyCode = vbKeyPageUp And ((Shift And vbCtrlMask) > 0) Or (KeyCode = vbKeyLeft) Then
    'if left key or the Ctrl+PageUp Key is pressed
  
    If m_iActiveTab > 0 Then    'if standing on some middle tab
      If m_iTabCount > 1 Then
        If m_aryTabs(m_iActiveTab - 1).Enabled Then
          ActiveTab = ActiveTab - 1   'then decrement tab by 1
        End If
      End If
      
    Else                        'we are standing on the first tab
        
      If m_iTabCount > 1 Then
        If m_aryTabs(m_iTabCount - 1).Enabled Then
          ActiveTab = m_iTabCount - 1   'then set it to last tab
        End If
      End If
    End If
  End If
  
  AutoRedraw = False
  m_bIsRecursive = False
  
  
    '<EhFooter>
    Exit Sub

UserControl_KeyDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_KeyDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyPress_Err
    '</EhHeader>
  RaiseEvent KeyPress(KeyAscii)
    '<EhFooter>
    Exit Sub

UserControl_KeyPress_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_KeyPress " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo UserControl_KeyUp_Err
    '</EhHeader>
  RaiseEvent KeyUp(KeyCode, Shift)
    '<EhFooter>
    Exit Sub

UserControl_KeyUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_KeyUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseDown_Err
    '</EhHeader>
    
  If m_bIsRecursive Then Exit Sub   'prevent recursion
  
  ' raise event, note: ByRef arguments user can change there value to
  ' control how tabs behave on Mouse Down
  RaiseEvent MouseDown(Button, Shift, X, Y)
  
  If Button = vbLeftButton Then     'only if left mouse down
  
    m_bIsRecursive = True
    AutoRedraw = True
    
    Call pHandleMouseDown(Button, Shift, X, Y)  'call theme specific HandleMouseDown
                                                'function to allow drawing.
    
    AutoRedraw = False
    m_bIsRecursive = False
  
  End If
  
    '<EhFooter>
    Exit Sub

UserControl_MouseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_MouseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' raise event, note: ByRef arguments user can change there value to
  ' control how tabs behave on Mouse Down
    '<EhHeader>
    On Error GoTo UserControl_MouseMove_Err
    '</EhHeader>
  RaiseEvent MouseMove(Button, Shift, X, Y)
  
  Call pHandleMouseMove(Button, Shift, X, Y)  'call theme specific HandleMouseDown
                                              'function to allow drawing.
  
    '<EhFooter>
    Exit Sub

UserControl_MouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_MouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '<EhHeader>
    On Error GoTo UserControl_MouseUp_Err
    '</EhHeader>
  
  If m_bIsRecursive Then Exit Sub
  
  ' raise event, note: ByRef arguments user can change there value to
  ' control how tabs behave on Mouse Up
  RaiseEvent MouseUp(Button, Shift, X, Y)
  
  If Button = vbLeftButton Then     'only if left button
    m_bIsRecursive = True
    AutoRedraw = True
    
    Call pHandleMouseUp(Button, Shift, X, Y)  'call theme specific HandleMouseDown
                                              'function to allow drawing.
      
    AutoRedraw = False
    m_bIsRecursive = False
  End If
    '<EhFooter>
    Exit Sub

UserControl_MouseUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_MouseUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


Private Sub UserControl_Paint()
    '<EhHeader>
    On Error GoTo UserControl_Paint_Err
    '</EhHeader>
  
  
  If m_bIsRecursive Then Exit Sub     'prevent recursion. (flag set from calling function)

  'If m_bIgnoreRedraw Then Exit Sub
  
  m_bIsRecursive = True
  AutoRedraw = True
  
    
  If Not m_bAreControlsAdded Then     'if the controls have not been added
  ' This is a big problem with Container ActiveX controls. :(
  ' untill show/paint they does not provide you information about contained controls
  ' that's why i'd provided the "AfterCompleteInit" event so that application can
  ' know that the control finished loading.
  
  
  
    If Ambient.UserMode Then
      Call pStoreOriginalTabStopValues      'store the original tabstop values
      
      m_bAreControlsAdded = True            'indicate control is initialised
      
      RaiseEvent AfterCompleteInit          'tell container that we've finished loading
    Else
      m_bAreControlsAdded = True            'indicate that control is initialised
    End If
  End If
  
  Call pDrawMe                              'now draw ourselves
  
  AutoRedraw = False
  m_bIsRecursive = False
    '<EhFooter>
    Exit Sub

UserControl_Paint_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_Paint " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo UserControl_AccessKeyPress_Err
    '</EhHeader>
  Dim iCnt As Integer
  
  'since we are using the access keys in uppercase, convert any lowercase keys to uppercase before comparision
  If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32  'convert to uppercase
  End If
  
  'compare each with the stored access keys
  For iCnt = 0 To m_iTabCount - 1
    If m_aryTabs(iCnt).AccessKey = KeyAscii And iCnt <> m_iActiveTab And m_aryTabs(iCnt).Enabled Then
      
      'if we find the pressed key as access key for any tab,
      ' simply make that tab active
      
      If m_bIsRecursive Then Exit Sub
      m_bIsRecursive = True
      AutoRedraw = True
  
      ActiveTab = iCnt
      
      AutoRedraw = False
      m_bIsRecursive = False
       
      Exit For
    End If
  Next
    '<EhFooter>
    Exit Sub

UserControl_AccessKeyPress_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_AccessKeyPress " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_GotFocus()
    '<EhHeader>
    On Error GoTo UserControl_GotFocus_Err
    '</EhHeader>
  m_bIsFocused = True
  Call pShowHideFocus
    '<EhFooter>
    Exit Sub

UserControl_GotFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_GotFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub UserControl_LostFocus()
    '<EhHeader>
    On Error GoTo UserControl_LostFocus_Err
    '</EhHeader>
  m_bIsFocused = False
  Call pShowHideFocus
    '<EhFooter>
    Exit Sub

UserControl_LostFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.UserControl_LostFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub
'=====================================================================================================================

'====Private functions================================================================================================

' Function draws both Background and the Tabs
Private Sub pDrawMe()
    '<EhHeader>
    On Error GoTo pDrawMe_Err
    '</EhHeader>
  Call pDrawTabBackground
  Call pDrawTabs
    '<EhFooter>
    Exit Sub

pDrawMe_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pDrawMe " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Function draws only Background
Private Sub pDrawTabBackground()
    '<EhHeader>
    On Error GoTo pDrawTabBackground_Err
    '</EhHeader>
  m_oTheme.DrawBackground
    '<EhFooter>
    Exit Sub

pDrawTabBackground_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pDrawTabBackground " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Function draws only Tabs
Private Sub pDrawTabs()
    '<EhHeader>
    On Error GoTo pDrawTabs_Err
    '</EhHeader>
  m_oTheme.DrawTabs
    '<EhFooter>
    Exit Sub

pDrawTabs_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pDrawTabs " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Function calls MouseDownHanlder of current theme
Private Sub pHandleMouseDown(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    '<EhHeader>
    On Error GoTo pHandleMouseDown_Err
    '</EhHeader>
  Call m_oTheme.MouseDownHanlder(iButton, iShift, sngX, sngY)
    '<EhFooter>
    Exit Sub

pHandleMouseDown_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pHandleMouseDown " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Function calls MouseUpHanlder of current theme
Private Sub pHandleMouseUp(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    '<EhHeader>
    On Error GoTo pHandleMouseUp_Err
    '</EhHeader>
  Call m_oTheme.MouseUpHanlder(iButton, iShift, sngX, sngY)
    '<EhFooter>
    Exit Sub

pHandleMouseUp_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pHandleMouseUp " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Function calls MouseMove of current theme
Private Sub pHandleMouseMove(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    '<EhHeader>
    On Error GoTo pHandleMouseMove_Err
    '</EhHeader>
  Call m_oTheme.MouseMoveHanlder(iButton, iShift, sngX, sngY)
    '<EhFooter>
    Exit Sub

pHandleMouseMove_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pHandleMouseMove " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


' VERY IMPORTANT FUNCTION:
' Handles the appearing and disappearing of controls for the current tab and
' last active tab
Private Sub pHandleContainedControls()
    '<EhHeader>
    On Error GoTo pHandleContainedControls_Err
    '</EhHeader>
  Dim oCtl As Control
  Dim sControlId As String
  Dim lLeft As Long
  Dim oControlDetails As ControlDetails
  
  On Error Resume Next
  
  If m_aryTabs(m_iLastActiveTab).ContainedControlsDetails.Count > 0 Then
    Set m_aryTabs(m_iLastActiveTab).ContainedControlsDetails = Nothing          'clear the existing collection, since there is no Col.clear method :)
    Set m_aryTabs(m_iLastActiveTab).ContainedControlsDetails = New Collection
  End If
  
  lLeft = -75000    'the decrement value for contained control's Left Property(inactive tabs)
  
  For Each oCtl In ContainedControls
    
    sControlId = pGetControlId(oCtl)    'Get ControlID (control.name & control.index)
    
    lLeft = oCtl.Left
    If lLeft > -10000 Then    ' if there was some error in the above line the value of iLeft will be -75000,
                              ' even if there is no error this means that the control is in the current tab
                              ' and must be added to the tab's contained controls array
      
      
      oControlDetails.ControlID = sControlId
      oControlDetails.TabStop = oCtl.TabStop      'store original tab stop
      
      m_aryTabs(m_iLastActiveTab).ContainedControlsDetails.Add oControlDetails, oControlDetails.ControlID
     
      
      
      oCtl.Left = oCtl.Left - 75000
            
      If Ambient.UserMode Then
        oCtl.TabStop = False
      End If
    Else
      If pIsControlAdded(m_iActiveTab, sControlId) Then
      
        oCtl.Left = oCtl.Left + 75000
        
        If Ambient.UserMode Then
          oCtl.TabStop = m_aryTabs(m_iActiveTab).ContainedControlsDetails(sControlId).TabStop
        End If
      End If
    End If
    lLeft = -75000
  Next
  Err.Clear
    '<EhFooter>
    Exit Sub

pHandleContainedControls_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pHandleContainedControls " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Function adds remainig controls to the active tab. Remainig controls means the
' controls which have not been added till now.

Private Sub pAddRemainingControlsToActiveTab()
    '<EhHeader>
    On Error GoTo pAddRemainingControlsToActiveTab_Err
    '</EhHeader>
  Dim oCtl As Control
  Dim sControlId As String
  Dim lLeft As Long
  Dim oControlDetails As ControlDetails
  
  On Error Resume Next
  
  If m_aryTabs(m_iActiveTab).ContainedControlsDetails.Count > 0 Then
    Set m_aryTabs(m_iActiveTab).ContainedControlsDetails = Nothing          'clear the existing collection
    Set m_aryTabs(m_iActiveTab).ContainedControlsDetails = New Collection
  End If
  
  lLeft = -75000
  
  For Each oCtl In ContainedControls
    
    sControlId = pGetControlId(oCtl)
    lLeft = oCtl.Left
    If lLeft > -10000 Then    'if there was some error in the above line the value of iLeft will be -75000,even if there is no error this means that the control is in the current tab and must be added to the tab's contained controls array
      
      oControlDetails.ControlID = sControlId
      oControlDetails.TabStop = oCtl.TabStop      'store original tab stop
      
      m_aryTabs(m_iActiveTab).ContainedControlsDetails.Add oControlDetails, oControlDetails.ControlID
     
    End If
    lLeft = -75000
  Next
  Err.Clear
    '<EhFooter>
    Exit Sub

pAddRemainingControlsToActiveTab_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pAddRemainingControlsToActiveTab " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' Function returns control's name & control's index combination
Private Function pGetControlId(ByRef oCtl As Control) As String
    '<EhHeader>
    On Error GoTo pGetControlId_Err
    '</EhHeader>
  On Error Resume Next
  
  Static sCtlName As String
  Static iCtlIndex As Integer
  
  iCtlIndex = -1
  
  sCtlName = oCtl.Name
  iCtlIndex = oCtl.Index
  pGetControlId = sCtlName & IIf(iCtlIndex <> -1, iCtlIndex, "")
    '<EhFooter>
    Exit Function

pGetControlId_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pGetControlId " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'Determines if a control is added in a specific tab or not
Private Function pIsControlAdded(iTabIndex As Integer, sCtrlName As String) As Boolean
    '<EhHeader>
    On Error GoTo pIsControlAdded_Err
    '</EhHeader>
  On Error GoTo Err_Handler
  
  If m_aryTabs(iTabIndex).ContainedControlsDetails.Count = 0 Then
    pIsControlAdded = False
    Exit Function
  End If
  
  pIsControlAdded = m_aryTabs(iTabIndex).ContainedControlsDetails.Item(sCtrlName).TabStop
  
  'if no error occured while accessing the value that means the control is already added
  pIsControlAdded = True
  Exit Function

Err_Handler:                      'error occured while accessing the control..... that means it is not added
  pIsControlAdded = False
    '<EhFooter>
    Exit Function

pIsControlAdded_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pIsControlAdded " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

' called only once to initialize the tab stop values
' that is store the original tab stop values for the controls
' and set there tab stop to false
Private Sub pStoreOriginalTabStopValues()
    '<EhHeader>
    On Error GoTo pStoreOriginalTabStopValues_Err
    '</EhHeader>
  Dim oCtl As Control
  Dim iCnt As Integer
  Dim sControlId As String
  Dim oControlDetails As ControlDetails
  
  On Error Resume Next    'used to prevent errors (
  
  For Each oCtl In ContainedControls
    sControlId = pGetControlId(oCtl)  'Get Control's Id (i.e. Control Name & Control Index)
    For iCnt = 0 To m_iTabCount - 1
      'Now see if the control is already added (it should be)
      If pIsControlAdded(iCnt, sControlId) Then
        oControlDetails.ControlID = sControlId
        oControlDetails.TabStop = oCtl.TabStop
        
        If iCnt <> m_iActiveTab Then    'if not the active tab then set the control's tabstop to false
          oCtl.TabStop = False
        End If
        m_aryTabs(iCnt).ContainedControlsDetails.Remove (sControlId)
        Call m_aryTabs(iCnt).ContainedControlsDetails.Add(oControlDetails, sControlId)
        Exit For
      End If
    Next
  Next
    '<EhFooter>
    Exit Sub

pStoreOriginalTabStopValues_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pStoreOriginalTabStopValues " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Function Called to handle the addition or deletion of tabs
Private Sub pHandleTabCount()
    '<EhHeader>
    On Error GoTo pHandleTabCount_Err
    '</EhHeader>
  Dim iCnt As Integer

  
  If m_iTabCount > UBound(m_aryTabs) Then      'tabs added
    iCnt = UBound(m_aryTabs) + 1
    
    ReDim Preserve m_aryTabs(m_iTabCount - 1)            'redim the tabs array
    
    'initialize the added tabs
    For iCnt = iCnt To m_iTabCount - 1
      m_aryTabs(iCnt).Caption = m_def_sCaption & " " & iCnt
      m_aryTabs(iCnt).Enabled = m_def_bTabEnabled
      
      Set m_aryTabs(iCnt).ContainedControlsDetails = New Collection
    Next
  ElseIf m_iTabCount <= UBound(m_aryTabs) Then   'tabs removed
    iCnt = UBound(m_aryTabs)
    
    'free memory of the deleted tabs
    For iCnt = iCnt To m_iTabCount Step -1
      Set m_aryTabs(iCnt).ContainedControlsDetails = Nothing
      Set m_aryTabs(iCnt).TabPicture = Nothing
      
    Next
    
    ReDim Preserve m_aryTabs(m_iTabCount - 1)            'redim the tabs array
    
    
    'make sure if the active tab is within the tab count
    If m_iActiveTab >= m_iTabCount Then
      ActiveTab = m_iTabCount - 1
    End If
    
    'Else ' :No need since this means that the number of tabs has not changed
  End If
    '<EhFooter>
    Exit Sub

pHandleTabCount_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pHandleTabCount " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub


' function used to extact the access keys from tabs and reassign them to AccessKeys property
Private Sub pAssignAccessKeys()
    '<EhHeader>
    On Error GoTo pAssignAccessKeys_Err
    '</EhHeader>
    
  Dim iCnt As Integer
  Dim sTmp As String
  
  For iCnt = 0 To m_iTabCount - 1
    If m_aryTabs(iCnt).AccessKey <> 0 Then
      sTmp = sTmp & Chr$(m_aryTabs(iCnt).AccessKey)
    End If
  Next
  AccessKeys = sTmp
    '<EhFooter>
    Exit Sub

pAssignAccessKeys_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pAssignAccessKeys " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'function called as a result of show/hide focus
Private Sub pShowHideFocus()
    '<EhHeader>
    On Error GoTo pShowHideFocus_Err
    '</EhHeader>
  If m_bIsRecursive Then Exit Sub
  
  m_bIsRecursive = True         'Set Recursive flag
  
  If m_bAreControlsAdded Then   'This is done to allow the control to draw properly first time
    AutoRedraw = True
    
    m_oTheme.ShowHideFocus
    
    Refresh                     'this will not cause repaint (since the AutoreDraw is true..
                                'This is an optimzed way to draw only the focus rect and prevent
                                'complete repaint)
    AutoRedraw = False
  Else
    
    m_oTheme.ShowHideFocus      'This also, will only draw the focus rectangle and
                                'prevent complete repaint
  End If
  
  m_bIsRecursive = False
  
    '<EhFooter>
    Exit Sub

pShowHideFocus_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pShowHideFocus " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'fuction called when Active tab changes to draw new changes
Private Sub pDrawOnActiveTabChange()
    '<EhHeader>
    On Error GoTo pDrawOnActiveTabChange_Err
    '</EhHeader>
  m_oTheme.DrawOnActiveTabChange
    '<EhFooter>
    Exit Sub

pDrawOnActiveTabChange_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pDrawOnActiveTabChange " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'Set's the theme object according to the name of the theme
Private Sub pSetThemeObject(iEnumValue As Integer)
    '<EhHeader>
    On Error GoTo pSetThemeObject_Err
    '</EhHeader>
  Set m_oTheme = Nothing
  
  Select Case iEnumValue
    Case xThemeWin9x:
      Set m_oTheme = New xThemeWin9x
    Case xThemeWinXP:
      Set m_oTheme = New xThemeWinXP
    Case xThemeVisualStudio2003:
      Set m_oTheme = New xThemeVisualStudio2003
    Case xThemeRoundTabs:
      Set m_oTheme = New xThemeRoundTabs
    Case xThemeOwnerDrawn:
      Set m_oTheme = New xThemeOwnerDrawn
  End Select
  
  Call m_oTheme.SetControl(Me)        'set the control reference
  
  Call m_oTheme.ResetColorsToDefault  'reset the colors to reflect current theme's default colors
  
    '<EhFooter>
    Exit Sub

pSetThemeObject_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pSetThemeObject " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'extract Red component from a color
Private Function pGetRValue(RGBValue As Long) As Integer
    '<EhHeader>
    On Error GoTo pGetRValue_Err
    '</EhHeader>
  pGetRValue = RGBValue And &HFF
    '<EhFooter>
    Exit Function

pGetRValue_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pGetRValue " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function


'extract Green component from a color
Private Function pGetGValue(RGBValue As Long) As Integer
    '<EhHeader>
    On Error GoTo pGetGValue_Err
    '</EhHeader>
  pGetGValue = ((RGBValue And &HFF00) / &H100) And &HFF
    '<EhFooter>
    Exit Function

pGetGValue_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pGetGValue " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'extract Blue component from a color
Private Function pGetBValue(RGBValue As Long) As Integer
    '<EhHeader>
    On Error GoTo pGetBValue_Err
    '</EhHeader>
  pGetBValue = ((RGBValue And &HFF0000) / &H10000) And &HFF
    '<EhFooter>
    Exit Function

pGetBValue_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pGetBValue " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

'function deletes the pictures etc and frees up the res
Private Sub pDestroyResources()
    '<EhHeader>
    On Error GoTo pDestroyResources_Err
    '</EhHeader>
  On Error Resume Next

  Dim iCnt As Integer

  For iCnt = 0 To m_iTabCount - 1           'free up the memory
    Set m_aryTabs(iCnt).ContainedControlsDetails = Nothing
    Set m_aryTabs(iCnt).TabPicture = Nothing
  Next
    '<EhFooter>
    Exit Sub

pDestroyResources_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pDestroyResources " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' function starts subclassing in desing mode.Allowing us to click and switch tab
Private Sub pStartSubClass()
    '<EhHeader>
    On Error GoTo pStartSubClass_Err
    '</EhHeader>
#If ISDEBUG = 0 Then
  'Exit if the window is already subclassed.
  If m_lOrigWndProc Then Exit Sub

  'Redirect the window's messages from this control's default
  'Window Procedure to the SubWndProc function in your .BAS
  'module and record the address of the previous Window
  'Procedure for this window in m_lOrigWndProc.
  m_lOrigWndProc = A_SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)

  'Record your window handle in case SetWindowLong gave you a
  'new one. You will need this handle so that you can unsubclass.
  m_lhWndOrigSubClassed = hwnd

  'Store a pointer to this object in the UserData section of
  'this window that will be used later to get the pointer to
  'the control based on the handle (hwnd) of the window getting
  'the message.
  Call A_SetWindowLong(hwnd, GWL_USERDATA, ObjPtr(Me))
#End If
    '<EhFooter>
    Exit Sub

pStartSubClass_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pStartSubClass " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

' function used to end Subclassing
Private Sub pEndSubClass()
  '-----------------------------------------------------------
  'Unsubclasses this UserControl's window (hwnd), setting the
  'address of the Windows Procedure back to the address it was
  'at before it was subclassed.
  '-----------------------------------------------------------
    '<EhHeader>
    On Error GoTo pEndSubClass_Err
    '</EhHeader>

  'Ensures that you don't try to unsubclass the window when
  'it is not subclassed.
  If m_lOrigWndProc = 0 Then Exit Sub

  'Reset the window's function back to the original address.
  A_SetWindowLong m_lhWndOrigSubClassed, GWL_WNDPROC, m_lOrigWndProc
  '0 Indicates that you are no longer subclassed.
  m_lOrigWndProc = 0
    '<EhFooter>
    Exit Sub

pEndSubClass_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ThunVBCC_v1_0.XTab.pEndSubClass " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

'=====================================================================================================================

