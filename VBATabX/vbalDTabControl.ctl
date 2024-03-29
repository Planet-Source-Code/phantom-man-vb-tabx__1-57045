VERSION 5.00
Begin VB.UserControl vbalDTabControlX 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vbalDTabControl.ctx":0000
   Begin VB.Timer m_tmrPinButton 
      Interval        =   25
      Left            =   1560
      Top             =   780
   End
   Begin VB.Timer m_tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1020
      Top             =   780
   End
   Begin VB.PictureBox picUnpinned 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "vbalDTabControlX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ======================================================================================
' Name:     vbalDTabControlX
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     7 January 2003
'
' Requires: -
'
' Copyright © 2003 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Control implementing a Visual Studio style tab interface.
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
' ======================================================================================

' Updates 27/04/03
' 1) Remove method ate any subscript out of range errors
' 2) Inserting a tab did not work correctly
' 3) If a tab panel was removed and then added again a subscript out of range occurred.
' Thanks to Michael Elashoff, Julien Margail and Matt Funnell

' Updates 06/02/03
' 1) Control is now alignable
' 2) InitProperties now initialises control, as well as ReadProperties
' 3) Flicker-Free Drawing.
' 4) Better scrolling for tabs: on MouseDown rather than MouseUp
' 5) When you remove a tab, if it has a panel it is now hidden
' 6) Clicking on tabs didn't always work when the tab was scrolled and
'    near the buttons
' 7) Added pinnable function.

' Updates 13/09/04 - Gary Noble
' 1) Added 2 Different DrawStyles To Bring The Control More Upto Date
'    Styles Added - Office 2003 And Office 2003 Hot
'    Office 2003: Simulates Office Colours WithOUT The Hot Orange
'    Office 2003 Hot: Simulates Office Colours With The Hot Orange
' 2) Fixed The Bug With The Caption Drawing On The Pinned Tabs
'    Tab Header On Pinned Tabs Was Not Bold When The Selected Font Was Bold
' 3) Minor Cosmetic Changes


Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
' lfQuality Constants:
Private Const DEFAULT_QUALITY = 0    ' Appearance of the font is set to default
Private Const DRAFT_QUALITY = 1    ' Appearance is less important that PROOF_QUALITY.
Private Const PROOF_QUALITY = 2    ' Best character quality
Private Const NONANTIALIASED_QUALITY = 3    ' Don't smooth font edges even if system is set to smooth font edges
Private Const ANTIALIASED_QUALITY = 4    ' Ensure font edges are smoothed if system is set to smooth font edges

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Enum ESetWindowPosFlags
    HWND_TOPMOST = -1
    HWND_DESKTOP = 0
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOREDRAW = &H8
    SWP_SHOWWINDOW = &H40
    SWP_FRAMECHANGED = &H20    '  The frame changed: send WM_NCCALCSIZE
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOOWNERZORDER = &H200    '  Don't do owner Z ordering
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOZORDER = &H4
End Enum
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Enum EWindowLongIndexes
    GWL_EXSTYLE = (-20)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_USERDATA = (-21)
    GWL_WNDPROC = (-4)
End Enum
' General window styles:
Private Enum EExWindowStyles
    WS_EX_DLGMODALFRAME = &H1
    WS_EX_NOPARENTNOTIFY = &H4
    WS_EX_TOPMOST = &H8
    WS_EX_ACCEPTFILES = &H10
    WS_EX_TRANSPARENT = &H20
    WS_EX_MDICHILD = &H40
    WS_EX_TOOLWINDOW = &H80
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum

Private Const WM_DESTROY = &H2

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
        lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const BITSPIXEL = 12
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Private Const DCX_WINDOW = &H1&
Private Const DCX_INTERSECTRGN = &H80&
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

#If Win32 Then
    Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
#Else
    Private Declare Function CreateFont Lib "gdi.dll" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
#End If

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

' DrawText
Private Enum EDrawTextFormat
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
    DT_EDITCONTROL = &H2000&
    DT_PATH_ELLIPSIS = &H4000&
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

' Image list functions:
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal fStyle As Long _
        ) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
        ) As Long
Private Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
        ) As Long
Private Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
        ) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal cx As Long, _
        ByVal cy As Long _
        ) As Long
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_OVERLAYMASK = 3840

Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
        (ByVal hdc As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lparam As Long, _
        ByVal wParam As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal fuFlags As Long) As Long
'/* Image type */
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4
' /* State type */
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10    ' /* Gray string appearance */
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadImageString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8
Private Const LR_LOADFROMFILE = &H10


'//---------------------------------------------------------------------------------------
'-- Start Of Additions By Gary Noble
'//---------------------------------------------------------------------------------------
'//-- Theme
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private m_sCurrentSystemThemename                         As String

'//-- Gradient
Private Type TRIVERTEX
    x                                                      As Long
    y                                                      As Long
    Red                                                    As Integer
    Green                                                  As Integer
    Blue                                                   As Integer
    Alpha                                                  As Integer
End Type
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, _
        pVertex As TRIVERTEX, _
        ByVal dwNumVertex As Long, _
        pMesh As GRADIENT_RECT, _
        ByVal dwNumMesh As Long, _
        ByVal dwMode As Long) As Long

Private Enum GradientFillRectType
    GRADIENT_FILL_RECT_H = 0
    GRADIENT_FILL_RECT_V = 1
End Enum

Private Type GRADIENT_RECT
    UpperLeft                                              As Long
    LowerRight                                             As Long
End Type

#If False Then
    Private GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V
#End If

'//-- Os Version info
Private Type OSVERSIONINFO
    dwVersionInfoSize                                      As Long
    dwMajorVersion                                         As Long
    dwMinorVersion                                         As Long
    dwBuildNumber                                          As Long
    dwPlatformId                                           As Long
    szCSDVersion(0 To 127)                                 As Byte
End Type

Private Const VER_PLATFORM_WIN32_NT                      As Integer = 2

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

'//-- System  Settings
Private m_bIsXp                                          As Boolean
Private m_bIsNt                                          As Boolean
Private m_bIs2000OrAbove                                 As Boolean
Private m_bHasGradientAndTransparency                    As Boolean

'//-- Draw Style
Public Enum E_DrawStyle
    EDS_DefaultNet = 0
    EDS_Office2003 = 1
    EDS_Office2003Hot = 2
End Enum

Private m_iDrawStyle As E_DrawStyle

'//-- Custom Draw Style Colours
Private m_lCustColorOneNormal                  As OLE_COLOR
Private m_lCustColorTwoNormal                  As OLE_COLOR
Private m_lCustColorOneSelected                As OLE_COLOR
Private m_lCustColorTwoSelected                As OLE_COLOR
Private m_lCustColorHeaderColorOne             As OLE_COLOR
Private m_lCustColorHeaderColorTwo             As OLE_COLOR
Private m_lCustColorHeaderForeColor            As OLE_COLOR
Private m_bCustUseGradient                     As Boolean

Private m_lColorOneSelectedNormal          As OLE_COLOR
Private m_lColorTwoSelectedNormal          As OLE_COLOR
Private m_lColorOneNormal                  As OLE_COLOR
Private m_lColorTwoNormal                  As OLE_COLOR
Private m_lColorOneSelected                As OLE_COLOR
Private m_lColorTwoSelected                As OLE_COLOR
Private m_lColorHeaderColorOne             As OLE_COLOR
Private m_lColorHeaderColorTwo             As OLE_COLOR
Private m_lColorHeaderForeColor            As OLE_COLOR
Private m_lColorHotOne                     As OLE_COLOR
Private m_lColorHotTwo                     As OLE_COLOR
Private m_lColorBorder                     As OLE_COLOR
'//---------------------------------------------------------------------------------------
'-- End Of Additions By Gary Noble
'//---------------------------------------------------------------------------------------

Private Type TabInfo
    sCaption As String
    sKey As String
    sToolTipText As String
    lItemData As Long
    sTag As String
    lIconIndex As Long
    bCanClose As Boolean
    bEnabled As Boolean
    lObjPtrPanel As Long
    lId As Long
    tTabR As RECT
    tPinnedR As RECT
End Type

Public Enum EMDITabAlign
    TabAlignTop
    TabAlignBottom
End Enum

Private m_lIdGenerator As Long

Private m_iDraggingTab As Long
Private m_bJustReplaced As Boolean
Private m_tJustReplacedPoint As POINTAPI
Private m_iTrackButton As Long
Private m_iPressButton As Long

Private m_eTabAlign As EMDITabAlign
Private m_bAllowScroll As Boolean
Private m_bAllowSelectDisabledTabs As Boolean
Private m_bShowCloseButton As Boolean
Private m_lTabHeight As Long
Private m_lButtonSize As Long
Private m_font As iFont
Private m_fontSelected As iFont
Private m_bShowTabs As Boolean
Private m_lOffsetX As Long
Private m_oBackColor As OLE_COLOR
Private m_oForeColor As OLE_COLOR
Private m_bPinnable As Boolean
Private m_bPinned As Boolean
Private m_bOut As Boolean
Private m_lUnpinnedWidth As Long
Private m_lSlideOutWidth As Long
Private m_lTitleBarHeight As Long
Private m_lSplitSize As Long

Private m_sLastToolTip As String

Private m_hIml As Long
Private m_ptrVb6ImageList As Long
Private m_lIconWidth As Long
Private m_lIconHeight As Long

Private m_cMemDC As New pcMemDC

Private m_tTab() As TabInfo
Private m_iTabCount As Long
Private m_iSelTab As Long
Private m_iLastSelTab As Long
Private m_tButtonR As RECT
Private m_tClientR As RECT
Private m_tUnpinCloseR As RECT
Private m_tUnpinPinR As RECT
Private m_bUnpinPinTrack As Boolean
Private m_bUnpinPinDown As Boolean
Private m_bUnpinCloseTrack As Boolean
Private m_bUnpinCloseDown As Boolean
Private m_hIconPin As Long
Private m_hIconUnpin As Long
Private m_hIconClose As Long

Private m_hWnd As Long
Private m_bDesignMode As Boolean
Private m_bInIde  As Boolean


Public Event Resize()
Public Event Pinned()
Public Event TabDoubleClick(theTab As cTab)
Public Event TabClose(theTab As cTab, ByRef bCancel As Boolean)
Public Event TabClick(theTab As cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
Public Event TabBarClick(ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
Public Event TabSelected(theTab As cTab)
Public Event UnPinned()

Public Property Get AllowScroll() As Boolean
    AllowScroll = m_bAllowScroll
End Property

Public Property Let AllowScroll(ByVal value As Boolean)
    If (m_bAllowScroll <> value) Then
        m_bAllowScroll = value
        If (m_bAllowScroll = False) Then
            m_lOffsetX = 0
        End If
        drawTabs
        PropertyChanged "AllowScroll"
    End If
End Property





Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oBackColor
End Property

Public Property Let BackColor(ByVal oColor As OLE_COLOR)
    If (m_oBackColor <> oColor) Then
        m_oBackColor = oColor
        drawTabs
        PropertyChanged "BackColor"
    End If
End Property

Private Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal Alpha As Long = 128 _
   ) As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
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

End Property

Public Property Get ClientHeight() As Long
    ClientHeight = m_tClientR.Bottom - m_tClientR.Top
End Property

Public Property Get ClientLeft() As Long
    ClientLeft = m_tClientR.Left
End Property

Public Property Get ClientTop() As Long
    ClientTop = m_tClientR.Top
End Property

Public Property Get ClientWidth() As Long
    ClientWidth = m_tClientR.Right - m_tClientR.Left
End Property

Private Sub drawButtonBorder( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal bTrack As Boolean, _
      ByVal bDown As Boolean _
   )
    ' up = down or track
    ' down = down & track
    ' else none
    Dim tJunk As POINTAPI
    If (bDown Or bTrack) Then

        Dim hPenBottomRight As Long
        Dim hPenTopLeft As Long
        Dim hPenOld As Long

        If (bDown And bTrack) Then
            hPenTopLeft = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
            hPenBottomRight = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
        Else
            hPenTopLeft = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
            hPenBottomRight = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
        End If
        hPenOld = SelectObject(lHDC, hPenTopLeft)
        MoveToEx lHDC, tR.Left, tR.Bottom - 2, tJunk
        LineTo lHDC, tR.Left, tR.Top
        LineTo lHDC, tR.Right - 1, tR.Top
        SelectObject lHDC, hPenOld
        hPenOld = SelectObject(lHDC, hPenBottomRight)
        MoveToEx lHDC, tR.Right - 1, tR.Top + 1, tJunk
        LineTo lHDC, tR.Right - 1, tR.Bottom - 1
        LineTo lHDC, tR.Left, tR.Bottom - 1
        SelectObject lHDC, hPenOld
        DeleteObject hPenTopLeft
        DeleteObject hPenBottomRight
    End If

End Sub

Private Sub drawButtons(ByVal lHDC As Long)
    Dim tR As RECT

    If (m_bAllowScroll) Then
        ' Left & Right Buttons
        drawOneButton lHDC, 1

        drawOneButton lHDC, 2

    End If

    If (m_bShowCloseButton And Not (m_bPinnable)) Then
        ' Close Button
        drawOneButton lHDC, 3
    End If

End Sub

Private Sub drawControl()
    '
    Dim bNoTx As Boolean


    GetGradientColors

    If Not (m_bPinned) Then
        ' draw the tabs into picUnpinned
        drawUnpinnedTabs

        If (m_bOut) Then
            ' draw the title:
            drawTitleBar
            ' draw the border of the unpinned area:
            drawUnpinnedBorder
        End If

    Else

        ' Draw the tabs:
        Dim lHDC As Long
        lHDC = m_cMemDC.hdc
        If (lHDC = 0) Then    ' out of memory
            lHDC = UserControl.hdc
            bNoTx = True
        End If

        Dim tR As RECT
        GetTabWindowRect tR
        LSet m_tClientR = tR
        If (m_bShowTabs) Then
            m_tClientR.Left = m_tClientR.Left + 1
            m_tClientR.Right = m_tClientR.Right - 1
            If (m_eTabAlign = TabAlignBottom) Then
                m_tClientR.Bottom = tR.Bottom - m_lTabHeight
                m_tClientR.Top = m_tClientR.Top + 1
            Else
                m_tClientR.Top = tR.Top + m_lTabHeight
                m_tClientR.Bottom = m_tClientR.Bottom - 1
            End If
        End If
        Dim hBrush As Long


        '//---------------------------------------------------------------------------------------
        '-- Added DrawStyle Params
        '-- Gary Noble
        '//---------------------------------------------------------------------------------------
        If Me.DrawStyle = EDS_DefaultNet Then
            hBrush = CreateSolidBrush(TranslateColor(m_oBackColor))
        Else
            hBrush = CreateSolidBrush(BlendColor(m_lColorOneNormal, vbWhite, 150))
        End If

        FillRect lHDC, m_tClientR, hBrush
        DeleteObject hBrush

        If (m_bPinnable And m_bPinned) Then
            m_tClientR.Top = m_tClientR.Top + m_lTitleBarHeight
            drawTitleBar
        End If

        If (m_bShowTabs) Then
            drawTabs lHDC
        End If

        If Not bNoTx Then
            BitBlt UserControl.hdc, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, lHDC, 0, 0, vbSrcCopy
        End If

    End If
End Sub

Private Sub drawOneButton( _
      ByVal lHDC As Long, _
      ByVal lGlyph As Long _
   )
    Dim tR As RECT
    Dim bEnabled As Boolean
    Dim bPressed As Boolean
    Select Case lGlyph
        Case 1
            getLeftButtonRect tR
            bEnabled = IsLeftButtonEnabled()
        Case 2
            getRightButtonRect tR
            bEnabled = IsRightButtonEnabled()
        Case 3
            bEnabled = IsCloseButtonEnabled()
            getCloseButtonRect tR
    End Select
    bPressed = ((m_iPressButton = lGlyph) And (m_iTrackButton = lGlyph))

    Dim tTextR As RECT
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tJunk As POINTAPI

    LSet tTextR = tR
    InflateRect tTextR, -2, -2

    If bEnabled Then
        If bPressed Then

            ' draw down border:
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
            hPenOld = SelectObject(lHDC, hPen)
            MoveToEx lHDC, tR.Left, tR.Bottom - 1, tJunk
            LineTo lHDC, tR.Left, tR.Top
            LineTo lHDC, tR.Right - 1, tR.Top
            SelectObject lHDC, hPenOld
            DeleteObject hPen

            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
            hPenOld = SelectObject(lHDC, hPen)
            MoveToEx lHDC, tR.Right - 1, tR.Top + 1, tJunk
            LineTo lHDC, tR.Right - 1, tR.Bottom - 1
            LineTo lHDC, tR.Left + 1, tR.Bottom - 1
            SelectObject lHDC, hPenOld
            DeleteObject hPen

            ' Move text down
            OffsetRect tTextR, 1, 1

        ElseIf (m_iTrackButton = lGlyph) Then

            ' draw up border:
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
            hPenOld = SelectObject(lHDC, hPen)
            MoveToEx lHDC, tR.Left, tR.Bottom - 1, tJunk
            LineTo lHDC, tR.Left, tR.Top
            LineTo lHDC, tR.Right - 1, tR.Top
            SelectObject lHDC, hPenOld
            DeleteObject hPen

            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
            hPenOld = SelectObject(lHDC, hPen)
            MoveToEx lHDC, tR.Right - 1, tR.Top + 1, tJunk
            LineTo lHDC, tR.Right - 1, tR.Bottom - 1
            LineTo lHDC, tR.Left + 1, tR.Bottom - 1
            SelectObject lHDC, hPenOld
            DeleteObject hPen

        End If
    End If

    Dim sFont As New StdFont
    sFont.Name = "Marlett"
    If (lGlyph = 3) Then
        sFont.Size = 8
    Else
        sFont.Size = 10
    End If
    Dim iFont As iFont
    Set iFont = sFont
    Dim hFontOld As Long

    hFontOld = SelectObject(lHDC, iFont.hFont)
    If (bEnabled) Then
        If Me.DrawStyle = EDS_DefaultNet Then
            SetTextColor lHDC, GetSysColor(vb3DDKShadow And &H1F&)
        Else
            SetTextColor lHDC, m_lColorBorder
        End If

    Else
        If Me.DrawStyle = EDS_DefaultNet Then
            SetTextColor lHDC, BlendColor(vbButtonFace, vb3DDKShadow, 192)
        Else
            SetTextColor lHDC, m_lColorOneNormal
        End If
    End If
    ' Draw the glyph:
    Select Case lGlyph
        Case 1    ' left
            DrawText lHDC, "3", -1, tTextR, DT_CENTER Or DT_VCENTER
        Case 2    ' right
            DrawText lHDC, "4", -1, tTextR, DT_CENTER Or DT_VCENTER
        Case 3    ' close
            DrawText lHDC, "r", -1, tTextR, DT_CENTER Or DT_VCENTER
    End Select

    SelectObject lHDC, hFontOld

End Sub

Public Property Get DrawStyle() As E_DrawStyle
    DrawStyle = m_iDrawStyle
End Property


'//---------------------------------------------------------------------------------------
' Procedure : DrawStyle
' Type      : Property
' DateTime  : 13/08/2004
' Author    : Gary Noble
' Purpose   : Sets The Drawing Style
' Returns   : E_DrawStyle
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  13/08/2004
'//---------------------------------------------------------------------------------------
Public Property Let DrawStyle(iStyle As E_DrawStyle)
    m_iDrawStyle = iStyle
    PropertyChanged "DrawStyle"
    drawControl
End Property

Private Sub drawTabs(Optional ByVal lhDCTo As Long = 0)

    If (m_bShowTabs) Then


        Dim tR As RECT
        GetTabWindowRect tR

        Dim lHDC As Long
        If (lhDCTo = 0) Then
            lHDC = m_cMemDC.hdc
            If (lHDC = 0) Then    ' out of memory
                lHDC = UserControl.hdc
                lhDCTo = lHDC    ' don't redraw
            End If
        Else
            lHDC = lhDCTo
        End If

        ' Draw all the borders:
        Dim hPenOld As Long
        Dim hPen As Long
        Dim tJunk As POINTAPI
        Dim lPen As Long

        '//---------------------------------------------------------------------------------------
        '-- Set The Pen Color
        '-- Ammended By: Gary Noble
        '//---------------------------------------------------------------------------------------
        If Me.DrawStyle = EDS_DefaultNet Then
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DShadow And &H1F&))
        Else
            hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
        End If
        hPenOld = SelectObject(lHDC, hPen)

        MoveToEx lHDC, tR.Left, tR.Top, tJunk
        LineTo lHDC, tR.Right - 1, tR.Top
        LineTo lHDC, tR.Right - 1, tR.Bottom - 1
        LineTo lHDC, tR.Left, tR.Bottom - 1
        LineTo lHDC, tR.Left, tR.Top

        SelectObject lHDC, hPenOld
        DeleteObject hPen

        If Me.DrawStyle = EDS_DefaultNet Then
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonFace And &H1F&))
        Else
            hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorOneNormal))
        End If

        hPenOld = SelectObject(lHDC, hPen)

        MoveToEx lHDC, tR.Left + 1, tR.Top + 1, tJunk
        LineTo lHDC, tR.Left + 1, tR.Bottom - 2
        MoveToEx lHDC, tR.Right - 2, tR.Top + 1, tJunk
        LineTo lHDC, tR.Right - 2, tR.Bottom - 2
        If (m_eTabAlign = TabAlignBottom) Then
            MoveToEx lHDC, tR.Left + 1, tR.Top + 1, tJunk
            LineTo lHDC, tR.Right - 1, tR.Top + 1
        Else
            MoveToEx lHDC, tR.Left + 1, tR.Bottom - 2, tJunk
            LineTo lHDC, tR.Right - 1, tR.Bottom - 2
        End If

        Dim tTabR As RECT
        LSet tTabR = tR
        tTabR.Left = tTabR.Left + 1
        tTabR.Right = tTabR.Right - 1
        If (m_eTabAlign = TabAlignBottom) Then
            tTabR.Top = tR.Bottom - m_lTabHeight
            tTabR.Bottom = tTabR.Bottom - 1
            MoveToEx lHDC, tTabR.Left, tTabR.Top, tJunk
            LineTo lHDC, tTabR.Right - 1, tTabR.Top
            MoveToEx lHDC, tTabR.Left, tTabR.Top + 1, tJunk
            LineTo lHDC, tTabR.Right - 1, tTabR.Top + 1
            tTabR.Top = tTabR.Top + 1

        Else
            tTabR.Bottom = tR.Top + m_lTabHeight
            tTabR.Top = tTabR.Top + 1
            MoveToEx lHDC, tTabR.Left, tTabR.Bottom - 1, tJunk
            LineTo lHDC, tTabR.Right - 1, tTabR.Bottom - 1
            MoveToEx lHDC, tTabR.Left, tTabR.Bottom - 2, tJunk
            LineTo lHDC, tTabR.Right - 1, tTabR.Bottom - 2
            tTabR.Bottom = tTabR.Bottom - 2
        End If

        SelectObject lHDC, hPenOld
        DeleteObject hPen

        ' Fill with generic back colour:
        Dim hBr As Long

        '//---------------------------------------------------------------------------------------
        '-- Included DrawStyle
        '-- Ammended By: Gary Noble
        '//---------------------------------------------------------------------------------------
        If Me.DrawStyle = EDS_DefaultNet Then
            hBr = CreateSolidBrush(BlendColor(vbButtonFace, vbWindowBackground, 64))
            FillRect lHDC, tTabR, hBr
            DeleteObject hBr
        Else

            UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 250), tTabR
        End If


        ' Now evaluate the positioning of the tabs (calculate
        ' using left to right layout, then when we draw we can
        ' subtract the width of the layout).

        ' If the tab is set to have allow scroll then we will draw
        ' in these positions until we get to the scroll point,
        ' otherwise we will need to squeeze them up until they
        ' fit.

        If (m_iTabCount > 0) Then


            Dim hFontOld As Long
            hFontOld = SelectObject(lHDC, m_font.hFont)

            Dim iC As Long
            Dim tCalcR As RECT

            For iC = 1 To m_iTabCount
                If (iC = 1) Then
                    m_tTab(iC).tTabR.Left = tTabR.Left + 2
                Else
                    m_tTab(iC).tTabR.Left = m_tTab(iC - 1).tTabR.Right
                End If
                m_tTab(iC).tTabR.Right = m_tTab(iC).tTabR.Left + 8    ' min tab size
                If (m_eTabAlign = TabAlignBottom) Then
                    m_tTab(iC).tTabR.Top = tTabR.Top
                    m_tTab(iC).tTabR.Bottom = tTabR.Bottom - 2
                Else
                    m_tTab(iC).tTabR.Top = tTabR.Top + 2
                    m_tTab(iC).tTabR.Bottom = tTabR.Bottom
                End If
                If (iC = m_iSelTab) Then

                    SelectObject lHDC, hFontOld

                    hFontOld = SelectObject(lHDC, m_fontSelected.hFont)
                End If
                If m_bIsNt Then
                    DrawTextW lHDC, StrPtr(m_tTab(iC).sCaption), -1, tCalcR, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE
                Else
                    DrawText lHDC, m_tTab(iC).sCaption, -1, tCalcR, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE
                End If
                m_tTab(iC).tTabR.Right = m_tTab(iC).tTabR.Left + 16 + tCalcR.Right - tCalcR.Left
                If (iC = m_iSelTab) Then
                    SelectObject lHDC, hFontOld
                    hFontOld = SelectObject(lHDC, m_font.hFont)
                End If
                If (m_tTab(iC).lIconIndex > -1) Then
                    If Not (m_hIml = 0) Or Not (m_ptrVb6ImageList = 0) Then
                        ' Add the size of the icon:
                        m_tTab(iC).tTabR.Right = m_tTab(iC).tTabR.Right + m_lIconWidth + 4
                    End If
                End If
            Next iC

            Dim lMaxRight As Long

            lMaxRight = tTabR.Right
            'Debug.Print lMaxRight
            If (m_bShowCloseButton And Not (m_bPinnable)) Then
                lMaxRight = lMaxRight - m_lButtonSize
            End If
            If (m_bAllowScroll) Then
                lMaxRight = lMaxRight - m_lButtonSize * 2
            End If

            Dim bDoesNotFit As Boolean

            If Not (m_bAllowScroll) Then
                If (m_tTab(m_iTabCount).tTabR.Right > lMaxRight) Then
                    bDoesNotFit = True
                    ' we don't fit, need to squash all the tabs up
                    Dim lActualSize As Long
                    lActualSize = (lMaxRight - 4) \ m_iTabCount
                    m_tTab(1).tTabR.Right = m_tTab(1).tTabR.Left + lActualSize
                    For iC = 2 To m_iTabCount
                        m_tTab(iC).tTabR.Left = m_tTab(iC - 1).tTabR.Right
                        m_tTab(iC).tTabR.Right = m_tTab(iC).tTabR.Left + lActualSize
                    Next iC
                End If
            End If

            Dim bChangedWindow As Boolean

            If (m_iSelTab <> m_iLastSelTab) Then
                If (m_iSelTab > 0) Then
                    m_iLastSelTab = m_iSelTab
                    bChangedWindow = True
                    ' ensure that a newly selected tab is scrolled into view
                    If (m_bAllowScroll) Then
                        If (m_tTab(m_iSelTab).tTabR.Right - m_lOffsetX) > (tTabR.Right - m_lButtonSize * 3) Then
                            m_lOffsetX = m_tTab(m_iSelTab).tTabR.Left - 16
                        ElseIf (m_tTab(m_iSelTab).tTabR.Left - m_lOffsetX < tTabR.Left) Then
                            m_lOffsetX = m_tTab(m_iSelTab).tTabR.Left - 16
                        End If
                        If (m_lOffsetX <= 16) Then
                            m_lOffsetX = 0
                        End If
                    End If
                End If
            End If

            Dim wFormat As Long
            Dim tTextR As RECT
            wFormat = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
            If (bDoesNotFit) Then
                wFormat = wFormat Or DT_END_ELLIPSIS
            End If
            SetBkMode lHDC, TRANSPARENT

            '//---------------------------------------------------------------------------------------
            '-- Set The Pen Color
            '-- Ammended By: Gary Noble
            '//---------------------------------------------------------------------------------------
            If Me.DrawStyle = EDS_DefaultNet Then
                hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DShadow And &H1F&))
            Else
                hPen = CreatePen(PS_SOLID, 1, BlendColor(m_lColorOneNormal, vbWhite, 150))
            End If
            hPenOld = SelectObject(lHDC, hPen)

            ensureEndTabOffset

            ' Actually do the drawing:
            Dim tActualR As RECT
            Dim tFillR As RECT
            Dim bClippedLeft As Boolean
            Dim bClippedRight As Boolean
            Dim bTabOffscreen As Boolean

            bTabOffscreen = True
            For iC = 1 To m_iTabCount

                LSet tActualR = m_tTab(iC).tTabR
                OffsetRect tActualR, -m_lOffsetX, 0

                If (tActualR.Right > lMaxRight) Then
                    tActualR.Right = lMaxRight
                    bClippedRight = True
                Else
                    bClippedRight = False
                End If
                If (tActualR.Left < 0) Then
                    bClippedLeft = True
                Else
                    bClippedLeft = False
                End If
                If (tActualR.Left > lMaxRight) Then
                    ' nothing to do
                    Exit For
                End If

                If (iC = m_iSelTab) Then
                    If (tActualR.Right < 0) Or (tActualR.Left > lMaxRight) Then
                        bTabOffscreen = True
                    Else
                        bTabOffscreen = False
                    End If

                    SelectObject lHDC, hFontOld
                    hFontOld = SelectObject(lHDC, m_fontSelected.hFont)
                    LSet tFillR = tActualR
                    If bClippedLeft Then
                        'Debug.Print tFillR.Left
                        tFillR.Left = 1
                    End If

                    '//---------------------------------------------------------------------------------------
                    '-- Draw The Tab BackGround
                    '-- Ammended By: Gary Noble
                    '//---------------------------------------------------------------------------------------

                    If Me.DrawStyle = EDS_DefaultNet Then
                        hBr = GetSysColorBrush(vbButtonFace And &H1F&)
                        FillRect lHDC, tFillR, hBr
                        DeleteObject hBr
                    Else
                        If Me.DrawStyle = EDS_Office2003Hot Then
                            If AppThemed Then
                                UtilDrawBackground lHDC, BlendColor(m_lColorTwoSelected, vbWhite, 100), m_lColorOneSelected, tFillR
                            Else
                                UtilDrawBackground lHDC, BlendColor(m_lColorTwoSelected, vbWhite, 100), BlendColor(vbApplicationWorkspace, vbWhite, 200), tFillR
                            End If
                        Else
                            If Me.DrawStyle = EDS_Office2003 Then
                                If AppThemed Then
                                    UtilDrawBackground lHDC, BlendColor(vbInactiveTitleBar, vbWhite, 100), BlendColor(vbActiveTitleBar, vbWhite, 250), tFillR
                                Else
                                    UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(vbActiveTitleBar, vbWhite, 200), tFillR
                                End If
                            Else
                                UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(vbActiveTitleBar, vbWhite, 150), tFillR
                            End If
                        End If
                    End If


                    SelectObject lHDC, hPenOld
                    DeleteObject hPen

                    ' replace pen:
                    If (m_eTabAlign = TabAlignBottom) Then
                        ' darkest 3d pen:
                        If Me.DrawStyle = EDS_DefaultNet Then
                            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
                        Else
                            hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
                        End If

                    Else
                        ' lightest 3d pen:
                        If Me.DrawStyle = EDS_DefaultNet Then
                            hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbWhite))
                        Else
                            hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
                        End If

                    End If
                    hPenOld = SelectObject(lHDC, hPen)

                    If (m_eTabAlign = TabAlignBottom) Then
                        MoveToEx lHDC, tTabR.Left, tActualR.Top, tJunk
                        LineTo lHDC, tActualR.Left, tActualR.Top
                        MoveToEx lHDC, tActualR.Left, tActualR.Bottom - 1, tJunk
                        LineTo lHDC, tActualR.Right - 1, tActualR.Bottom - 1
                        If Not (bClippedRight) Then
                            LineTo lHDC, tActualR.Right - 1, tActualR.Top
                        End If
                    Else
                        MoveToEx lHDC, tTabR.Left, tActualR.Bottom - 1, tJunk
                        LineTo lHDC, tActualR.Left, tActualR.Bottom - 1
                        LineTo lHDC, tActualR.Left, tActualR.Top
                        LineTo lHDC, tActualR.Right - 1, tActualR.Top
                    End If

                    If Not bClippedRight Then
                        If (m_eTabAlign = TabAlignBottom) Then
                            MoveToEx lHDC, tActualR.Right - 1, tActualR.Top, tJunk
                            LineTo lHDC, tTabR.Right - 1, tActualR.Top
                        Else
                            MoveToEx lHDC, tActualR.Right - 1, tActualR.Bottom - 1, tJunk
                            LineTo lHDC, tTabR.Right - 1, tActualR.Bottom - 1
                        End If
                    End If

                    SelectObject lHDC, hPenOld
                    DeleteObject hPen

                    If (m_eTabAlign = TabAlignBottom) Then
                        ' lightest 3d pen:
                        If Me.DrawStyle = EDS_DefaultNet Then
                            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
                        Else
                            hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
                        End If
                        hPenOld = SelectObject(lHDC, hPen)

                        MoveToEx lHDC, tActualR.Left, tActualR.Top, tJunk
                        LineTo lHDC, tActualR.Left, tActualR.Bottom - 1

                        SelectObject lHDC, hPenOld
                        DeleteObject hPen
                    Else
                        If Not bClippedRight Then
                            ' darkest 3d pen:
                            If Me.DrawStyle = EDS_DefaultNet Then
                                hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
                            Else
                                hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
                            End If
                            hPenOld = SelectObject(lHDC, hPen)

                            MoveToEx lHDC, tActualR.Right - 1, tActualR.Top + 1, tJunk
                            LineTo lHDC, tActualR.Right - 1, tActualR.Bottom

                            SelectObject lHDC, hPenOld
                            DeleteObject hPen
                        End If
                    End If

                    '//---------------------------------------------------------------------------------------
                    '-- Set The Pen Color
                    '-- Ammended By: Gary Noble
                    '//---------------------------------------------------------------------------------------
                    If Me.DrawStyle = EDS_DefaultNet Then
                        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DShadow And &H1F&))
                    Else
                        hPen = CreatePen(PS_SOLID, 1, BlendColor(TranslateColor(m_lColorBorder), m_lColorTwoNormal, 175))
                    End If
                    hPenOld = SelectObject(lHDC, hPen)

                ElseIf Not ((iC + 1) = m_iSelTab) Then

                    '//---------------------------------------------------------------------------------------
                    '-- Set The Pen Color
                    '-- Ammended By: Gary Noble
                    '//---------------------------------------------------------------------------------------
                    If Me.DrawStyle = EDS_Office2003 Or EDS_Office2003Hot Then
                        lPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
                        SelectObject lHDC, lPen
                        If Not bClippedRight Then
                            MoveToEx lHDC, tActualR.Right - 1, tActualR.Top + 3, tJunk
                            LineTo lHDC, tActualR.Right - 1, tActualR.Bottom - 2
                        End If
                        SelectObject lHDC, lPen
                        DeleteObject lPen
                    Else
                        If Not bClippedRight Then
                            MoveToEx lHDC, tActualR.Right - 1, tActualR.Top + 3, tJunk
                            LineTo lHDC, tActualR.Right - 1, tActualR.Bottom - 2
                        End If

                    End If

                End If


                LSet tTextR = tActualR
                tTextR.Left = tTextR.Left + 8
                tTextR.Right = tTextR.Right - 8

                If (m_tTab(iC).lIconIndex > -1) Then
                    If Not (m_hIml = 0) Or Not (m_ptrVb6ImageList = 0) Then
                        If (tTextR.Right - tTextR.Left > m_lIconWidth + 4) Then
                            If (m_tTab(iC).bEnabled) Then
                                ImageListDrawIcon m_ptrVb6ImageList, lHDC, m_hIml, _
                                        m_tTab(iC).lIconIndex, _
                                        tTextR.Left + 2, _
                                        tTextR.Top + ((tTextR.Bottom - tTextR.Top) - m_lIconHeight) \ 2
                            Else
                                ImageListDrawIconDisabled m_ptrVb6ImageList, lHDC, m_hIml, _
                                        m_tTab(iC).lIconIndex, _
                                        tTextR.Left + 2, _
                                        tTextR.Top + ((tTextR.Bottom - tTextR.Top) - m_lIconHeight) \ 2, _
                                        m_lIconWidth
                            End If
                            tTextR.Left = tTextR.Left + m_lIconWidth + 4
                        End If
                    End If
                End If

                If (iC = m_iSelTab) And (m_tTab(iC).bEnabled) Then
                    SetTextColor lHDC, GetSysColor(vbWindowText And &H1F&)
                Else
                    SetTextColor lHDC, GetSysColor(vb3DDKShadow And &H1F&)
                End If
                If m_bIsNt Then
                    DrawTextW lHDC, StrPtr(m_tTab(iC).sCaption), -1, tTextR, wFormat
                Else
                    DrawText lHDC, m_tTab(iC).sCaption, -1, tTextR, wFormat
                End If
                If (iC = m_iSelTab) Then
                    SelectObject lHDC, hFontOld
                    hFontOld = SelectObject(lHDC, m_font.hFont)
                End If

            Next iC

            ' Clear up
            SelectObject lHDC, hPenOld
            DeleteObject hPen

            '//---------------------------------------------------------------------------------------
            '-- Set The Pen Color
            '-- Ammended By: Gary Noble
            '//---------------------------------------------------------------------------------------
            If (m_eTabAlign = TabAlignBottom) Then
                ' darkest 3d pen:
                If Me.DrawStyle = EDS_DefaultNet Then
                    hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
                Else
                    hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
                End If
                hPenOld = SelectObject(lHDC, hPen)
            Else
                ' lightest 3d pen:
                If Me.DrawStyle = EDS_DefaultNet Then
                    hPen = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
                Else
                    hPen = CreatePen(PS_SOLID, 1, TranslateColor(m_lColorBorder))
                End If
                hPenOld = SelectObject(lHDC, hPen)
            End If

            If (bTabOffscreen) Then
                If (m_eTabAlign = TabAlignBottom) Then
                    MoveToEx lHDC, tTabR.Left, tTabR.Top, tJunk
                    LineTo lHDC, tTabR.Right, tTabR.Top
                Else
                    MoveToEx lHDC, tTabR.Left, tTabR.Bottom - 1, tJunk
                    LineTo lHDC, tTabR.Right, tTabR.Bottom - 1
                End If
            End If

            ' The buttons always have a line above them:
            If (tTabR.Right > lMaxRight) Then
                If (m_eTabAlign = TabAlignBottom) Then
                    MoveToEx lHDC, lMaxRight, tTabR.Top, tJunk
                    LineTo lHDC, tTabR.Right, tTabR.Top
                Else
                    MoveToEx lHDC, lMaxRight, tTabR.Bottom - 1, tJunk
                    LineTo lHDC, tTabR.Right, tTabR.Bottom - 1
                End If
            End If

            SelectObject lHDC, hPenOld
            DeleteObject hPen


            SelectObject lHDC, hFontOld

            ' Now draw the buttons
            LSet m_tButtonR = tTabR
            m_tButtonR.Left = lMaxRight
            OffsetRect m_tButtonR, 0, 3
            drawButtons lHDC

        End If

        If (m_iSelTab <> 0) And (m_iTabCount = 0) Then
            bChangedWindow = True
            m_iSelTab = 0
        End If

        If (bChangedWindow) Then
            If (m_iTabCount = 0) Or (m_iSelTab = 0) Then
                pPanelSize
                RaiseEvent TabSelected(Nothing)
            Else
                Dim cT As New cTab
                cT.fInit ObjPtr(Me), m_hWnd, m_tTab(m_iSelTab).lId
                pPanelSize
                drawTitleBar
                RaiseEvent TabSelected(cT)
            End If
        End If

        If (lhDCTo = 0) Then
            ' Transfer to control:
            BitBlt UserControl.hdc, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, lHDC, 0, 0, vbSrcCopy
        End If

    End If

End Sub

Private Sub drawTitleBar()
    Dim sCap As String
    Dim tTR As RECT

    If (m_bPinnable) Then
        If (m_iSelTab > 0) Then
            sCap = m_tTab(m_iSelTab).sCaption
        End If

        GetClientRect m_hWnd, tTR
        m_cMemDC.Width = tTR.Right - tTR.Left
        If Not (m_bPinned) Then
            If (UserControl.Extender.Align = vbAlignLeft) Then
                tTR.Right = tTR.Right - m_lSplitSize
                tTR.Left = tTR.Right - m_lSlideOutWidth + m_lSplitSize
            ElseIf (UserControl.Extender.Align = vbAlignRight) Then
                tTR.Left = m_lSplitSize
                tTR.Right = tTR.Right - m_lUnpinnedWidth - 2
            Else
                '
            End If
        Else
            tTR.Top = tTR.Top + 2
            tTR.Left = tTR.Left + 2
            tTR.Right = tTR.Right - 2
        End If
        tTR.Bottom = tTR.Top + m_lTitleBarHeight

        Dim tCapR As RECT
        LSet tCapR = tTR
        tCapR.Top = tCapR.Top + 1
        tCapR.Left = tCapR.Left + 1
        tCapR.Right = tCapR.Right - 1
        tCapR.Bottom = tCapR.Bottom - 1

        Dim hPen As Long
        Dim hPenOld As Long
        Dim lHDC As Long
        Dim hFontOld As Long
        Dim bNoTx As Boolean
        Dim hBr As Long
        Dim tJunk As POINTAPI

        lHDC = m_cMemDC.hdc
        If (lHDC = 0) Then
            lHDC = UserControl.hdc
            bNoTx = True
        Else
            '//---------------------------------------------------------------------------------------
            '-- Draw The background
            '-- Ammended By: Gary Noble
            '//---------------------------------------------------------------------------------------

            If Me.DrawStyle = EDS_DefaultNet Then
                hBr = GetSysColorBrush(vbButtonFace And &H1F&)
                FillRect lHDC, tCapR, hBr
                DeleteObject hBr
            Else
                UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), m_lColorTwoNormal, tCapR

            End If
        End If

        '//---------------------------------------------------------------------------------------
        '-- Set The Pen Color
        '-- Ammended By: Gary Noble
        '//---------------------------------------------------------------------------------------
        If Me.DrawStyle = EDS_DefaultNet Then
            hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonShadow And &H1F&))
        Else
            hPen = CreatePen(PS_SOLID, 1, m_lColorBorder)
        End If

        hPenOld = SelectObject(lHDC, hPen)
        m_font.Bold = Me.SelectedFont.Bold

        hFontOld = SelectObject(lHDC, m_font.hFont)

        ' Draw the title bar border:
        SetBkColor lHDC, GetSysColor(vbButtonFace And &H1F&)
        MoveToEx lHDC, tCapR.Left, tCapR.Top, tJunk
        LineTo lHDC, tCapR.Right - 1, tCapR.Top
        LineTo lHDC, tCapR.Right - 1, tCapR.Bottom - 1
        LineTo lHDC, tCapR.Left, tCapR.Bottom - 1
        LineTo lHDC, tCapR.Left, tCapR.Top
        SetTextColor lHDC, GetSysColor(vbWindowText And &H1F&)
        SetBkMode lHDC, TRANSPARENT

        Dim tTextR As RECT
        LSet tTextR = tCapR
        m_tUnpinCloseR.Left = 0
        m_tUnpinCloseR.Top = 0
        m_tUnpinCloseR.Right = 0
        m_tUnpinCloseR.Bottom = 0
        If (m_iSelTab > 0) And (m_bShowCloseButton) Then
            If m_tTab(m_iSelTab).bCanClose Then
                ' close button:
                LSet m_tUnpinCloseR = tCapR
                m_tUnpinCloseR.Left = m_tUnpinCloseR.Right - (m_tUnpinCloseR.Bottom - m_tUnpinCloseR.Top)
                tTextR.Right = m_tUnpinCloseR.Left

                m_tUnpinCloseR.Left = m_tUnpinCloseR.Left + 2
                m_tUnpinCloseR.Right = m_tUnpinCloseR.Right - 2
                m_tUnpinCloseR.Top = m_tUnpinCloseR.Top + 2
                m_tUnpinCloseR.Bottom = m_tUnpinCloseR.Bottom - 2

                ' Draw it:
                drawTitleBarButtons lHDC
            End If
        End If

        If (m_bPinnable) Then
            LSet m_tUnpinPinR = tCapR
            m_tUnpinPinR.Right = m_tUnpinPinR.Right - (m_tUnpinCloseR.Right - m_tUnpinCloseR.Left)
            If ((m_tUnpinCloseR.Right - m_tUnpinCloseR.Left) > 0) Then
                m_tUnpinPinR.Right = m_tUnpinPinR.Right - 1
            End If
            m_tUnpinPinR.Left = m_tUnpinPinR.Right - (m_tUnpinPinR.Bottom - m_tUnpinCloseR.Top)
            tTextR.Right = m_tUnpinPinR.Left

            m_tUnpinPinR.Left = m_tUnpinPinR.Left + 2
            m_tUnpinPinR.Right = m_tUnpinPinR.Right - 2
            m_tUnpinPinR.Top = m_tUnpinPinR.Top + 2
            m_tUnpinPinR.Bottom = m_tUnpinPinR.Bottom - 2

            ' Draw it:
            drawTitleBarButtons lHDC

        End If

        ' Draw the caption:
        SetTextColor lHDC, GetSysColor(vbWindowText And &H1F&)
        If (m_iSelTab > 0) Then
            If Not (m_tTab(m_iSelTab).bEnabled) Then
                SetTextColor lHDC, GetSysColor(vb3DDKShadow And &H1F&)
            End If
        End If

        If m_bIsNt Then
            DrawTextW lHDC, StrPtr(" " & sCap), -1, tTextR, DT_SINGLELINE Or DT_VCENTER Or DT_LEFT Or DT_WORD_ELLIPSIS
        Else
            DrawText lHDC, " " & sCap, -1, tTextR, DT_SINGLELINE Or DT_VCENTER Or DT_LEFT Or DT_WORD_ELLIPSIS
        End If

        If Not (hFontOld = 0) Then
            SelectObject lHDC, hFontOld
        End If
        If Not (hPenOld = 0) Then
            SelectObject lHDC, hPenOld
        End If
        If Not (hPen = 0) Then
            DeleteObject hPen
        End If

        If Not bNoTx Then
            BitBlt UserControl.hdc, tCapR.Left, tCapR.Top, tCapR.Right - tCapR.Left, tCapR.Bottom - tCapR.Top, lHDC, tCapR.Left, tCapR.Top, vbSrcCopy
        End If
    End If

End Sub

Private Sub drawTitleBarButtons(Optional ByVal lhDCTo As Long = 0)
    Dim lHDC As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim lSize As Long

    If (lhDCTo = 0) Then
        lHDC = m_cMemDC.hdc
        If (lHDC = 0) Then
            lHDC = UserControl.hdc
            lhDCTo = lHDC
        End If
    Else
        lHDC = lhDCTo
    End If

    '//---------------------------------------------------------------------------------------
    '-- Draw The Background Depending On the DrawStyle
    '-- Ammended By: Gary Noble
    '//---------------------------------------------------------------------------------------
    If (m_tUnpinCloseR.Right - m_tUnpinCloseR.Top) > 0 Then
        If Me.DrawStyle = EDS_DefaultNet Then
            FillRect lHDC, m_tUnpinCloseR, GetSysColorBrush(vbButtonFace And &H1F&)
        Else
            UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), m_lColorTwoNormal, m_tUnpinCloseR
        End If

        drawButtonBorder lHDC, m_tUnpinCloseR, m_bUnpinCloseTrack, m_bUnpinCloseDown
        If (m_tUnpinCloseR.Bottom - m_tUnpinCloseR.Top > 40) Then
            lSize = 32
        Else
            lSize = 16
        End If
        lLeft = m_tUnpinCloseR.Left + ((m_tUnpinCloseR.Right - m_tUnpinCloseR.Left) - lSize) \ 2 + 1
        lTop = m_tUnpinCloseR.Top + ((m_tUnpinCloseR.Bottom - m_tUnpinCloseR.Top) - lSize) \ 2 + 1
        If (m_bUnpinCloseTrack And m_bUnpinCloseDown) Then
            lLeft = lLeft + 1
            lTop = lTop + 1
        End If
        DrawIconEx lHDC, _
                lLeft, _
                lTop, _
                m_hIconClose, lSize, lSize, 0, 0, DI_NORMAL

    End If

    If (m_tUnpinPinR.Right - m_tUnpinPinR.Left) > 0 Then

        '//---------------------------------------------------------------------------------------
        '-- Draw The Background Depending On the DrawStyle
        '-- Ammended By: Gary Noble
        '//---------------------------------------------------------------------------------------
        If Me.DrawStyle = EDS_DefaultNet Then
            FillRect lHDC, m_tUnpinPinR, GetSysColorBrush(vbButtonFace And &H1F&)
        Else
            UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), m_lColorTwoNormal, m_tUnpinPinR
        End If

        drawButtonBorder lHDC, m_tUnpinPinR, m_bUnpinPinTrack, m_bUnpinPinDown
        Dim hIcon As Long
        If (m_bPinned) Then
            hIcon = m_hIconPin
        Else
            hIcon = m_hIconUnpin
        End If
        If (m_tUnpinPinR.Bottom - m_tUnpinPinR.Top > 40) Then
            lSize = 32
        Else
            lSize = 16
        End If
        lLeft = m_tUnpinPinR.Left + ((m_tUnpinPinR.Right - m_tUnpinPinR.Left) - lSize) \ 2 + 1
        lTop = m_tUnpinPinR.Top + ((m_tUnpinPinR.Bottom - m_tUnpinPinR.Top) - lSize) \ 2 + 1
        If (m_bUnpinPinTrack And m_bUnpinPinDown) Then
            lLeft = lLeft + 1
            lTop = lTop + 1
        End If
        DrawIconEx lHDC, _
                lLeft, _
                lTop, _
                hIcon, lSize, lSize, 0, 0, DI_NORMAL
    End If

    If (lhDCTo = 0) Then
        BitBlt UserControl.hdc, m_tUnpinCloseR.Left, m_tUnpinCloseR.Top, m_tUnpinCloseR.Right - m_tUnpinCloseR.Left, m_tUnpinCloseR.Bottom - m_tUnpinCloseR.Top, lHDC, m_tUnpinCloseR.Left, m_tUnpinCloseR.Top, vbSrcCopy
        BitBlt UserControl.hdc, m_tUnpinPinR.Left, m_tUnpinPinR.Top, m_tUnpinPinR.Right - m_tUnpinPinR.Left, m_tUnpinPinR.Bottom - m_tUnpinPinR.Top, lHDC, m_tUnpinPinR.Left, m_tUnpinPinR.Top, vbSrcCopy
    End If
End Sub

Private Sub drawUnpinnedBorder()
    Dim tTR As RECT
    Dim hPenOld As Long
    Dim lHDC As Long
    Dim tJunk As POINTAPI
    Dim hPenLeft As Long
    Dim hPenRight As Long

    If (m_bPinnable And m_bOut And Not (m_bPinned)) Then
        lHDC = UserControl.hdc

        GetClientRect m_hWnd, tTR
        If (UserControl.Extender.Align = vbAlignLeft) Then
            tTR.Left = tTR.Right - 2
            If Me.DrawStyle = EDS_DefaultNet Then
                hPenLeft = CreatePen(PS_SOLID, 1, GetSysColor(vb3DShadow And &H1F&))
                hPenRight = CreatePen(PS_SOLID, 1, GetSysColor(vb3DDKShadow And &H1F&))
            Else
                hPenLeft = CreatePen(PS_SOLID, 1, m_lColorBorder)
                hPenRight = CreatePen(PS_SOLID, 1, m_lColorBorder)
            End If
        ElseIf (UserControl.Extender.Align = vbAlignRight) Then
            tTR.Right = 2
            If Me.DrawStyle = EDS_DefaultNet Then
                hPenLeft = CreatePen(PS_SOLID, 1, GetSysColor(vb3DHighlight And &H1F&))
                hPenRight = CreatePen(PS_SOLID, 1, GetSysColor(vb3DLight And &H1F&))
            Else
                hPenLeft = CreatePen(PS_SOLID, 1, m_lColorBorder)
                hPenRight = CreatePen(PS_SOLID, 1, m_lColorBorder)
            End If
        Else
            '
        End If

        ' Draw the borders
        hPenOld = SelectObject(lHDC, hPenLeft)
        MoveToEx lHDC, tTR.Left, tTR.Top, tJunk
        LineTo lHDC, tTR.Left, tTR.Bottom
        SelectObject lHDC, hPenOld

        hPenOld = SelectObject(lHDC, hPenRight)
        MoveToEx lHDC, tTR.Left + 1, tTR.Top, tJunk
        LineTo lHDC, tTR.Left + 1, tTR.Bottom
        SelectObject lHDC, hPenOld

        DeleteObject hPenLeft
        DeleteObject hPenRight

    End If
End Sub

Private Sub drawUnpinnedTabs()
    '
    ' Draw the unpinned titlebar:

    ' Draw the unpinned tabs:
    Dim lHDC As Long
    lHDC = picUnpinned.hdc

    ' Fill the background:
    Dim tR As RECT
    Dim hBr As Long
    GetClientRect picUnpinned.hWnd, tR

    '//---------------------------------------------------------------------------------------
    '-- Draw The Background Depending On the DrawStyle
    '-- Ammended By: Gary Noble
    '//---------------------------------------------------------------------------------------

    If Me.DrawStyle = EDS_DefaultNet Then
        hBr = CreateSolidBrush(BlendColor(vbButtonFace, vbWindowBackground, 80))
        FillRect lHDC, tR, hBr
        DeleteObject hBr
    Else
        UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), m_lColorTwoNormal, tR, True
    End If

    Dim hPen As Long
    Dim hPenOld As Long
    If Me.DrawStyle = EDS_DefaultNet Then
        hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonShadow And &H1F&))
    Else
        hPen = CreatePen(PS_SOLID, 1, m_lColorBorder)
    End If
    hPenOld = SelectObject(lHDC, hPen)

    ' Get the font to draw with
    Dim bVertical As Boolean
    If (UserControl.Extender.Align = vbAlignLeft Or UserControl.Extender.Align = vbAlignRight) Then

        bVertical = True
        Debug.Print bVertical
        ' Draw vertically
        Dim hFnt As Long
        Dim hFntOld As Long
        Dim tLF As LOGFONT


        pOLEFontToLogFont Me.SelectedFont, lHDC, tLF
        tLF.lfEscapement = 2700

        hFnt = CreateFont(tLF.lfHeight, tLF.lfWidth, tLF.lfEscapement, tLF.lfEscapement, tLF.lfWeight, 0, 0, 0, tLF.lfCharSet, tLF.lfClipPrecision, tLF.lfOutPrecision, tLF.lfQuality, tLF.lfPitchAndFamily, tLF.lfFaceName)

        If Not (hFnt = 0) Then
            hFntOld = SelectObject(lHDC, hFnt)
        End If

    Else
        ' Draw horizontally:
        hFntOld = SelectObject(lHDC, m_font.hFont)
        bVertical = False
    End If

    ' Now draw the tabs:
    Dim iC As Long
    Dim tTabR As RECT
    Dim tTextR As RECT
    Dim lIconLeft As Long
    Dim lIconTop As Long
    Dim tJunk As POINTAPI
    Dim lMaxTextSize As Long

    ' work out the maximum text size:

    For iC = 1 To m_iTabCount
        If m_bIsNt Then
            DrawTextW lHDC, StrPtr(m_tTab(iC).sCaption), -1, tTextR, DT_SINGLELINE Or DT_CALCRECT
        Else
            DrawText lHDC, m_tTab(iC).sCaption, -1, tTextR, DT_SINGLELINE Or DT_CALCRECT
        End If
        If (tTextR.Right - tTextR.Left + 8) > lMaxTextSize Then
            lMaxTextSize = (tTextR.Right - tTextR.Left + 8)
        End If
    Next iC

    LSet tTabR = tR
    For iC = 1 To m_iTabCount
        If (bVertical) Then
            tTabR.Bottom = tTabR.Top + m_lIconHeight + 8
        Else
            tTabR.Right = tTabR.Left + m_lIconWidth + 8
        End If

        ' Get the tab size:
        If (iC = m_iSelTab) Then
            ' we draw the text too
            If (bVertical) Then
                tTabR.Bottom = tTabR.Bottom + lMaxTextSize
            Else
                tTabR.Left = tTabR.Right + lMaxTextSize
            End If
        End If

        '//---------------------------------------------------------------------------------------
        '-- Draw The Background Depending On the DrawStyle
        '-- Ammended By: Gary Noble
        '//---------------------------------------------------------------------------------------
        If Me.DrawStyle = EDS_DefaultNet Then
            FillRect lHDC, tTabR, GetSysColorBrush(vbButtonFace And &H1F&)
        Else

            If Me.DrawStyle = EDS_Office2003Hot Then
                If (iC = m_iSelTab) Then
                    If AppThemed Then
                        UtilDrawBackground lHDC, BlendColor(m_lColorTwoSelected, vbWhite, 100), m_lColorOneSelected, tTabR, True
                    Else
                        UtilDrawBackground lHDC, BlendColor(m_lColorTwoSelected, vbWhite, 60), BlendColor(vbApplicationWorkspace, vbWhite, 200), tTabR, True
                    End If
                Else
                    UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 250), tTabR, True

                End If
            Else
                If (iC = m_iSelTab) Then
                    If AppThemed Then
                        UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(vbActiveTitleBar, vbWhite, 250), tTabR, True
                    Else
                        UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 150), BlendColor(vbActiveTitleBar, vbWhite, 150), tTabR, True
                    End If

                Else
                    UtilDrawBackground lHDC, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 250), tTabR, True

                End If
            End If

        End If

        If bVertical Then
            lIconLeft = ((tTabR.Right - tTabR.Left) - m_lIconWidth) \ 2
            lIconTop = tTabR.Top + 4
            tTextR.Top = lIconTop + m_lIconHeight + 8
        Else
            lIconLeft = tTabR.Left + 4
            lIconTop = ((tTabR.Bottom - tTabR.Top) - m_lIconHeight) \ 2
            tTextR.Left = lIconLeft + m_lIconWidth + 8
        End If

        If (m_tTab(iC).lIconIndex > -1) Then
            If (m_tTab(iC).bEnabled) Then
                ImageListDrawIcon m_ptrVb6ImageList, lHDC, m_hIml, _
                        m_tTab(iC).lIconIndex, _
                        lIconLeft, _
                        lIconTop
            Else
                ImageListDrawIconDisabled m_ptrVb6ImageList, lHDC, m_hIml, _
                        m_tTab(iC).lIconIndex, _
                        lIconLeft, _
                        lIconTop, _
                        m_lIconWidth
            End If
        End If

        If (iC = m_iSelTab) Then
            SetTextColor lHDC, GetSysColor(vb3DDKShadow And &H1F&)

            If (bVertical) Then
                Dim tSwap As RECT

                tSwap.Left = tTabR.Right - 4
                tSwap.Top = tTextR.Top
                tSwap.Right = 4
                tSwap.Bottom = tTabR.Bottom + (tTextR.Right - tTextR.Left)
                ' LSet tSwap = tTabR
                LSet tTextR = tSwap
                If m_bIsNt Then
                    DrawTextW lHDC, StrPtr(m_tTab(iC).sCaption), -1, tTextR, DT_SINGLELINE
                Else
                    DrawText lHDC, m_tTab(iC).sCaption, -1, tTextR, DT_SINGLELINE
                End If
            End If
        End If
        MoveToEx lHDC, tTabR.Left, tTabR.Top, tJunk
        LineTo lHDC, tTabR.Right - 1, tTabR.Top
        LineTo lHDC, tTabR.Right - 1, tTabR.Bottom
        LineTo lHDC, tTabR.Left, tTabR.Bottom
        LineTo lHDC, tTabR.Left, tTabR.Top

        LSet m_tTab(iC).tPinnedR = tTabR

        If (bVertical) Then
            tTabR.Top = tTabR.Bottom
        Else
            tTabR.Left = tTabR.Right
        End If

    Next iC

    If Not (hPenOld = 0) Then
        SelectObject lHDC, hPenOld
    End If
    If Not (hPen = 0) Then
        DeleteObject hPen
    End If

    If Not (hFntOld = 0) Then
        SelectObject lHDC, hFntOld
    End If
    If Not (hFnt = 0) Then
        DeleteObject hFnt
    End If

    ' Show the changes:
    picUnpinned.Refresh
    '
End Sub

Private Function ensureEndTabOffset()
    Dim lMaxRight As Long
    Dim lSize As Long
    Dim tR As RECT
    If (m_iTabCount > 0) Then
        GetTabWindowRect tR

        lMaxRight = m_tTab(m_iTabCount).tTabR.Right
        lSize = tR.Right - tR.Left
        If (m_bAllowScroll) Then
            lSize = lSize - m_lButtonSize * 2
        End If
        lSize = lSize - m_lButtonSize

        If (lMaxRight > lSize) Then
            If (lMaxRight - m_lOffsetX < lSize) Then
                m_lOffsetX = lMaxRight - lSize + 4
            End If
        ElseIf (lSize > lMaxRight) Then
            If (m_lOffsetX > 0) Then
                m_lOffsetX = 0
            End If
        End If
    End If
End Function

Friend Function fAdd( _
      Optional Key As Variant, _
      Optional KeyBefore As Variant, _
      Optional Caption As String, _
      Optional IconIndex As Long = -1 _
      ) As cTab
    ' Check key:
    Dim sKey As String

    If Not IsMissing(Key) Then
        ' validate key.
        If IsNumeric(Key) Then
            ' invalid key
            Err.Raise 13, App.EXEName & ".vbalDTabControlX"
            Exit Function
        End If
        On Error Resume Next
        sKey = Key
        If (Err.Number <> 0) Then
            ' invalid key
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
        Dim i As Long
        For i = 1 To m_iTabCount
            If (m_tTab(i).sKey = sKey) Then
                ' duplicate key
                Err.Raise 457, App.EXEName & ".vbalDTabControlX"
                Exit Function
            End If
        Next i
    End If

    ' Check KeyBefore:
    Dim iIndexBefore As Long
    iIndexBefore = 0
    If Not IsMissing(KeyBefore) Then
        On Error Resume Next
        iIndexBefore = tabForKey(KeyBefore)
        If (Err.Number <> 0) Then
            Err.Raise Err.Number, App.EXEName & ".vbalDTabControlX", Err.Description
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If

    ' Ok all checks passed. We can add the item.
    ' Check if this is an insert:
    Dim iTabIndex As Long
    m_iTabCount = m_iTabCount + 1
    If (m_iTabCount = 1) Then
        m_iSelTab = 1
    End If
    ReDim Preserve m_tTab(1 To m_iTabCount) As TabInfo
    If (iIndexBefore > 0) Then
        ' Fix: should step backwards!
        For i = m_iTabCount - 1 To iIndexBefore Step -1
            LSet m_tTab(i + 1) = m_tTab(i)
        Next i
        iTabIndex = iIndexBefore
    Else
        iTabIndex = m_iTabCount
    End If

    ' set the info:
    m_tTab(iTabIndex).sCaption = Caption
    m_tTab(iTabIndex).lIconIndex = IconIndex
    m_tTab(iTabIndex).bCanClose = True
    m_tTab(iTabIndex).bEnabled = True
    m_tTab(iTabIndex).lId = nextId()
    If (sKey = "") Then
        m_tTab(iTabIndex).sKey = "I" & m_tTab(iTabIndex).lId
    Else
        m_tTab(iTabIndex).sKey = sKey
    End If
    drawTabs

    Dim cT As New cTab
    cT.fInit ObjPtr(Me), m_hWnd, m_tTab(iTabIndex).lId

    Set fAdd = cT

End Function

Friend Function fItem( _
      Key As Variant _
   )
    Dim iIndex As Long
    On Error Resume Next
    iIndex = tabForKey(Key)
    If (Err.Number <> 0) Then
        Err.Raise Err.Number, App.EXEName & ".vbalDTabControlX", Err.Description
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    Dim cT As New cTab
    cT.fInit ObjPtr(Me), m_hWnd, m_tTab(iIndex).lId
    Set fItem = cT

End Function

Public Property Get Font() As iFont
    Dim iFnt As iFont
    Dim iFntC As iFont
    Set iFnt = m_font
    iFnt.Clone iFntC
    Set Font = iFntC
End Property

Public Property Let Font(iFnt As iFont)
    pSetFont iFnt
End Property

Public Property Set Font(iFnt As iFont)
    pSetFont iFnt
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_oForeColor
End Property

Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
    If (m_oForeColor <> oColor) Then
        m_oForeColor = oColor
        drawTabs
        PropertyChanged "ForeColor"
    End If
End Property

Friend Function fRemove(Key As Variant)

    ' Get tab to remove:
    Dim iToRemove As Long
    On Error Resume Next
    iToRemove = tabForKey(Key)
    If (Err.Number <> 0) Then
        On Error GoTo 0
        Err.Raise Err.Number, App.EXEName & ".vbalDTabControlX", Err.Description
        Exit Function
    End If
    On Error GoTo 0

    ' its valid.
    Dim ctl As Control
    If (pbGetTabPanel(iToRemove, ctl)) Then
        pbPanelVisible ctl, False
    End If

    If (m_iTabCount = 1) Then
        m_iTabCount = 0
        m_iSelTab = 0
        Erase m_tTab
    Else
        If (m_iSelTab = iToRemove) Then
            If (m_iSelTab = m_iTabCount) Then
                m_iSelTab = m_iTabCount - 1
            End If
        End If
        Dim i As Long
        For i = iToRemove + 1 To m_iTabCount
            LSet m_tTab(i - 1) = m_tTab(i)
        Next i
        m_iTabCount = m_iTabCount - 1
        ReDim Preserve m_tTab(1 To m_iTabCount) As TabInfo
    End If
    drawTabs

End Function

Friend Property Get fTabCanClose(ByVal lId As Long) As Boolean
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabCanClose = m_tTab(lIndex).bCanClose
    End If
End Property

Friend Property Let fTabCanClose(ByVal lId As Long, ByVal bCanClose As Boolean)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        m_tTab(lIndex).bCanClose = bCanClose
        drawTabs
        If (m_bPinnable) Then
            If (m_bPinned) Then
                drawTitleBar
            Else
                drawUnpinnedTabs
            End If
        End If
    End If
End Property

Friend Property Get fTabCaption(ByVal lId As Long) As String
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabCaption = m_tTab(lIndex).sCaption
    End If
End Property

Friend Property Let fTabCaption(ByVal lId As Long, ByVal sCaption As String)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        m_tTab(lIndex).sCaption = sCaption
        drawTabs
        If (m_bPinnable) Then
            If (m_bPinned) Then
                drawTitleBar
            Else
                drawUnpinnedTabs
            End If
        End If
    End If
End Property

Friend Property Get fTabCount() As Long
    fTabCount = m_iTabCount
End Property

Friend Property Get fTabEnabled(ByVal lId As Long) As Boolean
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabEnabled = m_tTab(lIndex).bEnabled
    End If
End Property

Friend Property Let fTabEnabled(ByVal lId As Long, ByVal bEnabled As Boolean)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        m_tTab(lIndex).bEnabled = bEnabled
        drawTabs
        If (m_bPinnable) Then
            If (m_bPinned) Then
                drawTitleBar
            Else
                drawUnpinnedTabs
            End If
        End If
    End If
End Property

Friend Property Get fTabIconIndex(ByVal lId As Long) As Long
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabIconIndex = m_tTab(lIndex).lIconIndex
    End If
End Property

Friend Property Let fTabIconIndex(ByVal lId As Long, ByVal lIconIndex As Long)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        m_tTab(lIndex).lIconIndex = lIconIndex
        drawTabs
        If (m_bPinnable) Then
            If (m_bPinned) Then
                drawTitleBar
            Else
                drawUnpinnedTabs
            End If
        End If
    End If
End Property

Friend Property Get fTabIndex(ByVal lId As Long) As Long
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabIndex = lIndex
    End If
End Property

Friend Property Let fTabIndex(ByVal lId As Long, ByVal lIndex As Long)
    Dim lCurrentIndex As Long
    If (getTabForId(lId, lCurrentIndex)) Then
        If Not (lIndex = lCurrentIndex) Then
            If (lIndex > 0) And (lIndex <= m_iTabCount) Then
                replaceWithCandidate lCurrentIndex, lIndex
            Else
                ' New index out of range
                Err.Raise 9, App.EXEName & ".vbalDTabControlX"
            End If
        End If
    End If
End Property

Friend Property Get fTabItemData(ByVal lId As Long) As Long
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabItemData = m_tTab(lIndex).lItemData
    End If
End Property

Friend Property Let fTabItemData(ByVal lId As Long, ByVal lItemData As Long)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        m_tTab(lIndex).lItemData = lItemData
    End If
End Property

Friend Property Get fTabKey(ByVal lId As Long) As String
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabKey = m_tTab(lIndex).sKey
    End If
End Property

Friend Property Get fTabPanel(ByVal lId As Long) As Object
    Dim ctlThis As Object
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        ' Fix thanks to Matt Funnell: use lIndex not lId to find panel:
        If pbGetTabPanel(lIndex, ctlThis) Then
            Set fTabPanel = ctlThis
        End If
    End If
End Property

Friend Property Let fTabPanel(ByVal lId As Long, ByVal ctlThis As Object)
    Dim ctlPanel As Object
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        If pbGetTabPanel(lIndex, ctlPanel) Then
            pbPanelVisible ctlPanel, False
        End If
        Set ctlThis.Container = UserControl.Extender
        m_tTab(lIndex).lObjPtrPanel = ObjPtr(ctlThis)
        If (lIndex = m_iSelTab) Then
            pPanelSize
        Else
            pbPanelVisible ctlThis, False
        End If
    End If
End Property

Friend Property Get fTabSelected(ByVal lId As Long) As Boolean
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabSelected = (lIndex = m_iSelTab)
    End If
End Property

Friend Property Let fTabSelected(ByVal lId As Long, ByVal bSelected As Boolean)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        If Not (lIndex = m_iSelTab) Then
            m_iSelTab = lIndex
            drawTabs
            pPanelSize
            If (m_bPinnable) Then
                If (m_bPinned) Then
                    drawTitleBar
                Else
                    drawUnpinnedTabs
                End If
            End If
        End If
    End If
End Property

Friend Property Get fTabTag(ByVal lId As Long) As String
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabTag = m_tTab(lIndex).sTag
    End If
End Property

Friend Property Let fTabTag(ByVal lId As Long, ByVal sTag As String)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        m_tTab(lIndex).sTag = sTag
    End If
End Property

Friend Property Get fTabToolTipText(ByVal lId As Long) As String
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        fTabToolTipText = m_tTab(lIndex).sToolTipText
    End If
End Property

Friend Property Let fTabToolTipText(ByVal lId As Long, ByVal sToolTipText As String)
    Dim lIndex As Long
    If (getTabForId(lId, lIndex)) Then
        m_tTab(lIndex).sToolTipText = sToolTipText
    End If
End Property

Private Sub getCloseButtonRect(tRClose As RECT)
    LSet tRClose = m_tButtonR
    tRClose.Top = tRClose.Top
    tRClose.Bottom = tRClose.Top + m_lButtonSize
    If (m_bAllowScroll) Then
        OffsetRect tRClose, m_lButtonSize * 2, 0
        tRClose.Right = tRClose.Left + m_lButtonSize
    End If
End Sub




Private Sub getLeftButtonRect(tRLeft As RECT)
    LSet tRLeft = m_tButtonR
    tRLeft.Top = tRLeft.Top
    tRLeft.Bottom = tRLeft.Top + m_lButtonSize
    tRLeft.Right = tRLeft.Left + m_lButtonSize
End Sub

Private Sub getRightButtonRect(tRRight As RECT)
    LSet tRRight = m_tButtonR
    tRRight.Top = tRRight.Top
    tRRight.Bottom = tRRight.Top + m_lButtonSize
    tRRight.Left = tRRight.Left + m_lButtonSize
    tRRight.Right = tRRight.Left + m_lButtonSize
End Sub

Private Function getTabForId(ByVal lId As Long, ByRef lIndex As Long) As Boolean
    Dim i As Long
    For i = 1 To m_iTabCount
        If (m_tTab(i).lId = lId) Then
            lIndex = i
            getTabForId = True
            Exit Function
        End If
    Next i
    Err.Raise 9, App.EXEName & ".vbalDTabControlX"
End Function

Private Sub GetTabWindowRect(tR As RECT)
    GetClientRect m_hWnd, tR
End Sub

Public Function GetThemeName(hWnd As Long) As String
    'Gett the current Theme name, ans Scheme Color
    Dim hTheme As Long
    Dim sShellStyle As String
    Dim sThemeFile As String
    Dim lPtrThemeFile As Long, lPtrColorName As Long, hres As Long
    Dim iPos As Long
    On Error Resume Next
    hTheme = OpenThemeData(hWnd, StrPtr("ExplorerBar"))

    If Not hTheme = 0 Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColorName = VarPtr(bColorName(0))
        hres = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)

        sThemeFile = bThemeFile
        iPos = InStr(sThemeFile, vbNullChar)
        If (iPos > 1) Then sThemeFile = Left(sThemeFile, iPos - 1)
        m_sCurrentSystemThemename = bColorName
        iPos = InStr(m_sCurrentSystemThemename, vbNullChar)
        If (iPos > 1) Then m_sCurrentSystemThemename = Left(m_sCurrentSystemThemename, iPos - 1)

        sShellStyle = sThemeFile
        For iPos = Len(sThemeFile) To 1 Step -1
            If (Mid(sThemeFile, iPos, 1) = "\") Then
                sShellStyle = Left(sThemeFile, iPos)
                Exit For
            End If
        Next iPos
        sShellStyle = sShellStyle & "Shell\" & m_sCurrentSystemThemename & "\ShellStyle.dll"
        CloseThemeData hTheme
    Else
        m_sCurrentSystemThemename = "Classic"
    End If
    ' Debug.Print m_sCurrentSystemThemename

End Function

Private Function getTypicalScrollDistance() As Long
    Dim tR As RECT
    Dim lDist As Long
    Dim i As Long
    Dim lTabAvg As Long
    If (m_iTabCount > 0) Then
        For i = 1 To m_iTabCount
            lTabAvg = lTabAvg + (m_tTab(i).tTabR.Right - m_tTab(i).tTabR.Left)
        Next i
        lTabAvg = lTabAvg \ m_iTabCount
        GetTabWindowRect tR
        lDist = (tR.Right - tR.Left)
        If (lDist > lTabAvg * 2) Then
            lDist = lDist - lTabAvg
        End If
        If (lDist < 0) Then
            lDist = lTabAvg \ 2
        End If
        If (lDist < 0) Then
            lDist = 32
        End If
        getTypicalScrollDistance = lDist
    End If
End Function

Private Sub GradientFillRect(ByVal lHDC As Long, _
                             tR As RECT, _
                             ByVal oStartColor As OLE_COLOR, _
                             ByVal oEndColor As OLE_COLOR, _
                             ByVal eDir As GradientFillRectType)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR         As GRADIENT_RECT
    Dim hBrush      As Long
    Dim lStartColor As Long
    Dim lEndColor   As Long

    'Dim lR As Long
    ' Use GradientFill:
    If (HasGradientAndTransparency) Then
        lStartColor = TranslateColor(oStartColor)
        lEndColor = TranslateColor(oEndColor)
        setTriVertexColor tTV(0), lStartColor
        tTV(0).x = tR.Left
        tTV(0).y = tR.Top
        setTriVertexColor tTV(1), lEndColor
        tTV(1).x = tR.Right
        tTV(1).y = tR.Bottom
        tGR.UpperLeft = 0
        tGR.LowerRight = 1
        GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
    Else
        ' Fill with solid brush:
        hBrush = CreateSolidBrush(TranslateColor(oEndColor))
        FillRect lHDC, tR, hBrush
        DeleteObject hBrush
    End If

End Sub

Public Property Get HasGradientAndTransparency()

    HasGradientAndTransparency = m_bHasGradientAndTransparency

End Property

Private Function hitTestButton() As Long
    Dim tR As RECT
    Dim tP As POINTAPI

    GetCursorPos tP
    ScreenToClient m_hWnd, tP
    If (m_bAllowScroll) Then
        If IsLeftButtonEnabled() Then
            getLeftButtonRect tR
            If Not (PtInRect(tR, tP.x, tP.y) = 0) Then
                hitTestButton = 1
                Exit Function
            End If
        End If
        If IsRightButtonEnabled() Then
            getRightButtonRect tR
            If Not (PtInRect(tR, tP.x, tP.y) = 0) Then
                hitTestButton = 2
                Exit Function
            End If
        End If
    End If
    If IsCloseButtonEnabled() Then
        getCloseButtonRect tR
        If Not (PtInRect(tR, tP.x, tP.y) = 0) Then
            hitTestButton = 3
        End If
    End If
End Function

Private Function hitTestTab() As Long
    '
    Dim tP As POINTAPI
    GetCursorPos tP
    ScreenToClient m_hWnd, tP
    tP.x = tP.x + m_lOffsetX

    Dim i As Long
    For i = 1 To m_iTabCount
        If Not (PtInRect(m_tTab(i).tTabR, tP.x, tP.y) = 0) Then
            If (PtInRect(m_tButtonR, tP.x - m_lOffsetX, tP.y) = 0) Then
                If (m_tTab(i).bEnabled) Or (m_bAllowSelectDisabledTabs) Then
                    hitTestTab = i
                    Exit For
                End If
            End If
        End If
    Next i
    '
End Function

Public Property Let ImageList( _
        ByRef vImageList As Variant _
    )
    m_hIml = 0
    m_ptrVb6ImageList = 0
    If (VarType(vImageList) = vbLong) Then
        ' Assume a handle to an image list:
        m_hIml = vImageList
    ElseIf (VarType(vImageList) = vbObject) Then
        ' Assume a VB image list:
        On Error Resume Next
        ' Get the image list initialised..
        vImageList.ListImages(1).draw 0, 0, 0, 1
        m_hIml = vImageList.hImageList
        If (Err.Number = 0) Then
            ' Check for VB6 image list:
            If (TypeName(vImageList) = "ImageList") Then
                If (vImageList.ListImages.Count <> ImageList_GetImageCount(m_hIml)) Then
                    Dim o As Object
                    Set o = vImageList
                    m_ptrVb6ImageList = ObjPtr(o)
                End If
            End If
        Else
            Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
        End If
        On Error GoTo 0
    End If
    If (m_hIml <> 0) Then
        If (m_ptrVb6ImageList <> 0) Then
            m_lIconWidth = vImageList.ImageWidth
            m_lIconHeight = vImageList.ImageHeight
            If (UserControl.Extender.Align = vbAlignLeft) Or (UserControl.Extender.Align = vbAlignRight) Then
                pSetTabHeight
                UserControl_Resize
            End If
        Else
            Dim rc As RECT
            ImageList_GetImageRect m_hIml, 0, rc
            m_lIconWidth = rc.Right - rc.Left
            m_lIconHeight = rc.Bottom - rc.Top
            If (UserControl.Extender.Align = vbAlignLeft) Or (UserControl.Extender.Align = vbAlignRight) Then
                pSetTabHeight
                UserControl_Resize
            End If
        End If
    End If
    drawTabs
End Property

Private Sub ImageListDrawIcon( _
        ByVal ptrVb6ImageList As Long, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal bSelected As Boolean = False, _
        Optional ByVal bBlend25 As Boolean = False _
    )
    Dim lFlags As Long
    Dim lR As Long

    lFlags = ILD_TRANSPARENT
    If (bSelected) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If (bBlend25) Then
        lFlags = lFlags Or ILD_BLEND25
    End If
    If (ptrVb6ImageList <> 0) Then
        Dim o As Object
        On Error Resume Next
        Set o = ObjectFromPtr(ptrVb6ImageList)
        If Not (o Is Nothing) Then
            If ((lFlags And ILD_SELECTED) = ILD_SELECTED) Then
                lFlags = 2    ' best we can do in VB6
            End If
            o.ListImages(iIconIndex + 1).draw hdc, lX * Screen.TwipsPerPixelX, lY * Screen.TwipsPerPixelY, lFlags
        End If
        On Error GoTo 0
    Else
        lR = ImageList_Draw( _
                hIml, _
                iIconIndex, _
                hdc, _
                lX, _
                lY, _
                lFlags)
        If (lR = 0) Then
            'Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
        End If
    End If
End Sub

Private Sub ImageListDrawIconDisabled( _
        ByVal ptrVb6ImageList As Long, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long, _
        Optional ByVal asShadow As Boolean _
    )
    Dim lR As Long
    Dim hIcon As Long

    hIcon = 0
    If (ptrVb6ImageList <> 0) Then
        Dim o As Object
        On Error Resume Next
        Set o = ObjectFromPtr(ptrVb6ImageList)
        If Not (o Is Nothing) Then
            hIcon = o.ListImages(iIconIndex + 1).ExtractIcon()
        End If
        On Error GoTo 0
    Else
        hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
    End If
    If (hIcon <> 0) Then
        If (asShadow) Then
            Dim hBr As Long
            hBr = GetSysColorBrush(vb3DShadow And &H1F)
            lR = DrawState(hdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO)
            DeleteObject hBr
        Else
            lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
        End If
        DestroyIcon hIcon
    End If

End Sub

Private Function inIde() As Boolean
    m_bInIde = True
    inIde = m_bInIde
End Function

Public Property Get Is2000OrAbove() As Boolean

    Is2000OrAbove = m_bIs2000OrAbove

End Property

Private Function IsCloseButtonEnabled() As Boolean
    Dim bR As Boolean
    bR = False
    If (m_bShowCloseButton) Then
        If (m_iTabCount > 0) Then
            If (m_iSelTab > 0) Then
                bR = (m_tTab(m_iSelTab).bCanClose)
            End If
        End If
    End If
    IsCloseButtonEnabled = bR
End Function

Private Function IsLeftButtonEnabled() As Boolean
    IsLeftButtonEnabled = (m_lOffsetX > 0)
End Function

Public Property Get IsNt() As Boolean

    IsNt = m_bIsNt

End Property

Private Function IsRightButtonEnabled() As Boolean
    If (m_iTabCount > 0) Then
        IsRightButtonEnabled = ((m_tTab(m_iTabCount).tTabR.Right - m_lOffsetX) > m_tButtonR.Left)
    End If
End Function

Public Property Get IsXp() As Boolean

    IsXp = m_bIsXp

End Property

Private Function KeyExists(ByVal c As Collection, ByVal Key As String) As Boolean
    On Error Resume Next
    Dim oItem As Variant
    oItem = c(Key)
    If (Err.Number = 0) Then
        KeyExists = True
    End If
End Function

Private Sub loadResources()
    Debug.Assert inIde
    If (m_bInIde) Then
        m_hIconPin = LoadImageString(App.hInstance, App.Path & "\res\pinned.ico", IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        m_hIconUnpin = LoadImageString(App.hInstance, App.Path & "\res\unpinned.ico", IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        m_hIconClose = LoadImageString(App.hInstance, App.Path & "\res\close.ico", IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
    Else
        m_hIconPin = LoadImageLong(App.hInstance, 64, IMAGE_ICON, 16, 16, 0)
        m_hIconUnpin = LoadImageLong(App.hInstance, 65, IMAGE_ICON, 16, 16, 0)
        m_hIconClose = LoadImageLong(App.hInstance, 66, IMAGE_ICON, 16, 16, 0)
    End If
End Sub

Private Sub m_tmr_Timer()
    '
    If Not (m_bUnpinCloseDown Or m_bUnpinPinDown) Then
        If GetAsyncKeyState(vbLeftButton) = 0 Then
            unshowPinned
        End If
    End If
    '
End Sub

Private Sub m_tmrPinButton_Timer()
    Dim tP As POINTAPI
    If m_bUnpinCloseTrack Then
        GetCursorPos tP
        ScreenToClient m_hWnd, tP
        If (PtInRect(m_tUnpinCloseR, tP.x, tP.y) = 0) Then
            m_tmrPinButton.Enabled = False
            m_bUnpinCloseTrack = False
            drawTitleBarButtons
        End If
    ElseIf m_bUnpinPinTrack Then
        GetCursorPos tP
        ScreenToClient m_hWnd, tP
        If (PtInRect(m_tUnpinPinR, tP.x, tP.y) = 0) Then
            m_tmrPinButton.Enabled = False
            m_bUnpinPinTrack = False
            drawTitleBarButtons
        End If
    Else
        m_tmrPinButton.Enabled = False
    End If
End Sub

Private Function nextId() As Long
    m_lIdGenerator = m_lIdGenerator + 1
    nextId = m_lIdGenerator
End Function

Private Property Get NoPalette(Optional ByVal bForce As Boolean = False) As Boolean
    Static bOnce As Boolean
    Static bNoPalette As Boolean
    Dim lHDC As Long
    Dim lBits As Long
    If (bForce) Then
        bOnce = False
    End If
    If Not (bOnce) Then
        lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        If (lHDC <> 0) Then
            lBits = GetDeviceCaps(lHDC, BITSPIXEL)
            If (lBits <> 0) Then
                bOnce = True
            End If
            bNoPalette = (lBits > 8)
            DeleteDC lHDC
        End If
    End If
    NoPalette = bNoPalette
End Property

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oTemp As Object
    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oTemp, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oTemp
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oTemp, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but it will be because of the subclass, not the uncounted reference
End Property

Private Function pbGetTabPanel(ByVal lIndex As Long, ByRef ctlThis As Object) As Boolean
    Dim ctl As Control
    Dim lPtr As Long
    Dim i As Long
    For Each ctl In UserControl.ContainedControls
        lPtr = ObjPtr(ctl)
        If lPtr = m_tTab(lIndex).lObjPtrPanel Then
            Set ctlThis = ctl
            pbGetTabPanel = True
        End If
    Next
End Function

Private Function pbPanelVisible(ByRef ctlThis As Object, ByVal bState As Boolean)
    ctlThis.Visible = bState
End Function

Private Sub picUnpinned_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
End Sub

Private Sub picUnpinned_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    Dim tP As POINTAPI
    GetCursorPos tP
    pSetToolTipText tP

    If (Button = 0) Then
        ScreenToClient picUnpinned.hWnd, tP
        Dim i As Long
        For i = 1 To m_iTabCount
            If Not (PtInRect(m_tTab(i).tPinnedR, tP.x, tP.y) = 0) Then
                If Not (i = m_iSelTab) Or Not (m_bOut) Then
                    m_iSelTab = i
                    drawUnpinnedTabs
                    showUnpinned
                End If
            End If
        Next i
    End If
    '
End Sub

Private Sub picUnpinned_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
End Sub

Private Sub picUnpinned_Resize()
    '
    drawUnpinnedTabs
    '
End Sub

Public Property Get Pinnable() As Boolean
    Pinnable = m_bPinnable
End Property

Public Property Let Pinnable(ByVal bState As Boolean)
    m_bPinnable = bState
    PropertyChanged "Pinnable"
End Property

Public Property Get Pinned() As Boolean
    Pinned = m_bPinned
End Property

Public Property Let Pinned(ByVal bState As Boolean)
    m_bPinned = bState
    If (m_bPinnable) Then
        UserControl_Resize
    Else
        ' not relevant
    End If
    PropertyChanged "Pinned"
End Property

Private Sub pOLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)
    Dim sFont As String
    Dim iChar As Integer

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name
        ' There is a quicker way involving StrConv and CopyMemory, but
        ' this is simpler!:
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
        Next iChar
        ' Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        .lfCharSet = fntThis.Charset
        .lfQuality = ANTIALIASED_QUALITY
    End With

End Sub

Private Function pPanelSize()
    Dim ctlPanel As Control
    Dim ctl As Control
    Dim rc As RECT
    Dim fL As Single, fT As Single, fW As Single, fH As Single
    Dim lTab As Long, lOffset As Long

    If m_iTabCount > 0 Then
        lTab = m_iSelTab
        If lTab > 0 Then
            If pbGetTabPanel(lTab, ctlPanel) Then
                LSet rc = m_tClientR
                fL = ScaleX(rc.Left, vbPixels, UserControl.ScaleMode)
                fT = ScaleY(rc.Top, vbPixels, UserControl.ScaleMode)
                fW = ScaleX(rc.Right - rc.Left - 2, vbPixels, UserControl.ScaleMode)
                fH = ScaleY(rc.Bottom - rc.Top, vbPixels, UserControl.ScaleMode)
                If (m_bPinnable And Not m_bPinned) Then
                    pbPanelVisible ctlPanel, False
                Else
                    On Error Resume Next
                    ctlPanel.Move fL, fT, fW, fH
                    On Error GoTo 0
                    pbPanelVisible ctlPanel, True
                End If
            End If
        End If
    End If
    For Each ctl In UserControl.ContainedControls
        If ctl Is ctlPanel Then
        Else
            pbPanelVisible ctl, False
        End If
    Next

End Function

Private Sub pSetFont(iFnt As iFont)
    Dim iFntC As iFont
    iFnt.Clone iFntC
    Set m_font = iFntC
    pSetTabHeight
    PropertyChanged "Font"
End Sub

Private Sub pSetSelectedFont(iFnt As iFont)
    Dim iFntC As iFont
    iFnt.Clone iFntC
    Set m_fontSelected = iFntC
    pSetTabHeight
    PropertyChanged "SelectedFont"
End Sub

Private Sub pSetTabHeight()
    Dim tR As RECT
    Dim lHeight As Long
    Dim lSelectedHeight As Long
    Dim hFontOld As Long
    Dim bResize As Boolean

    ' Bug reported by Andrea Batina (a_batina@hotmail.com):
    ' Need to configure the height of the items for the new
    ' font:

    ' First get the standard font:
    tR.Bottom = 128
    tR.Right = 128
    hFontOld = SelectObject(m_cMemDC.hdc, m_font.hFont)
    DrawText m_cMemDC.hdc, "Zg", -1, tR, DT_CALCRECT Or DT_SINGLELINE Or DT_LEFT
    SelectObject m_cMemDC.hdc, hFontOld
    lHeight = (tR.Bottom - tR.Top)

    ' Now the selected font:
    tR.Bottom = 128
    tR.Right = 128
    hFontOld = SelectObject(m_cMemDC.hdc, m_fontSelected.hFont)
    DrawText m_cMemDC.hdc, "Zg", -1, tR, DT_CALCRECT Or DT_SINGLELINE Or DT_LEFT
    SelectObject m_cMemDC.hdc, hFontOld
    lSelectedHeight = (tR.Bottom - tR.Top)

    If (lHeight >= lSelectedHeight) Then
        lHeight = lHeight + 11
    Else
        lHeight = lSelectedHeight + 11
    End If

    ' Now check the icon height:
    If (lHeight < m_lIconHeight + 4) Then
        lHeight = m_lIconHeight + 4
    End If

    If Not (m_lTabHeight = lHeight) Then
        m_lTabHeight = lHeight
        bResize = True
    End If
    If Not (m_lTitleBarHeight = lHeight) Then
        m_lTitleBarHeight = lHeight
        bResize = True
    End If
    If Not (m_lUnpinnedWidth = lHeight) Then
        m_lUnpinnedWidth = lHeight
        bResize = True
    End If
    If (bResize) Then
        UserControl_Resize
    End If
    UserControl.Refresh

End Sub

Private Sub pSetToolTipText(tP As POINTAPI)
    ' Where are we?

    Dim tR As RECT
    Dim sToolTip As String
    Dim i As Long
    Dim tPC As POINTAPI
    LSet tPC = tP

    If (m_bPinnable And Not m_bPinned) Then
        GetWindowRect picUnpinned.hWnd, tR
        If Not (PtInRect(tR, tP.x, tP.y) = 0) Then
            ' Check tabs:
            ScreenToClient picUnpinned.hWnd, tPC
            For i = 1 To m_iTabCount
                If Not (PtInRect(m_tTab(i).tPinnedR, tPC.x, tPC.y) = 0) Then
                    picUnpinned.ToolTipText = m_tTab(i).sToolTipText
                    Exit Sub
                End If
            Next i
            picUnpinned.ToolTipText = ""
        Else
            ' Check title bar
            ScreenToClient m_hWnd, tPC
            If Not (PtInRect(m_tUnpinCloseR, tPC.x, tPC.y)) = 0 Then
                sToolTip = "Close"
            ElseIf Not (PtInRect(m_tUnpinPinR, tPC.x, tPC.y)) = 0 Then
                sToolTip = "Autohide"
            End If
        End If
    Else
        ' Check buttons
        i = hitTestButton()
        If (i > 0) Then
            Select Case i
                Case 1
                    sToolTip = "Scroll Left"
                Case 2
                    sToolTip = "Scroll Right"
                Case 3
                    sToolTip = "Close"
            End Select
        Else
            ' Check tabs:
            i = hitTestTab()
            If (i > 0) Then
                sToolTip = m_tTab(i).sToolTipText
            Else
                ' Check title bar:
                ScreenToClient m_hWnd, tPC
                If Not (PtInRect(m_tUnpinCloseR, tPC.x, tPC.y)) = 0 Then
                    sToolTip = "Close"
                ElseIf Not (PtInRect(m_tUnpinPinR, tPC.x, tPC.y)) = 0 Then
                    sToolTip = "Autohide"
                Else

                End If
            End If
        End If
    End If

    If Not (sToolTip = m_sLastToolTip) Then
        Debug.Print "Setting tooltip to:", sToolTip
        On Error Resume Next
        UserControl.Extender.ToolTipText = sToolTip
        m_sLastToolTip = sToolTip
    End If

End Sub

Private Function replaceWithCandidate(ByVal iDragging As Long, ByVal iCandidate As Long)

    ReDim tNew(1 To m_iTabCount) As TabInfo
    Dim i As Long
    Dim iPos As Long

    If (iCandidate < iDragging) Then
        For i = 1 To iCandidate - 1
            If (i <> iDragging) Then
                iPos = iPos + 1
                LSet tNew(iPos) = m_tTab(i)
            End If
        Next i
        iPos = iPos + 1
        LSet tNew(iPos) = m_tTab(iDragging)
        m_iDraggingTab = iPos
        m_iSelTab = iPos
        For i = iCandidate To m_iTabCount
            If (i <> iDragging) Then
                iPos = iPos + 1
                LSet tNew(iPos) = m_tTab(i)
            End If
        Next i
        'Debug.Print "Replaced:"; iDragging; " with"; iCandidate; " Dragging now at:"; m_iDraggingTab

    Else
        For i = 1 To iCandidate
            If (i <> iDragging) Then
                iPos = iPos + 1
                LSet tNew(iPos) = m_tTab(i)
            End If
        Next i
        iPos = iPos + 1
        LSet tNew(iPos) = m_tTab(iDragging)
        m_iDraggingTab = iPos
        m_iSelTab = iPos
        For i = iCandidate + 1 To m_iTabCount
            If (i <> iDragging) Then
                iPos = iPos + 1
                LSet tNew(iPos) = m_tTab(i)
            End If
        Next i
        'Debug.Print "Replaced:"; iDragging; " with"; iCandidate; " Dragging now at:"; m_iDraggingTab

    End If

    m_bJustReplaced = True
    GetCursorPos m_tJustReplacedPoint

    For i = 1 To m_iTabCount
        LSet m_tTab(i) = tNew(i)
    Next i
    drawTabs

End Function

Public Sub ScrollLeft()
    Dim lDist As Long
    ' determine how far to go:
    lDist = getTypicalScrollDistance()
    m_lOffsetX = m_lOffsetX - lDist
    If (m_lOffsetX < 0) Then
        m_lOffsetX = 0
    End If
    drawTabs
End Sub

Public Sub ScrollRight()
    Dim lDist As Long
    ' determine how far to go:
    lDist = getTypicalScrollDistance()
    m_lOffsetX = m_lOffsetX + lDist
    ' We only go as far so the rightmost tab is visible:
    ensureEndTabOffset

    If (m_lOffsetX < 0) Then
        m_lOffsetX = 0
    End If
    drawTabs

End Sub

Public Property Get SelectedFont() As iFont
    Dim iFnt As iFont
    Dim iFntC As iFont
    Set iFnt = m_fontSelected
    iFnt.Clone iFntC
    Set SelectedFont = iFntC
End Property

Public Property Let SelectedFont(iFnt As iFont)
    pSetSelectedFont iFnt
End Property

Public Property Set SelectedFont(iFnt As iFont)
    pSetSelectedFont iFnt
End Property

Public Property Get SelectedTab() As cTab
    If (m_iSelTab > 0) And (m_iTabCount > 0) Then
        Dim cT As New cTab
        cT.fInit ObjPtr(Me), m_hWnd, m_tTab(m_iSelTab).lId
        Set SelectedTab = cT
    End If
End Property

Private Sub setTriVertexColor(tTV As TRIVERTEX, _
                              ByVal lColor As Long)


    Dim lRed   As Long
    Dim lGreen As Long
    Dim lBlue  As Long

    lRed = (lColor And &HFF&) * &H100&
    lGreen = (lColor And &HFF00&)
    lBlue = (lColor And &HFF0000) \ &H100&
    With tTV
        setTriVertexColorComponent .Red, lRed
        setTriVertexColorComponent .Green, lGreen
        setTriVertexColorComponent .Blue, lBlue
    End With    'tTV

End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, _
                                       ByVal lComponent As Long)

    If (lComponent And &H8000&) = &H8000& Then
        iColor = (lComponent And &H7F00&)
        iColor = iColor Or &H8000
    Else
        iColor = lComponent
    End If

End Sub

Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = m_bShowCloseButton
End Property

Public Property Let ShowCloseButton(ByVal value As Boolean)
    If (m_bShowCloseButton <> value) Then
        m_bShowCloseButton = value
        drawControl
        pPanelSize
        PropertyChanged "ShowCloseButton"
    End If
End Property

Public Property Get Shown() As Boolean
    Shown = m_bOut
End Property

Public Property Let Shown(ByVal bState As Boolean)
    If (m_bPinnable And Not m_bPinned) Then
        If (m_bOut) Then
            If Not (bState) Then
                m_bOut = False
                m_bUnpinPinDown = False
                m_bUnpinCloseTrack = False
                m_bUnpinPinTrack = False
                unshowPinned
                UserControl_Resize
            End If
        Else
            If (bState) Then
                drawUnpinnedTabs
                showUnpinned
            End If
        End If
    End If
End Property

Public Property Get ShowTabs() As Boolean
    ShowTabs = m_bShowTabs
End Property

Public Property Let ShowTabs(ByVal value As Boolean)
    If (m_bShowTabs <> value) Then
        m_bShowTabs = value
        drawControl
        pPanelSize
        PropertyChanged "ShowTabs"
    End If
End Property

Private Sub showUnpinned()
    Dim i As Long
    Dim ctlPanel As Control

    ' Hide anything that's not the current panel:
    For i = 1 To m_iTabCount
        If Not (i = m_iSelTab) Then
            If pbGetTabPanel(i, ctlPanel) Then
                ctlPanel.Visible = False
            End If
        End If
    Next i

    ' show the current panel:
    If (pbGetTabPanel(m_iSelTab, ctlPanel)) Then
        Dim tR As RECT
        GetWindowRect picUnpinned.hWnd, tR
        Dim tP As POINTAPI
        tP.x = tR.Left
        tP.y = tR.Top
        ScreenToClient GetParent(m_hWnd), tP
        tR.Left = tP.x
        tR.Top = tP.y
        tP.x = tR.Right
        tP.y = tR.Bottom
        ScreenToClient GetParent(m_hWnd), tP
        tR.Right = tP.x
        tR.Bottom = tP.y

        If (UserControl.Extender.Align = vbAlignLeft) Then
            m_bOut = True

            ctlPanel.Move _
                    ctlPanel.ScaleX(-m_lSlideOutWidth + (tR.Right - tR.Left), vbPixels, ctlPanel.ScaleMode), _
                    ctlPanel.ScaleY(m_lTitleBarHeight, vbPixels, ctlPanel.ScaleMode), _
                    ctlPanel.ScaleX(m_lSlideOutWidth - m_lSplitSize, vbPixels, ctlPanel.ScaleMode), _
                    ctlPanel.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, ctlPanel.ScaleMode) - ctlPanel.ScaleY(m_lTitleBarHeight + 4, vbPixels, ctlPanel.ScaleMode)
            ctlPanel.Visible = True
            picUnpinned.ZOrder

            For i = 0 To m_lSlideOutWidth Step 8
                SetWindowPos m_hWnd, 0, 0, 0, i + (tR.Right - tR.Left), (tR.Bottom - tR.Top), SWP_NOMOVE    'UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 0
                drawTitleBar
                ctlPanel.Left = ctlPanel.ScaleX((i - m_lSlideOutWidth) + (tR.Right - tR.Left), vbPixels, ctlPanel.ScaleMode)
                ctlPanel.Refresh
            Next i
            ctlPanel.SetFocus
            m_tmr.Enabled = True

        ElseIf (UserControl.Extender.Align = vbAlignRight) Then
            m_bOut = True

            picUnpinned.Visible = False
            ctlPanel.Move _
                    ctlPanel.ScaleX(tR.Right + m_lSlideOutWidth, vbPixels, ctlPanel.ScaleMode), _
                    ctlPanel.ScaleY(m_lTitleBarHeight, vbPixels, ctlPanel.ScaleMode), _
                    ctlPanel.ScaleX(m_lSlideOutWidth - m_lSplitSize, vbPixels, ctlPanel.ScaleMode), _
                    ctlPanel.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, ctlPanel.ScaleMode) - ctlPanel.ScaleY(m_lTitleBarHeight + 4, vbPixels, ctlPanel.ScaleMode)

            ctlPanel.Visible = True

            picUnpinned.Visible = True
            picUnpinned.ZOrder
            For i = 0 To m_lSlideOutWidth Step 8
                SetWindowPos picUnpinned.hWnd, 0, i, 0, 0, 0, SWP_NOSIZE    '
                SetWindowPos m_hWnd, 0, tR.Left - i, tR.Top, (tR.Right - tR.Left) + i, (tR.Bottom - tR.Top), 0    ' UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 0
                drawTitleBar
                ctlPanel.Left = ctlPanel.ScaleX(m_lSplitSize, vbPixels, ctlPanel.ScaleMode)
                ctlPanel.Refresh
            Next i
            ctlPanel.SetFocus
            m_tmr.Enabled = True

        Else

        End If

    End If

End Sub

Public Property Get TabAlign() As EMDITabAlign
    TabAlign = m_eTabAlign
End Property

Public Property Let TabAlign(ByVal value As EMDITabAlign)
    m_eTabAlign = value
    drawControl
    pPanelSize
    PropertyChanged "TabAlign"
End Property

Private Function tabForKey(Key As Variant) As Long
    If IsNumeric(Key) Then
        Dim lCheckIndex As Long
        lCheckIndex = Key
        If (lCheckIndex < 0) Or (lCheckIndex > m_iTabCount) Then
            Err.Raise 9, App.EXEName & ".vbalDTabControlX"
        Else
            tabForKey = lCheckIndex
        End If
    Else
        Dim i As Long
        For i = 1 To m_iTabCount
            If (m_tTab(i).sKey = Key) Then
                tabForKey = i
                Exit Function
            End If
        Next i
        Err.Raise 9, App.EXEName & ".vbalDTabControlX"
    End If
End Function

Public Property Get Tabs() As cTabCollection
    Dim cT As New cTabCollection
    cT.Init ObjPtr(Me), m_hWnd
    Set Tabs = cT
End Property

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Property Get UnpinnedWidth() As Long
    UnpinnedWidth = m_lSlideOutWidth
End Property

Public Property Let UnpinnedWidth(ByVal lWidth As Long)
    m_lSlideOutWidth = lWidth
    If (m_bPinnable And Not m_bPinned) Then
        If (UserControl.Extender.Align = vbAlignLeft Or UserControl.Extender.Align = vbAlignRight) Then
            UserControl.Width = ScaleX(lWidth, vbPixels, UserControl.ScaleMode)
        End If
    End If
    PropertyChanged "UnpinnedWidth"
End Property

Private Sub unshowPinned()

    If (m_bOut) Then


        Dim tP As POINTAPI
        GetCursorPos tP
        Dim tR As RECT
        GetWindowRect m_hWnd, tR
        If (PtInRect(tR, tP.x, tP.y) = 0) Then

            m_tmr.Enabled = False

            Dim i As Long
            Dim ctlPanel As Control

            ' Hide all panels
            For i = 1 To m_iTabCount
                If pbGetTabPanel(i, ctlPanel) Then
                    ctlPanel.Visible = False
                End If
            Next i

            ' No longer out:
            m_bOut = False

            UserControl_Resize
            UserControl.Cls

        End If

    ElseIf Not (m_bPinned) Then
        ' Hide all panels
        For i = 1 To m_iTabCount
            If pbGetTabPanel(i, ctlPanel) Then
                ctlPanel.Visible = False
            End If
        Next i

        UserControl.Cls

    End If

End Sub

Private Sub UserControl_DblClick()
    Dim i As Long
    i = hitTestTab()
    If (i > 0) Then
        Dim c As New cTab
        c.fInit ObjPtr(Me), m_hWnd, m_tTab(i).lId
        RaiseEvent TabDoubleClick(c)
    End If
End Sub

Private Sub UserControl_Initialize()
    '
    Debug.Print "vbalDTabControlX.Initialize"
    '
    m_bShowTabs = True
    m_bShowCloseButton = True
    m_lTabHeight = 24
    m_lButtonSize = 16
    m_lUnpinnedWidth = m_lTabHeight
    m_bAllowScroll = True
    m_eTabAlign = TabAlignBottom
    Set m_font = UserControl.Font
    Set m_fontSelected = UserControl.Font
    m_oBackColor = vbButtonFace
    m_oForeColor = vbWindowText
    m_bPinned = True
    m_bOut = True
    m_lSlideOutWidth = 192
    m_lTitleBarHeight = 22
    m_lSplitSize = 6
    m_bAllowSelectDisabledTabs = False

    Dim lVer As Long
    lVer = GetVersion()
    m_bIsNt = ((lVer And &H80000000) = 0)
    VerInitialise
    GetThemeName hWnd
    GetGradientColors

    '
End Sub

Private Sub UserControl_InitProperties()
    '
    m_hWnd = UserControl.hWnd
    VerInitialise
    m_bDesignMode = Not (UserControl.Ambient.UserMode)
    Set m_cMemDC = New pcMemDC
    loadResources
    m_iDrawStyle = EDS_DefaultNet

    '
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '

    If (m_bPinnable And (m_bOut Or m_bPinned)) Then

        ' Unpinned title bar mouse move processing:
        If (Button = vbLeftButton) Then
            Dim tP As POINTAPI
            GetCursorPos tP
            ScreenToClient m_hWnd, tP
            If Not (PtInRect(m_tUnpinCloseR, tP.x, tP.y)) = 0 Then
                If Not m_bUnpinCloseDown Then
                    m_tmrPinButton.Enabled = False
                    m_bUnpinCloseDown = True
                    m_bUnpinPinTrack = False
                    drawTitleBarButtons
                End If
            ElseIf Not (PtInRect(m_tUnpinPinR, tP.x, tP.y)) = 0 Then
                If Not m_bUnpinPinDown Then
                    m_tmrPinButton.Enabled = False
                    m_bUnpinPinDown = True
                    m_bUnpinCloseTrack = False
                    drawTitleBarButtons
                End If
            Else
                If (m_bUnpinCloseTrack Or m_bUnpinPinTrack) Then
                    m_bUnpinCloseTrack = False
                    m_bUnpinPinTrack = False
                    m_tmrPinButton.Enabled = False
                    drawTitleBarButtons
                End If
            End If
        End If

    End If

    Dim i As Long
    i = hitTestButton()
    If (i > 0) Then
        If (Button = vbLeftButton) Then
            m_iPressButton = i
            m_iTrackButton = m_iPressButton
            drawTabs
            Select Case i
                Case 1
                    ' left scroll:
                    If IsLeftButtonEnabled Then
                        ScrollLeft
                    End If
                Case 2
                    ' right scroll:
                    If IsRightButtonEnabled Then
                        ScrollRight
                    End If
            End Select
        End If
    Else
        i = hitTestTab()
        If (i > 0) Then
            m_iSelTab = i
            m_iDraggingTab = i
            m_bJustReplaced = True
            GetCursorPos m_tJustReplacedPoint
            SetCapture m_hWnd
            drawTabs
        End If
    End If
    '
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    Dim i As Long
    Dim tP As POINTAPI
    GetCursorPos tP

    pSetToolTipText tP

    If (m_bPinnable And (m_bOut Or m_bPinned)) Then

        ' Unpinned title bar mouse move processing:
        ScreenToClient m_hWnd, tP
        If Not (PtInRect(m_tUnpinCloseR, tP.x, tP.y)) = 0 Then
            If Not m_bUnpinCloseTrack Then
                If Not m_bUnpinPinDown Then
                    m_bUnpinCloseTrack = True
                End If
                m_bUnpinPinTrack = False
                drawTitleBarButtons
                m_tmrPinButton.Enabled = True
            End If
        ElseIf Not (PtInRect(m_tUnpinPinR, tP.x, tP.y)) = 0 Then
            If Not m_bUnpinPinTrack Then
                If Not m_bUnpinCloseDown Then
                    m_bUnpinPinTrack = True
                End If
                m_bUnpinCloseTrack = False
                drawTitleBarButtons
                m_tmrPinButton.Enabled = True
            End If
        Else
            If (m_bUnpinCloseTrack Or m_bUnpinPinTrack) Then
                m_bUnpinCloseTrack = False
                m_bUnpinPinTrack = False
                m_tmrPinButton.Enabled = False
                drawTitleBarButtons
            End If
        End If
        '

    End If

    ' Tab mouse move processing:
    If (m_iDraggingTab > 0) And (Button = vbLeftButton) Then
        If (m_bJustReplaced) Then
            If m_iDraggingTab <> hitTestTab() Then
                If Abs(tP.x - m_tJustReplacedPoint.x) > (m_tTab(m_iDraggingTab).tTabR.Right - m_tTab(m_iDraggingTab).tTabR.Left) / 2 Then
                    m_bJustReplaced = False
                Else
                    Exit Sub
                End If
            Else
                m_bJustReplaced = False
            End If
        End If
        ScreenToClient m_hWnd, tP
        tP.x = tP.x + m_lOffsetX
        If (tP.y > m_tTab(1).tTabR.Top - 64) And (tP.y < m_tTab(1).tTabR.Bottom + 64) Then
            ' potential to place:
            Dim replaceCandidate As Long
            If (tP.x < m_tTab(1).tTabR.Left) Then
                ' replace the first one
                replaceCandidate = 1
            ElseIf (tP.x > m_tTab(m_iTabCount).tTabR.Right) Then
                ' replace the last one:
                replaceCandidate = m_iTabCount
            Else
                For i = 1 To m_iTabCount
                    If (tP.x > m_tTab(i).tTabR.Left) And (tP.x < m_tTab(i).tTabR.Right) Then
                        ' replacement a central item:
                        replaceCandidate = i
                        Exit For
                    End If
                Next i
            End If
            If (replaceCandidate > 0) Then
                If (replaceCandidate <> m_iDraggingTab) Then
                    'Debug.Print "Replacement Candidate:", replaceCandidate
                    replaceWithCandidate m_iDraggingTab, replaceCandidate
                End If
            End If
        End If
    Else

        If (m_iTrackButton > 0) Then
            i = hitTestButton()
            If Not (i = m_iTrackButton) Then
                If (i = 0) Then
                    If (m_iPressButton = 0) Then
                        ' end of capture
                        ReleaseCapture
                    End If
                    m_iTrackButton = 0
                    drawTabs
                Else
                    ' change of capture
                    m_iTrackButton = i
                    drawTabs
                End If
            ElseIf (i = m_iPressButton) Then
                Select Case i
                    Case 1
                        ' left scroll:
                        If IsLeftButtonEnabled Then
                            ScrollLeft
                        End If
                    Case 2
                        ' right scroll:
                        If IsRightButtonEnabled Then
                            ScrollRight
                        End If
                End Select
            End If
        Else

            i = hitTestButton()
            If Not (i = m_iTrackButton) Then
                ' change of capture:
                If (m_iTrackButton = 0) Then
                    SetCapture m_hWnd
                End If
                m_iTrackButton = i
                drawTabs
            End If

        End If
    End If
    '
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '

    If (m_bPinnable And (m_bOut Or m_bPinned)) Then

        ' Unpinned title bar mouse move processing:
        Dim tP As POINTAPI
        GetCursorPos tP
        ScreenToClient m_hWnd, tP
        If Not (PtInRect(m_tUnpinCloseR, tP.x, tP.y)) = 0 Then
            If m_bUnpinCloseDown Then
                ' close button pressed:
                m_tmrPinButton.Enabled = False
                m_bUnpinCloseDown = False
                m_bUnpinPinTrack = False

                ' close window:
                Dim bCancel As Boolean
                Dim cT As New cTab
                cT.fInit ObjPtr(Me), m_hWnd, m_tTab(m_iSelTab).lId
                RaiseEvent TabClose(cT, bCancel)
                If Not (bCancel) Then
                    fRemove m_iSelTab
                    If m_iTabCount > 0 Then
                        If Not (m_bPinned) Then
                            showUnpinned
                        End If
                    Else
                        unshowPinned
                        UserControl_Resize
                    End If
                End If

                drawTitleBarButtons

            End If
        ElseIf Not (PtInRect(m_tUnpinPinR, tP.x, tP.y)) = 0 Then

            If m_bUnpinPinDown Then
                ' Pin button pressed:
                m_tmrPinButton.Enabled = False
                m_bUnpinPinDown = False
                m_bUnpinCloseTrack = False
                m_bUnpinPinTrack = False

                If (m_bPinned) Then
                    ' UnPin the tabs:
                    m_bPinned = False
                    m_bOut = False
                    unshowPinned
                    UserControl_Resize

                    RaiseEvent UnPinned
                Else
                    ' Pin the tabs:
                    m_bPinned = True
                    m_bOut = True
                    UserControl.Extender.Width = UserControl.ScaleX(m_lSlideOutWidth, vbPixels, UserControl.ScaleMode)
                    ' Ensure selected tab in view:
                    m_iLastSelTab = 0
                    UserControl_Resize

                    drawTitleBarButtons

                    RaiseEvent Pinned
                End If
            End If

        Else
            If (m_bUnpinCloseTrack Or m_bUnpinPinTrack Or m_bUnpinCloseDown Or m_bUnpinPinDown) Then
                m_bUnpinCloseTrack = False
                m_bUnpinPinTrack = False
                m_bUnpinCloseDown = False
                m_bUnpinPinDown = False
                m_tmrPinButton.Enabled = False
                drawTitleBarButtons
            End If
        End If
    End If

    Dim i As Long
    ReleaseCapture

    If (m_iDraggingTab > 0) Then
        i = hitTestTab()
        If (i > 0) Then
            Dim c As New cTab
            c.fInit ObjPtr(Me), m_hWnd, m_tTab(i).lId
            RaiseEvent TabClick(c, Button, Shift, x, y)
        End If
        m_iDraggingTab = 0

    Else

        If (m_iPressButton > 0) Then
            i = hitTestButton()
            If (i = m_iPressButton) Then
                m_iTrackButton = 0
                m_iPressButton = 0
                ReleaseCapture
                drawTabs
                Select Case i
                    Case 1
                        ' left scroll:
                        If IsLeftButtonEnabled Then
                            ScrollLeft
                        End If
                    Case 2
                        ' right scroll:
                        If IsRightButtonEnabled Then
                            ScrollRight
                        End If
                    Case 3
                        ' close window:
                        cT.fInit ObjPtr(Me), m_hWnd, m_tTab(m_iSelTab).lId
                        RaiseEvent TabClose(cT, bCancel)
                        If Not (bCancel) Then
                            fRemove m_iSelTab
                        End If
                End Select
            Else
                ' not a press:
                m_iTrackButton = 0
                m_iPressButton = 0
                ReleaseCapture
                drawTabs
            End If
        Else
            RaiseEvent TabBarClick(Button, Shift, x, y)
        End If
    End If
    '
End Sub

Private Sub UserControl_Paint()
    '
    GetThemeName hWnd
    drawControl
    '
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    AllowScroll = PropBag.ReadProperty("AllowScroll", True)
    TabAlign = PropBag.ReadProperty("TabAlign", TabAlignBottom)
    Font = PropBag.ReadProperty("Font", UserControl.Font)
    SelectedFont = PropBag.ReadProperty("SelectedFont", UserControl.Font)
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    ShowTabs = PropBag.ReadProperty("ShowTabs", True)
    ShowCloseButton = PropBag.ReadProperty("ShowCloseButton", True)
    Pinnable = PropBag.ReadProperty("Pinnable", False)
    Pinned = PropBag.ReadProperty("Pinned", True)
    UnpinnedWidth = PropBag.ReadProperty("UnpinnedWidth", 192)
    m_bAllowSelectDisabledTabs = PropBag.ReadProperty("AllowSelectDisabledTabs", False)
    m_iDrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    m_hWnd = UserControl.hWnd
    m_bDesignMode = Not (UserControl.Ambient.UserMode)
    Set m_cMemDC = New pcMemDC

    loadResources

    UserControl_Resize
    '
End Sub

Private Sub UserControl_Resize()
    Dim tR As RECT
    Dim hWndMdiClient As Long
    '
    m_bOut = False

    ' set up the memory DC:
    m_cMemDC.Width = UserControl.ScaleX(UserControl.Width, UserControl.ScaleMode, vbPixels) + 8
    m_cMemDC.Height = UserControl.ScaleY(UserControl.Height, UserControl.ScaleMode, vbPixels) + 8

    ' do any sizing necessary
    If (m_bPinnable And Not m_bDesignMode) Then
        If Not (m_bPinned) Then
            If (UserControl.Extender.Align = vbAlignLeft) Then
                picUnpinned.Left = 0
                picUnpinned.Height = UserControl.ScaleHeight
                UserControl.Extender.Width = UserControl.ScaleX(m_lUnpinnedWidth, vbPixels, UserControl.ScaleMode)
                picUnpinned.Width = UserControl.ScaleX(m_lUnpinnedWidth, vbPixels, UserControl.ScaleMode)
                GetWindowRect m_hWnd, tR
                SetWindowPos m_hWnd, 0, 0, 0, m_lUnpinnedWidth, tR.Bottom - tR.Top, SWP_NOMOVE

            ElseIf (UserControl.Extender.Align = vbAlignRight) Then

                GetWindowRect m_hWnd, tR
                Dim tP As POINTAPI
                tP.x = tR.Left
                tP.y = tR.Top
                ScreenToClient GetParent(m_hWnd), tP
                tR.Left = tP.x
                tR.Top = tP.y
                tP.x = tR.Right
                tP.y = tR.Bottom
                ScreenToClient GetParent(m_hWnd), tP
                tR.Right = tP.x
                tR.Bottom = tP.y

                SetWindowPos m_hWnd, 0, tR.Right - m_lUnpinnedWidth, tR.Top, m_lUnpinnedWidth, tR.Bottom - tR.Top, 0
                UserControl.Extender.Width = UserControl.ScaleX(m_lUnpinnedWidth, vbPixels, UserControl.ScaleMode)
                picUnpinned.Left = 0
                picUnpinned.Height = UserControl.ScaleHeight
                picUnpinned.Width = UserControl.ScaleX(m_lUnpinnedWidth, vbPixels, UserControl.ScaleMode)
            Else
                UserControl.Extender.Width = UserControl.ScaleY(m_lUnpinnedWidth, vbPixels, UserControl.ScaleMode)
            End If
            If Not picUnpinned.Visible Then
                picUnpinned.Visible = True
            End If
        Else
            If picUnpinned.Visible Then
                picUnpinned.Visible = False
            End If
        End If
    Else
        If picUnpinned.Visible Then
            picUnpinned.Visible = False
        End If
    End If

    ' Draw the control:
    drawControl
    ' Resize the panels:
    If (m_bPinned Or m_bOut) Then
        pPanelSize
    End If
    ' Now for the user:
    RaiseEvent Resize
    '
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    '
    Debug.Print "vbalDTabControlX.Terminate"
    Set m_cMemDC = Nothing
    If Not (m_hIconPin = 0) Then
        DestroyIcon m_hIconPin
    End If
    If Not (m_hIconUnpin = 0) Then
        DestroyIcon m_hIconUnpin
    End If
    If Not (m_hIconClose = 0) Then
        DestroyIcon m_hIconClose
    End If
    '
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    PropBag.WriteProperty "AllowScroll", m_bAllowScroll, True
    PropBag.WriteProperty "TabAlign", m_eTabAlign, TabAlignBottom
    PropBag.WriteProperty "Font", Font
    PropBag.WriteProperty "SelectedFont", SelectedFont
    PropBag.WriteProperty "BackColor", BackColor, vbButtonFace
    PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
    PropBag.WriteProperty "ShowTabs", m_bShowTabs, True
    PropBag.WriteProperty "ShowCloseButton", m_bShowCloseButton, True
    PropBag.WriteProperty "Pinnable", m_bPinnable, False
    PropBag.WriteProperty "Pinned", m_bPinned, True
    PropBag.WriteProperty "UnpinnedWidth", UnpinnedWidth, 192
    PropBag.WriteProperty "AllowSelectDisabledTabs", m_bAllowSelectDisabledTabs, False
    PropBag.WriteProperty "DrawStyle", m_iDrawStyle, E_DrawStyle.EDS_DefaultNet
    '
End Sub



'//---------------------------------------------------------------------------------------
' Procedure : UtilDrawBackground
' Type      : Sub
' DateTime  : 13/08/2004 16:13
' Author    : Gary Noble
' Purpose   : Gradient Drawing Sub
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  13/08/2004
'//---------------------------------------------------------------------------------------
Private Sub UtilDrawBackground(ByVal lngHdc As Long, _
                              ByVal colorStart As Long, _
                              ByVal colorEnd As Long, _
                              tR As RECT, _
                              Optional ByVal horizontal As Boolean = False)


    GradientFillRect lngHdc, tR, colorStart, colorEnd, IIf(horizontal, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub


'//---------------------------------------------------------------------------------------
'-- Start Of Additions By Gary Noble
'//---------------------------------------------------------------------------------------

'//---------------------------------------------------------------------------------------
' Procedure : VerInitialise
' Type      : Sub
' DateTime  : 13/08/2004 16:13
' Author    : Gary Noble
' Purpose   : Initialise The System Params
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  13/08/2004
'//---------------------------------------------------------------------------------------
Public Sub VerInitialise()

    Dim tOSV As OSVERSIONINFO

    tOSV.dwVersionInfoSize = Len(tOSV)
    GetVersionEx tOSV
    m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    If (tOSV.dwMajorVersion > 5) Then
        m_bHasGradientAndTransparency = True
        m_bIsXp = True
        m_bIs2000OrAbove = True
    ElseIf (tOSV.dwMajorVersion = 5) Then
        m_bHasGradientAndTransparency = True
        m_bIs2000OrAbove = True
        If (tOSV.dwMinorVersion >= 1) Then
            m_bIsXp = True
        End If
    ElseIf (tOSV.dwMajorVersion = 4) Then
        If (tOSV.dwMinorVersion >= 10) Then
            m_bHasGradientAndTransparency = True
        End If
    Else
    End If

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : GetGradientColors
' Type      : Sub
' DateTime  : 13/08/2004 16:14
' Author    : Gary Noble
' Purpose   : Sets The Custom Gradient Colors
'             Note That The Colours Also Get Used By The BlendColor Call
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  13/08/2004
'//---------------------------------------------------------------------------------------
Sub GetGradientColors()

    m_lColorOneSelected = 1
    m_lColorTwoSelected = 1
    m_lColorHeaderColorOne = 1
    m_lColorHeaderColorTwo = 1
    m_lColorHeaderForeColor = 1
    m_lColorHotOne = 1
    m_lColorHotTwo = 1

    If AppThemed Then

        Select Case m_sCurrentSystemThemename
            Case "HomeStead"
                m_lColorOneNormal = RGB(228, 235, 200)
                m_lColorTwoNormal = RGB(175, 194, 142)
                m_lColorBorder = RGB(100, 144, 88)
                m_lColorHeaderColorOne = RGB(165, 182, 121)
                m_lColorHeaderColorTwo = BlendColor(RGB(99, 122, 68), vbBlack, 200)
            Case "NormalColor"
                m_lColorOneNormal = RGB(197, 221, 250)
                m_lColorTwoNormal = RGB(128, 167, 225)
                m_lColorBorder = RGB(0, 45, 150)
                m_lColorHeaderColorOne = RGB(81, 128, 208)
                m_lColorHeaderColorTwo = BlendColor(RGB(11, 63, 153), vbBlack, 230)
            Case "Metallic"
                m_lColorOneNormal = RGB(219, 220, 232)
                m_lColorTwoNormal = RGB(149, 147, 177)
                m_lColorBorder = RGB(119, 118, 151)
                m_lColorHeaderColorOne = RGB(163, 162, 187)
                m_lColorHeaderColorTwo = BlendColor(RGB(112, 111, 145), vbBlack, 200)
            Case Else

                m_lColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
                m_lColorTwoNormal = vbButtonFace
                m_lColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
                m_lColorHeaderColorOne = vbButtonFace
                m_lColorHeaderColorTwo = BlendColor(vbInactiveTitleBar, vbBlack, 200)
                m_lColorBorder = TranslateColor(vbInactiveTitleBar)

        End Select
        m_lColorOneSelectedNormal = RGB(248, 216, 126)
        m_lColorTwoSelectedNormal = RGB(240, 160, 38)

        m_lColorHotOne = BlendColor(vbWindowBackground, vbButtonFace, 220)
        m_lColorHotTwo = RGB(248, 216, 126)

        m_lColorOneSelected = RGB(240, 160, 38)
        m_lColorTwoSelected = RGB(248, 216, 126)

    Else
        m_lColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
        m_lColorTwoNormal = vbButtonFace
        m_lColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
        m_lColorHeaderColorOne = vbButtonFace
        m_lColorHeaderColorTwo = BlendColor(vbInactiveTitleBar, BlendColor(vbBlack, vbButtonFace, 10), 200)
        m_lColorBorder = TranslateColor(vbInactiveTitleBar)
        m_lColorHotTwo = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 50), 10)
        m_lColorHotOne = m_lColorHotTwo
        m_lColorOneSelected = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 100)
        m_lColorTwoSelected = m_lColorOneSelected
        m_lColorOneSelectedNormal = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 130)
        m_lColorTwoSelectedNormal = m_lColorOneSelectedNormal
    End If


End Sub

'//---------------------------------------------------------------------------------------
' Procedure : AppThemed
' Type      : Function
' DateTime  : 13/08/2004 16:15
' Author    : Gary Noble
' Purpose   : Tells Us If The Sysytem Is Using A Theme
' Returns   : Boolean
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  13/08/2004
'//---------------------------------------------------------------------------------------
Public Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

'//---------------------------------------------------------------------------------------
'-- End Of Additions By Gary Noble
'//---------------------------------------------------------------------------------------

