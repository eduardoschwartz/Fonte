VERSION 5.00
Begin VB.UserControl RMComboView 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   KeyPreview      =   -1  'True
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   253
   ToolboxBitmap   =   "RMComboView.ctx":0000
   Begin VB.PictureBox picImages 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3120
      Picture         =   "RMComboView.ctx":00FA
      ScaleHeight     =   495
      ScaleWidth      =   765
      TabIndex        =   3
      Top             =   690
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Timer tmrRelease 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   630
      Top             =   1410
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   150
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   2925
      Begin VB.HScrollBar hscItem 
         Height          =   225
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   405
      End
      Begin VB.VScrollBar vscItem 
         Height          =   675
         Left            =   2640
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Value           =   1
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "RMComboView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#############################################################################################################################
'Title:     ComboView (An Owner-drawn Multi-Column/Checkbox enabled ComboBox replacement)
'Author:    Richard Mewett
'Created:   01/06/05
'Version:   0.2.7 (27th July 2005)

'Copyright: ©2005, Richard Mewett, All Rights Reserved

'IMPORTANT NOTICE:
'ComboView may be used in personal and commercial environments, but may not be sold as a program,
'source, modified source or as a program derived from the source code of this program without prior
'permission.
'
'Credits:   Paul Caton - Subclassing
'           Matt Usner - Rounded Region Code (+ feedback on beta version - thanks!)
'           Heriberto Mantilla Santamaria - Creating Window from PictureBox
'           Anders Lyman - Mouse Wheel Support
'           Carles PV - WM_CTLCOLORSCROLLBAR Tip!
'           Fred.cpp - Solid Part of DrawRect
'           Dana Seaman - DrawText wrapper & Unicode Res file for Demo

'        +  Phantom Man - For his concise bug reports

'Updates (dd/mm/yy):
'04/07/05   Added ColumnResize property for end-user Column resizing
'           Subclassed WM_CTLCOLORSCROLLBAR for scrollbars
'05/07/05   Added CustomSort Event, FocusRectColor + FocusRectStyle Properties
'07/07/05   Removed ColumnHeadingHeight Property (always calculated from DropDownFont size) seems -
'           unneccessary to be able to change it!
'           Tweaked Sort Arrow position
'           ListIndex adjusted when Text Property changed and in Standard Style
'08/07/05   Custom Border Style (+ Border Colours, Button Colors)
'15/07/05   Added DrawText wrapper function (to support DrawTextW on NT based systems)
'22/07/05   vscItem.LargeChange =  PageScrollItems property (thanks Jeff Mayes)
'27/07/05   Enabled State tweaks
'#############################################################################################################################

'Windows API Declarations
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
 
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
 
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_BTNFACE = 15

Private Const CLR_INVALID = &HFFFF

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_SINGLELINE = &H20

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_FLAT = &H4000
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const SWP_FRAMECHANGED          As Long = &H20
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOSIZE                As Long = &H1

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_TOOLWINDOW = &H80
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2

Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const MERGEPAINT = &HBB0226

Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X           As Long
    Y           As Long
End Type
'#############################################################################################################################
'Subclassing Code (all credits to Paul Caton!)
Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private mInCtrl                      As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                                   'Subclass data type
    hWnd                               As Long                                            'Handle of the window being subclassed
    nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                            'Msg after table entry count
    nMsgCntB                           As Long                                            'Msg before table entry count
    aMsgTblA()                         As Long                                            'Msg after table array
    aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WM_SETFOCUS            As Long = &H7
Private Const WM_KILLFOCUS           As Long = &H8
Private Const WM_MOUSELEAVE          As Long = &H2A3
Private Const WM_MOUSEMOVE           As Long = &H200
Private Const WM_MOUSEHOVER          As Long = &H2A1
Private Const WM_MOUSEWHEEL          As Long = &H20A
Private Const WM_LBUTTONDOWN         As Long = &H201
Private Const WM_RBUTTONDOWN         As Long = &H204
Private Const WM_GETMINMAXINFO       As Long = &H24
Private Const WM_SIZE                As Long = &H5
Private Const WM_WINDOWPOSCHANGED    As Long = &H47
Private Const WM_WINDOWPOSCHANGING   As Long = &H46
Private Const WM_CTLCOLORSCROLLBAR   As Long = &H137

'#############################################################################################################################
'User Control Declarations
Private Const BORDER_LEFT = 4
Private Const BORDER_TOP = 3
Private Const BUTTON_WIDTH = 15

Private Const DEF_ALIGNMENT = vbLeftJustify
Private Const DEF_AUTOCOMPLETE = False
Private Const DEF_BACKCOLOR = vbWindowBackground
Private Const DEF_BORDERCOLOR = vbBlack
Private Const DEF_BORDERCURVE = 5
Private Const DEF_BORDERSTYLE = 1
Private Const DEF_BORDERWIDTH = 1
Private Const DEF_BUTTONBACKCOLOR = vbButtonFace
Private Const DEF_COLS = 1
Private Const DEF_COLUMNHEADERS = False
Private Const DEF_COLUMNRESIZE = False
Private Const DEF_COLUMNSORT = False
Private Const DEF_DROPDOWNAUTOWIDTH = False
Private Const DEF_DROPDOWNITEMSVISIBLE = 8
Private Const DEF_DROPDOWNWIDTH = 0
Private Const DEF_DEFAULTITEMFORECOLOR = vbWindowText
Private Const DEF_EDITABLE = False
Private Const DEF_ENABLED = True
Private Const DEF_FOCUSRECTCOLOR = &HFFFF&
Private Const DEF_FOCUSRECTSTYLE = 1
Private Const DEF_FORECOLOR = vbWindowText
Private Const DEF_INTEGRALHEIGHT = False
Private Const DEF_LOCKED = False
Private Const DEF_PAGESCROLLITEMS = 8
Private Const DEF_REQUIRECHECKEDITEM = False
Private Const DEF_ROWHEIGHTMIN = 0
Private Const DEF_SCALEUNITS = vbTwips
Private Const DEF_SEARCHCOLUMN = 0
Private Const DEF_STYLE = 0
Private Const DEF_TEXTALL = "-- All --"
Private Const DEF_TEXTNONE = "-- None --"
Private Const DEF_TEXTSELECTION = "-- Selection --"

Private Const CACHE_INCREMENT = 10
Private Const EVENT_TIMEOUT = 500
Private Const AUTOSCROLL_TIMEOUT = 50
Private Const NULL_RESULT = -1

Private Enum FlagsEnum
    flgChecked = 2
    flgSelected = 4
    flgBold = 8
End Enum

Private Enum SearchEnum
    cvEqual = 0
    cvGreaterEqual = 1
    cvLike = 2
End Enum

Public Enum ColAlignmentEnum
    AlignLeftTop = DT_LEFT Or DT_TOP
    AlignLeftCenter = DT_LEFT Or DT_VCENTER
    AlignLeftBottom = DT_LEFT Or DT_BOTTOM
    AlignCenterTop = DT_CENTER Or DT_TOP
    AlignCenterCenter = DT_CENTER Or DT_VCENTER
    AlignCenterBottom = DT_CENTER Or DT_BOTTOM
    AlignRightTop = DT_RIGHT Or DT_TOP
    AlignRightCenter = DT_RIGHT Or DT_VCENTER
    AlignRightBottom = DT_RIGHT Or DT_BOTTOM
End Enum

Public Enum BorderStyleEnum
    BorderNone = 0
    BorderSunken = 1
    BorderRaised = 2
    BorderFlat = 3
    BorderCustom = 4
End Enum

Public Enum ColTypeEnum
    TypeString = 0
    TypeNumeric = 1
    TypeDate = 2
    TypeCustom = 3
End Enum

Public Enum FocusRectStyleEnum
    FocusRectNone = 0
    FocusRectLight = 1
    FocusRectHeavy = 2
End Enum

Public Enum SortOrderEnum
    Ascending = 1
    Descending = 0
End Enum

Public Enum StyleEnum
    Standard = 0
    Checkboxes = 1
    OptionButtons = 2
End Enum

#If False Then
    Private flgChecked, flgSelected, flgBold, ctString, ctNumeric, ctNumeric
#End If

Private Type ColType
    nAlignment As ColAlignmentEnum
    dCustomWidth As Single
    lWidth As Long
    nSortOrder As Integer
    nType As Integer
    bVisible As Boolean
    sHeading As String
End Type

Private Type ItemType
    vImage As Variant
    lForeColor As Long
    lItemData As Long
    nFlags As Byte
    sValue() As String
End Type

'Data
Private mCols() As ColType
Private mItems() As ItemType
Private mPositions() As Integer
Private mItemCount As Integer
Private mListIndex As Integer

'Misc
Private mImageList As Object
Private mHotImageList As Object

Private mInFocus As Boolean
Private mMouseDown As Boolean
Private mButtonIndex As Integer
Private mResizeCol As Integer
Private mButtonRect As RECT
Private mButtonClickTick As Long
Private mScrollTick As Long
Private mIgnoreKeyPress As Boolean
Private mLockTextBoxEvent As Boolean
Private mWindowsNT As Boolean
Private mSelectedText As String

'Properties
Private mAlignment As AlignmentConstants
Private mAutoComplete As Boolean
Private mBackColor As OLE_COLOR
Private mBorderColor As OLE_COLOR
Private mBorderCurve As Long
Private mBorderStyle As BorderStyleEnum
Private mBorderWidth As Long
Private mButtonBackColor As OLE_COLOR
Private mColumnHeaders As Boolean
Private mColumnResize As Boolean
Private mColumnSort As Boolean
Private mDefaultItemForeColor As OLE_COLOR
Private mDisplayEllipsis  As Boolean
Private mDropDownAutoWidth As Boolean
Private mDropDownFont As Font
Private mDropDownItemsVisible As Integer
Private mDropDownWidth As Single
Private mEditable As Boolean
Private mEnabled As Boolean
Private mFocusRectColor As OLE_COLOR
Private mFocusRectStyle As FocusRectStyleEnum
Private mFont As Font
Private mFormatString As String
Private mForeColor As OLE_COLOR
Private mHighlighted As Integer
Private mHotBorderColor As OLE_COLOR
Private mHotButtonBackColor As OLE_COLOR
Private mIntegralHeight As Boolean
Private mLocked As Boolean
Private mMaxLength As Integer
Private mPageScrollItems As Integer
Private mRequireCheckedItem As Boolean
Private mRowHeightMin As Single
Private mScaleUnits As ScaleModeConstants
Private mSortColumn As Integer
Private mSortSubColumn As Integer
Private mSearchColumn As Integer
Private mStyle As StyleEnum
Private mTextAll As String
Private mTextNone As String
Private mTextSelection As String

'Events
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Scroll()

Public Event ClickItem(nListIndex As Integer, Button As Integer, Shift As Integer)
Public Event CustomSort(bAscending As Boolean, nCol As Integer, sValue1 As String, sValue2 As String, bSwap As Boolean)
Public Event DropDownClose()
Public Event DropDownOpen()
Public Event RequestItemChecked(nListIndex As Integer, bValue As Boolean, bCancel As Boolean)
Public Event RequestListChecked(bValue As Boolean, bCancel As Boolean)
Public Event SelectionChanged()
Public Event SortComplete()

Private Sub DrawText(ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)
    If mWindowsNT Then
        DrawTextW hdc, StrPtr(lpString), nCount, lpRect, wFormat
    Else
        DrawTextA hdc, lpString, nCount, lpRect, wFormat
    End If
End Sub

'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'THIS MUST BE THE FIRST PUBLIC ROUTINE IN THIS FILE.
'That includes public properties also
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data

    Select Case uMsg
    Case WM_KILLFOCUS
        'Another Control has got the focus
        DoKillFocus
        
    Case WM_MOUSEMOVE
        SetTimer False
        
        If mEnabled Then
            If Not mInCtrl Then
                mInCtrl = True
                DrawComboBorder
                
                Call TrackMouseLeave(lng_hWnd)
                Call TrackMouseHover(lng_hWnd, 0)
            End If
        End If
        
        If IsMouseInScrollArea() Then
            DoAutoScroll
        End If
    
    Case WM_MOUSELEAVE
        mInCtrl = False
        
        If mBorderStyle = BorderCustom Then
            DrawComboBorder
        End If
            
        If picList.Visible Then
            Call GetAsyncKeyState(VK_LBUTTON)
            Call GetAsyncKeyState(VK_RBUTTON)
    
            SetTimer True
        End If
    
    Case WM_MOUSEHOVER
        If mEnabled Then
            mInCtrl = False
        End If
    
    Case WM_MOUSEWHEEL
        If mInFocus Then
            Select Case wParam
            Case Is > False
                If vscItem.Value > vscItem.Min Then
                    vscItem.Value = vscItem.Value - 1
                End If
            
            Case Else
                If vscItem.Value < vscItem.Max Then
                    vscItem.Value = vscItem.Value + 1
                End If
                
            End Select
        End If
          
    Case WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED, WM_GETMINMAXINFO, WM_SIZE, WM_LBUTTONDOWN, WM_RBUTTONDOWN
        'If Parent form is changing we want to close!
        DoKillFocus
        
    Case WM_CTLCOLORSCROLLBAR
        bHandled = True
        
    End Select
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    MsgBox "ComboView Control 0.2.7 ©2005, Richard Mewett", vbInformation
End Sub

Public Sub AddItem(ByVal Item As String, Optional Index As Integer = -1, Optional Checked As Boolean, Optional Image As Variant)
    Dim nCol As Integer
    Dim nCount As Integer
    Dim sText() As String
    
    '#############################################################################################################################
    'mItems() is an array of the Items in the ComboBox
    'mPositions() is an array of "pointers" to mItems()
    
    'The pointer technique is used to allow much faster Inserts & Sorts
    'since we only need to swap an Integer (2 bytes) rather than a large
    'data structure (a UDT in this case)
    
    'The mItems() is resized incrementally to reduce the Redim Preserve
    'overhead. Since we will only ever be too large by CACHE_INCREMENT (10)
    'the potential unused allocated memory is minimal
    '#############################################################################################################################
    
    mItemCount = mItemCount + 1
    If mItemCount > UBound(mItems) Then
        ReDim Preserve mItems(mItemCount + CACHE_INCREMENT)
        ReDim Preserve mPositions(mItemCount + CACHE_INCREMENT)
    End If
    
    If (Index >= 0) And (Index < mItemCount) Then
        If mItemCount > 1 Then
            For nCount = mItemCount To Index + 1 Step -1
                mPositions(nCount) = mPositions(nCount - 1)
            Next nCount
            mPositions(Index) = mItemCount
        End If
    Else
        mPositions(mItemCount) = mItemCount
    End If
    
    ReDim mItems(mItemCount).sValue(UBound(mCols))
    
    If UBound(mCols) > 0 Then
        sText() = Split(Item, vbTab)
        For nCount = LBound(sText) To UBound(sText)
            mItems(mItemCount).sValue(nCol) = sText(nCount)
            nCol = nCol + 1
            If nCol > UBound(mCols) Then
                Exit For
            End If
        Next nCount
    Else
        mItems(mItemCount).sValue(0) = Item
    End If
    
    With mItems(mItemCount)
        .lForeColor = mDefaultItemForeColor
        .vImage = Image
    
        If Checked Then
            SetFlag mItemCount, flgChecked, True
        End If
        
        'Default Bold
        If mDropDownFont.Bold Then
            SetFlag mItemCount, flgBold, True
        End If
    End With
End Sub

Public Property Get Alignment() As AlignmentConstants
    Alignment = mAlignment
End Property

Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
    mAlignment = NewValue
    txtCombo.Alignment = mAlignment
    
    PropertyChanged "Alignment"
End Property

Public Property Get AutoComplete() As Boolean
    AutoComplete = mAutoComplete
End Property

Public Property Let AutoComplete(ByVal NewValue As Boolean)
    mAutoComplete = NewValue
    
    PropertyChanged "AutoComplete"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    mBackColor = NewValue
    
    With UserControl
        .BackColor = mBackColor
        .Picture = .Image
    End With
    
    txtCombo.BackColor = mBackColor
    picList.BackColor = mBackColor
    ShowText mInFocus
    
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As BorderStyleEnum
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As BorderStyleEnum)
    mBorderStyle = NewValue
    DrawComboBorder
    ShowText mInFocus
    
    PropertyChanged "BorderStyle"
End Property

Public Sub Clear()
    ReDim mItems(0)
    ReDim mPositions(0)
    
    mItemCount = -1
    mListIndex = -1
    
    mButtonIndex = NULL_RESULT
    mSortColumn = NULL_RESULT
    mSortSubColumn = NULL_RESULT
    
    mResizeCol = NULL_RESULT
End Sub

Public Property Get ColAlignment(ByVal Index As Integer) As ColAlignmentEnum
    ColAlignment = mCols(Index).nAlignment
End Property

Public Property Let ColAlignment(ByVal Index As Integer, ByVal NewValue As ColAlignmentEnum)
    mCols(Index).nAlignment = NewValue
End Property

Public Property Get ColHeading(ByVal Index As Integer) As String
    ColHeading = mCols(Index).sHeading
End Property

Public Property Let ColHeading(ByVal Index As Integer, ByVal NewValue As String)
    mCols(Index).sHeading = NewValue
End Property

Public Property Get Cols() As Integer
    Cols = UBound(mCols) + 1
End Property

Public Property Let Cols(ByVal NewValue As Integer)
    Dim nCol As Integer
    
    If NewValue > 0 Then
        ReDim mCols(0 To NewValue - 1)
        For nCol = LBound(mCols) To UBound(mCols)
            mCols(nCol).dCustomWidth = 1000
            mCols(nCol).lWidth = ScaleX(mCols(nCol).dCustomWidth, mScaleUnits, vbPixels)
            mCols(nCol).bVisible = True
        Next nCol
    Else
        ReDim mCols(0)
    End If
End Property

Public Property Get ColType(ByVal Index As Integer) As ColTypeEnum
    ColType = mCols(Index).nType
End Property

Public Property Let ColType(ByVal Index As Integer, ByVal NewValue As ColTypeEnum)
    mCols(Index).nType = NewValue
End Property

Public Property Get ColumnHeaders() As Boolean
    ColumnHeaders = mColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal NewValue As Boolean)
    mColumnHeaders = NewValue
    
    PropertyChanged "ColumnHeaders"
End Property

Public Property Get ColumnResize() As Boolean
    ColumnResize = mColumnResize
End Property

Public Property Let ColumnResize(ByVal NewValue As Boolean)
    mColumnResize = NewValue
    
    PropertyChanged "ColumnResize"
End Property

Public Property Get ColumnSort() As Boolean
    ColumnSort = mColumnSort
End Property

Public Property Let ColumnSort(ByVal NewValue As Boolean)
    mColumnSort = NewValue
    
    PropertyChanged "ColumnSort"
End Property

Public Property Get ColVisible(ByVal Index As Integer) As Boolean
    ColVisible = mCols(Index).bVisible
End Property

Public Property Let ColVisible(ByVal Index As Integer, ByVal NewValue As Boolean)
    mCols(Index).bVisible = NewValue
End Property

Public Property Get ColWidth(ByVal Index As Integer) As Single
    ColWidth = mCols(Index).dCustomWidth
End Property

Public Property Let ColWidth(ByVal Index As Integer, ByVal NewValue As Single)
    'dCustomWidth is in the Units the Control is operating in
    mCols(Index).dCustomWidth = NewValue
    
    'lWidth is always Pixels (because thats what API functions require) and
    'is calculated to prevent repeated Width Scaling calculations
    mCols(Index).lWidth = ScaleX(mCols(Index).dCustomWidth, mScaleUnits, vbPixels)
End Property

Public Property Get DefaultItemForeColor() As OLE_COLOR
    DefaultItemForeColor = mDefaultItemForeColor
End Property

Public Property Let DefaultItemForeColor(ByVal NewValue As OLE_COLOR)
    mDefaultItemForeColor = NewValue
    
    PropertyChanged "DefaultItemForeColor"
End Property

Public Property Get DisplayEllipsis() As Boolean
    DisplayEllipsis = mDisplayEllipsis
End Property

Public Property Let DisplayEllipsis(ByVal NewValue As Boolean)
    mDisplayEllipsis = NewValue
    
    PropertyChanged "DisplayEllipsis"
End Property

Private Sub DoAutoScroll()
    Const MAX_COUNT As Long = 2147483647
    
    Static bActive As Boolean
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    Dim lCount As Long
    
    'This scrolls the list up/down when the mouse moves outside the DropDown
    'and the left button is pressed. It will terminate as soon as the mouse
    'moves back into the DropDown or the control loses focus
    
    'Prevent recursion
    If Not bActive Then
        bActive = True
        'Debug.Print "DoAutoScroll >"
        
        Call GetWindowRect(picList.hWnd, uRect)
        
        Do While mInFocus
            If (GetTickCount() - mScrollTick) > AUTOSCROLL_TIMEOUT Then
                mScrollTick = GetTickCount()
                
                Call GetCursorPos(uPoint)
                
                If (uPoint.Y < uRect.Top) Then
                    If vscItem.Value > vscItem.Min Then
                        mHighlighted = mHighlighted - 1
                        vscItem.Value = vscItem.Value - 1
                    End If
                ElseIf (uPoint.Y > uRect.Bottom) Then
                    If vscItem.Value < vscItem.Max Then
                        mHighlighted = mHighlighted + 1
                        vscItem.Value = vscItem.Value + 1
                    End If
                Else
                    Exit Do
                End If
            End If
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
            ElseIf lCount = MAX_COUNT Then
                lCount = 0
            End If
        Loop
        
        bActive = False
        'Debug.Print "DoAutoScroll <"
    End If
End Sub

Private Sub DoKillFocus()
    If picList.Visible Then
        SetDropDown
    End If

    If mInFocus Then
        mInFocus = False
        ShowText False
    End If
End Sub

Private Sub DoSort()
    If (mSortColumn = NULL_RESULT) And (mSortSubColumn <> NULL_RESULT) Then
        mSortColumn = mSortSubColumn
        mSortSubColumn = NULL_RESULT
    ElseIf mSortColumn = mSortSubColumn Then
        mSortSubColumn = NULL_RESULT
    End If
    
    SortArray LBound(mItems), mItemCount, mSortColumn, mCols(mSortColumn).nSortOrder
    SortSubList
    
    RaiseEvent SortComplete
End Sub

Private Sub DrawComboBorder()
    Const ARROW_HEIGHT = 3
    Const ARROW_WIDTH = 5

    Static bResetRegion As Boolean
    
    Dim R As RECT
    Dim lColor As Long
    Dim hBrush As Long
    Dim hRgn1  As Long
    Dim hRgn2  As Long
    Dim lX As Long
    Dim lY As Long
    
    '#############################################################################################################################
    'This draws the Border of the ComboBox and the Dropdown Button
    '#############################################################################################################################
    
    On Local Error GoTo DrawComboBorderError
    
    With mButtonRect
        .Left = txtCombo.Width + BORDER_LEFT + 1
        .Top = BORDER_TOP - 1
        .Right = .Left + BUTTON_WIDTH
        .Bottom = .Top + UserControl.ScaleHeight - BORDER_TOP - 1
    End With
    
    With UserControl
        Call SetRect(R, 0, 0, .ScaleWidth, .ScaleHeight)
        DrawRect .hdc, mButtonRect, TranslateColor(mBackColor), True
        
        If mBorderStyle = BorderCustom Then
            Call SetRect(R, txtCombo.Width + BORDER_LEFT + 1, 0, .ScaleWidth, .ScaleHeight)
            If mInCtrl Then
                DrawRect .hdc, R, TranslateColor(mHotButtonBackColor), True
            Else
                DrawRect .hdc, R, TranslateColor(mButtonBackColor), True
            End If
        Else
            DrawRect .hdc, mButtonRect, TranslateColor(mButtonBackColor), True
        
            If bResetRegion Then
                hRgn1 = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
                SetWindowRgn hWnd, hRgn1, True
                
                SetWindowRgn picList.hWnd, hRgn1, True
                DeleteObject hRgn1
                
                bResetRegion = False
            End If
            
            Call DrawEdge(.hdc, mButtonRect, EDGE_RAISED, BF_RECT)
        End If

        Select Case mBorderStyle
        Case BorderSunken
            Call DrawEdge(.hdc, R, EDGE_SUNKEN, BF_RECT)
        
        Case BorderRaised
            Call DrawEdge(.hdc, R, EDGE_RAISED, BF_RECT)
        
        Case BorderFlat
            Call DrawEdge(.hdc, R, EDGE_SUNKEN, BF_RECT Or BF_FLAT)
        
        Case BorderCustom
            If mInCtrl Then
                lColor = TranslateColor(mHotBorderColor)
            Else
                lColor = TranslateColor(mBorderColor)
            End If
            
            hRgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, mBorderCurve, mBorderCurve)
            hRgn2 = CreateRoundRectRgn(mBorderWidth, mBorderWidth, ScaleWidth - mBorderWidth, ScaleHeight - mBorderWidth, mBorderCurve, mBorderCurve)
            CombineRgn hRgn2, hRgn1, hRgn2, 3
            
            hBrush = CreateSolidBrush(lColor)
            FillRgn hdc, hRgn2, hBrush
            
            SetWindowRgn hWnd, hRgn1, True
            SetWindowRgn picList.hWnd, hRgn1, True
            
            DeleteObject hRgn2
            DeleteObject hBrush
            DeleteObject hRgn1
            
            bResetRegion = True
        
        Case Else
            .Picture = Nothing
        
        End Select
         
        lX = mButtonRect.Left + (BUTTON_WIDTH / 2) - (ARROW_WIDTH / 2)
        lY = (.ScaleHeight / 2) - (ARROW_HEIGHT / 2)
    
        If mEnabled Then
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 0, MERGEPAINT)
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 0, SRCAND)
        Else
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 4, MERGEPAINT)
            Call BitBlt(.hdc, lX, lY, ARROW_WIDTH, ARROW_HEIGHT, picImages.hdc, 42, 4, SRCAND)
        End If
        
        .Picture = .Image
    End With
    Exit Sub
    
DrawComboBorderError:
    Exit Sub
End Sub

Private Sub DrawRect(hdc As Long, rc As RECT, lColor As Long, bFilled As Boolean)
    Dim lNewBrush As Long
  
    lNewBrush = CreateSolidBrush(lColor)
    
    If bFilled Then
        Call FillRect(hdc, rc, lNewBrush)
    Else
        Call FrameRect(hdc, rc, lNewBrush)
    End If

    Call DeleteObject(lNewBrush)
End Sub

Public Property Get DropDownAutoWidth() As Boolean
    DropDownAutoWidth = mDropDownAutoWidth
End Property

Public Property Let DropDownAutoWidth(ByVal NewValue As Boolean)
    mDropDownAutoWidth = NewValue
End Property

Public Property Get DropDownFont() As Font
   Set DropDownFont = mDropDownFont
End Property

Public Property Set DropDownFont(ByVal NewValue As StdFont)
    Set mDropDownFont = NewValue
    
    PropertyChanged "DropDownFont"
End Property

Public Property Get DropDownItemsVisible() As Integer
    DropDownItemsVisible = mDropDownItemsVisible
End Property

Public Property Let DropDownItemsVisible(ByVal NewValue As Integer)
    mDropDownItemsVisible = NewValue
    
    PropertyChanged "DropDownItemsVisible"
End Property

Public Property Get DropDownWidth() As Single
    DropDownWidth = mDropDownWidth
End Property

Public Property Let DropDownWidth(ByVal NewValue As Single)
    mDropDownWidth = NewValue
    
    PropertyChanged "DropDownWidth"
End Property

Public Property Get Editable() As Boolean
    Editable = mEditable
End Property

Public Property Let Editable(ByVal NewValue As Boolean)
    mEditable = NewValue
    txtCombo.Visible = mEditable
    
    ShowText mInFocus
    
    PropertyChanged "Editable"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    mEnabled = NewValue
    txtCombo.Enabled = mEnabled
    
    DrawComboBorder
    ShowText mInFocus
End Property

Public Property Get FocusRectColor() As OLE_COLOR
    FocusRectColor = mFocusRectColor
End Property

Public Property Let FocusRectColor(ByVal NewValue As OLE_COLOR)
    mFocusRectColor = NewValue
    
    PropertyChanged "FocusRectColor"
End Property

Public Property Get FocusRectStyle() As FocusRectStyleEnum
    FocusRectStyle = mFocusRectStyle
End Property

Public Property Let FocusRectStyle(ByVal NewValue As FocusRectStyleEnum)
    mFocusRectStyle = NewValue
    
    PropertyChanged "FocusRectStyle"
End Property

Public Property Get Font() As Font
   Set Font = mFont
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Set mFont = NewValue
    
    Set UserControl.Font = mFont
    Set txtCombo.Font = mFont
   
    If mIntegralHeight Then
        UserControl.Height = ScaleY(UserControl.TextHeight("A") + (BORDER_TOP * 2), vbPixels, vbTwips)
        UserControl_Resize
    End If
    ShowText mInFocus
   
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    mForeColor = NewValue
    
    With UserControl
        .ForeColor = mForeColor
    End With
    
    txtCombo.ForeColor = mForeColor
    ShowText mInFocus
    
    PropertyChanged "ForeColor"
End Property

Public Property Get FormatString() As String
    FormatString = mFormatString
End Property

Public Property Let FormatString(ByVal NewValue As String)
    Dim lCol As Long
    Dim sCols() As String
    
    mFormatString = NewValue
    
    sCols() = Split(NewValue, "|")
    If UBound(sCols()) > UBound(mCols) Then
        Cols = UBound(sCols()) + 1
    End If
    
    For lCol = LBound(sCols) To UBound(sCols)
        Select Case Mid$(sCols(lCol), 1, 1)
        Case "^"
            mCols(lCol).sHeading = Mid$(sCols(lCol), 2)
            mCols(lCol).nAlignment = AlignCenterTop
        Case "<"
            mCols(lCol).sHeading = Mid$(sCols(lCol), 2)
            mCols(lCol).nAlignment = AlignLeftTop
        Case ">"
            mCols(lCol).sHeading = Mid$(sCols(lCol), 2)
            mCols(lCol).nAlignment = AlignRightTop
        Case Else
            mCols(lCol).sHeading = sCols(lCol)
        End Select
        
        mCols(lCol).dCustomWidth = 1000
        mCols(lCol).lWidth = ScaleX(mCols(lCol).dCustomWidth, mScaleUnits, vbPixels)
        mCols(lCol).bVisible = True
    Next lCol
    
    PropertyChanged "FormatString"
End Property

Private Function GetColFromX(X As Single, Optional lColPosX As Long) As Integer
    Dim lX As Long
    Dim nCol As Integer
    
    GetColFromX = -1
    
    For nCol = hscItem.Value To UBound(mCols)
        If (X > lX) And (X < lX + mCols(nCol).lWidth) Then
            lColPosX = lX
            GetColFromX = nCol
        End If
        
        lX = lX + mCols(nCol).lWidth
    Next nCol
End Function

Private Function GetColumnHeadingHeight() As Long
    With picList
        GetColumnHeadingHeight = .TextHeight("A") + .ScaleY(4, vbPixels, .ScaleMode)
    End With
End Function

Private Function GetFlag(ByVal nIndex As Integer, nFlag As FlagsEnum) As Boolean
    'Gets information by bit flags for a ListItem.
    
    If mItems(nIndex).nFlags And nFlag Then
        GetFlag = True
    End If
End Function

Private Function GetRowFromY(Y As Single) As Integer
    Dim lColumnHeadingHeight As Long
    Dim nRow As Integer
    
    With picList
        If mColumnHeaders Then
            lColumnHeadingHeight = GetColumnHeadingHeight()
            
            If Y > lColumnHeadingHeight Then
                nRow = ((Y - lColumnHeadingHeight) \ GetRowHeight()) + vscItem.Value
            Else
                nRow = -1
            End If
        Else
            nRow = (Y \ GetRowHeight()) + vscItem.Value
        End If
    End With
    
    If nRow <= mItemCount Then
        GetRowFromY = nRow
    Else
        GetRowFromY = -1
    End If
End Function

Private Function GetRowHeight() As Long
    If mRowHeightMin > 0 Then
        GetRowHeight = picList.ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
    Else
        GetRowHeight = picList.TextHeight("A")
    End If
End Function

Public Property Get HotImageList() As Object
    Set HotImageList = mHotImageList
End Property

Public Property Let HotImageList(ByVal NewValue As Object)
    Set mHotImageList = NewValue
End Property

Private Sub hscItem_Change()
    ShowItems
End Sub

Private Sub hscItem_Scroll()
    hscItem_Change
    picList.Refresh
End Sub

Public Property Let ImageList(ByVal NewValue As Object)
    Set mImageList = NewValue
End Property

Public Property Get IntegralHeight() As Boolean
    IntegralHeight = mIntegralHeight
End Property

Public Property Let IntegralHeight(ByVal NewValue As Boolean)
    mIntegralHeight = NewValue
    
    If mIntegralHeight Then
        UserControl.Height = ScaleY(UserControl.TextHeight("A") + (BORDER_TOP * 2), vbPixels, vbTwips)
        UserControl_Resize
        ShowText mInFocus
    End If
    
    PropertyChanged "IntegralHeight"
End Property

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    Call FreeLibrary(hMod)
  End If
End Function
'END Subclassing Code===================================================================================

Private Function IsMouseInScrollArea() As Boolean
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    
    Call GetWindowRect(picList.hWnd, uRect)
    Call GetCursorPos(uPoint)
    
    If (uPoint.Y < uRect.Top) Or (uPoint.Y > uRect.Bottom) Then
        IsMouseInScrollArea = True
    End If
End Function

Public Property Get ItemChecked(ByVal Index As Integer) As Boolean
    ItemChecked = GetFlag(mPositions(Index), flgChecked)
End Property

Public Property Let ItemChecked(ByVal Index As Integer, ByVal NewValue As Boolean)
    SetFlag mPositions(Index), flgChecked, NewValue
End Property

Public Property Let ItemData(ByVal Index As Integer, NewValue As Long)
    mItems(mPositions(Index)).lItemData = NewValue
End Property

Public Property Get ItemData(ByVal Index As Integer) As Long
    ItemData = mItems(mPositions(Index)).lItemData
End Property

Public Property Get ItemFontBold(ByVal Index As Integer) As Boolean
    ItemFontBold = mItems(mPositions(Index)).nFlags And flgBold
End Property

Public Property Let ItemFontBold(ByVal Index As Integer, ByVal NewValue As Boolean)
    SetFlag Index, flgBold, NewValue
End Property

Public Property Get ItemForeColor(ByVal Index As Integer) As Long
    ItemForeColor = mItems(mPositions(Index)).lForeColor
End Property

Public Property Let ItemForeColor(ByVal Index As Integer, ByVal NewValue As Long)
    mItems(mPositions(Index)).lForeColor = NewValue
End Property

Public Property Let ItemImage(ByVal Index As Integer, NewValue As Variant)
    mItems(mPositions(Index)).vImage = NewValue
End Property

Public Property Get ItemImage(ByVal Index As Integer) As Variant
    ItemImage = mItems(mPositions(Index)).vImage
End Property

Public Property Get ItemText(ByVal Index As Integer, ByVal Item As Integer) As String
    If UBound(mItems(mPositions(Index)).sValue) >= Item Then
        ItemText = mItems(mPositions(Index)).sValue(Item)
    End If
End Property

Public Property Let ItemText(ByVal Index As Integer, ByVal Item As Integer, NewValue As String)
    If UBound(mItems(mPositions(Index)).sValue) >= Item Then
        mItems(mPositions(Index)).sValue(Item) = NewValue
    End If
End Property

Public Property Get List(ByVal Index As Integer) As String
    List = mItems(mPositions(Index)).sValue(0)
End Property

Public Property Get ListCount() As Integer
    ListCount = mItemCount + 1
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = mListIndex
End Property

Public Property Let ListIndex(ByVal NewValue As Integer)
    mListIndex = NewValue
    
    If mListIndex >= 0 Then
        mSelectedText = mItems(mPositions(mListIndex)).sValue(0)
        ShowText mInFocus
    Else
        mSelectedText = ""
        ShowText mInFocus
    End If
    
    RaiseEvent Click
End Property

Public Property Get Locked() As Boolean
    Locked = mLocked
End Property

Public Property Let Locked(ByVal NewValue As Boolean)
    mLocked = NewValue
    txtCombo.Locked = mLocked
    
    PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Integer
    MaxLength = mMaxLength
End Property

Public Property Let MaxLength(ByVal NewValue As Integer)
    mMaxLength = NewValue
    txtCombo.MaxLength = mMaxLength
    
    PropertyChanged "MaxLength"
End Property

Private Function NavigateDown() As Boolean
    If mHighlighted < mItemCount Then
        NavigateDown = True
        
        mHighlighted = mHighlighted + 1
        If mHighlighted >= (vscItem.Value + mDropDownItemsVisible) Then
            vscItem.Value = vscItem.Value + 1
        Else
            ShowItems
        End If
    End If
End Function

Private Function NavigateUp() As Boolean
    If mHighlighted > 0 Then
        NavigateUp = True
        
        mHighlighted = mHighlighted - 1
        If mHighlighted < vscItem.Value Then
            vscItem.Value = vscItem.Value - 1
        Else
            ShowItems
        End If
    End If
End Function

Public Property Get NewIndex() As Integer
    NewIndex = mItemCount
End Property

Public Property Get PageScrollItems() As Integer
    PageScrollItems = mPageScrollItems
End Property

Public Property Let PageScrollItems(ByVal NewValue As Integer)
    mPageScrollItems = NewValue
    vscItem.LargeChange = mPageScrollItems
    
    PropertyChanged "PageScrollItems"
End Property

Private Sub picList_Click()
    RaiseEvent Click
End Sub

Private Sub picList_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As RECT
    Dim lX As Long
    Dim nIndex As Integer
    Dim bCancel As Boolean
    Dim bValue As Boolean
    
    If (Button = vbLeftButton) And Not mLocked Then
        Call SetCapture(picList.hWnd)
        mMouseDown = True
        
        nIndex = GetRowFromY(Y)
        
        If nIndex >= 0 Then
            mListIndex = nIndex
            RaiseEvent ClickItem(nIndex, Button, Shift)
        
            Select Case mStyle
            Case Checkboxes
                bValue = Not GetFlag(mPositions(mListIndex), flgChecked)
                RaiseEvent RequestItemChecked(mPositions(mListIndex), bValue, bCancel)
            Case OptionButtons
                bValue = True
                RaiseEvent RequestItemChecked(mPositions(mListIndex), bValue, bCancel)
            End Select

            If Not bCancel Then
                SetFlag mPositions(mListIndex), flgChecked, bValue
                ShowItems
                SetText mListIndex

                RaiseEvent SelectionChanged
            End If
        ElseIf mColumnSort And (picList.MousePointer <> vbSizeWE) Then
            mButtonIndex = GetColFromX(X, lX)
            If mButtonIndex <> NULL_RESULT Then
                With picList
                    Call SetRect(R, lX, 0, lX + mCols(mButtonIndex).lWidth, GetColumnHeadingHeight())
                    Call DrawEdge(.hdc, R, EDGE_SUNKEN, BF_RECT)
                    
                    .Refresh
                End With
            End If
        End If
    End If
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lResizeX As Long
    
    Dim lWidth As Long
    Dim nIndex As Integer
    Dim nPointer As Integer
    
    If Not mLocked Then
        If (Button = vbLeftButton) And (mResizeCol >= 0) Then
            'We are resizing a Column
            lWidth = (X - lResizeX)
            If lWidth > 1 Then
                mCols(mResizeCol).lWidth = lWidth
                mCols(mResizeCol).dCustomWidth = ScaleX(mCols(mResizeCol).lWidth, vbPixels, mScaleUnits)
                
                ShowItems
                SetScrollBars
            End If
        ElseIf Button = 0 Then
            'Only check for resize cursor if no buttons depressed
            lResizeX = 0
            mResizeCol = NULL_RESULT
            
            nIndex = GetRowFromY(Y)
            nPointer = vbDefault
            
            If (nIndex >= 0) Then
                If (mHighlighted <> nIndex) Then
                    mHighlighted = nIndex
                    ShowItems
                End If
            ElseIf mColumnResize Then
                 For nIndex = LBound(mCols) To UBound(mCols)
                    lWidth = lWidth + mCols(nIndex).lWidth
                    
                    If (X < lWidth + 2) And (X > lWidth - 2) Then
                        nPointer = vbSizeWE
                        mResizeCol = nIndex
                        Exit For
                    End If
                    
                    lResizeX = lResizeX + mCols(nIndex).lWidth
                Next nIndex
            End If
        
            With picList
                If .MousePointer <> nPointer Then
                    .MousePointer = nPointer
                End If
            End With
        End If
    End If
End Sub

Private Sub picList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nMouseRow As Integer
    
    If Button = vbLeftButton Then
        Call ReleaseCapture
        
        mMouseDown = False
        nMouseRow = GetRowFromY(Y)
        
        If mLocked Then
            SetDropDown
        ElseIf (nMouseRow < 0) Then
            If (GetColFromX(X) = mButtonIndex) And (mButtonIndex <> NULL_RESULT) Then
                If (Shift And vbCtrlMask) And (mSortColumn <> NULL_RESULT) Then
                    If mSortSubColumn <> mButtonIndex Then
                        mCols(mButtonIndex).nSortOrder = 0
                    End If
                    mSortSubColumn = mButtonIndex
                Else
                    If mSortColumn <> mButtonIndex Then
                        mCols(mButtonIndex).nSortOrder = 0
                        mSortSubColumn = NULL_RESULT
                    End If
                    mSortColumn = mButtonIndex
                End If
                
                If mCols(mButtonIndex).nSortOrder = 0 Then
                    mCols(mButtonIndex).nSortOrder = 1
                Else
                    mCols(mButtonIndex).nSortOrder = 0
                End If
                
                DoSort
                ShowItems
            ElseIf mButtonIndex >= 0 Then
                ShowItems
            End If
        ElseIf (mStyle = Standard) And (mResizeCol < 0) Then
            mListIndex = nMouseRow
            
            SetDropDown
        End If
    End If
End Sub

Public Sub Refresh()
    SetText mListIndex
    
    If picList.Visible Then
        ShowItems
    End If
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
    Dim nCount As Integer
    Dim nPosition As Integer
   
    '#############################################################################################################################
    'See AddItem for details of the Arrays used
    '#############################################################################################################################
   
    nPosition = mPositions(Index)
    
    'Reset Item Data
    For nCount = mPositions(Index) To mItemCount - 1
        mItems(nCount) = mItems(nCount + 1)
    Next nCount
    
    'Adjust Item Pointers
    For nCount = Index To mItemCount - 1
        mPositions(nCount) = mPositions(nCount + 1)
    Next nCount
    
    'Validate Pointers for Items after deleted Item
    For nCount = 1 To mItemCount - 1
        If mPositions(nCount) > nPosition Then
            mPositions(nCount) = mPositions(nCount) - 1
        End If
    Next nCount
    
    mItemCount = mItemCount - 1
    If (mItemCount + CACHE_INCREMENT) < UBound(mItems) Then
        ReDim Preserve mItems(mItemCount)
        ReDim Preserve mPositions(mItemCount)
    End If
End Sub

Public Property Get RequireCheckedItem() As Boolean
    RequireCheckedItem = mRequireCheckedItem
End Property

Public Property Let RequireCheckedItem(ByVal NewValue As Boolean)
    mRequireCheckedItem = NewValue
    
    PropertyChanged "RequireCheckedItem"
End Property

Public Property Get RowHeightMin() As Single
    RowHeightMin = mRowHeightMin
End Property

Public Property Let RowHeightMin(ByVal NewValue As Single)
    mRowHeightMin = NewValue
    
    PropertyChanged "RowHeightMin"
End Property

Public Property Get ScaleUnits() As ScaleModeConstants
    ScaleUnits = mScaleUnits
End Property

Public Property Let ScaleUnits(ByVal NewValue As ScaleModeConstants)
    mScaleUnits = NewValue
    
    PropertyChanged "ScaleUnits"
End Property

Private Function ScaleValue(ByVal lValue As Long, ByVal lMin As Long, ByVal lMax As Long) As Long
    If lValue > lMax Then
        ScaleValue = lMax
    ElseIf lValue < lMin Then
        ScaleValue = lMin
    Else
        ScaleValue = lValue
    End If
End Function

Private Function SearchCode(sCode As String, nMode As SearchEnum) As Integer
    Dim nCount As Integer
    
    SearchCode = NULL_RESULT
    
    For nCount = LBound(mItems) To mItemCount
        Select Case nMode
        Case cvEqual
            If UCase$(mItems(mPositions(nCount)).sValue(mSearchColumn)) = sCode Then
                SearchCode = nCount
                Exit For
            End If
        
        Case cvGreaterEqual
            If UCase$(Left$(mItems(mPositions(nCount)).sValue(mSearchColumn), Len(sCode))) >= sCode Then
                SearchCode = nCount
                Exit For
            End If
        
        Case cvLike
            If UCase$(mItems(mPositions(nCount)).sValue(mSearchColumn)) Like sCode & "*" Then
                SearchCode = nCount
                Exit For
            End If

        End Select
        
    Next nCount
End Function

Public Property Get SelCount() As Integer
    Dim nCount As Integer
    Dim nSelected As Integer
    
    For nCount = LBound(mItems) To mItemCount
        If GetFlag(nCount, flgChecked) Then
            nSelected = nSelected + 1
        End If
    Next nCount
    
    SelCount = nSelected
End Property

Public Property Get SelLength() As Integer
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txtCombo.SelLength
End Property

Public Property Let SelLength(ByVal NewValue As Integer)
    txtCombo.SelLength = NewValue
End Property

Public Property Get SelStart() As Integer
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txtCombo.SelStart
End Property

Public Property Let SelStart(ByVal NewValue As Integer)
    txtCombo.SelStart = NewValue
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    SelText = txtCombo.SelText
End Property

Public Property Let SelText(ByVal NewValue As String)
    txtCombo.SelText = NewValue
End Property

Private Sub SetDropDown(Optional bDrawButton As Boolean)
    Dim oRect As RECT
    Dim dRowHeight As Single
    Dim dHeight As Single
    Dim dColumnsWidth As Single
    Dim dWidth As Single
    Dim dLeft As Single
    Dim dTop As Single
    Dim nCount As Integer
    
    With picList
        .ScaleMode = vbTwips
        
        If .Visible Then
            SetTimer False
            SetText mListIndex
            
            If (mStyle = Standard) And (mListIndex >= 0) Then
                RaiseEvent Click
            End If
            
            .Visible = False
            
            With txtCombo
                If .Visible Then
                    .SetFocus
                End If
            End With
            
            RaiseEvent DropDownClose
        ElseIf ListCount() > 0 Then
            RaiseEvent DropDownOpen
            
            If bDrawButton And (mBorderStyle <> BorderCustom) Then
                With UserControl
                    Call DrawEdge(.hdc, mButtonRect, EDGE_SUNKEN, BF_RECT)
                    .Picture = .Image
                End With
                mButtonClickTick = GetTickCount()
            End If
        
            mHighlighted = mListIndex
            
            For nCount = LBound(mCols) To UBound(mCols)
                If mCols(nCount).bVisible Then
                    dColumnsWidth = dColumnsWidth + mCols(nCount).dCustomWidth
                End If
            Next nCount
            
            If mDropDownAutoWidth Then
                dWidth = dColumnsWidth + vscItem.Width
            ElseIf mDropDownWidth > 0 Then
                dWidth = mDropDownWidth
            Else
                dWidth = UserControl.Width
            End If
            
            If dWidth > Screen.Width Then
                .Width = Screen.Width
            Else
                .Width = dWidth
            End If
            
            Set .Font = mDropDownFont

            If mRowHeightMin > 0 Then
                dRowHeight = ScaleY(mRowHeightMin, mScaleUnits, vbTwips)
            Else
                dRowHeight = .TextHeight("A")
            End If
            
            If ListCount() > mDropDownItemsVisible Then
                If dColumnsWidth > dWidth Then
                    dHeight = (dRowHeight * mDropDownItemsVisible) + (Screen.TwipsPerPixelY * 2) + (hscItem.Height + (Screen.TwipsPerPixelY * 2))
                Else
                    dHeight = (dRowHeight * mDropDownItemsVisible) + (Screen.TwipsPerPixelY * 2)
                End If
            Else
                dHeight = (dRowHeight * ListCount()) + (Screen.TwipsPerPixelY * 2)
            End If
            
            If mColumnHeaders Then
                dHeight = dHeight + GetColumnHeadingHeight()
            End If
            .Height = dHeight
            
            vscItem.Left = (.Width - vscItem.Width) - (Screen.TwipsPerPixelY * 2)
            
            hscItem.Top = (.Height - hscItem.Height) - (Screen.TwipsPerPixelY * 2)
            hscItem.Width = .Width - (Screen.TwipsPerPixelX * 2)
            hscItem.Value = hscItem.Min
            hscItem.Max = UBound(mCols)
            
            SetScrollBars
            
            Call GetWindowRect(hWnd, oRect)
            dLeft = oRect.Left * Screen.TwipsPerPixelX
            If (dLeft + .Width) > Screen.Width Then
                dLeft = Screen.Width - .Width
            End If
            
            dTop = oRect.Bottom * Screen.TwipsPerPixelY
            If (dTop + dHeight) > Screen.Height Then
                dTop = (oRect.Bottom * Screen.TwipsPerPixelY) - (UserControl.Height + dHeight)
            End If
            
            Call picList.Move(dLeft, dTop)
            Call SetWindowPos(.hWnd, -1, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
            
            .ScaleMode = vbPixels
            
            If mEditable Then
                mListIndex = SearchCode(UCase$(mSelectedText), cvEqual)
            End If
            mHighlighted = mListIndex
            
            If ListCount() > mDropDownItemsVisible Then
                If mListIndex > vscItem.Max Then
                    vscItem.Value = vscItem.Max
                ElseIf mListIndex > 0 Then
                    vscItem.Value = mListIndex
                Else
                    vscItem.Value = 0
                End If
            Else
                vscItem.Value = 0
            End If
            
            ShowItems
            
            ShowText False
            .Visible = True
            .SetFocus
            
            SetTimer True
        End If
    End With
End Sub

Private Sub SetFlag(ByVal nIndex As Long, nFlag As FlagsEnum, bValue As Boolean)
    Dim nCount As Integer
    
    'Sets information by bit flags for a ListItem.
    
    If (nFlag = flgChecked) And mRequireCheckedItem Then
        If SelCount() = 1 And Not bValue Then
            bValue = True
        End If
    End If
    
    If bValue Then
        If nFlag = flgChecked And (mStyle <> Checkboxes) Then
            For nCount = LBound(mItems) To UBound(mItems)
                If mItems(nCount).nFlags And nFlag Then
                    mItems(nCount).nFlags = mItems(nCount).nFlags Xor nFlag
                End If
            Next nCount
        End If
        
        mItems(nIndex).nFlags = mItems(nIndex).nFlags Or nFlag
    Else
        If mItems(nIndex).nFlags And nFlag Then
            mItems(nIndex).nFlags = mItems(nIndex).nFlags Xor nFlag
        End If
    End If
End Sub

Private Sub SetFlags(nFlag As FlagsEnum, bValue As Boolean)
    Dim nCount As Integer
    
    For nCount = LBound(mItems) To UBound(mItems)
        If bValue Then
            mItems(nCount).nFlags = mItems(nCount).nFlags Or nFlag
        Else
            If mItems(nCount).nFlags And nFlag Then
                mItems(nCount).nFlags = mItems(nCount).nFlags Xor nFlag
            End If
        End If
    Next nCount
End Sub

Public Sub SetItem(ByVal vData As Variant, Optional ByVal nDefault As Integer = -1)
    Dim nCount As Long
    Dim bFound As Boolean
    Dim bItemData As Boolean
    
    If VarType(vData) = vbLong Then
        bItemData = True
    End If

    For nCount = 0 To mItemCount
        If bItemData Then
            If vData = mItems(nCount).lItemData Then
                bFound = True
                ListIndex = nCount
                Exit For
            End If
        Else
            If vData = mItems(nCount).sValue(0) Then
                bFound = True
                ListIndex = nCount
                Exit For
            End If
        End If
    Next nCount
    
    If Not bFound And nDefault >= 0 Then
        ListIndex = nDefault
    End If
End Sub

Private Sub SetScrollBars()
    Dim dHeight As Single
    Dim dWidth As Single
    Dim dRowHeight As Single
    Dim nCount As Integer
    Dim nDropDownItemsVisible As Integer
    
    '#############################################################################################################################
    'Sets the visibilty of scroll bars and sets max scroll values
    '#############################################################################################################################
    
    picList.ScaleMode = vbTwips
    
    'Calculate total width of columns
    For nCount = LBound(mCols) To UBound(mCols)
        If mCols(nCount).bVisible Then
            dWidth = dWidth + mCols(nCount).dCustomWidth
        End If
    Next nCount
    
    If vscItem.Visible Then
        dWidth = dWidth + vscItem.Width
    End If
    
    With hscItem
        .Visible = (dWidth > picList.Width)
    End With
    
    'Calculate total height available for drawing Items
    dHeight = picList.Height
    If mColumnHeaders Then
        dHeight = dHeight - ScaleY(GetColumnHeadingHeight(), mScaleUnits, vbTwips)
    End If
    
    If mRowHeightMin > 0 Then
        dRowHeight = ScaleY(mRowHeightMin, mScaleUnits, vbTwips)
    Else
        dRowHeight = picList.TextHeight("A")
    End If
    
    With vscItem
        If (dWidth > picList.Width) Then
            dHeight = dHeight - (hscItem.Height + Screen.TwipsPerPixelY * 2)
            .Height = picList.Height - (hscItem.Height + Screen.TwipsPerPixelY * 2)
        Else
            .Height = picList.Height - (Screen.TwipsPerPixelY * 2)
        End If
        
        'This may differ from the DropDownItemsVisible Property if scroll bars
        'have been forced by user dragging a column wider
        nDropDownItemsVisible = (dHeight / dRowHeight)
        
        If ListCount() > nDropDownItemsVisible Then
            vscItem.Max = mItemCount - (nDropDownItemsVisible - 1)
            vscItem.Visible = True
        Else
            vscItem.Max = mItemCount
            vscItem.Visible = False
        End If
    End With
    
    picList.ScaleMode = vbPixels
End Sub

Private Sub SetText(nIndex As Integer)
    Dim nCol As Integer
    Dim nCount As Integer
    Dim nSelCount(1) As Integer
    Dim nFirstSelected As Integer
    
    If mStyle = Standard Then
        If nIndex >= 0 Then
            mSelectedText = mItems(mPositions(nIndex)).sValue(0)
            ShowText mInFocus
        End If
    Else
        For nCount = LBound(mCols) To UBound(mCols)
            If mCols(nCount).bVisible And mCols(nCount).dCustomWidth > 0 Then
                nCol = nCount
                Exit For
            End If
        Next nCount
        
        nFirstSelected = -1
        For nCount = LBound(mItems) To mItemCount
            If GetFlag(nCount, flgChecked) Then
                nSelCount(0) = nSelCount(0) + 1
                If nFirstSelected < 0 Then
                    nFirstSelected = nCount
                End If
            Else
                nSelCount(1) = nSelCount(1) + 1
            End If
        Next nCount
        
        If nSelCount(0) = 1 Then
            mListIndex = nFirstSelected
            mSelectedText = mItems(nFirstSelected).sValue(nCol)
        ElseIf (nSelCount(0) > 0) And (nSelCount(1) = 0) Then
            mSelectedText = mTextAll
        ElseIf (nSelCount(0) = 0) Then
            mSelectedText = mTextNone
        Else
            mSelectedText = mTextSelection
        End If
        
        ShowText mInFocus
    End If
End Sub

Private Sub SetTimer(bEnabled As Boolean)
    If tmrRelease.Enabled <> bEnabled Then
        If bEnabled Then
            tmrRelease.Enabled = True
            'Debug.Print "Timer ON"
        Else
            tmrRelease.Enabled = False
            'Debug.Print "Timer OFF"
        End If
    End If
End Sub

Private Sub ShowItems()
    Const CHECKBOX_SIZE = 11
    Const OPTIONBUTTON_SIZE = 10
    Const SORTARROW_SIZE = 8
    Const SMALL_SORTARROW_SIZE = 6
    
    Const HEADER_LEFT = 3
    Const IMAGE_LEFT = 2
    
    Dim R As RECT
    Dim lX As Long
    Dim lY As Long
    
    Dim lLeftImage As Long
    Dim lLeftText As Long
    Dim lTextHeight As Long
    Dim lColumnHeadingHeight As Long
    
    Dim lCBSpace As Long
    Dim lImageSpace As Long
    Dim lSortSpace As Long
    Dim nCount As Integer
    Dim nItem As Integer
    Dim bRenderImages As Boolean
    Dim sText As String
    
    'Left Position to Draw Text
    If mStyle <> Standard Then
        lLeftText = 15
    Else
        lLeftText = 3
    End If
    
    'Left Position to Draw Images
    lLeftImage = ScaleX(lLeftText, vbPixels, vbTwips)
    
    'Adjust Text Position for Images
    If Not mImageList Is Nothing Then
        bRenderImages = True
        lImageSpace = ((GetRowHeight() - mImageList.ImageHeight) / 2)
        lLeftText = lLeftText + mImageList.ImageWidth + 2
    End If
    
    lCBSpace = ((GetRowHeight() - CHECKBOX_SIZE) / 2)
    
    With picList
        .Cls
        .DrawWidth = 1
        .ForeColor = vbWindowText
        
        lColumnHeadingHeight = GetColumnHeadingHeight()
        lTextHeight = .TextHeight("A")
        
        '#############################################################################################################################
        'Column Headers
        If mColumnHeaders Then
            Call SetRect(R, 0, 0, .ScaleWidth, lColumnHeadingHeight)
            DrawRect .hdc, R, GetSysColor(COLOR_BTNFACE), True
            
            For nCount = hscItem.Value To UBound(mCols)
                 If mCols(nCount).bVisible Then
                    'Draw the Column Header Buttons
                    Call SetRect(R, lX, 0, lX + mCols(nCount).lWidth, lColumnHeadingHeight)
                    Call DrawEdge(.hdc, R, EDGE_RAISED, BF_RECT)
                 
                    Call SetRect(R, lX + HEADER_LEFT, (lColumnHeadingHeight / 2) - (lTextHeight / 2), (lX + mCols(nCount).lWidth) - HEADER_LEFT, lColumnHeadingHeight)
                    
                    sText = mCols(nCount).sHeading
                    
                    'Format/Render Text
                    If mDisplayEllipsis Then
                        Call DrawText(.hdc, sText, -1, R, mCols(nCount).nAlignment Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
                    Else
                        Call DrawText(.hdc, sText, -1, R, mCols(nCount).nAlignment Or DT_SINGLELINE)
                    End If
                    
                    'Render Sort Arrows
                    If mCols(nCount).lWidth > SORTARROW_SIZE Then
                        If nCount = mSortColumn Then
                            lSortSpace = (lColumnHeadingHeight / 2) - (SORTARROW_SIZE / 2)
                            If mCols(nCount).nSortOrder = 1 Then
                                Call BitBlt(.hdc, R.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 25, 14, MERGEPAINT)
                                Call BitBlt(.hdc, R.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 1, 14, SRCAND)
                            Else
                                Call BitBlt(.hdc, R.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 37, 14, MERGEPAINT)
                                Call BitBlt(.hdc, R.Right - SORTARROW_SIZE, lY + lSortSpace, SORTARROW_SIZE, 7, picImages.hdc, 13, 14, SRCAND)
                            End If
                        ElseIf nCount = mSortSubColumn Then
                            lSortSpace = (lColumnHeadingHeight / 2) - (SMALL_SORTARROW_SIZE / 2)
                            If mCols(nCount).nSortOrder = 1 Then
                                Call BitBlt(.hdc, R.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 26, 23, MERGEPAINT)
                                Call BitBlt(.hdc, R.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 2, 23, SRCAND)
                            Else
                                Call BitBlt(.hdc, R.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 38, 23, MERGEPAINT)
                                Call BitBlt(.hdc, R.Right - SMALL_SORTARROW_SIZE, lY + lSortSpace, SMALL_SORTARROW_SIZE, 5, picImages.hdc, 14, 23, SRCAND)
                            End If
                        End If
                    End If
                    
                    lX = lX + mCols(nCount).lWidth
                End If
            Next nCount
            
            lY = lColumnHeadingHeight
        End If
        
        lTextHeight = GetRowHeight()
        
        '#############################################################################################################################
        'List Items
        For nItem = vscItem.Value To (vscItem.Value + mDropDownItemsVisible) - 1
            If nItem > mItemCount Then
                Exit For
            End If
            
            If nItem = mHighlighted Then
                'Draw Highlight & Focus Rectangles
                Call SetRect(R, lLeftText, lY, .ScaleWidth, lY + lTextHeight)
                DrawRect .hdc, R, GetSysColor(COLOR_HIGHLIGHT), True
                
                Select Case mFocusRectStyle
                Case FocusRectLight
                    Call DrawFocusRect(.hdc, R)
                Case FocusRectHeavy
                    DrawRect .hdc, R, TranslateColor(mFocusRectColor), False
                End Select
                
                .ForeColor = vbHighlightText
            Else
                .ForeColor = mItems(mPositions(nItem)).lForeColor
            End If
            .FontBold = mItems(mPositions(nItem)).nFlags And flgBold
            
            'Blit appropriate Checkbox Image
            Select Case mStyle
            Case Checkboxes
                If mItems(mPositions(nItem)).nFlags And flgChecked Then
                    Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, CHECKBOX_SIZE, CHECKBOX_SIZE, picImages.hdc, 11, 0, SRCCOPY)
                Else
                    Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, CHECKBOX_SIZE, CHECKBOX_SIZE, picImages.hdc, 0, 0, SRCCOPY)
                End If
            
            Case OptionButtons
                If mItems(mPositions(nItem)).nFlags And flgChecked Then
                    Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, OPTIONBUTTON_SIZE, OPTIONBUTTON_SIZE, picImages.hdc, 32, 0, SRCCOPY)
                Else
                    Call BitBlt(.hdc, IMAGE_LEFT, lY + lCBSpace, OPTIONBUTTON_SIZE, OPTIONBUTTON_SIZE, picImages.hdc, 22, 0, SRCCOPY)
                End If

            End Select
            
            If bRenderImages Then
                'If we have an Image Index then Draw it
                If mItems(mPositions(nItem)).vImage <> Empty Then
                    If nItem = mHighlighted Then
                        mImageList.ListImages(mItems(mPositions(nItem)).vImage).Draw .hdc, lLeftImage, ScaleY(lY + lImageSpace, vbPixels, vbTwips), 2
                    Else
                        mImageList.ListImages(mItems(mPositions(nItem)).vImage).Draw .hdc, lLeftImage, ScaleY(lY + lImageSpace, vbPixels, vbTwips), 1
                    End If
                End If
            End If
            
            lX = -1
            For nCount = hscItem.Value To UBound(mCols)
                If mCols(nCount).bVisible Then
                    If lX < 0 Then
                        lX = 1
                        Call SetRect(R, lLeftText, lY, (lLeftText + mCols(nCount).lWidth) - lLeftText, lY + lTextHeight)
                     Else
                        Call SetRect(R, lX, lY, (lX + mCols(nCount).lWidth) - 3, lY + lTextHeight)
                     End If
                
                    sText = mItems(mPositions(nItem)).sValue(nCount)
                    
                    'Format/Render Text
                    If mDisplayEllipsis Then
                        Call DrawText(.hdc, sText, -1, R, mCols(nCount).nAlignment Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
                    Else
                        Call DrawText(.hdc, sText, -1, R, mCols(nCount).nAlignment Or DT_SINGLELINE)
                    End If
                    
                    lX = lX + mCols(nCount).lWidth
                End If
            Next nCount
            
            lY = lY + lTextHeight
        Next nItem
    End With
End Sub

Private Sub ShowText(Optional bFocus As Boolean)
    Dim R As RECT
    
    With UserControl
        .Cls
        
        'Are are using a Textbox or drawing Text?
        If mEditable Then
            mLockTextBoxEvent = True
            txtCombo.Text = mSelectedText
            mLockTextBoxEvent = False
            
            If bFocus Then
                txtCombo.SelStart = 0
                txtCombo.SelLength = Len(txtCombo.Text)
            End If
        Else
            If (mBorderStyle = BorderCustom) And (mBorderCurve > 0) Then
                Call SetRect(R, BORDER_LEFT + 2, BORDER_TOP, BORDER_LEFT + txtCombo.Width, BORDER_TOP + txtCombo.Height)
            Else
                Call SetRect(R, BORDER_LEFT - 1, BORDER_TOP, BORDER_LEFT + txtCombo.Width, BORDER_TOP + txtCombo.Height)
            End If
                        
            If mEnabled Then
                If bFocus Then
                    'Draw Highlight & Focus Rectangles
                    DrawRect .hdc, R, GetSysColor(COLOR_HIGHLIGHT), True
                    Call DrawFocusRect(.hdc, R)
                    .ForeColor = vbHighlightText
                Else
                    'Clear any previous Highlight/Focus Rectangles
                    DrawRect .hdc, R, TranslateColor(mBackColor), True
                End If
            Else
                DrawRect .hdc, R, TranslateColor(mBackColor), True
                .ForeColor = vbGrayText
            End If
                         
            R.Left = R.Left + 1
            R.Top = R.Top + 1
            
            Select Case txtCombo.Alignment
            Case vbRightJustify
                Call DrawText(.hdc, mSelectedText, -1, R, DT_RIGHT)
            Case vbCenter
                Call DrawText(.hdc, mSelectedText, -1, R, DT_CENTER)
            Case Else
                Call DrawText(.hdc, mSelectedText, -1, R, DT_LEFT)
            End Select
            
            .ForeColor = mForeColor
        End If
    End With
End Sub

Public Sub Sort(Column As Integer, SortOrder As SortOrderEnum, Optional SubColumn As Integer = -1, Optional SubSortOrder As Integer)
    mSortColumn = Column
    mCols(Column).nSortOrder = SortOrder
    
    mSortSubColumn = SubColumn
    If SubColumn >= 0 Then
        mCols(SubColumn).nSortOrder = SubSortOrder
    End If
    
    DoSort
End Sub

Private Sub SortArray(ByVal lFirst As Long, ByVal lLast As Long, nSortColumn As Integer, nSortType As Integer)
    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapInt mPositions(lFirst), mPositions((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            Select Case mCols(nSortColumn).nType
            Case TypeDate
                bSwap = CDate(mItems(mPositions(lIndex)).sValue(nSortColumn)) > CDate(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeNumeric
                bSwap = Val(mItems(mPositions(lIndex)).sValue(nSortColumn)) > Val(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeCustom
                RaiseEvent CustomSort(True, nSortColumn, mItems(mPositions(lIndex)).sValue(nSortColumn), mItems(mPositions(lFirst)).sValue(nSortColumn), bSwap)
            
            Case Else
                bSwap = mItems(mPositions(lIndex)).sValue(nSortColumn) > mItems(mPositions(lFirst)).sValue(nSortColumn)
            End Select
        Else
            Select Case mCols(nSortColumn).nType
            Case TypeDate
                bSwap = CDate(mItems(mPositions(lIndex)).sValue(nSortColumn)) < CDate(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeNumeric
                bSwap = Val(mItems(mPositions(lIndex)).sValue(nSortColumn)) < Val(mItems(mPositions(lFirst)).sValue(nSortColumn))
            Case TypeCustom
                RaiseEvent CustomSort(False, nSortColumn, mItems(mPositions(lIndex)).sValue(nSortColumn), mItems(mPositions(lFirst)).sValue(nSortColumn), bSwap)
            
            Case Else
                bSwap = mItems(mPositions(lIndex)).sValue(nSortColumn) < mItems(mPositions(lFirst)).sValue(nSortColumn)
            End Select
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapInt mPositions(lBoundary), mPositions(lIndex)
        End If
    Next lIndex

    SwapInt mPositions(lFirst), mPositions(lBoundary)
    SortArray lFirst, lBoundary - 1, nSortColumn, nSortType
    SortArray lBoundary + 1, lLast, nSortColumn, nSortType
End Sub

Private Sub SortSubList()
    Dim sMajorSort As String
    Dim lStartSort As Long
    Dim nCount As Integer
    Dim bDifferent As Boolean

    If mSortSubColumn > NULL_RESULT Then
        'Re-Sort the Items by a secondary column, preserving the sort sequence of the
        'primary sort
        
        lStartSort = LBound(mItems)
        For nCount = LBound(mItems) To mItemCount
            bDifferent = mItems(mPositions(nCount)).sValue(mSortColumn) <> sMajorSort
            If bDifferent Or nCount = mItemCount Then
                If nCount > 1 Then
                    If nCount - lStartSort > 1 Then
                        If nCount = mItemCount And Not bDifferent Then
                            SortArray lStartSort, nCount, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        Else
                            SortArray lStartSort, nCount - 1, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        End If
                    End If
                    lStartSort = nCount
                End If
                
                sMajorSort = mItems(mPositions(nCount)).sValue(mSortColumn)
            End If
        Next nCount
    End If
End Sub

Public Property Get Style() As StyleEnum
    Style = mStyle
End Property

Public Property Let Style(ByVal NewValue As StyleEnum)
    mStyle = NewValue
    
    If mStyle = OptionButtons Then
        SetFlags flgChecked, False
    End If
    SetText mListIndex
    
    PropertyChanged "Style"
End Property

'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    I = I - 1                                                                           'Next element
  Loop
Errs:
End Sub

Private Sub SwapInt(Int1 As Integer, Int2 As Integer)
    Dim sTemp As Integer

    sTemp = Int1
    Int1 = Int2
    Int2 = sTemp
End Sub

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = mSelectedText
End Property

Public Property Let Text(ByVal NewValue As String)
    Dim nCount As Integer
    
    If mEditable Then
        If mStyle = Standard Then
            mListIndex = NULL_RESULT
            
            For nCount = LBound(mItems) To mItemCount
                If mItems(mPositions(nCount)).sValue(mSearchColumn) = NewValue Then
                    mListIndex = nCount
                    Exit For
                End If
            Next nCount
        End If
        
        mSelectedText = NewValue
        ShowText mInFocus
    Else
        Err.Raise 383, "ComboView", "Text is Read-Only"
    End If
End Property

Public Property Get TextAll() As String
    TextAll = mTextAll
End Property

Public Property Let TextAll(ByVal NewValue As String)
    mTextAll = NewValue
    
    PropertyChanged "TextAll"
End Property

Public Property Get TextNone() As String
    TextNone = mTextNone
End Property

Public Property Let TextNone(ByVal NewValue As String)
    mTextNone = NewValue
    
    PropertyChanged "TextNone"
End Property

Public Property Get TextSelection() As String
    TextSelection = mTextSelection
End Property

Public Property Let TextSelection(ByVal NewValue As String)
    mTextSelection = NewValue
    
    PropertyChanged "TextSelection"
End Property

Private Sub tmrRelease_Timer()
    Dim uPoint  As POINTAPI
    Dim uRect As RECT
    Dim nLB As Integer
    Dim nRB As Integer
    
    '#############################################################################################################################
    'This is soley for detecting if we have clicked on a container which does not generate
    'WM_KILLFOCUS message for us to detect. i.e. the parent Form or a Frame
    
    'I don't like Timers in UserControls but wanted to make the Control behave as a normal Combo which
    'closes DropDown when the above situation occurs. I may still remove this "feature"!
    
    'NOTE: This Timer is only Enabled when we detect a WM_MOUSELEAVE so it does not fire unneccessarily
    'while the DropDown is displayed. It is Disabled as soon as the mouse re-enters the DropDown.
    '#############################################################################################################################
    
    Call GetCursorPos(uPoint)
    Call GetWindowRect(picList.hWnd, uRect)
        
    nLB = GetAsyncKeyState(VK_LBUTTON)
    nRB = GetAsyncKeyState(VK_RBUTTON)
    
    If (uPoint.X >= uRect.Left) And (uPoint.X <= uRect.Right) And (uPoint.Y >= uRect.Top) And (uPoint.Y <= uRect.Bottom) Then
        'The mouse pointer is within the Dropdown list
    ElseIf nLB Or nRB Then
        Select Case WindowFromPoint(uPoint.X, uPoint.Y)
        Case UserControl.hWnd
            'The mouse pointer is within the Control
        Case Else
            If (GetTickCount() - mScrollTick) > EVENT_TIMEOUT Then
                If picList.Visible Then
                    SetDropDown
                Else
                    SetTimer False
                End If
            End If
        
        End Select
    End If
End Sub

'Track the mouse hovering the indicated window
Private Sub TrackMouseHover(ByVal lng_hWnd As Long, lHoverTime As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_HOVER
      .hwndTrack = lng_hWnd
      .dwHoverTime = lHoverTime
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional hPalette As Long = 0) As Long
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub txtCombo_Change()
    Dim nResult As Integer
    Dim nStart As Integer
    Dim sText As String
    
    If Not mLockTextBoxEvent Then
        If mAutoComplete Then
            With txtCombo
                nStart = .SelStart
                sText = Left$(.Text, nStart)
                If Len(sText) > 0 Then
                    'nResult = SearchCode(UCase$(sText), cvLike)
                    nResult = SearchCode(UCase$(sText), cvGreaterEqual)
                    If (nResult > NULL_RESULT) Then
                        mLockTextBoxEvent = True
                        .SelText = Mid$(mItems(mPositions(nResult)).sValue(0), nStart + 1)
                        .SelStart = nStart
                        .SelLength = Len(.Text) - nStart
                        mLockTextBoxEvent = False
                    End If
                End If
            End With
        End If
        
        RaiseEvent Change
    End If
End Sub

Private Sub txtCombo_Click()
    RaiseEvent Click
End Sub

Private Sub txtCombo_DblClick()
    Dim bCancel As Boolean
    Dim bValue As Boolean
    
    If mStyle = Checkboxes Then
        bValue = (SelCount() <> ListCount())
        RaiseEvent RequestListChecked(bValue, bCancel)
    Else
        bCancel = True
    End If
    
    If bCancel Then
        RaiseEvent DblClick
    Else
        SetFlags flgChecked, bValue
        Refresh
        
        RaiseEvent SelectionChanged
    End If
End Sub

Private Sub txtCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nResult As Integer
    
    If mAutoComplete And (txtCombo.SelLength > 0) Then
        Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            mLockTextBoxEvent = True
            txtCombo.SelText = ""
            mLockTextBoxEvent = False
        
        Case vbKeyReturn
            nResult = SearchCode(UCase$(mSelectedText), cvEqual)
            If (nResult > NULL_RESULT) Then
                mSelectedText = mItems(mPositions(nResult)).sValue(0)
                If picList.Visible Then
                    SetDropDown
                End If
            End If

        End Select
    End If
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtCombo_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtCombo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtCombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtCombo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_DblClick()
    If mEnabled And (GetTickCount() - mButtonClickTick) > EVENT_TIMEOUT Then
        txtCombo_DblClick
    End If
End Sub

Private Sub UserControl_EnterFocus()
    'Debug.Print "UserControl_EnterFocus"
    
    mInFocus = True
    
    If Not picList.Visible Then
        ShowText True
    End If
End Sub

Private Sub UserControl_ExitFocus()
    DoKillFocus
End Sub

Private Sub UserControl_Initialize()
    Dim OS As OSVERSIONINFO
      
    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    
    mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    
    ReDim mCols(0)
    Clear
End Sub

Private Sub UserControl_InitProperties()
    Set mFont = Ambient.Font
    Set mDropDownFont = Ambient.Font
    
    mAlignment = DEF_ALIGNMENT
    mAutoComplete = DEF_AUTOCOMPLETE
    mBackColor = DEF_BACKCOLOR
    mBorderColor = DEF_BORDERCOLOR
    mBorderCurve = DEF_BORDERCURVE
    mBorderStyle = DEF_BORDERSTYLE
    mBorderWidth = DEF_BORDERWIDTH
    mButtonBackColor = DEF_BUTTONBACKCOLOR
    mColumnHeaders = DEF_COLUMNHEADERS
    mColumnResize = DEF_COLUMNRESIZE
    mColumnSort = DEF_COLUMNSORT
    mDefaultItemForeColor = DEF_DEFAULTITEMFORECOLOR
    mDropDownAutoWidth = DEF_DROPDOWNAUTOWIDTH
    mDropDownItemsVisible = DEF_DROPDOWNITEMSVISIBLE
    mDropDownWidth = DEF_DROPDOWNWIDTH
    mEditable = DEF_EDITABLE
    mEnabled = DEF_ENABLED
    mFocusRectColor = DEF_FOCUSRECTCOLOR
    mFocusRectStyle = DEF_FOCUSRECTSTYLE
    mForeColor = DEF_FORECOLOR
    mIntegralHeight = DEF_INTEGRALHEIGHT
    mLocked = DEF_LOCKED
    mMaxLength = 0
    mPageScrollItems = DEF_PAGESCROLLITEMS
    mRequireCheckedItem = DEF_REQUIRECHECKEDITEM
    mRowHeightMin = DEF_ROWHEIGHTMIN
    mScaleUnits = DEF_SCALEUNITS
    mSearchColumn = DEF_SEARCHCOLUMN
    mStyle = DEF_STYLE
    mTextAll = DEF_TEXTALL
    mTextNone = DEF_TEXTNONE
    mTextSelection = DEF_TEXTSELECTION
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If mEnabled Then
        If Shift And vbCtrlMask Then
            mIgnoreKeyPress = True
            
            If KeyCode = vbKeyA Then
                KeyCode = 0
                
                SetFlags flgChecked, (SelCount() <> ListCount())
                Refresh
                
                RaiseEvent SelectionChanged
            End If
        End If
        
        If picList.Visible Then
            Select Case KeyCode
            Case vbKeyF4
                mListIndex = mHighlighted
                SetDropDown
                
            Case vbKeyEscape
                SetDropDown
            
            Case vbKeyUp
                If NavigateUp() Then
                    KeyCode = 0
                    SetText mHighlighted
                End If
            Case vbKeyDown
                If NavigateDown() Then
                    KeyCode = 0
                    SetText mHighlighted
                End If
            
            Case vbKeyPageUp
                If mHighlighted > 0 Then
                    KeyCode = 0
                    mHighlighted = (mHighlighted - mDropDownItemsVisible) + 1
                    If mHighlighted < 0 Then
                        mHighlighted = 0
                    End If
                    
                    vscItem.Value = mHighlighted
                    ShowItems
                    
                    SetText mHighlighted
                End If
            
            Case vbKeyPageDown
                If mHighlighted < mItemCount Then
                    KeyCode = 0
                    mHighlighted = ScaleValue((mHighlighted + mDropDownItemsVisible) - 1, 0, mItemCount)
                    vscItem.Value = ScaleValue(mHighlighted, 0, vscItem.Max)
                    ShowItems
                    SetText mHighlighted
                End If
            
            Case vbKeySpace
                If mHighlighted >= 0 Then
                    mIgnoreKeyPress = True
                    KeyCode = 0
                    
                    SetFlag mHighlighted, flgChecked, Not GetFlag(mHighlighted, flgChecked)
                    ShowItems
                    
                    RaiseEvent SelectionChanged
                End If
            
            End Select
        Else
            Select Case KeyCode
            Case vbKeyF4
                SetDropDown
            
            Case vbKeyUp
                If mListIndex > 0 Then
                    KeyCode = 0
                    mListIndex = mListIndex - 1
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
            Case vbKeyDown
                If mListIndex < mItemCount Then
                    KeyCode = 0
                    mListIndex = mListIndex + 1
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
            
             Case vbKeyPageUp
                If mListIndex > 0 Then
                    KeyCode = 0
                    mListIndex = ScaleValue((mListIndex - mDropDownItemsVisible) + 1, 0, mItemCount)
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
           
             Case vbKeyPageDown
                If mListIndex < mItemCount Then
                    KeyCode = 0
                    mListIndex = ScaleValue((mListIndex + mDropDownItemsVisible) - 1, 0, mItemCount)
                    SetText mListIndex
                    
                    RaiseEvent Click
                End If
           
            End Select
        End If
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Static lTime As Long
    Static sCode As String
    Dim nResult As Integer
    
    If picList.Visible And Not mIgnoreKeyPress Then
        If (GetTickCount() - lTime) < 1000 Then
            sCode = sCode & Chr$(KeyAscii)
        Else
            sCode = Chr$(KeyAscii)
        End If
        
        lTime = GetTickCount()
        
        nResult = SearchCode(UCase$(sCode), cvGreaterEqual)
        If nResult > NULL_RESULT Then
            mListIndex = nResult
            mHighlighted = nResult
            If mListIndex > vscItem.Max Then
                vscItem.Value = vscItem.Max
            Else
                vscItem.Value = mListIndex
            End If
            ShowItems
        End If
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    mIgnoreKeyPress = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled And (Button = vbLeftButton) Then
        If mStyle = Checkboxes Then
            If (X > mButtonRect.Left) Then
                SetDropDown True
            End If
        Else
            SetDropDown True
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled Then
        If X > UserControl.ScaleWidth Or X < 0 Or Y > UserControl.ScaleHeight Or Y < 0 Then
            ReleaseCapture
            mInCtrl = False
        ElseIf mInCtrl Then
            RaiseEvent MouseMove(Button, Shift, X, Y)
        Else
            mInCtrl = True
            Call TrackMouseLeave(UserControl.hWnd)
 
            If mBorderStyle = BorderCustom Then
                DrawComboBorder
            End If
            
            RaiseEvent MouseMove(Button, Shift, X, Y)
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled And (mBorderStyle <> BorderCustom) Then
        With UserControl
            Call DrawEdge(.hdc, mButtonRect, EDGE_RAISED, BF_RECT)
            .Picture = .Image
        End With
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mAlignment = PropBag.ReadProperty("Alignment", DEF_ALIGNMENT)
    mAutoComplete = PropBag.ReadProperty("AutoComplete", DEF_AUTOCOMPLETE)
    mBackColor = PropBag.ReadProperty("BackColor", DEF_BACKCOLOR)
    mBorderColor = PropBag.ReadProperty("BorderColor", DEF_BORDERCOLOR)
    mBorderCurve = PropBag.ReadProperty("BorderCurve", DEF_BORDERCURVE)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", DEF_BORDERSTYLE)
    mBorderWidth = PropBag.ReadProperty("BorderWidth", DEF_BORDERWIDTH)
    mButtonBackColor = PropBag.ReadProperty("ButtonBackColor", DEF_BUTTONBACKCOLOR)
    mColumnHeaders = PropBag.ReadProperty("ColumnHeaders", DEF_COLUMNHEADERS)
    mColumnResize = PropBag.ReadProperty("ColumnResize", DEF_COLUMNRESIZE)
    mColumnSort = PropBag.ReadProperty("ColumnSort", DEF_COLUMNSORT)
    mDefaultItemForeColor = PropBag.ReadProperty("DefaultItemForeColor", DEF_DEFAULTITEMFORECOLOR)
    mDisplayEllipsis = PropBag.ReadProperty("DisplayEllipsis", DEF_EDITABLE)
    mDropDownAutoWidth = PropBag.ReadProperty("DropDownAutoWidth", DEF_DROPDOWNAUTOWIDTH)
    mDropDownItemsVisible = PropBag.ReadProperty("DropDownItemsVisible", DEF_DROPDOWNITEMSVISIBLE)
    mDropDownWidth = PropBag.ReadProperty("DropDownWidth", DEF_DROPDOWNWIDTH)
    mEditable = PropBag.ReadProperty("Editable", DEF_EDITABLE)
    mEnabled = PropBag.ReadProperty("Enabled", DEF_ENABLED)
    mFocusRectColor = PropBag.ReadProperty("FocusRectColor", DEF_FOCUSRECTCOLOR)
    mFocusRectStyle = PropBag.ReadProperty("FocusRectStyle", DEF_FOCUSRECTSTYLE)
    mForeColor = PropBag.ReadProperty("ForeColor", DEF_FORECOLOR)
    mHotBorderColor = PropBag.ReadProperty("HotBorderColor", DEF_BORDERCOLOR)
    mHotButtonBackColor = PropBag.ReadProperty("HotButtonBackColor", DEF_BUTTONBACKCOLOR)
    mIntegralHeight = PropBag.ReadProperty("IntegralHeight", DEF_INTEGRALHEIGHT)
    mLocked = PropBag.ReadProperty("Locked", DEF_LOCKED)
    mMaxLength = PropBag.ReadProperty("MaxLength", 0)
    mPageScrollItems = PropBag.ReadProperty("PageScrollItems", DEF_PAGESCROLLITEMS)
    mRequireCheckedItem = PropBag.ReadProperty("RequireCheckedItem", DEF_REQUIRECHECKEDITEM)
    mRowHeightMin = PropBag.ReadProperty("RowHeightMin", DEF_ROWHEIGHTMIN)
    mScaleUnits = PropBag.ReadProperty("ScaleUnits", DEF_SCALEUNITS)
    mSearchColumn = PropBag.ReadProperty("SearchColumn", DEF_SEARCHCOLUMN)
    mStyle = PropBag.ReadProperty("Style", DEF_STYLE)
    mTextAll = PropBag.ReadProperty("TextAll", DEF_TEXTALL)
    mTextNone = PropBag.ReadProperty("TextNone", DEF_TEXTNONE)
    mTextSelection = PropBag.ReadProperty("TextSelection", DEF_TEXTSELECTION)
    
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set DropDownFont = PropBag.ReadProperty("DropDownFont", Ambient.Font)
    Cols = PropBag.ReadProperty("Cols", DEF_COLS)

    '#############################################################################################################################    'Format Controls
    With txtCombo
        .Alignment = mAlignment
        .BackColor = mBackColor
        .ForeColor = mForeColor
        .MaxLength = mMaxLength
        .Visible = mEditable
    End With
    
    picList.BackColor = mBackColor
    
    With UserControl
        .BackColor = mBackColor
        .ForeColor = mForeColor
    End With
    
    vscItem.LargeChange = mPageScrollItems
    
    '#############################################################################################################################
    'Subclassing
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        
        With UserControl.Parent
            Call Subclass_Start(.hWnd)

            Call Subclass_AddMsg(.hWnd, WM_WINDOWPOSCHANGING, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_WINDOWPOSCHANGED, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_GETMINMAXINFO, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_SIZE, MSG_AFTER)
        End With

        With UserControl
            Call Subclass_Start(.hWnd)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_KILLFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_SETFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
        End With

        With picList
            Call Subclass_Start(.hWnd)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEHOVER, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_CTLCOLORSCROLLBAR, MSG_BEFORE)
        End With
      End If
End Sub

Private Sub UserControl_Resize()
    With txtCombo
        .Left = BORDER_LEFT
        .Top = BORDER_TOP
        .Height = UserControl.ScaleHeight - (BORDER_TOP * 2)
        .Width = (UserControl.ScaleWidth - (BORDER_LEFT + BORDER_TOP)) - BUTTON_WIDTH
    End With
    
    With UserControl
        .Picture = Nothing
    End With
    
    DrawComboBorder
End Sub

Private Sub UserControl_Show()
    Dim lResult As Long
    
    'This modifies the PictureBox control so that it is not bound by
    'its Container
    'Dropdown can render over any Container the control is in
    '(such as a Frame) and is not restricted by the Forms Boundaries
    
    lResult = GetWindowLong(picList.hWnd, GWL_EXSTYLE)
    Call SetWindowLong(picList.hWnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
    Call SetWindowPos(picList.hWnd, picList.hWnd, 0, 0, 0, 0, 39)
    Call SetWindowLong(picList.hWnd, -8, Parent.hWnd)
    Call SetParent(picList.hWnd, 0)
End Sub

Private Sub UserControl_Terminate()
    On Local Error GoTo UserControl_TerminateError
    
    Call Subclass_Stop(UserControl.Parent.hWnd)
    Call Subclass_Stop(UserControl.hWnd)
    Call Subclass_Stop(picList.hWnd)
  
UserControl_TerminateError:
    Exit Sub
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
    Call PropBag.WriteProperty("DropDownFont", mDropDownFont, Ambient.Font)
    Call PropBag.WriteProperty("Cols", UBound(mCols) + 1, DEF_COLS)
    
    Call PropBag.WriteProperty("Alignment", mAlignment, DEF_ALIGNMENT)
    Call PropBag.WriteProperty("AutoComplete", mAutoComplete, DEF_AUTOCOMPLETE)
    Call PropBag.WriteProperty("BackColor", mBackColor, DEF_BACKCOLOR)
    Call PropBag.WriteProperty("BorderColor", mBorderColor, DEF_BORDERCOLOR)
    Call PropBag.WriteProperty("BorderCurve", mBorderCurve, DEF_BORDERCURVE)
    Call PropBag.WriteProperty("BorderStyle", mBorderStyle, DEF_BORDERSTYLE)
    Call PropBag.WriteProperty("BorderWidth", mBorderWidth, DEF_BORDERWIDTH)
    Call PropBag.WriteProperty("ButtonBackColor", mButtonBackColor, DEF_BUTTONBACKCOLOR)
    Call PropBag.WriteProperty("ColumnHeaders", mColumnHeaders, DEF_COLUMNHEADERS)
    Call PropBag.WriteProperty("ColumnResize", mColumnResize, DEF_COLUMNRESIZE)
    Call PropBag.WriteProperty("ColumnSort", mColumnSort, DEF_COLUMNSORT)
    Call PropBag.WriteProperty("DefaultItemForeColor", mDefaultItemForeColor, DEF_DEFAULTITEMFORECOLOR)
    Call PropBag.WriteProperty("DisplayEllipsis", mDisplayEllipsis, DEF_EDITABLE)
    Call PropBag.WriteProperty("DropDownAutoWidth", mDropDownAutoWidth, DEF_DROPDOWNAUTOWIDTH)
    Call PropBag.WriteProperty("DropDownItemsVisible", mDropDownItemsVisible, DEF_DROPDOWNITEMSVISIBLE)
    Call PropBag.WriteProperty("DropDownWidth", mDropDownWidth, DEF_DROPDOWNWIDTH)
    Call PropBag.WriteProperty("Editable", mEditable, DEF_EDITABLE)
    Call PropBag.WriteProperty("Enabled", mEnabled, DEF_ENABLED)
    Call PropBag.WriteProperty("FocusRectColor", mFocusRectColor, DEF_FOCUSRECTCOLOR)
    Call PropBag.WriteProperty("FocusRectStyle", mFocusRectStyle, DEF_FOCUSRECTSTYLE)
    Call PropBag.WriteProperty("ForeColor", mForeColor, DEF_FORECOLOR)
    Call PropBag.WriteProperty("HotBorderColor", mHotBorderColor, DEF_BORDERCOLOR)
    Call PropBag.WriteProperty("HotButtonBackColor", mHotButtonBackColor, DEF_BUTTONBACKCOLOR)
    Call PropBag.WriteProperty("IntegralHeight", mIntegralHeight, DEF_INTEGRALHEIGHT)
    Call PropBag.WriteProperty("Locked", mLocked, DEF_LOCKED)
    Call PropBag.WriteProperty("MaxLength", mMaxLength, 0)
    Call PropBag.WriteProperty("PageScrollItems", mPageScrollItems, DEF_PAGESCROLLITEMS)
    Call PropBag.WriteProperty("RequireCheckedItem", mRequireCheckedItem, DEF_REQUIRECHECKEDITEM)
    Call PropBag.WriteProperty("RowHeightMin", mRowHeightMin, DEF_ROWHEIGHTMIN)
    Call PropBag.WriteProperty("ScaleUnits", mScaleUnits, DEF_SCALEUNITS)
    Call PropBag.WriteProperty("SearchColumn", mSearchColumn, DEF_SEARCHCOLUMN)
    Call PropBag.WriteProperty("Style", mStyle, DEF_STYLE)
    Call PropBag.WriteProperty("TextAll", mTextAll, DEF_TEXTALL)
    Call PropBag.WriteProperty("TextNone", mTextNone, DEF_TEXTNONE)
    Call PropBag.WriteProperty("TextSelection", mTextSelection, DEF_TEXTSELECTION)
End Sub

Private Sub vscItem_Change()
    mScrollTick = GetTickCount()
    ShowItems
    
    RaiseEvent Scroll
End Sub

Private Sub vscItem_Scroll()
    vscItem_Change
    picList.Refresh
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get hdc() As Long
   hdc = UserControl.hdc
End Property
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(ByVal NewValue As OLE_COLOR)
    mBorderColor = NewValue
    DrawComboBorder
    
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(ByVal NewValue As Long)
    mBorderWidth = NewValue
    DrawComboBorder
    
    PropertyChanged "BorderWidth"
End Property

Public Property Get BorderCurve() As Long
    BorderCurve = mBorderCurve
End Property

Public Property Let BorderCurve(ByVal NewValue As Long)
    mBorderCurve = NewValue
    DrawComboBorder
    
    PropertyChanged "BorderCurve"
End Property

Public Property Get ButtonBackColor() As OLE_COLOR
    ButtonBackColor = mButtonBackColor
End Property

Public Property Let ButtonBackColor(ByVal NewValue As OLE_COLOR)
    mButtonBackColor = NewValue
    DrawComboBorder
    
    PropertyChanged "ButtonBackColor"
End Property

Public Property Get HotBorderColor() As OLE_COLOR
    HotBorderColor = mHotBorderColor
End Property

Public Property Let HotBorderColor(ByVal NewValue As OLE_COLOR)
    mHotBorderColor = NewValue
    DrawComboBorder
    
    PropertyChanged "HotBorderColor"
End Property

Public Property Get HotButtonBackColor() As OLE_COLOR
    HotButtonBackColor = mHotButtonBackColor
End Property

Public Property Let HotButtonBackColor(ByVal NewValue As OLE_COLOR)
    mHotButtonBackColor = NewValue
    DrawComboBorder
    
    PropertyChanged "HotButtonBackColor"
End Property

