Attribute VB_Name = "mListView"
Option Explicit
' ===========================================================================
' Filename: mListView
' Author:   Steve McMahon (steve@dogma.demon.co.uk)
' Date:     15 February 1998
'
' Requires:
'
' Description:
' Declares to support advanced ListView features.
'
' ---------------------------------------------------------------------------
' Visit vbAccelerator - the VB Programmers resource for advanced,free VB
' source code.
'     http://vbaccelerator.com
' ===========================================================================

Public Const CLR_NONE = &HFFFFFFFF

Public Const WM_NOTIFY = &H4E
Public Const LVN_FIRST = -100&
Public Const LVN_BEGINDRAG = (LVN_FIRST - 9)

Public Const LVM_FIRST = &H1000                   '// ListView messages
Public Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const LVM_SETICONSPACING = (LVM_FIRST + 53)

Public Const LVM_SETHOVERTIME = (LVM_FIRST + 71)
Public Const LVM_GETHOVERTIME = (LVM_FIRST + 72)

Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54) '// optional wParam == mask

Public Const LVS_EX_GRIDLINES = &H1
Public Const LVS_EX_SUBITEMIMAGES = &H2
Public Const LVS_EX_CHECKBOXES = &H4
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_HEADERDRAGDROP = &H10
Public Const LVS_EX_FULLROWSELECT = &H20         '// applies to report mode only
Public Const LVS_EX_ONECLICKACTIVATE = &H40
Public Const LVS_EX_TWOCLICKACTIVATE = &H80
'#if (_WIN32_IE >= =&H0400)
Public Const LVS_EX_FLATSB = &H100
Public Const LVS_EX_REGIONAL = &H200
Public Const LVS_EX_INFOTIP = &H400              '// listview does InfoTips for you
Public Const LVS_EX_UNDERLINEHOT = &H800
Public Const LVS_EX_UNDERLINECOLD = &H1000
Public Const LVS_EX_MULTIWORKAREAS = &H2000
'#endif

' Bitmaps in list views!
Type LVBKIMAGE
    ulFlags As Long
    hbm As Long
    pszImage As String
    cchImageMax As Long
    xOffsetPercent As Long
    yOffsetPercent As Long
End Type

' 4.71:
Public Const LVBKIF_SOURCE_NONE = &H0
Public Const LVBKIF_SOURCE_HBITMAP = &H1    ' Not supported
Public Const LVBKIF_SOURCE_URL = &H2
Public Const LVBKIF_SOURCE_MASK = &H3
Public Const LVBKIF_STYLE_NORMAL = &H0
Public Const LVBKIF_STYLE_TILE = &H10
Public Const LVBKIF_STYLE_MASK = &H10

Public Const LVM_SETBKIMAGEA = (LVM_FIRST + 68)
Public Const LVM_GETBKIMAGEA = (LVM_FIRST + 69)
Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEA
Public Const LVM_GETBKIMAGE = LVM_GETBKIMAGEA

' Manipulating ListView Columns
Type LVCOLUMN
    Mask As Long
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
'#if (_WIN32_IE >= 0x0300)
    iImage As Long
    iOrder As Long
'#End If
End Type

' LVCOLUMN mask values:
Public Const LVCF_FMT = &H1
Public Const LVCF_WIDTH = &H2
Public Const LVCF_TEXT = &H4
Public Const LVCF_SUBITEM = &H8
'#if (_WIN32_IE >= =&H0300)
Public Const LVCF_IMAGE = &H10
Public Const LVCF_ORDER = &H20
'#End If

' LVCOLUMN fmt values:
Public Const LVCFMT_LEFT = &H0
Public Const LVCFMT_RIGHT = &H1
Public Const LVCFMT_CENTER = &H2
Public Const LVCFMT_JUSTIFYMASK = &H3
'#if (_WIN32_IE >= =&H0300)
Public Const LVCFMT_IMAGE = &H800
Public Const LVCFMT_BITMAP_ON_RIGHT = &H1000
Public Const LVCFMT_COL_HAS_IMAGES = &H8000
'#End If

Public Const LVM_GETCOLUMNA = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMN = LVM_GETCOLUMNA
Public Const LVM_SETCOLUMNA = (LVM_FIRST + 26)
Public Const LVM_SETCOLUMN = LVM_SETCOLUMNA

Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)

' Manipulating ListView items:
Type LVITEM
    Mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
'#if (_WIN32_IE >= 0x0300)
    iIndent As Long
'#End If
End Type

' LVITEM mask values:
Public Const LVIF_TEXT = &H1
Public Const LVIF_IMAGE = &H2
Public Const LVIF_PARAM = &H4
Public Const LVIF_STATE = &H8
'#if (_WIN32_IE >= =&H0300)
Public Const LVIF_INDENT = &H10
Public Const LVIF_NORECOMPUTE = &H800
'#End If

Public Const LVM_GETITEMA = (LVM_FIRST + 5)
Public Const LVM_GETITEM = LVM_GETITEMA
Public Const LVM_SETITEMA = (LVM_FIRST + 6)
Public Const LVM_SETITEM = LVM_SETITEMA

' Check boxes:
Public Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)

Public Const LVIS_FOCUSED = &H1&
Public Const LVIS_SELECTED = &H2&
Public Const LVIS_CUT = &H4&
Public Const LVIS_DROPHILITED = &H8&
Public Const LVIS_ACTIVATING = &H20&

Public Const LVIS_OVERLAYMASK = &HF00&
Public Const LVIS_STATEIMAGEMASK = &HF000&

' Finding:
'Public Type POINT
'  X As Long
'  Y As Long
'End Type
   
Public Type LVFINDINFO
  flags As Long
  psz As String
  lParam As Long
  PT As POINTAPI
  vkDirection As Long
End Type

Private Const LVFI_PARAM = 1
Public Const LVFI_STRING = &H2
Public Const LVFI_PARTIAL = &H8
Public Const LVFI_WRAP = &H20
Public Const LVFI_NEARESTXY = &H40

Private Const LVM_FINDITEM = LVM_FIRST + 13
Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
Public Const LVM_SORTITEMS = LVM_FIRST + 48

Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)

Public Const LVNI_ALL = &H0
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2
Public Const LVNI_CUT = &H4
Public Const LVNI_DROPHILITED = &H8

Public Const LVNI_ABOVE = &H100
Public Const LVNI_BELOW = &H200
Public Const LVNI_TOLEFT = &H400
Public Const LVNI_TORIGHT = &H800

' Header control styles
Public Const HDS_HOTTRACK = &H4 ' v 4.70
Public Const HDS_BUTTONS = &H2

' Message functions:
'Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

'Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' Show window styles
Public Const SW_SHOWNORMAL = 1
Public Const SW_ERASE = &H4
Public Const SW_HIDE = 0
Public Const SW_INVALIDATE = &H2
Public Const SW_MAX = 10
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_OTHERUNZOOM = 4
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTCLOSING = 1
Public Const SW_RESTORE = 9
Public Const SW_PARENTOPENING = 3
Public Const SW_SHOW = 5
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
'Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

' Window style bit functions:
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
    ) As Long
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long _
    ) As Long
' Window Long indexes:
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)

'Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_HIDEWINDOW = &H80
'Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
'Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
'Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
'Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Public Const HWND_NOTOPMOST = -2

'Public Declare Function CoInitialize Lib "OLE32.DLL" (ByVal pvReserved As Long) As Long
'Public Declare Sub CoUninitialize Lib "OLE32.DLL" ()
Public Const NOERROR = &H0&
Public Const S_OK = &H0&
Public Const S_FALSE = &H1&

'Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
'Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
'Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Const TRANSPARENT = 1

Public m_iSortCol As Long
Public m_eSortOrder As MSComctlLib.ListSortOrderConstants
Public Enum ELVSortTypes
    elvstText
    elvstInteger
    elvstFloat
    elvstDate
End Enum
Public m_eSortType As ELVSortTypes

Public Function LVWSortCompare( _
        ByVal lParam1 As Long, _
        ByVal lParam2 As Long, _
        ByVal hwnd As Long _
    ) As Long
Dim lI1 As Long, lI2 As Long
Dim v1 As Variant, v2 As Variant, vt As Variant

    lI1 = GetListViewIndexForlParam(hwnd, lParam1)
    lI2 = GetListViewIndexForlParam(hwnd, lParam2)
    
    'Compare the items
    'Return -ve if lI1<lI2,
    '       0 if lI1 = lI2
    '       +ve if lI1 > lI2
    On Error Resume Next
    Select Case m_eSortType
    Case elvstInteger
        v1 = CLng(GetListViewItem(hwnd, lI1))
        v2 = CLng(GetListViewItem(hwnd, lI2))
    Case elvstFloat
        v1 = CDbl(GetListViewItem(hwnd, lI1))
        v2 = CDbl(GetListViewItem(hwnd, lI2))
    Case elvstDate
        v1 = CDate(GetListViewItem(hwnd, lI1))
        v2 = CDate(GetListViewItem(hwnd, lI2))
    Case elvstText
        v1 = GetListViewItem(hwnd, lI1)
        v2 = GetListViewItem(hwnd, lI2)
    End Select
        
    If (m_eSortOrder = lvwDescending) Then
        vt = v2
        v2 = v1
        v1 = vt
    End If
    If (v1 < v2) Then
        LVWSortCompare = -1
    ElseIf (v1 = v2) Then
        LVWSortCompare = 0
    Else
        LVWSortCompare = 1
    End If
    
End Function
Private Function GetListViewIndexForlParam(ByVal hwnd As Long, ByVal lParam As Long)
Dim tLV As LVFINDINFO

    ' Convert the input parameter to an index in the list view
    tLV.flags = LVFI_PARAM
    tLV.lParam = lParam
    GetListViewIndexForlParam = SendMessage(hwnd, LVM_FINDITEM, -1, tLV)
    
End Function
Private Function GetListViewItem(ByVal hwnd As Long, ByVal lIndex As Long) As String
Dim tLV As LVITEM
Dim lLength As Long

    tLV.Mask = LVIF_TEXT
    tLV.iSubItem = m_iSortCol - 1
    tLV.pszText = String$(32, Chr$(0))
    tLV.cchTextMax = 32
    lLength = SendMessage(hwnd, LVM_GETITEMTEXT, lIndex, tLV)
    If lLength > 0 Then
      GetListViewItem = Left$(tLV.pszText, lLength)
    End If
   
End Function
