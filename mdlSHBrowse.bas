Attribute VB_Name = "mdlSHBrowse"
Option Explicit

Type SHITEMID   ' mkid
    cb As Long       ' Size of the ID (including cb itself)
    abID() As Byte  ' The item ID (variable length)
End Type

Type ITEMIDLIST   ' idl
    mkid As SHITEMID
End Type

'Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pIdl As Long, ByVal pszPath As String) As Long

'Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                              (ByVal hWndOwner As Long, ByVal nFolder As Long, _
                              pIdl As ITEMIDLIST) As Long

Public Const NOERROR = 0
Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5   ' (Documents folder)
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8   ' (Recent folder)
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15   ' (ShellNew folder)

'Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal pv As Long)
'Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long ' ITEMIDLIST
Public Type BROWSEINFO   ' bi
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000

'Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, _
                                                           ByVal X As Long, _
                                                           ByVal Y As Long, _
                                                           ByVal hIcon As Long) As Boolean

'Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
                                                              ByVal xLeft As Long, _
                                                              ByVal yTop As Long, _
                                                              ByVal hIcon As Long, _
                                                              ByVal cxWidth As Long, _
                                                              ByVal cyWidth As Long, _
                                                              ByVal istepIfAniCur As Long, _
                                                              ByVal hbrFlickerFreeDraw As Long, _
                                                              ByVal diFlags As Long) As Boolean
' DrawIconEx() diFlags values:
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

'Declare Function SHGetFileInfo Lib "Shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As Long) As Long


Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800

Public Const MAX_PATH = 260

Type SHFILEINFO   ' shfi
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Const SHGFI_LARGEICON = &H0&
Public Const SHGFI_SMALLICON = &H1&
Public Const SHGFI_OPENICON = &H2&
Public Const SHGFI_SHELLICONSIZE = &H4&
Public Const SHGFI_PIDL = &H8&
Public Const SHGFI_USEFILEATTRIBUTES = &H10&
Public Const SHGFI_ICON = &H100&
Public Const SHGFI_DISPLAYNAME = &H200&
Public Const SHGFI_TYPENAME = &H400&
Public Const SHGFI_ATTRIBUTES = &H800&
Public Const SHGFI_ICONLOCATION = &H1000&
Public Const SHGFI_EXETYPE = &H2000&
Public Const SHGFI_SYSICONINDEX = &H4000&
Public Const SHGFI_LINKOVERLAY = &H8000&
Public Const SHGFI_SELECTED = &H10000

