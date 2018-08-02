Attribute VB_Name = "mdlGeral"
Private Const WM_QUIT As Long = &H12


Public bDBInternet As Boolean
Public LoginDSN As String
Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 1
Public Const MIB_IF_TYPE_TOKENRING = 2
Public Const MIB_IF_TYPE_FDDI = 3
Public Const MIB_IF_TYPE_PPP = 4
Public Const MIB_IF_TYPE_LOOPBACK = 5
Public Const MIB_IF_TYPE_SLIP = 6
Public Const REG_SZ = 1    'Constant for a string variable type.
Public Const HKEY_LOCAL_MACHINE = &H80000002
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Private Const LVS_EX_DOUBLEBUFFER = &H10000
Private Const LVS_EX_BORDERSELECT = &H8000

Public Enum SeqEndereco
    Imobiliario = 0
    Mobiliario = 1
    cidadao = 2
End Enum

Public Enum TipoEndereco
    Localizacao = 0
    Entrega = 1
    cadastrocidadao = 2
End Enum


Public Enum Direct
     chomp_left = 0
     chomp_righT = 1
End Enum

Public Enum elvSearch
    elvSearchText = 1
    elvSearchSub = 2
    elvSearchTag = 4
End Enum

Public Const MIB_DB = "07/10/1997"
Public Const UL = "gtisys"
Public Const UP = "everest"

Private Type RELMOB1
    nCodReduz As Long
    sRazao As String
    nAno As Integer
    sTx1 As String
    sTx2 As String
    sTx3 As String
    sVs1 As String
    sVs2 As String
    sVs3 As String
    sVs4 As String
End Type

'processos
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Type TSTARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Public Type IP_ADDR_STRING
        Next As Long
        IpAddress As String * 16
        IpMask As String * 16
        Context As Long
End Type

Public Type IP_ADAPTER_INFO
        Next As Long
        ComboIndex As Long
        AdapterName As String * MAX_ADAPTER_NAME_LENGTH
        Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
        AddressLength As Long
        Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
        Index As Long
        Type As Long
        DhcpEnabled As Long
        CurrentIpAddress As Long
        IpAddressList As IP_ADDR_STRING
        GatewayList As IP_ADDR_STRING
        DhcpServer As IP_ADDR_STRING
        HaveWins As Boolean
        PrimaryWinsServer As IP_ADDR_STRING
        SecondaryWinsServer As IP_ADDR_STRING
        LeaseObtained As Long
        LeaseExpires As Long
End Type
Public Type FIXED_INFO
        HostName As String * MAX_HOSTNAME_LEN
        DomainName As String * MAX_DOMAIN_NAME_LEN
        CurrentDnsServer As Long
        DnsServerList As IP_ADDR_STRING
        NodeType As Long
        ScopeId  As String * MAX_SCOPE_ID_LEN
        EnableRouting As Long
        EnableProxy As Long
        EnableDns As Long
End Type

Public Type Bairro
    Codigo As Integer
    Nome As String
End Type


Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal _
uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As _
Long) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function GetDesktopWindow Lib "user32" () As Long


Declare Function CreateProcess Lib "KERNEL32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As String, lpStartupInfo As TSTARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function WaitForSingleObjectEx Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetExitCodeProcess Lib "KERNEL32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Const CREATE_NEW_CONSOLE = &H10
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const STILL_ACTIVE = &H103
Public Const INFINITE = &HFFFFFFFF
Public Const STATUS_TIMEOUT = &H102
Public Const STARTF_USESHOWWINDOW = &H1
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_MINIMIZE = 6
Public Const LB_ADDSTRING = &H180
Public Const LB_SETITEMDATA = &H19A
Public Const NovoProtocolo = 1
Public ProcessInformation(1 To 10) As PROCESS_INFORMATION
'Declaração de Variaveis

Public sDataFormat As String
Public IPServer As String
Public bCloseChat As Boolean
Public nGuiche As Integer
Public aDocDAM() As Long
Public bLocal As Boolean
Public bAnistia As Boolean
Public bComercioEletronico As Boolean
Public dcJuros As New clsDictionary
Public dcUfir As New clsDictionary
Public aMulta() As Multa
Public dcFeriado As New Dictionary
Public bSkin As Boolean
Public sPathBin As String
Public ArqBinImg As String
Public ArqBinImgTmp As String
Public NomeCidade As String
Public NomeBaseDados As String
Public NomeDoComputador As String
Public NomeDoUsuario As String
Public NomeDeLogin As String
Public PwdDeLogin As String
Public sWd As String
Public en As rdoEnvironment
Public cn As New rdoConnection
Public cnInt As New rdoConnection
Public cnGti As New rdoConnection
Public enInt As rdoEnvironment
Public cnEicon As New rdoConnection
Public cnEicon2 As New rdoConnection
Public enEicon As rdoEnvironment
Public enBkp As rdoEnvironment
Public en2 As rdoEnvironment
Public cn2 As New rdoConnection
Public cnBkp As New rdoConnection
Public nCodLastUser As Integer
Public LastUser As String
Public UserPwd As String
Public sPathAutoUpdate As String
Public sPathArqBanco As String
Public sPathArqDA As String
Public sPrintBottom As String
Public sPathHelp As String
Public sPathImage As String
Public sPathFoto As String
Public sPathReport As String
Public MBI_LG As String
Public MI As Boolean         'usado para informar a DAM sobre Multa de Infração
Public CodCidadao As Long     'Usado no Cadastro de Cidadão
Public CodImovel As String       'Usado no Cadastro de Imóvel
Public CodEmpresa As String       'Usado no Cadastro de Empresa
Public CodCond As Integer        'Usado na Selecao de Condominio
Public NomeCond As String       'Usado na Selecao de Condominio
Public CodProcesso As Long
Public AnoProcesso As Integer
Public CodProcessoCP As Long
Public AnoProcessoCP As Integer
Public NumeroProcesso As String
Public NumRegAtend As Long
Public AnoRegAtend As Integer
Public CodRural As Double      'Usado no Cadastro Rural
Public sItemEdit As String       'Usado na Edição de Imovel
Public sForm As String              'Usado na Consulta de Imovel
Public sFormMob As String           'Usado na Consulta de Empresa
Public sFormFoto As String      'Usado na exibição da Foto
Public sParamForm As String     'Usado nas telas de parametros
Public aArrayVazio()                 'Se o report nao tiver formulas passadas
Public FiltroE As Integer 'Usado no Form Filtro de Parcela
Public FiltroE2 As Integer 'Usado no Form Filtro de Parcela
Public FiltroS As Integer 'Usado no Form Filtro de Parcela
Public FiltroL As Integer 'Usado no Form Filtro de Parcela
Public FiltroLP As String 'Usado no Form Filtro de Parcela
Public FiltroD As String 'Usado no Form Filtro de Parcela
Public FiltroA As String 'Usado no Form Filtro de Parcela
Public FiltroSEQ As Integer 'Usado no Form Filtro de Parcela
Public bFiltroSEQ As Boolean 'Usado no Form Filtro de Parcela
Public dDataAtualiza As Date 'usado no extrato
Public nSeqFator As Integer
Public nSeqFator2 As Integer
Public nIndexFind As Long 'usado na busca de item em matrizes(carregaiss)
Public NewSec As Boolean
Public SecId As String
Public bBoleto As Boolean
Public nMargem_Top As Integer
Public nMargem_Left As Integer
Public nMargem_Bottom As Integer
Public nMargem_Right As Integer

Private cle(17) As Long
Private x1a0(9) As Long
Private x1a2 As Long
Public InDebug As Boolean
Public aPlanoDesconto() As PlanoDesconto
Private inter As Long, res As Long, ax As Long, bx As Long
Private cx As Long, dX As Long, si As Long, tmp As Long
Global lpPrevWndProc As Long
Global gHW As Long

'Classes
Public gtiObj As gtiProc.Tmuna
Dim gtiObjR As New gtiProc.Registry

'Constantes
'Public oSQLServer As SQLDMO.SQLServer
Public Const pRegMain = "HKEY_LOCAL_MACHINE\Software\GTI"
Public Const pRegPath = pRegMain & "\Path"
Public Const Cinza = vbButtonFace
Public Const Branco = vbWhite
Public Const Kde = &HEEEEEE
Public Const Azul2000 = &HCC9966
Public Const AzulClaro = &HF8EBC2
Public Const Vinho = &H80&
Public Const Tzahov = &HE4FEFC
Public Const AmareloClaro = &HC0FFFF
Public Const CinzaCaixao = &H808080
Public Const CinzaEscuro = &H404040
Public Const Marrom = &H80&
Public Const Roxo = &H800080
Public Const VerdeEscuro = &H8000&
Public Const VerdeAccess = &H808000
Public Const Preto = &H80000012
Public Const LightBlue = &HFFC0C0
Public Const TPM_LEFTALIGN = &H0&
Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)
Public Const CB_SETCURSEL = &H14E
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOZORDER = &H4
Public Const GWL_WNDPROC = -4
'Public Const GWL_STYLE = (-16)
Public Const GRADIENT_FILL_RECT_V  As Long = &H1
Public Const GRADIENT_FILL_RECT_H As Long = &H0
Public Const LF_FACESIZE = 32
Public Const WS_CAPTION = &HC00000
Public Const WM_PAINT = &HF
Public Const WM_TIMER = &H113
Public Const WM_MOUSEMOVE = &H200
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_COMMAND = &H111
Public Const WM_CLOSE = &H10

Public Const QS_HOTKEY = &H80
Public Const QS_KEY = &H1
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_MOUSEMOVE = &H2
Public Const QS_PAINT = &H20
Public Const QS_POSTMESSAGE = &H8
Public Const QS_SENDMESSAGE = &H40
Public Const QS_TIMER = &H10
Public Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Public Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Public Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

Public Const FLASHW_STOP = 0
Public Const FLASHW_CAPTION = 1
Public Const FLASHW_TRAY = 2
Public Const FLASHW_ALL = FLASHW_CAPTION Or FLASHW_TRAY
Public Const FLASHW_TIMER = 4
Public Const FLASHW_TIMEROFG = 12

'Tipos de Dados
Type POINT
  x As Long
  Y As Long
End Type
Type POINTAPI
        x As Long
        Y As Long
End Type
Type TRegiao
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
    PosMes As Integer
End Type
Type Botao
    Regiao As Long
    Numero As Integer
    Ativo As Boolean
End Type
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type LOGFONT
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
'     lfFaceName As String * 32
   End Type
Type NMHDR
   hwndFrom As Long ' Window handle of control sending message
   idfrom As Long ' Identifier of control sending message
   code As Long ' Specifies the notification code
End Type
Type TRIVERTEX
    x As Long
    Y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Type GRADIENT_RECT
    UPPERLEFT As Long  'In reality this is a UNSIGNED Long
    LOWERRIGHT As Long 'In reality this is a UNSIGNED Long
End Type
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
Type MSGBOXPARAMS
  cbSize As Long
  hwndOwner As Long
  hInstance As Long
  lpszText As String
  lpszCaption As String
  dwStyle As Long
  lpszIcon As String
  dwContextHelpId As Long
  lpfnMsgBoxCallback As Long
  dwLanguageId As Long
End Type
Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type
Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type
Type HH_REG_VALUES
  pszFileName     As String
  pszFilePath     As String
End Type

Public Type FLASHWINFO
    cbSize As Long
    hwnd As Long
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type

Public Type PlanoDesconto
    nMin As Integer
    nMax As Integer
    nValor As Double
End Type

Public Type Multa
    nAno As Integer
    nMin As Integer
    nMax As Integer
    nValor As Double
End Type

Private Type AliquotaISS
    dData As Date
    nAliquota As Double
End Type



'Enum
Public Enum Elg
    logon = 1
    LogOff = 2
    Form = 3
    Configuração = 4
End Enum
Public Enum EvFrm
    Nenhum = 0
    Inclusão = 1
    Alteração = 2
    Exclusão = 3
    Impressão = 4
End Enum
Public Enum sbAlign
   sbAlignLeft = 1
   sbAlignRight = 2
End Enum
Public Enum sbFillStyle
   sbFilled = 1
   sbSmooth = 2
End Enum

'Declaração de Api

Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)
Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Declare Function CoInitialize Lib "OLE32.DLL" (ByVal pvReserved As Long) As Long
Declare Sub CoUninitialize Lib "OLE32.DLL" ()
Declare Sub CopyMem Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal pv As Long)
Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Boolean
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "KERNEL32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Declare Function FlashWindowEx Lib "user32" (FWInfo As FLASHWINFO) As Boolean
Declare Function GetAdaptersInfo Lib "IPHlpApi.dll" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&) As Long
Declare Sub GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetLongPathName Lib "KERNEL32" (ByRef pszShortPath As String, ByRef lpszLongPath As String, ByVal cchBuffer As Long) As Long
Declare Function GetNetworkParams Lib "IpHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetVersionExA Lib "KERNEL32" (lpVersionInformation As OSVERSIONINFO) As Integer
Declare Function GetVolumeInformation Lib "KERNEL32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long                               ' ITEMIDLIST
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Sub Hook()
   'Establish a hook to capture messages to this window
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
     Dim TEMP As Long
     'Reset the message handler for this window
     TEMP = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim nmh As NMHDR

   Select Case uMsg
   Case WM_NOTIFY
      ' Fill the NMHDR struct from the lParam pointer.
      ' (for any WM_NOTIFY msg, lParam always points to a struct which is
      ' either the NMHDR struct, or whose 1st member is the NMHDR struct)
      Call MoveMemory(nmh, ByVal lParam, Len(nmh))

      Select Case nmh.code
      Case LVN_BEGINDRAG
         'Notifies a list view control's
         'parent window that a drag-and-drop
         'operation involving the left
         'mouse button is being initiated.
         WindowProc = 1
         Exit Function
      End Select
   End Select

   'Pass message on to the original window message handler
   WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)

End Function
Sub Main()
On Error GoTo Erro
Dim uwid As Long, uhgt As Long
Dim FS As FileSystemObject, sIP As String, Data1 As String, Data2 As String
Dim sPathOrigem As String, sPathDest As String, fo As File, fso As Folder
Dim handle As Long
If App.PrevInstance Then
     MsgBox "Já existe uma cópia deste programa rodando.", vbCritical, "ATENÇÃO"
     End
End If

uwid = Screen.Width
uhgt = Screen.Height
uwid = uwid / 2 / 7.5
uhgt = uhgt / 2 / 7.5
MBI_LG = Chr(67) & Chr(86) & Chr(82) & Chr(67) & Chr(83) & Chr(48) & Chr(52) & Chr(49)

'Inicialização do Sistema
NomeDoComputador = RetornaNomeDoComputador
NomeDoUsuario = GetUser
dDataAtualiza = Now


Data1 = "": Data2 = ""
Set FS = New FileSystemObject
sPathOrigem = "\\192.168.200.130\atualizagti\"
sPathDest = App.Path & "\"
If FS.FileExists(sPathOrigem & "\GTI_SERVER.EXE") Then
    Set fso = FS.GetFolder(sPathOrigem)
    For Each fo In fso.Files
        If UCase$(fo.Name) = "GTI_SERVER.EXE" Then
            Data1 = fo.DateLastModified
            Exit For
        End If
    Next
End If
If FS.FileExists(sPathDest & "\GTI_SERVER.EXE") Then
    Set fso = FS.GetFolder(sPathDest)
    For Each fo In fso.Files
        If UCase$(fo.Name) = "GTI_SERVER.EXE" Then
            Data2 = fo.DateLastModified
            Exit For
        End If
    Next
End If
If Data2 = "" Then
    Data2 = "01/01/1980"
End If
If CDate(Data1) > CDate(Data2) Then
    FS.CopyFile sPathOrigem & "\GTI_SERVER.EXE", sPathDest & "\GTI_SERVER.EXE", True
    If Err.Number = 70 Then
        handle = FindWindow("GTI_SERVER.EXE", vbNullString)
        If handle Then
            PostMessage handle, WM_QUIT, 0&, 0&
        End If
        FS.CopyFile sPathOrigem & "\GTI_SERVER.EXE", sPathDest & "\GTI_SERVER.EXE", True
    End If
End If

Set FS = Nothing

'If NomeDeLogin = "LEONARDO.DINIZ" Or NomeDeLogin = "LEANDRO" Or NomeDeLogin = "LUIZH" Then
'   Shell sPathDest & "\GTI_SERVER.exe", vbNormalFocus
'End If



'CheckDSN
Inicio:
LoadReg

CarregaPlanoDesconto

frmMdi.show

Exit Sub
Erro:
Resume Next

End Sub

Private Sub CarregaPlanoDesconto()
ReDim aPlanoDesconto(10)
aPlanoDesconto(0).nMin = 1
aPlanoDesconto(0).nMax = 48
aPlanoDesconto(0).nValor = 0
aPlanoDesconto(1).nMin = 1
aPlanoDesconto(1).nMax = 12
aPlanoDesconto(1).nValor = 80
aPlanoDesconto(2).nMin = 13
aPlanoDesconto(2).nMax = 24
aPlanoDesconto(2).nValor = 60
aPlanoDesconto(3).nMin = 25
aPlanoDesconto(3).nMax = 36
aPlanoDesconto(3).nValor = 40
aPlanoDesconto(4).nMin = 37
aPlanoDesconto(4).nMax = 48
aPlanoDesconto(4).nValor = 20
aPlanoDesconto(5).nMin = 1
aPlanoDesconto(5).nMax = 12
aPlanoDesconto(5).nValor = 90
aPlanoDesconto(6).nMin = 13
aPlanoDesconto(6).nMax = 24
aPlanoDesconto(6).nValor = 80
aPlanoDesconto(7).nMin = 25
aPlanoDesconto(7).nMax = 36
aPlanoDesconto(7).nValor = 70
aPlanoDesconto(8).nMin = 37
aPlanoDesconto(8).nMax = 48
aPlanoDesconto(8).nValor = 60
aPlanoDesconto(9).nMin = 49
aPlanoDesconto(9).nMax = 120
aPlanoDesconto(9).nValor = 50



End Sub

Public Function TipoDePlano(nNumParcela) As Integer
Dim x As Integer

For x = 1 To 4
    If nNumParcela >= aPlanoDesconto(x).nMin And nNumParcela <= aPlanoDesconto(x).nMax Then
        TipoDePlano = x
    End If
Next

End Function

Sub CheckDSN()
Dim DriverODBC As String
Dim NameDSN As String

DriverODBC = String(255, Chr(32))
'NOME DO DSN
NameDSN = "odbcTributacao"

'VERIFICA SE OS DRIVERS DO SQL ESTAO INSTALADOS
If Not gtiObjR.xCheckSqlDriver(DriverODBC) Then
    MsgBox "VOCE DEVE INSTALAR O SQLServer ODBC Drivers ANTES DE USAR ESTE PROGRAMA.", vbOKOnly + vbCritical
    MsgBox "PROGRAMA ENCERRADO.", vbOKOnly + vbCritical
    End
End If

'JA EXISTE DSN?
If (gtiObjR.xSQLDSNWanted(NameDSN)) = True Then
'        ***** "ESTE DSN JA FOI CRIADO." ****
    Else
        If Not gtiObjR.xMakeSQLDSN(DriverODBC, NameDSN) Then
            MsgBox "NÃO FOI POSSIVEL CRIAR O DSN NESTA MAQUINA.", vbOKOnly + vbCritical
        End If
End If

End Sub

Public Function Security() As Boolean
Dim RdoAux As rdoResultset, Sql As String
Dim RdoAux2 As rdoResultset

Dim volume_name As String
Dim serial_number As Long
Dim max_component_length As Long
Dim file_system_flags As Long
Dim file_system_name As String
Dim sKey As String
Security = False

If GetVolumeInformation("C:\", volume_name, _
    Len(volume_name), serial_number, _
    max_component_length, file_system_flags, _
    file_system_name, Len(file_system_name)) = 0 _
Then
    MsgBox "No Disk In Drive!", vbInformation, "Error Reading Disk"
    Security = False
    Exit Function
End If

Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='" & Left$(NomeDoComputador, 2) & serial_number & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    If .RowCount = 0 Then
       Security = False
       Exit Function
    Else
       Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='CODDATAANT'"
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
       sKey = KeyGen(Left$(NomeDoComputador, 2) & serial_number, RdoAux2!valparam, 3)
       If Decrypt128(!valparam, "sysTribut") <> sKey Then
          Security = False
          Exit Function
       End If
    End If
End With
Security = True

End Function

Public Sub LoadReg()

'**********************


'**********************

'sPathArqDA = GetSetting("GTI", "PATH", "DEBITOAUTOMATICO")
'If sPathArqDA = "" Then
'   SaveSetting "GTI", "PATH", "DEBITOAUTOMATICO", "\\192.168.200.130\ATUALIZAGTI\DEBITOAUTOMATICO"
'   sPathArqDA = GetSetting("GTI", "PATH", "DEBITOAUTOMATICO")
'End If
'If bLocal Then
'    sPathArqBanco = "C:\Trabalho\GTI\Bancos"
'Elses
    
    sPathArqDA = "\\192.168.200.130\ATUALIZAGTI\DEBITOAUTOMATICO"
    If NomeDoComputador = "SKYNET" Then
        sPathArqBanco = "D:\TRABALHO\GTI\BANCO"
    Else
        sPathArqBanco = "\\192.168.200.130\ATUALIZAGTI"
    End If
'End If

'sPathArqBanco = GetSetting("GTI", "PATH", "ARQUIVOBANCO")
'If sPathArqBanco = "" Then
'   SaveSetting "GTI", "PATH", "ARQUIVOBANCO", "\\APPSERVER\TESTES IPTU\ARQUIVOBANCO"
'   sPathArqBanco = GetSetting("GTI", "PATH", "ARQUIVOBANCO")
'End If
sPathAutoUpdate = GetSetting("GTI", "PATH", "AUTOUPDATE")
If sPathAutoUpdate = "" Then
   SaveSetting "GTI", "PATH", "AUTOUPDATE", "\\192.168.200.130\TESTES IPTU"
   sPathAutoUpdate = GetSetting("GTI", "PATH", "AUTOUPDATE")
End If
sPathHelp = GetSetting("GTI", "PATH", "HELP")
If sPathHelp = "" Then
   SaveSetting "GTI", "PATH", "HELP", App.Path & "\HELP"
   sPathHelp = GetSetting("GTI", "PATH", "HELP")
End If
sPathImage = GetSetting("GTI", "PATH", "IMAGE")
If sPathImage = "" Then
   SaveSetting "GTI", "PATH", "IMAGE", App.Path & "\IMAGE"
   sPathImage = GetSetting("GTI", "PATH", "IMAGE")
End If
sPathFoto = GetSetting("GTI", "PATH", "FOTO")
'sPathFoto = "\\APPSERVER\RECADDOC\QUADRAS_E_FOTOS"
'If sPathFoto = "" Then
   SaveSetting "GTI", "PATH", "FOTO", "\\192.168.200.130\FOTOSGTI"
   sPathFoto = GetSetting("GTI", "PATH", "FOTO") '
'End If
' Mudei aqui o Report
sPathReport = App.Path & "\REPORT"
'sPathReport = GetSetting("GTI", "PATH", "REPORT")
'If sPathReport = "" Then
'   SaveSetting "GTI", "PATH", "REPORT", App.Path & "\REPORT"
'   sPathReport = GetSetting("GTI", "PATH", "REPORT")
'End If
sPathBin = GetSetting("GTI", "PATH", "BIN")
If sPathBin = "" Then
   SaveSetting "GTI", "PATH", "BIN", App.Path & "\BIN"
   sPathBin = GetSetting("GTI", "PATH", "BIN")
End If
sPathBin = App.Path & "\BIN"

sPrintBottom = GetSetting("GTI", "PRINT", "BOTTOM")
If sPrintBottom = "" Then
   SaveSetting "GTI", "PRINT", "BOTTOM", "N"
   sPrintBottom = GetSetting("GTI", "PRINT", "BOTTOM")
End If

If InDebug Then
    ArqBinImg = "c:\tmp\TbImgBlb.bin"
    ArqBinImgTmp = "c:\tmp\TbImgTmp"
Else
    ArqBinImg = sPathBin & "\TbImgBlb.bin"
    ArqBinImgTmp = sPathBin & "\TbImgTmp"
End If
End Sub

Private Function RetornaNomeDoComputador() As String

'Retorna o Nome do Computador
Dim kk As Long, TmpName As String, tmpCompName As String * 200, x As Integer
kk = GetComputerName(tmpCompName, 200)
tmpCompName = Trim$(tmpCompName)
tmpCompName = Left$(tmpCompName, Len(tmpCompName) - 1)

For x = 1 To Len(tmpCompName)
   If Asc(Mid(tmpCompName, x, 1)) <> 0 Then
      TmpName = TmpName & Mid(tmpCompName, x, 1)
   Else
      tmpCompName = TmpName
      Exit For
   End If
Next
RetornaNomeDoComputador = Trim$(TmpName)

End Function

Private Function GetUser() As String
Dim TmpName As String, tmpUserName As String * 200, x As Integer
    
'Retorna o Nome do Usuario
Dim lpUserID As String
Dim nBuffer As Long
Dim ret As Long
lpUserID = String(25, 0)
nBuffer = 25
ret = GetUserName(lpUserID, nBuffer)

tmpUserName = Trim$(lpUserID$)
tmpUserName = Left$(tmpUserName, Len(tmpUserName) - 1)

For x = 1 To Len(tmpUserName)
   If Asc(Mid(tmpUserName, x, 1)) <> 0 Then
      TmpName = TmpName & Mid(tmpUserName, x, 1)
   Else
      tmpUserName = TmpName
      Exit For
   End If
Next
GetUser$ = Trim$(tmpUserName)

End Function

Public Function Conecta(User As String, Pwd As String, Optional sParam As String) As Boolean
Dim FS As FileSystemObject
On Error GoTo Erro


If sParam = "-T" Then
    LoginDSN = "odbcTribTeste"
ElseIf sParam = "-L" Then
    LoginDSN = "odbcTribLocal"
ElseIf bDBInternet Then
    LoginDSN = "odbcTribInternet"
Else
   LoginDSN = "odbcTributacao"
End If

If Trim$(User) = "" Then
     Conecta = False
     Exit Function
End If

Screen.MousePointer = vbHourglass

'    Set cn = en.OpenConnection(dsname:=LoginDSN, _
        Prompt:=rdDriverNoPrompt, _
        Connect:="uid=" & User & ";PWD=" & Pwd & ";driver={SQL Server};")


Set en = rdoEngine.rdoEnvironments(0)
en.CursorDriver = rdUseOdbc
With en
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 60
     
     
sIP = ""
Set FS = New FileSystemObject
If FS.FileExists(App.Path & "\gti.ini") Then
    Open App.Path & "\gti.ini" For Input As #1
    Do While Not EOF(1)
        Line Input #1, strLinha
        If Left(strLinha, 6) = "SERVER" Then
            sIP = Mid(strLinha, 8, Len(strLinha) - 7)
        End If
    Loop
    Close #1
    If sIP = "" Then
        Open App.Path & "\gti.ini" For Append As #1
        Print #1, "SERVER=192.168.15.160"
        Close #1
    End If
Else
    Open App.Path & "\gti.ini" For Output As #1
    Print #1, "SERVER=192.168.15.160"
    sIP = "192.168.15.160"
    Close #1
End If

IPServer = sIP

    
     
     
If LoginDSN = "odbcTributacao" Then
   Conn$ = "UID=gtisys;PWD=everest;" _
    & "DATABASE=tributacao;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cn = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)
ElseIf LoginDSN = "odbcTribTeste" Then
   Conn$ = "UID=gtisys;PWD=everest;" _
    & "DATABASE=TributacaoTeste;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cn = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)
ElseIf LoginDSN = "odbcTribInternet" Then
   Conn$ = "UID=gtisys;PWD=everest;" _
    & "DATABASE=tributacao;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cn = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)
ElseIf LoginDSN = "odbcTribLocal" Then
    If NomeDoComputador = "GTI-PC" Then
        Conn$ = "UID=gtisys;PWD=everest;" _
        & "DATABASE=tributacaoBKP;" _
        & "SERVER=" & IPServer & ";" _
        & "DRIVER={SQL SERVER};DSN='';"
    Else
        Conn$ = "UID=gtisys;PWD=everest;" _
        & "DATABASE=tributacao;" _
        & "SERVER=" & IPServer & ";" _
        & "DRIVER={SQL SERVER};DSN='';"
    End If
    Set cn = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)
End If

    '**** CONEXÃO COM O POSTGRESQL ****
'     LoginDSN = "GTIPG"
'     Set cn = en.OpenConnection(dsname:=LoginDSN, _
'        Prompt:=rdDriverNoPrompt, _
'        Connect:="Driver={PostgreSQL};SERVER=localhost;DATABASE=teste;UID=postgres;PWD=aranja")
'     Set cn = en.OpenConnection(dsname:=LoginDSN, _
        Prompt:=rdDriverNoPrompt, _
        Connect:="Driver={PostgreSQL30};SERVER=200.168.187.43;DATABASE=Tributacao;UID=pgsql;PWD=kcb2psf@,thor.7gw")
         
     
End With

Conecta = True
cn.QueryTimeout = 180
Screen.MousePointer = vbDefault
Exit Function
Erro:
Liberado
MsgBox "Não é possivel conectar em " & IPServer, vbCritical, "error"
Conecta = False

End Function

Public Function ConectaIntegrativa() As Boolean

On Error GoTo Erro

Screen.MousePointer = vbHourglass

Set enInt = rdoEngine.rdoEnvironments(0)
enInt.CursorDriver = rdUseOdbc
With enInt
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 20
     
    Conn$ = "UID=integrativa;PWD=integrativa;" _
    & "DATABASE=GTI_INTEGRATIVA;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cnInt = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)

End With

ConectaIntegrativa = True
cnInt.QueryTimeout = 180
Screen.MousePointer = vbDefault
Exit Function
Erro:
'MsgBox Err.Description
ConectaIntegrativa = False

End Function

Public Function ConectaGTI() As Boolean

On Error GoTo Erro

Screen.MousePointer = vbHourglass

Set enInt = rdoEngine.rdoEnvironments(0)
enInt.CursorDriver = rdUseOdbc
With enInt
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 20
     
    Conn$ = "UID=gtisys;PWD=everest;" _
    & "DATABASE=GTI;" _
    & "SERVER=GTI-PC\sqlexpress;" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cnGti = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)

End With

ConectaGTI = True
cnInt.QueryTimeout = 180
Screen.MousePointer = vbDefault
Exit Function
Erro:
'MsgBox Err.Description
ConectaGTI = False

End Function



Public Function ConectaBkp() As Boolean

On Error GoTo Erro

Screen.MousePointer = vbHourglass

Set enBkp = rdoEngine.rdoEnvironments(0)
enBkp.CursorDriver = rdUseOdbc
With enBkp
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 20
     
    Conn$ = "UID=" & UL & ";PWD=" & UP & ";" _
    & "DATABASE=TRIBUTACAOTESTE;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cnBkp = enBkp.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)

End With

ConectaBkp = True
cnBkp.QueryTimeout = 180
Screen.MousePointer = vbDefault
Exit Function
Erro:
'MsgBox Err.Description
ConectaBkp = False

End Function

Public Function ConectaEicon() As Boolean

On Error GoTo Erro

Screen.MousePointer = vbHourglass

Set enEicon = rdoEngine.rdoEnvironments(0)
enEicon.CursorDriver = rdUseOdbc
With enEicon
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 20
     
    Conn$ = "UID=" & UL & ";PWD=" & UP & ";" _
    & "DATABASE=GTI_EICON;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cnEicon = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)

End With

ConectaEicon = True
cnEicon.QueryTimeout = 180
Screen.MousePointer = vbDefault
Exit Function
Erro:
'MsgBox Err.Description
ConectaEicon = False

End Function

Public Function ConectaEicon2() As Boolean

On Error GoTo Erro

Screen.MousePointer = vbHourglass

Set enEicon = rdoEngine.rdoEnvironments(0)
enEicon.CursorDriver = rdUseOdbc
With enEicon
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 20
     
    Conn$ = "UID=" & UL & ";PWD=" & UP & ";" _
    & "DATABASE=GTI_EICON;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cnEicon2 = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)

End With

ConectaEicon2 = True
cnEicon2.QueryTimeout = 180
Screen.MousePointer = vbDefault
Exit Function
Erro:
'MsgBox Err.Description
ConectaEicon2 = False

End Function


Public Function SubNull(dado)

'Substitui Nulo por Vazio
On Error Resume Next
If IsNull(dado) Then
    SubNull = ""
Else
    SubNull = dado
End If

End Function

Function Mask(texto As String)

'Substitui aspas em string Sql
Dim x As Long, Letra As String, NovaString As String
NovaString = ""
For x = 1 To Len(texto)
   Letra = Mid(texto, x, 1)
   If Asc(Letra) = 39 Then
      Letra = Chr(39) & Chr(39)
   End If
   NovaString = NovaString & Letra
Next
Mask = NovaString

End Function

Function Virg2Ponto(sValor As String) As String

'Troca Virgula por Ponto
Dim i As Integer
Do
    i = InStr(sValor, ",")

    If (i <> 0) Then
        Mid(sValor, i, 1) = "."
    End If
Loop Until (i = 0)

Virg2Ponto = sValor

End Function

Function Ponto2Virg(sValor As String) As String

'Troca Ponto por Virgula
Dim i As Integer
Do
    i = InStr(sValor, ".")

    If (i <> 0) Then
        Mid(sValor, i, 1) = ","
    End If

Loop Until (i = 0)

Ponto2Virg = sValor

End Function

Function RemovePonto(sValor As String) As String

Dim x As Integer, Letra As String, NovaString As String
NovaString = ""
For x = 1 To Len(sValor)
   Letra = Mid(sValor, x, 1)
   If Letra = "." Then
      Letra = ""
   End If
   NovaString = NovaString & Letra
Next
RemovePonto = NovaString

End Function

Function RemoveSpace(sValor As String) As String

Dim x As Integer, Letra As String, NovaString As String
NovaString = ""
For x = 1 To Len(sValor)
   Letra = Mid(sValor, x, 1)
   If Letra = " " Then
      Letra = "_"
    ElseIf Letra = "/" Or Letra = "." Then
      Letra = ""
   End If
   NovaString = NovaString & Letra
Next
RemoveSpace = NovaString

End Function

Function removeAcentos(ByVal texto As String) As String
    Dim vPos As Byte
    
    Const vComAcento = "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜàáâãäåçèéêëìíîïòóôõöùúûü"
    Const vSemAcento = "AAAAAACEEEEIIIIOOOOOUUUUaaaaaaceeeeiiiiooooouuuu"
    
    For i = 1 To Len(texto)
        vPos = InStr(1, vComAcento, Mid(texto, i, 1))
        If vPos > 0 Then
           Mid(texto, i, 1) = Mid(vSemAcento, vPos, 1)
        End If
    Next
    removeAcentos = texto
End Function


Public Sub LimpaMascara(Controle As esMaskedEdit)

'Limpa MaskEdit
Dim OldMask As String
OldMask = Controle.Mask
Controle.Mask = ""
Controle.Text = ""
Controle.Mask = OldMask

End Sub

Public Sub Log(nEvento As Elg, sNomeForm As String, nSecEvento As EvFrm, Desc As String)
Dim MaxCod As Long

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM LOGEVENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

'Sql = "INSERT LOGEVENTO (SEQ,DATAHORAEVENTO,COMPUTADOR,USUARIO,FORM,EVENTO,SECEVENTO,LOGEVENTO) VALUES("
'Sql = Sql & MaxCod & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & NomeDoComputador & "','" & NomeDeLogin & "','"
'Sql = Sql & Left$(sNomeForm, 30) & "'," & nEvento & "," & nSecEvento & ",'" & Left$(Mask(Desc), 500) & "')"
Sql = "INSERT LOGEVENTO (SEQ,DATAHORAEVENTO,COMPUTADOR,USERID,FORM,EVENTO,SECEVENTO,LOGEVENTO) VALUES("
Sql = Sql & MaxCod & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & NomeDoComputador & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
Sql = Sql & Left$(sNomeForm, 30) & "'," & nEvento & "," & nSecEvento & ",'" & Left$(Mask(Desc), 500) & "')"
'cn.Execute Sql, rdExecDirect

End Sub

Public Function RetornaDVCodReduzido(CodImovel As Long) As String

Dim sFromN As String        'Converte Num to String
Dim nTotPosAtual As Long    'Total da Posição
Dim nTotalGeral As Long     'Total da Somatoria
Dim sTotal As String        'Total em String
Dim nDV As Integer          'Digito verificador

nTotPosAtual = 0

sFromN = Format(CodImovel, "0000000")

nTotPosAtual = Val(Mid(sFromN, 1, 1)) * 8
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 2, 1)) * 7)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 3, 1)) * 6)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 4, 1)) * 5)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 5, 1)) * 4)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 6, 1)) * 3)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 7, 1)) * 2)

nTotalGeral = nTotPosAtual Mod 11
If nTotalGeral = 11 Then
   nDV = 1
ElseIf nTotalGeral = 10 Then
   nDV = 0
Else
   sTotal = Format(nTotalGeral, "0000000")
   nDV = Val(Right$(sTotal, 1))
End If

RetornaDVCodReduzido = CStr(nDV)

End Function

Public Function Encrypt128(ByVal Plaintext As String, ByVal Key As String) As String
    Dim sData As String
    
    si = 0
    x1a2 = 0
    i = 0
    
    For fois = 1 To 16
        cle(fois) = 0
    Next fois
    
    champ1 = Key
    lngchamp1 = Len(champ1)
    
    For fois = 1 To lngchamp1
        cle(fois) = Asc(Mid(champ1, fois, 1))
    Next fois
    
    champ1 = Plaintext
    lngchamp1 = Len(champ1)
    For fois = 1 To lngchamp1
        c = Asc(Mid(champ1, fois, 1))
        
        Assemble128
        
        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256
        
        For compte = 1 To 16
        
            cle(compte) = cle(compte) Xor c
        
        Next compte
        
        c = c Xor (cfc Xor cfd)
        
        d = (((c / 16) * 16) - (c Mod 16)) / 16
        e = c Mod 16
        
        sData = sData & Chr$(&H61 + d) ' d+&h61 give one letter range from a to p for the 4 high bits of c
        sData = sData & Chr$(&H61 + e) ' e+&h61 give one letter range from a to p for the 4 low bits of c
        
    Next fois
    Encrypt128 = sData
End Function

Public Function Decrypt128(ByVal Text As String, ByVal Key As String) As String
    Dim sData As String
    si = 0
    x1a2 = 0
    i = 0
    
    For fois = 1 To 16
        cle(fois) = 0
    Next fois
    
    champ1 = Key
    lngchamp1 = Len(champ1)
    
    For fois = 1 To lngchamp1
    cle(fois) = Asc(Mid(champ1, fois, 1))
    Next fois
    
    champ1 = Text
    lngchamp1 = Len(champ1)
    
    For fois = 1 To lngchamp1
    
        d = Asc(Mid(champ1, fois, 1))
        If (d - &H61) >= 0 Then
            d = d - &H61  ' to transform the letter to the 4 high bits of c
            If (d >= 0) And (d <= 15) Then
                d = d * 16
            End If
        End If
        If (fois <> lngchamp1) Then
            fois = fois + 1
        End If
        e = Asc(Mid(champ1, fois, 1))
        If (e - &H61) >= 0 Then
            e = e - &H61 ' to transform the letter to the 4 low bits of c
            If (e >= 0) And (e <= 15) Then
                c = d + e
            End If
        End If
        
        Assemble128
        
        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256
        
        c = c Xor (cfc Xor cfd)
        
        For compte = 1 To 16
        
            cle(compte) = cle(compte) Xor c
        
        Next compte
        
        sData = sData & Chr$(c)
    
    Next fois
    Decrypt128 = sData
End Function

Private Sub Assemble128()
    
    x1a0(0) = ((cle(1) * 256) + cle(2)) Mod 65536
    code128
    inter = res
    
    x1a0(1) = x1a0(0) Xor ((cle(3) * 256) + cle(4))
    code128
    inter = inter Xor res
    
    x1a0(2) = x1a0(1) Xor ((cle(5) * 256) + cle(6))
    code128
    inter = inter Xor res
    
    x1a0(3) = x1a0(2) Xor ((cle(7) * 256) + cle(8))
    code128
    inter = inter Xor res
    
    x1a0(4) = x1a0(3) Xor ((cle(9) * 256) + cle(10))
    code128
    inter = inter Xor res
    
    x1a0(5) = x1a0(4) Xor ((cle(11) * 256) + cle(12))
    code128
    inter = inter Xor res
    
    x1a0(6) = x1a0(5) Xor ((cle(13) * 256) + cle(14))
    code128
    inter = inter Xor res
    
    x1a0(7) = x1a0(6) Xor ((cle(15) * 256) + cle(16))
    code128
    inter = inter Xor res
    
    i = 0

End Sub

Private Sub code128()
    dX = (x1a2 + i) Mod 65536
    ax = x1a0(i)
    cx = &H15A
    bx = &H4E35
    
    tmp = ax
    ax = si
    si = tmp
    
    tmp = ax
    ax = dX
    dX = tmp
    
    If (ax <> 0) Then
        ax = (ax * bx) Mod 65536
    End If
    
    tmp = ax
    ax = cx
    cx = tmp
    
    If (ax <> 0) Then
        ax = (ax * si) Mod 65536
        cx = (ax + cx) Mod 65536
    End If
    
    tmp = ax
    ax = si
    si = tmp
    ax = (ax * bx) Mod 65536
    dX = (cx + dX) Mod 65536
    
    ax = ax + 1
    
    x1a2 = dX
    x1a0(i) = ax
    
    res = ax Xor dX
    i = i + 1

End Sub

Function ValidaCGC(cgc As String) As Integer
        Dim a, j, i, d1, d2
        
        If Len(cgc) = 0 Then
            ValidaCGC = True
            Exit Function
        End If
        
        If Len(cgc) = 8 And Val(cgc) > 0 Then
           a = 0
           j = 0
           d1 = 0
           For i = 1 To 7
               a = Val(Mid(cgc, i, 1))
               If (i Mod 2) <> 0 Then
                  a = a * 2
               End If
               If a > 9 Then
                  j = j + Int(a / 10) + (a Mod 10)
               Else
                  j = j + a
               End If
           Next i
           d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
           If d1 = Val(Mid(cgc, 8, 1)) Then
              ValidaCGC = True
           Else
              ValidaCGC = False
           End If
        Else
           If Len(cgc) = 14 And Val(cgc) > 0 Then
              a = 0
              i = 0
              d1 = 0
              d2 = 0
              j = 5
              For i = 1 To 12 Step 1
                  a = a + (Val(Mid(cgc, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
              a = a Mod 11
              d1 = IIf(a > 1, 11 - a, 0)
              a = 0
              i = 0
              j = 6
              For i = 1 To 13 Step 1
                  a = a + (Val(Mid(cgc, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
              a = a Mod 11
              d2 = IIf(a > 1, 11 - a, 0)
              If (d1 = Val(Mid(cgc, 13, 1)) And d2 = Val(Mid(cgc, 14, 1))) Then
                 ValidaCGC = True
              Else
                 ValidaCGC = False
              End If
           Else
              ValidaCGC = False
           End If
        End If
End Function

Public Function ValidaCPF(CPF As String) As Integer
    
    
    Dim soma As Integer
    Dim Resto As Integer
    Dim i As Integer
    
    'Valida argumento
    If Len(CPF) = 0 Then
         ValidaCPF = False
         Exit Function
    End If
    If Len(CPF) <> 11 Then
        ValidaCPF = False
        Exit Function
    End If
        
    If (CPF = "11111111111") Then
        ValidaCPF = False
        Exit Function
    End If
        
    soma = 0
    For i = 1 To 9
        soma = soma + Val(Mid$(CPF, i, 1)) * (11 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 10, 1)) Then
        ValidaCPF = False
        Exit Function
    End If
        
    soma = 0
    For i = 1 To 10
        soma = soma + Val(Mid$(CPF, i, 1)) * (12 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 11, 1)) Then
        ValidaCPF = False
        Exit Function
    End If
    
    ValidaCPF = True

End Function

Public Sub Ocupado()
frmMdi.imOK.Picture = frmMdi.imStatus(2)
frmMdi.imWorking.Picture = frmMdi.imStatus(1)
frmMdi.Picture2.Refresh
Screen.MousePointer = vbHourglass
End Sub

Public Sub Liberado()
frmMdi.imOK.Picture = frmMdi.imStatus(0)
frmMdi.imWorking.Picture = frmMdi.imStatus(2)
frmMdi.Picture2.Refresh
Screen.MousePointer = vbDefault
End Sub

Public Function RetornaBairro(Cep As String) As Bairro
Dim RdoAux As rdoResultset, Sql As String

Sql = "SELECT cep.codbairro, bairro.descbairro FROM cep INNER JOIN bairro ON cep.codbairro = bairro.codbairro "
Sql = Sql & "WHERE (bairro.siglauf = 'SP') AND (bairro.codcidade = 413) AND (cep.cep = '" & Cep & "')"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        RetornaBairro.Codigo = !CodBairro
        RetornaBairro.Nome = !DescBairro
    Else
        RetornaBairro.Codigo = 0
        RetornaBairro.Nome = ""
    End If
   .Close
End With

End Function

Public Function RetornaCEP(Logradouro As Long, Numero As Integer) As String
Dim Sql As String
Dim RdoS As rdoResultset
Dim nConta As Integer
Dim Impar As Integer, Par As Integer
Dim Num1 As Long, Num2 As Long
Dim sCep As String
If Numero Mod 2 = 0 Then
     Par = 1
     Impar = 0
Else
     Par = 0
     Impar = 1
End If
sCep = ""
Sql = "SELECT CEP,VALOR1,VALOR2,IMPAR,PAR FROM CEP "
Sql = Sql & "WHERE CODLOGR=" & Logradouro
Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoS
       nConta = .RowCount
       If nConta = 0 Then
            RetornaCEP = "14870-000"
       ElseIf nConta = 1 Then
            RetornaCEP = Left$(!Cep, 5) & "-" & Right$(!Cep, 3)
       ElseIf nConta > 1 Then
            Do Until .EOF
                  Num1 = !VALOR1
                  Num2 = Val(SubNull(!VALOR2))
                  If Numero >= Num1 And Numero <= Num2 Then
                       If Impar = 1 And !Impar = True Then
                            sCep = !Cep
                            Exit Do
                       Else
                            If Par = 1 And !Par = True Then
                                 sCep = !Cep
                                 Exit Do
                            End If
                       End If
                  ElseIf Numero >= Num1 And Num2 = 0 Then
                       If Impar = 1 And !Impar = True Then
                            sCep = !Cep
                            Exit Do
                       Else
                            If Par = 1 And !Par = True Then
                                 sCep = !Cep
                                 Exit Do
                            End If
                       End If
                  End If
                 .MoveNext
            Loop
            If sCep = "" Then sCep = "     -   "
            RetornaCEP = Left$(sCep, 5) & "-" & Right$(sCep, 3)
       End If
End With

End Function

'Public Function RetornaIdUsuario(Username As String) As Integer
'Dim RdoAux As rdoResultset
'
'Sql = "SELECT IDUSUARIO FROM SEG_USUARIO WHERE NOMEUSUARIO='" & Username & "'"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'If RdoAux.RowCount > 0 Then
'   RetornaIdUsuario = RdoAux!IDUSUARIO
'Else
'   RetornaIdUsuario = 0
'End If
'RdoAux.Close
'
'End Function

Public Function RetEventUserForm(sNomeForm As String) As String

Dim sRetorno As String
Dim RdoAux As rdoResultset, Sql As String
Sql = "SELECT DISTINCT SEG_USERACESS.CODEVENTO "
Sql = Sql & "FROM SEG_USERACESS INNER JOIN  SEG_TELASISTEMA ON "
Sql = Sql & "SEG_USERACESS.CODTELA = SEG_TELASISTEMA.CODTELA  Inner Join SEG_EVENTO ON "
Sql = Sql & "SEG_USERACESS.CODEVENTO = SEG_EVENTO.CODEVENTO WHERE "
Sql = Sql & "SEG_USERACESS.NOMEUSUARIO='" & LastUser & "' AND "
Sql = Sql & "SEG_TELASISTEMA.NOMEFORM='" & sNomeForm & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
       Do Until .EOF
            sRetorno = sRetorno & "#" & Format(!CODEVENTO, "000")
           .MoveNext
       Loop
      .Close
End With
RetEventUserForm = sRetorno

End Function

Public Function ValidaFeriado(sDataFeriado As String) As Integer
Dim RdoAux As rdoResultset, Sql As String
Dim dDataFeriado As Date
On Error Resume Next
ValidaFeriado = 0

If Not IsDate(sDataFeriado) Then
   ValidaFeriado = 4
   Exit Function
End If

dDataFeriado = CDate(sDataFeriado)

If Weekday(dDataFeriado) = 1 Then 'Domingo
    ValidaFeriado = 1
    Exit Function
ElseIf Weekday(dDataFeriado) = 7 Then 'Sabado
    ValidaFeriado = 2
    Exit Function
End If

If dcFeriado.Exists(Format(dDataFeriado, "dd/mm/yyyy")) Then
    ValidaFeriado = 3
End If

End Function

Public Function RetornaDiaUtil(dDataFeriado As Date) As Date

Inicio:
If ValidaFeriado(Format(dDataFeriado, "dd/mm/yyyy")) = 0 Then
    RetornaDiaUtil = Format(dDataFeriado, "dd/mm/yyyy")
    Exit Function
ElseIf ValidaFeriado(Format(dDataFeriado, "dd/mm/yyyy")) = 1 Then
    RetornaDiaUtil = Format(dDataFeriado + 1, "dd/mm/yyyy")
    Exit Function
ElseIf ValidaFeriado(Format(dDataFeriado, "dd/mm/yyyy")) = 2 Then
    RetornaDiaUtil = Format(dDataFeriado + 2, "dd/mm/yyyy")
    Exit Function
ElseIf ValidaFeriado(Format(dDataFeriado, "dd/mm/yyyy")) = 3 Then
    dDataFeriado = dDataFeriado + 1
    GoTo Inicio
Else
    dDataFeriado = dDataFeriado - 1
    GoTo Inicio
End If

End Function

Public Function RetornaDVNumDoc(NumDoc As Long) As String

Dim sFromN As String        'Converte Num to String
Dim nTotPosAtual As Long    'Total da Posição

nTotPosAtual = 0

sFromN = Format(NumDoc, "00000000")

nTotPosAtual = Val(Mid(sFromN, 1, 1)) * 7
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 2, 1)) * 3)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 3, 1)) * 9)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 4, 1)) * 7)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 5, 1)) * 3)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 6, 1)) * 9)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 7, 1)) * 7)
nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 8, 1)) * 3)

RetornaDVNumDoc = Right$(CStr(nTotPosAtual), 1)

End Function

Public Function CalculaJuros(nValorDebito As Double, dDataVencto As Date, Optional dDataNow As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String
Dim sDataVencto As String, nDia As Integer, nMes As Integer, nAno As Integer

If dDataNow = "00:00:00" Then
    dDataNow = Now
End If

'SE O VENCIMENTO FOR MAIOR OU IGUAL A DATA ATUAL, NÃO EXISTE JUROS
If dDataVencto >= dDataNow Then
    CalculaJuros = 0
    Exit Function
End If

'SE ESTIVER NO MESMO MES E ANO QUE A DATA ATUAL, NAO EXISTE JUROS
If Month(dDataVencto) = Month(dDataNow) And Year(dDataVencto) = Year(dDataNow) Then
    CalculaJuros = 0
    Exit Function
End If

If Not dcJuros.Exists(Year(dDataNow)) Then
'   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJuros = 1
   'Exit Function
End If

'MONTA O NOVO VENCIMENTO A PARTIR DO DIA 1 DO MES SUBSEQUENTE
nDia = Day(dDataVencto)
nMes = Month(dDataVencto)
nAno = Year(dDataVencto)
nDia = 1
If nMes = 12 Then
    nMes = 1
    nAno = nAno + 1
Else
    nMes = nMes + 1
End If

'sDataVencto = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
'dDataVencto = Format(sDataVencto, "dd/mm/yyyy")
''nNumMes = Int(DateDiff("d", dDataVencto, dDataNow) / 30) + 1
'nNumMes = DateDiff("m", dDataVencto, dDataNow) + 1

'If Month(dDataVencto) = Month(dDataNow) And Year(dDataVencto) = Year(dDataNow) Then
'    nNumMes = 1
'Else
'    sDataVencto = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
'    dDataVencto = Format(sDataVencto, "dd/mm/yyyy")
    'nNumMes = Int(DateDiff("d", dDataVencto, dDataNow) / 30) + 1
    nNumMes = DateDiff("m", dDataVencto, dDataNow)
'End If

nValorPerc = dcJuros.Item(Year(dDataNow))
nValorPerc = nValorPerc / 100

CalculaJuros = nValorDebito * nValorPerc * nNumMes
If CalculaJuros > 0 Then
   CalculaJuros = FormatNumber(CalculaJuros, 3)
End If

End Function

Public Function CalculaMulta(nValorDebito As Double, dDataVencto As Date, Optional dDataNow As Date, Optional bDA As Boolean) As Double
Dim nNumDia As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String
If dDataNow = "00:00:00" Then
    dDataNow = Now
End If

If dDataVencto >= dDataNow Then
    CalculaMulta = 0
    Exit Function
End If
On Error Resume Next
nNumDia = Abs(DateDiff("d", dDataNow, dDataVencto))

If nNumDia = 0 Then
   CalculaMulta = 0
   Exit Function
End If

'bDA = False
'If Year(dDataVencto) >= 2007 And bDA Then
'    nValorPerc = 20
'Else
    For x = 1 To UBound(aMulta)
        If aMulta(x).nAno = Year(dDataVencto) Then
            If nNumDia >= aMulta(x).nMin And nNumDia <= aMulta(x).nMax Then
                nValorPerc = aMulta(x).nValor
                Exit For
            ElseIf nNumDia >= aMulta(x).nMin And aMulta(x).nMax = 0 Then
                nValorPerc = aMulta(x).nValor
                Exit For
            End If
        End If
    Next
'End If

nValorPerc = nValorPerc / 100
CalculaMulta = Round(nValorDebito, 2) * nValorPerc
If CalculaMulta > 0 Then
   CalculaMulta = FormatNumber(CalculaMulta, 2)
End If

End Function

Public Function CalculaCorrecao(nValorDebito As Double, dDataBase As Date, Optional dDataNow As Date) As Double

Dim RdoAux As rdoResultset, Sql As String
Dim UfirAtual As Double
Dim UfirBase As Double

If dDataNow = "00:00:00" Then
    dDataNow = Now
End If

If Year(dDataBase) > Year(dDataNow) Then
    CalculaCorrecao = 0
    Exit Function
End If

UfirAtual = RetornaUFIR(Year(dDataNow))
If UfirAtual = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano " & Year(dDataNow), vbCritical, "Alerta !!!"
    CalculaCorrecao = 0
    Exit Function
End If

UfirBase = RetornaUFIR(Year(dDataBase))
If UfirBase = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano base.", vbCritical, "Alerta !!!"
    CalculaCorrecao = 0
    Exit Function
End If

CalculaCorrecao = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecao > 0 Then
   CalculaCorrecao = FormatNumber(CalculaCorrecao, 2)
End If

End Function

Function KeyGen(kNamev As Variant, kPass As String, kType As Integer) As String
'****************************************************************************
'* KeyGen v2.01 Build 01                                                    *
'* Copyright © 2000 W.G.Griffiths                                           *
'*                                                                          *
'* Url: http://www.webdreams.org.uk                                         *
'* E-Mail: w.g.griffiths@telinco.co.uk                                      *
'*                                                                          *
'* kNamev = Any text String, Object, String()                               *
'* kPass = Developer Password as String                                     *
'*                                                                          *
'* kType = 1  Numeric Key                                                   *
'* ktype = 2  Alphanumeric Key                                              *
'* kType = 3  Hex Key                                                       *
'*                                                                          *
'* This function returns a Software Key for a given                         *
'* name and password                                                        *
'*                                                                          *
'* NOTE: Watch www.webdreams.org.uk over the next few months....            *
'****************************************************************************

On Error Resume Next         'still here just as a precaution

Dim cTable(512) As Integer   'character map
Dim nKeys(16) As Integer     'xor keys used for pArray(x) xor nkeys(x)
Dim s0(512) As Integer       'swap-box data used to map character table
Dim nArray(16) As Integer    'name array data
Dim pArray(16) As Integer    'password array data
Dim n As Integer             'for next loop counter
Dim nPtr As Integer          'name pointer (used for counting)
Dim cPtr As Integer          'character pointer (used for counting)
Dim cFlip As Boolean         'character flip (used to flip between numeric and alpha)
Dim sIni As Integer          'holds s-box values
Dim TEMP As Integer          'holds s-box values
Dim Rtn As Integer           'holds generated key values used agains chr map
Dim gKey As String           'generated key as string
Dim nLen As Integer          'number of chr's in name
Dim pLen As Integer          'number of chr's in password
Dim kPtr As Integer          'key pointer
Dim sPtr As Integer          'space pointer (used in hex key)
Dim nOffset As Integer       'name offset
Dim pOffset As Integer       'password offset
Dim tOffset As Integer       'total offset
Dim KeySize As Integer       'the size of the key to make

Const nXor As Integer = 18   'name xor value
Const pXor As Integer = 25   'password xor value
Const cLw As Integer = 65    'character lower limit 65 = A ** do not change **
Const nLw As Integer = 48    'number lower limit 48 = 0 ** do not change **
Const sOffset As Integer = 0 'character map offset

'****************************************************************************
'Thanks to Chris Fournier for his suggestions for adding support for        *
'Strings, Objects and String() as arrays                                    *
'Your comments please                                                       *
'****************************************************************************
Dim VarType As String
Dim kName As String
Dim AryCtl As Integer
Dim AryCtrl As Control

VarType = TypeName(kNamev)

Select Case VarType
    Case "String"
        kName = kNamev
    Case "TextBox"
        kName = kNamev.Text
    Case "Object"
        For Each AryCtrl In kNamev
            If AryCtrl.Text <> "" Then
                kName = kName & AryCtrl.Text & "|"
            End If
        Next
        kName = Left$(kName, Len(kName) - 1)
    Case "String()"
        For AryCtl = LBound(kNamev) To UBound(kNamev)
            If kNamev(AryCtl) <> "" Then
                kName = kName & kNamev(AryCtl) & "|"
            End If
        Next
        kName = Left$(kName, Len(kName) - 1)
        Case Else
            MsgBox VarType & " is an unsupported type to be passed to KeyGen"
End Select
'****************************************************************************

nLen = Len(kName)
pLen = Len(kPass)

'password xor keys ** change to make keygen unique **
nKeys(1) = 46
nKeys(2) = 89
nKeys(3) = 142
nKeys(4) = 63
nKeys(5) = 231
nKeys(6) = 32
nKeys(7) = 129
nKeys(8) = 51
nKeys(9) = 28
nKeys(10) = 97
nKeys(11) = 248
nKeys(12) = 41
nKeys(13) = 136
nKeys(14) = 53
nKeys(15) = 78
nKeys(16) = 164

sIni = 0

'set s boxes
For n = 0 To 512
    s0(n) = n
Next n

For n = 0 To 512
    sIni = (sOffset + sIni + n) Mod 256
    TEMP = s0(n)
    s0(n) = s0(sIni)
    s0(sIni) = TEMP
Next n

If kType = 1 Then       '(numeric)
    
    nPtr = 0
    KeySize = 16
    gKey = String(16, " ")
    
    For n = 0 To 512
        cTable(s0(n)) = (nLw + (nPtr))
        nPtr = nPtr + 1
        If nPtr = 10 Then nPtr = 0
    Next n
    
    

ElseIf kType = 2 Then   '(alphanumeric)
    
    nPtr = 0
    cPtr = 0
    KeySize = 16
    gKey = String(16, " ")
    
    cFlip = False
    For n = 0 To 512
        If cFlip Then
            cTable(s0(n)) = (nLw + nPtr)
            nPtr = nPtr + 1
            If nPtr = 10 Then nPtr = 0
            cFlip = False
        Else
            cTable(s0(n)) = (cLw + cPtr)
            cPtr = cPtr + 1
            If cPtr = 26 Then cPtr = 0
            cFlip = True
        End If
    Next n
    
Else  '(hex)

    KeySize = 8
    gKey = String(19, " ")
    
End If

kPtr = 1

For n = 1 To nLen 'name
  nArray(kPtr) = nArray(kPtr) + Asc(Mid(kName, n, 1)) Xor nXor
  nOffset = nOffset + nArray(kPtr)
  kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n

For n = 1 To pLen 'password
  pArray(kPtr) = pArray(kPtr) + Asc(Mid(kPass, n, 1)) Xor pXor
  pOffset = pOffset + pArray(kPtr)
  kPtr = kPtr + 1
    If kPtr = 9 Then kPtr = 1
Next n

tOffset = (nOffset + pOffset) Mod 512

kPtr = 1
sPtr = 1
For n = 1 To KeySize
  pArray(n) = pArray(n) Xor nKeys(n)
  Rtn = Abs(((nArray(n) Xor pArray(n)) Mod 512) - tOffset)
  
  If kType = 3 Then 'hex key
        If Rtn < 16 Then
            Mid(gKey, kPtr, 2) = "0" & Hex(Rtn)
        Else
            Mid(gKey, kPtr, 2) = Hex(Rtn)
        End If
            If sPtr = 2 And kPtr < 18 Then
                kPtr = kPtr + 1
                Mid(gKey, kPtr + 1, 1) = "-"
            End If
        kPtr = kPtr + 2
        sPtr = sPtr + 1
        If sPtr = 3 Then sPtr = 1
  Else  'numeric - alphanumeric key
    Mid(gKey, n, 1) = Chr(cTable(Rtn))
  End If
Next

KeyGen = gKey

End Function

Public Sub Centraliza(sfrm As Form)
sfrm.Left = frmMdi.ScaleWidth / 2 - sfrm.Width / 2
sfrm.Top = frmMdi.ScaleHeight / 2 - sfrm.Height / 2
End Sub

Public Function RetornaNumero(sString As String) As String
Dim x As Integer
Dim sLetra As String, sNewString As String

For x = 1 To Len(sString)
    sLetra = Mid(sString, x, 1)
    If IsNumeric(sLetra) Then
        sNewString = sNewString & sLetra
    End If
Next
RetornaNumero = sNewString
End Function

Public Function LongToUShort(Unsigned As Long) As Integer
    LongToUShort = CInt(Unsigned - &H10000)
End Function

Public Function Gera2of5Str(sDado As String) As String

Dim CurrentChar As Integer
DataToEncode = sDado
If (Len(DataToEncode) Mod 2) = 1 Then DataToEncode = "0" & DataToEncode
StartCode = Chr(203)
StopCode = Chr(204)
StringLeng = Len(DataToEncode)
For i = 1 To StringLeng Step 2
    CurrentChar = (Mid(DataToEncode, i, 2))
    If CurrentChar < 94 Then DataToPrint = DataToPrint & Chr(CurrentChar + 33)
    If CurrentChar > 93 Then DataToPrint = DataToPrint & Chr(CurrentChar + 103)
Next i
Gera2of5Str = StartCode + DataToPrint + StopCode

End Function

Public Function Gera2of5Cod(sValorParc As String, dDataVencimento As Date, nNumDocumento As Long, nCodReduz As Long) As String

Dim sDv0 As String
Dim sBloco1 As String, sBloco2 As String, sBloco3 As String, sBloco4 As String
Dim sValorParcela As String, sAno As String, sMes As String, sDia As String, sNumDoc As String
Dim sCodReduz As String
Dim c As Integer

sValorParc = FormatNumber(sValorParc, 2)
For c = 1 To Len(sValorParc)
      If Mid(sValorParc, c, 1) <> "," Then
         sValorParcela = sValorParcela & Mid(sValorParc, c, 1)
      End If
Next

sValorParcela = Format(sValorParcela, "00000000000")
sDia = Format(Day(dDataVencimento), "00")
sMes = Format(Month(dDataVencimento), "00")
sAno = Format(Year(dDataVencimento), "00")
sNumDoc = Format(nNumDocumento, "000000000") 'VERIFICAR
sCodReduz = Format(nCodReduz, "00000000")

sBloco1 = "816" & Left$(sValorParcela, 7)
sBloco2 = Right$(sValorParcela, 4) & "2177" & Left$(sAno, 3)
sBloco3 = Right$(sAno, 1) & sMes & sDia & Left$(sNumDoc, 6)
sBloco4 = Right$(sNumDoc, 3) & sCodReduz

sDv0 = CStr(RetornaDV2of5(sBloco1 & sBloco2 & sBloco3 & sBloco4))
sBloco1 = Left$(sBloco1, 3) & sDv0 & Right$(sBloco1, 7)

sBloco1 = sBloco1 & "-" & CStr(RetornaDV2of5(sBloco1))
sBloco2 = sBloco2 & "-" & CStr(RetornaDV2of5(sBloco2))
sBloco3 = sBloco3 & "-" & CStr(RetornaDV2of5(sBloco3))
sBloco4 = sBloco4 & "-" & CStr(RetornaDV2of5(sBloco4))

Gera2of5Cod = sBloco1 & sBloco2 & sBloco3 & sBloco4

End Function

Public Function RetornaDV2of5(sBloco As String) As Integer
Dim c As Integer
Dim d As Integer
Dim e As String
Dim nSoma As Integer
Dim nResto As Integer

For c = Len(sBloco) To 1 Step -1
      If c Mod 2 = 1 Then
         d = Val(Mid(sBloco, c, 1)) * 2
      Else
         d = Val(Mid(sBloco, c, 1)) * 1
      End If
      If d > 0 Then
         If d > 9 Then
            e = CStr(d)
            d = Val(Left$(e, 1)) + Val(Right$(e, 1))
         End If
         nSoma = nSoma + d
      End If
Next

nResto = nSoma Mod 10
If nResto = 0 Then
   RetornaDV2of5 = 0
Else
   RetornaDV2of5 = 10 - nResto
End If

End Function

Public Function ExtraiNumero(sDado As String) As String
Dim c As Integer
Dim NewStr As String

For c = 1 To Len(sDado)
      If Asc(Mid$(sDado, c, 1)) <> 47 Then
         If IsNumeric(Mid(sDado, c, 1)) Then
            NewStr = NewStr & Mid(sDado, c, 1)
         End If
      End If
Next

ExtraiNumero = NewStr

End Function

Public Function Chomp(s As String, Side2Chomp As Direct, NumChar2Chomp As Integer) As String
    'trim leading/trailing spaces
    s = Trim(s)
    
    'raise error if len of string to chomp is shorter
    'than the amount you wish to chomp it by
    If Len(s) < NumChar2Chomp Then
    '    Err.Raise 10101, , "Length of string is shorter than " & _
                         "the amount you wish to chomp."
        Chomp = s
        Exit Function
    End If
    
    'remove characters from left(if side2chomp = left),
    'or characters from right(if side2chomp= Right)
    If Side2Chomp = chomp_left Then
         Chomp = Mid(s, (NumChar2Chomp + 1), (Len(s) - NumChar2Chomp))
    ElseIf Side2Chomp = chomp_righT Then
         Chomp = Mid(s, 1, (Len(s) - (NumChar2Chomp)))
    End If
End Function

Public Sub DefaultAccess(sUser As String)
' If Not bAdmin Then Exit Sub
'PARAMETROS
If Left$(sUser, 1) <> "[" Then
    sUser = "[" & sUser & "]"
End If
Sql = "USE Tributacao"
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT,INSERT,UPDATE ON PARAMETROS TO " & sUser
'cn.Execute Sql, rdExecDirect

'LOG
Sql = "GRANT SELECT,INSERT ON LOGEVENTO TO " & sUser
'cn.Execute Sql, rdExecDirect
'Sql = "GRANT EXEC ON spGRAVALOG TO " & sUser
'cn.Execute Sql, rdExecDirect

'SEGURANCA
Sql = "GRANT SELECT ON SEG_EVENTO TO " & sUser
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT ON SEG_EVENTOACESSO TO " & sUser
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT ON SEG_USUARIO TO " & sUser
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT ON SEG_GRUPO TO " & sUser
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT ON SEG_USERACESS TO " & sUser
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT ON SEG_GRUPOACESSO TO " & sUser
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT ON SEG_TELASISTEMA TO " & sUser
'cn.Execute Sql, rdExecDirect
Sql = "GRANT SELECT ON SEG_MENUACESSO TO " & sUser
'cn.Execute Sql, rdExecDirect

'PARAMETROS
Sql = "GRANT SELECT,INSERT ON PARAMETROS TO " & sUser
'cn.Execute Sql, rdExecDirect

End Sub

' =- Ferramentas VB
'    http://www.geocities.com/SiliconValley/Sector/1496/
'
'    FUNÇÃO PARA RETORNAR O VALOR EM EXTENSO
'    EXEMPLO:
'    EXTENSO(100.01)  RESULTADO "CEM REAIS E UM CENTAVO"
'
'    DATA: 29/05/1998

Public Function Extenso(nValor)

'Valida Argumento
If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then Exit Function

'Variáveis
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), aTexto(4) As String

'Matrizes de extensos (Parciais)
ReDim aUnid(19) As String
aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
aUnid(19) = "dezenove "

ReDim aDezena(9) As String
aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
aDezena(9) = "noventa "

ReDim aCentena(9) As String
aCentena(1) = "cento ": aCentena(2) = "duzentos "
aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
aCentena(9) = "novecentos "

'Separa valor em grupos
cValor = Format$(nValor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)

'Calcula cada grupo
For nContador = 1 To 4
    cParte = aGrupo(nContador)
    nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
    If nTamanho = 3 Then
        If Right$(cParte, 2) <> "00" Then
            aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
            nTamanho = 2
        Else
            aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", aCentena(Left(cParte, 1)))
        End If
    End If
    If nTamanho = 2 Then
        If Val(Right(cParte, 2)) < 20 Then
            aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
        Else
            aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
            If Right$(cParte, 1) <> "0" Then
                aTexto(nContador) = aTexto(nContador) + "e "
                nTamanho = 1
            End If
        End If
    End If
    If nTamanho = 1 Then
        aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
    End If
Next

'Final
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
    cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos")
Else
    cFinal = ""
    cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "milhões ", "milhão "), "")
    If Val(aGrupo(2) + aGrupo(3)) = 0 Then
        cFinal = cFinal + "de "
    Else
        cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "mil ", "")
    End If
    cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "real ", "reais ")
    cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "e " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos"), "")
End If
Extenso = cFinal

End Function

Public Function ValidaProcesso(sNumProcesso As String) As String
Dim nAno As Integer, nNumproc As Long
Dim Sql As String, RdoAux As rdoResultset

If Trim$(sNumProcesso) = "" Then
    ValidaProcesso = "Nº de Processo inválido."
    Exit Function
End If

If InStr(1, sNumProcesso, "/", vbBinaryCompare) = 0 Then
    ValidaProcesso = "Nº do processo inválido. Formato deve ser: Nº do Processo/Ano."
    Exit Function
End If

If Not IsNumeric(Right$(sNumProcesso, 4)) Then
    ValidaProcesso = "Nº do processo inválido. O ano deve ter 4 digitos."
    Exit Function
End If

If IsNumeric(Right$(sNumProcesso, 5)) Then
    ValidaProcesso = "Nº do processo inválido. O ano deve ter 4 digitos."
    Exit Function
End If

If Not IsNumeric(Left$(sNumProcesso, 1)) Then
    ValidaProcesso = "Nº do processo inválido."
    Exit Function
End If

nAno = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
nNumproc = Val(Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 1))

If NovoProtocolo = 0 Then
    Sql = "SELECT * FROM PROCESSO WHERE ANOPROCESS=" & nAno & " AND NUMEROPROC=" & nNumproc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            ValidaProcesso = "Processo não Cadastrado."
            Exit Function
        Else
            If !DataCancel > CDate("01/01/1900") Then
                ValidaProcesso = "Este Processo esta CANCELADO desde " & Format(!DataCancel, "dd/mm/yyyy") & "."
                Exit Function
            Else
                If !DATAARQUIV > CDate("01/01/1900") Then
                    ValidaProcesso = "Este Processo esta ARQUIVADO desde " & Format(!DATAARQUIV, "dd/mm/yyyy") & "."
                    Exit Function
                Else
                    If !DATASUSPEN > CDate("01/01/1900") Then
                        ValidaProcesso = "Este Processo esta SUSPENSO desde " & Format(!DATASUSPEN, "dd/mm/yyyy") & "."
                        Exit Function
                    Else
                        ValidaProcesso = "OK"
                    End If
                End If
            End If
        End If
       .Close
    End With
Else
    Sql = "SELECT * FROM PROCESSOGTI WHERE ANO=" & nAno & " AND NUMERO=" & nNumproc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            ValidaProcesso = "Processo não Cadastrado."
            Exit Function
        Else
            If !DataCancel > CDate("01/01/1900") Then
                ValidaProcesso = "Este Processo esta CANCELADO desde " & Format(!DataCancel, "dd/mm/yyyy") & "."
                Exit Function
            Else
                If !DATAARQUIVA > CDate("01/01/1900") Then
                    ValidaProcesso = "Este Processo esta ARQUIVADO desde " & Format(!DATAARQUIVA, "dd/mm/yyyy") & "."
                    Exit Function
                Else
                    If !DATASUSPENSO > CDate("01/01/1900") Then
                        ValidaProcesso = "Este Processo esta SUSPENSO desde " & Format(!DATASUSPENSO, "dd/mm/yyyy") & "."
                        Exit Function
                    Else
                        ValidaProcesso = "OK"
                    End If
                End If
            End If
        End If
       .Close
    End With
End If

End Function

Public Function ValidaProcesso2(sNumProcesso As String) As String
Dim nAno As Integer, nNumproc As Long
Dim Sql As String, RdoAux As rdoResultset

If Trim$(sNumProcesso) = "" Then
    ValidaProcesso2 = "Nº de Processo inválido."
    Exit Function
End If

If InStr(1, sNumProcesso, "/", vbBinaryCompare) = 0 Then
    ValidaProcesso2 = "Nº do processo inválido. Formato deve ser: Nº do Processo/Ano."
    Exit Function
End If

If Not IsNumeric(Right$(sNumProcesso, 4)) Then
    ValidaProcesso2 = "Nº do processo inválido. O ano deve ter 4 digitos."
    Exit Function
End If

If IsNumeric(Right$(sNumProcesso, 5)) Then
    ValidaProcesso2 = "Nº do processo inválido. O ano deve ter 4 digitos."
    Exit Function
End If

If Not IsNumeric(Left$(sNumProcesso, 1)) Then
    ValidaProcesso2 = "Nº do processo inválido."
    Exit Function
End If

nAno = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
nNumproc = Val(Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 2))

If NovoProtocolo = 0 Then
    Sql = "SELECT * FROM PROCESSO WHERE ANOPROCESS=" & nAno & " AND NUMEROPROC=" & nNumproc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            ValidaProcesso2 = "Processo não Cadastrado."
            Exit Function
        Else
            If !DataCancel > CDate("01/01/1900") Then
                ValidaProcesso2 = "Este Processo esta CANCELADO desde " & Format(!DataCancel, "dd/mm/yyyy") & "."
                Exit Function
            Else
                If !DATAARQUIV > CDate("01/01/1900") Then
                    ValidaProcesso2 = "Este Processo esta ARQUIVADO desde " & Format(!DATAARQUIV, "dd/mm/yyyy") & "."
                    Exit Function
                Else
                    If !DATASUSPEN > CDate("01/01/1900") Then
                        ValidaProcesso2 = "Este Processo esta SUSPENSO desde " & Format(!DATASUSPEN, "dd/mm/yyyy") & "."
                        Exit Function
                    Else
                        ValidaProcesso2 = "OK"
                    End If
                End If
            End If
        End If
       .Close
    End With
Else
    Sql = "SELECT * FROM PROCESSOGTI WHERE ANO=" & nAno & " AND NUMERO=" & nNumproc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            ValidaProcesso2 = "Processo não Cadastrado."
            Exit Function
        Else
            If !DataCancel > CDate("01/01/1900") Then
                ValidaProcesso2 = "Este Processo esta CANCELADO desde " & Format(!DataCancel, "dd/mm/yyyy") & "."
                Exit Function
            Else
                If !DATAARQUIVA > CDate("01/01/1900") Then
                    ValidaProcesso2 = "Este Processo esta ARQUIVADO desde " & Format(!DATAARQUIVA, "dd/mm/yyyy") & "."
                    Exit Function
                Else
                    If !DATASUSPENSO > CDate("01/01/1900") Then
                        ValidaProcesso2 = "Este Processo esta SUSPENSO desde " & Format(!DATASUSPENSO, "dd/mm/yyyy") & "."
                        Exit Function
                    Else
                        ValidaProcesso2 = "OK"
                    End If
                End If
            End If
        End If
       .Close
    End With
End If

End Function

Public Sub modLg(sLg As String)
'Dim x As Integer, FF1 As Integer, sUser As String, sComputer As String, sData As String, sHora As String, ax As String
'On Error GoTo Erro
'
'FF1 = FreeFile()
'Open sPathBin & "\gti.000" For Append As FF1
'ax = NomeDeLogin & "#" & NomeDoComputador & "#" & Format(Now, "dd/mm/yyyy") & "#" & Format(Now, "hh:mm") & "#" & sLg
'ax = Encrypt128(ax, MBI_LG)
'Print #FF1, ax
'Close #FF1
'
'frmHist.lstLog.AddItem Format(Now, "hh:mm") & " - " & sLg
'
'Exit Sub
'Erro:
'MsgBox "Erro desconhecido.", vbCritical, "Crítico"
'Resume Next

End Sub

Public Sub modLg000()
Dim FF1 As Integer, FF2 As Integer, sReg As String

'Open sPathBin & "\gti.dat" For Append As #1
'Open sPathBin & "\gti.000" For Input As #2
'
'While Not EOF(2)
'    Input #2, sReg
'    Print #1, sReg
'Wend
'
'Close #2
'Close #1

'frmHist.lstLog.Clear
'FF1 = FreeFile()
'Open sPathBin & "\gti.000" For Output As FF1
'Close #FF1
'
'Exit Sub
'Erro:
'MsgBox "Erro desconhecido 000.", vbCritical, "Crítico"
'Resume Next

End Sub


Public Function RetornaDataProcesso(nNumproc As Long, nAno As Integer) As Date
Dim Sql As String, RdoAux As rdoResultset

If NovoProtocolo = 0 Then
    Sql = "SELECT * FROM PROCESSO WHERE ANOPROCESS=" & nAno & " AND NUMEROPROC=" & nNumproc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            RetornaDataProcesso = "01/01/1899"
        Else
            If !DATAREATIV > CDate("01/01/1900") Then
                RetornaDataProcesso = Format(!DATAREATIV, "dd/mm/yyyy")
            Else
                RetornaDataProcesso = Format(!DATAENTRAD, "dd/mm/yyyy")
            End If
        End If
       .Close
    End With
Else
    Sql = "SELECT * FROM PROCESSOGTI WHERE ANO=" & nAno & " AND NUMERO=" & nNumproc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            RetornaDataProcesso = "01/01/1899"
        Else
            If !DATAREATIVA > CDate("01/01/1900") Then
                RetornaDataProcesso = Format(!DATAREATIVA, "dd/mm/yyyy")
            Else
                RetornaDataProcesso = Format(!DATAENTRADA, "dd/mm/yyyy")
            End If
        End If
       .Close
    End With
End If
End Function

Public Function RetornaNumeroProcessoPlusDV(nNumproc As Long)
Dim sNumProc As String, nIndex As Integer, nSoma As Integer, nMult As Integer
Dim nDigAux As Integer, nDigito As Integer

sNumProc = CStr(nNumproc)
Do While Len(sNumProc) < 5
    sNumProc = "0" & sNumProc
Loop

nSoma = 0: nIndex = 1: nMult = 6

Do While nIndex <= 5
    nSoma = nSoma + (nMult * Val(Mid(sNumProc, nIndex, 1)))
    nMult = nMult - 1
    nIndex = nIndex + 1
Loop

nDigAux = nSoma Mod 11
nDigito = 11 - nDigAux

If nDigito = 10 Then
    nDigito = 0
ElseIf nDigito = 11 Then
    nDigito = 1
End If

sNumProc = sNumProc + CStr(nDigito)
RetornaNumeroProcessoPlusDV = Val(sNumProc)

End Function

Public Function ExtraiNumeroProcesso(sNumProcesso As String) As String
ExtraiNumeroProcesso = Left$(sNumProcesso, InStr(1, sNumProcesso, "/", vbBinaryCompare) - 1)
End Function

Public Function ExtraiAnoProcesso(sNumProcesso As String) As String
ExtraiAnoProcesso = Right$(sNumProcesso, 4)
End Function


Public Function RetornaDVProcesso(nNumproc As Long)
Dim sNumProc As String, nIndex As Integer, nSoma As Integer, nMult As Integer
Dim nDigAux As Integer, nDigito As Integer

sNumProc = CStr(nNumproc)
Do While Len(sNumProc) < 5
    sNumProc = "0" & sNumProc
Loop

nSoma = 0: nIndex = 1: nMult = 6

Do While nIndex <= 5
    nSoma = nSoma + (nMult * Val(Mid(sNumProc, nIndex, 1)))
    nMult = nMult - 1
    nIndex = nIndex + 1
Loop

nDigAux = nSoma Mod 11
nDigito = 11 - nDigAux

If nDigito = 10 Then
    nDigito = 0
ElseIf nDigito = 11 Then
    nDigito = 1
End If

RetornaDVProcesso = nDigito

End Function

Public Sub ConsertaProcesso()
Dim Sql As String, RdoAux As rdoResultset, RdoGenesio As rdoResultset, RdoAux2 As rdoResultset
Dim sNumProcOld As String, nNumProcNew As Long, nAno As Integer, sNumProcNew As String

Sql = "SELECT NUMPROCESSO FROM PROCESSOREPARC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       'NUMERO DO PROCESSO ERRADO
       nNumProcOld = Val(Left$(Trim$(!NUMPROCESSO), InStr(1, Trim$(!NUMPROCESSO), "/", vbBinaryCompare) - 1))
       nAno = Val(Mid(Trim$(!NUMPROCESSO), InStr(1, Trim$(!NUMPROCESSO), "/", vbBinaryCompare) + 1, 4))
       'PROCURA NO GENÉSIO
       Sql = "SELECT ANOPROCESS,NUMEROPROC FROM PROCESSO WHERE NUMEROPROC=" & nNumProcOld & " AND ANOPROCESS=" & nAno
       Set RdoGenesio = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoGenesio
            If .RowCount = 0 Then
                'NÃO ENCONTROU, CALCULA DV
                 nNumProcNew = RetornaNumeroProcessoPlusDV(CLng(nNumProcOld))
                'RECONSTROI NUMERO DO PROCESSO
                 sNumProcNew = CStr(nNumProcNew) & "/" & CStr(nAno)
                'VERIFICA SE JA EXISTE NA PROCESSOREPARC
                 Sql = "SELECT * FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProcNew & "'"
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If .RowCount = 0 Then
                       'REGRAVA NA TABELA PROCESSOREPARC
                        Sql = "INSERT PROCESSOREPARC "
                        Sql = Sql & "SELECT '" & sNumProcNew & "',DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CODIGORESP,FUNCIONARIO,CANCELADO,DATACANCEL,FUNCIONARIOCANCEL,NUMprotocolo,PLANO FROM PROCESSOREPARC "
                        Sql = Sql & "WHERE NUMPROCESSO='" & RdoAux!NUMPROCESSO & "'"
                        cn.Execute Sql, rdExecDirect
                     End If
                    .Close
                 End With
                'VERIFICA SE JA EXISTE NA ORIGEMREPARC
                 Sql = "SELECT * FROM ORIGEMREPARC WHERE NUMPROCESSO='" & sNumProcNew & "'"
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If .RowCount = 0 Then
                       'REGRAVA NA TABELA ORIGEMREPARC
                        Sql = "INSERT ORIGEMREPARC "
                        Sql = Sql & "SELECT '" & sNumProcNew & "',CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO FROM ORIGEMREPARC "
                        Sql = Sql & "WHERE NUMPROCESSO='" & RdoAux!NUMPROCESSO & "'"
                        cn.Execute Sql, rdExecDirect
                     End If
                    .Close
                 End With
                'APAGA DA TABELA ORIGEMREPARC
                 Sql = "DELETE FROM ORIGEMREPARC WHERE NUMPROCESSO='" & CStr(nNumProcOld) & "/" & CStr(nAno) & "'"
                 cn.Execute Sql, rdExecDirect
                'VERIFICA SE JA EXISTE NA DESTINOREPARC
                 Sql = "SELECT * FROM DESTINOREPARC WHERE NUMPROCESSO='" & sNumProcNew & "'"
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If .RowCount = 0 Then
                       'REGRAVA NA TABELA DESTINOREPARC
                        Sql = "INSERT DESTINOREPARC "
                        Sql = Sql & "SELECT '" & sNumProcNew & "',CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO FROM DESTINOREPARC "
                        Sql = Sql & "WHERE NUMPROCESSO='" & RdoAux!NUMPROCESSO & "'"
                        cn.Execute Sql, rdExecDirect
                     End If
                    .Close
                 End With
                'APAGA DA TABELA DESTINOREPARC
                 Sql = "DELETE FROM DESTINOREPARC WHERE NUMPROCESSO='" & sNumProcOld & "'"
                 cn.Execute Sql, rdExecDirect
                'APAGA DA TABELA PROCESSOREPARC
                 Sql = "DELETE FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProcOld & "'"
                 cn.Execute Sql, rdExecDirect
                'CORRIGE PROCESSO EM DEBITOPARCELA
                 Sql = "UPDATE DEBITOPARCELA SET NUMPROCESSO='" & sNumProcNew & "' WHERE "
                 Sql = Sql & "NUMPROCESSO='" & RdoAux!NUMPROCESSO & "'"
                 cn.Execute Sql, rdExecDirect
            End If
           .Close
       End With
      .MoveNext
    Loop
   .Close
End With

'MsgBox "FINAL DA CORREÇÃO"

End Sub


Public Sub Add3DBorder(ByVal ControlorForm As Object)
'Add a 3D, office 2000 style border to a form or control
'Examples: Add3DBorder me ' for form
'          Add3DBorder text1 ' for control

On Error Resume Next
Dim lHwnd As Long
Dim lRet As Long

lHwnd = ControlorForm.hwnd
If lHwnd = 0 Then Exit Sub
ControlorForm.BorderStyle = 0
lRet = GetWindowLong(lHwnd, GWL_EXSTYLE)
lRet = lRet Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
SetWindowLong lHwnd, GWL_EXSTYLE, lRet
SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Public Function RetornaUFIR(nAnoUFIR) As Double

If dcUfir.Exists(nAnoUFIR) Then
    If nAnoUFIR > 1995 Then
       If nAnoUFIR = 1999 Or nAnoUFIR = 2005 Or nAnoUFIR = 2006 Or nAnoUFIR = 2007 Then
          RetornaUFIR = dcUfir.Item(nAnoUFIR) / 1000
       Else
          RetornaUFIR = dcUfir.Item(nAnoUFIR) / 10000
       End If
    Else
       RetornaUFIR = dcUfir.Item(nAnoUFIR) / 100
    End If
Else
    RetornaUFIR = 0
End If

End Function

Public Sub RemoveTitleBar(frmDest As Form)
SetWindowLong frmDest.hwnd, GWL_STYLE, GetWindowLong(frmDest.hwnd, GWL_STYLE) And Not WS_CAPTION
oldMode% = frmDest.ScaleMode
frmDest.ScaleMode = 1
frmDest.Height = frmDest.Height - 300
frmDest.ScaleMode = oldMode%
End Sub

Public Function CloseApplication() As Boolean
On Error GoTo Erro
    Dim blnRtn As Boolean, RdoAux As rdoResultset
    Dim error As Long
    Dim FixedInfoSize As Long
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String
    Dim NewTime As Date
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim AddrStr As IP_ADDR_STRING
    Dim FixedInfo As FIXED_INFO
    Dim Buffer As IP_ADDR_STRING
    Dim pAddrStr As Long
    Dim pAdapt As Long
    Dim Buffer2 As IP_ADAPTER_INFO
    Dim FixedInfoBuffer() As Byte
    Dim AdapterInfoBuffer() As Byte
    Dim sIP As String

    ' Get the main IP configuration information for this machine
    ' using a FIXED_INFO structure.
    FixedInfoSize = 0
    error = GetNetworkParams(ByVal 0&, FixedInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetNetworkParams sizing failed with error " & error
           Exit Function
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)

    AdapterInfoSize = 0
    error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetAdaptersInfo sizing failed with error " & error
           Exit Function
        End If
    End If
    ReDim AdapterInfoBuffer(AdapterInfoSize - 1)

    ' Get actual adapter information
    error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
    If error <> 0 Then
       MsgBox "GetAdaptersInfo failed with error " & error
       Exit Function
    End If

    ' Allocate memory
     CopyMemory AdapterInfo, AdapterInfoBuffer(0), AdapterInfoSize
    pAdapt = AdapterInfo.Next

    CopyMemory Buffer2, AdapterInfo, AdapterInfoSize
    sIP = Trim$(Buffer2.IpAddressList.IpAddress)
    For x = 1 To Len(sIP)
       If Asc(Mid(sIP, x, 1)) <> 0 Then
          TmpName = TmpName & Mid(sIP, x, 1)
       Else
          sIP = TmpName
          Exit For
       End If
    Next
    
    
    Sql = "select * from machines where usuario='" & NomeDeLogin & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        Sql = "insert machines (usuario,ip,computer,nome,data,gti_version) values('" & NomeDeLogin & "','" & sIP & "','" & NomeDoComputador & "','" & NomeDoUsuario & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & App.Major & "." & App.Minor & "." & App.Revision & "')"
        cn.Execute Sql, rdExecDirect
    Else
        Sql = "update machines set ip='" & sIP & "',computer='" & NomeDoComputador & "',nome='" & NomeDoUsuario & "',data='" & Format(Now, sDataFormat & " hh:mm:ss") & "',gti_version='" & App.Major & "." & App.Minor & "." & App.Revision & "' where usuario='" & NomeDeLogin & "'"
        cn.Execute Sql, rdExecDirect
        
    End If
    RdoAux.Close
'     On Error Resume Next
'     cn.Close
'     Conecta "usergti", "cvrcs04"
'     Sql = "SELECT IPSERVIDOR FROM INTRANET..ATUALIZACOES WHERE IPSERVIDOR='" & sIP & "'"
'     Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'     With RdoAux
'            If .RowCount > 0 Then
'                blnRtn = EnumWindows(AddressOf EnumCallBack, 0)
'            End If
'           .Close
'     End With
'    cn.Close
Exit Function
Erro:
MsgBox Err.Description
Resume Next
End Function

Public Function EnumCallBack(ByVal hWndChild As Long, ByVal lParam As Long) As Long

Dim lngSize As Long
Dim strPadString As String
Dim lHwnd As Long
Dim lRetVal As Long

strPadString = String(255, 0)
lngSize = GetWindowText(hWndChild, strPadString, Len(strPadString))
strPadString = Left$(strPadString, lngSize)
If UCase(Left$(strPadString, 5)) = "ATUAL" Then
    lHwnd = FindWindow(vbNullString, strPadString)
    If lHwnd <> 0 Then
        lRetVal = PostMessage(lHwnd, WM_CLOSE, 0&, 0&)
    End If
End If
EnumCallBack = True

End Function

Public Function Getpath_SYSTEM() As String
Dim WindirS As String * 255          'declares a full lenght string for DIR name(for getting the path)
                                        
Dim TEMP                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
Dim Result                          'a variable for holding the the output of the function
TEMP = GetSystemDirectory(WindirS, 255)      'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, TEMP)                 'holds final path
Getpath_SYSTEM = Result
End Function

Public Function pathOfFile(FileName As String) As String
    Dim posn As Integer
    posn = InStrRev(FileName, "\")
    If posn > 0 Then
        pathOfFile = Left$(FileName, posn)
    Else
        pathOfFile = ""
    End If
End Function





Public Function WordWrap(ByVal Text As String, Optional ByVal MaxLineLen As Integer = 70)

    Dim i As Integer

    For i = 1 To Len(Text) / MaxLineLen
        Text = Mid(Text, 1, MaxLineLen * i - 1) & Replace(Text, " ", vbCrLf, MaxLineLen * i, 1, vbTextCompare)
    Next i

    WordWrap = Text
End Function

Sub SetProgressBar(Progressbar As Object, Percent As Integer, Optional Style As Integer, Optional Style2 As Integer)

With Progressbar
    .AutoRedraw = True
    .Cls
    .FontTransparent = True
    .Tag = Percent
    .ScaleWidth = 100
    .ScaleHeight = 10
    .DrawStyle = Style2
    .DrawMode = 13
    .FillStyle = Style
     Progressbar.Line (0, 0)-(Percent, .ScaleHeight - 1), , BF
     Progressbar.Line (0, 0)-(Percent, .ScaleHeight - 1), , B
    .FontTransparent = False
    .CurrentX = 50 - .TextWidth(Percent & "%")
    .CurrentY = (.ScaleHeight / 2) - (.TextHeight(Percent & "%") / 2)
    .FontBold = True
    .FontSize = 7
    .FontName = "Tahoma"
    .Print " " & Percent & "% "
End With

End Sub

Public Function Modulo11(ByVal nNumero As Long) As Integer

    Dim sFromN As String        'Converte Num to String
    Dim nTotPosAtual As Long    'Total da Posição
    Dim nDV As Integer

    nTotPosAtual = 0

    sFromN = Format(nNumero, "0000000000000")

    nTotPosAtual = Val(Mid(sFromN, 1, 1)) * 6
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 2, 1)) * 5)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 3, 1)) * 4)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 4, 1)) * 3)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 5, 1)) * 2)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 6, 1)) * 9)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 7, 1)) * 8)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 8, 1)) * 7)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 9, 1)) * 6)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 10, 1)) * 5)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 11, 1)) * 4)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 12, 1)) * 3)
    nTotPosAtual = nTotPosAtual + (Val(Mid(sFromN, 13, 1)) * 2)

    nDV = nTotPosAtual Mod 11
    nDV = 11 - nDV
    If nDV = 1 Then nDV = 0
    If nDV = 10 Then nDV = 1
    If nDV = 11 Then nDV = 0
    Modulo11 = nDV

End Function

Public Sub AtualizaPropDuplicado(nCodReduz As Long, nCodCidadao As Long)
Dim RdoAux As rdoResultset, Sql As String

'APAGA O IMOVEL DA TABELA DE DUPLICADOS
Sql = "DELETE FROM PROPRIETARIODUPLICADO WHERE CODREDUZIDO=" & nCodReduz
cn.Execute Sql, rdExecDirect

'SÓ VAI INSERIR SE JA TIVER ALGUM OUTRO IMOVEL CADASTRADO COM O MESMO CIDADAO
Sql = "SELECT CODREDUZIDO,CODCIDADAO FROM PROPRIETARIODUPLICADO WHERE CODCIDADAO=" & nCodCidadao
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Sql = "INSERT PROPRIETARIODUPLICADO(CODREDUZIDO,CODCIDADAO) VALUES(" & nCodReduz & "," & nCodCidadao & ")"
        cn.Execute Sql, rdExecDirect
    End If
   .Close
End With

End Sub

Public Function RetornaAliquotaISS(nCodigo As Integer, dData As Date)
Dim Sql As String, RdoAux As rdoResultset, aData() As AliquotaISS, x As Integer

ReDim aData(0)

Sql = "SELECT * FROM TABELAISS WHERE CODIGOATIV=" & nCodigo & " ORDER BY DATA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aData(UBound(aData) + 1)
        aData(UBound(aData)).dData = Format(!Data, "dd/mm/yyyy")
        aData(UBound(aData)).nAliquota = !Aliquota
       .MoveNext
    Loop
   .Close
End With

If UBound(aData) = 1 Then
   RetornaAliquotaISS = FormatNumber(aData(1).nAliquota, 3)
   Exit Function
Else
   For x = 1 To UBound(aData)
      If x = UBound(aData) Then
         RetornaAliquotaISS = FormatNumber(aData(x).nAliquota, 3)
         Exit Function
      Else
         If dData >= aData(x).dData And dData < aData(x + 1).dData Then
            RetornaAliquotaISS = FormatNumber(aData(x).nAliquota, 3)
            Exit Function
         End If
      End If
   Next
End If

End Function

Public Function ImovelAreaUnica(nCodCidadao As Long) As Boolean
Dim x As Integer, Sql As String, RdoTmp As rdoResultset, RdoTmp2 As rdoResultset

x = 0
Sql = "SELECT CODREDUZIDO FROM PROPRIETARIO WHERE CODCIDADAO=" & nCodCidadao
Set RdoTmp = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoTmp
    Do Until .EOF
        x = x + 1
       .MoveNext
    Loop
   .Close
End With

If x = 1 Then
    ImovelAreaUnica = True
Else
    ImovelAreaUnica = False
End If

End Function

Public Function cGetInputState()
Dim qsRet As Long
qsRet = GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT)
                   cGetInputState = qsRet
End Function

Public Sub SaveFormImageToFile(ByRef ContainerForm As Form, ByRef PictureBoxControl As picturebox, ByVal ImageFileName As String)
  Dim FormInsideWidth As Long
  Dim FormInsideHeight As Long
  Dim PictureBoxLeft As Long
  Dim PictureBoxTop As Long
  Dim PictureBoxWidth As Long
  Dim PictureBoxHeight As Long
  Dim FormAutoRedrawValue As Boolean
  
  With PictureBoxControl
    'Set PictureBox properties
    .Visible = False
    .AutoRedraw = True
    .Appearance = 0 ' Flat
    .AutoSize = False
    .BorderStyle = 0 'No border
    
    'Store PictureBox Original Size and location Values
    PictureBoxHeight = .Height: PictureBoxWidth = .Width: PictureBoxLeft = .Left: PictureBoxTop = .Top
    
    'Make PictureBox to size to inside of form.
    .Align = vbAlignTop: .Align = vbAlignLeft
    DoEvents
    
    FormInsideHeight = .Height: FormInsideWidth = .Width
    
    'Restore PictureBox Original Size and location Values
    .Align = vbAlignNone
    .Height = FormInsideHeight: .Width = FormInsideWidth: .Left = PictureBoxLeft: .Top = PictureBoxTop
    
    FormAutoRedrawValue = ContainerForm.AutoRedraw
    ContainerForm.AutoRedraw = False
    DoEvents
    
    'Copy Form Image to Picture Box
    BitBlt .hdc, 0, 0, FormInsideWidth / Screen.TwipsPerPixelX, FormInsideHeight / Screen.TwipsPerPixelY, ContainerForm.hdc, 0, 0, vbSrcCopy
    DoEvents
    SavePicture .Image, ImageFileName
    DoEvents
    
    ContainerForm.AutoRedraw = FormAutoRedrawValue
    DoEvents
  End With
End Sub




'Purpose     :  Retreview text from a web site
'Inputs      :  sURLFileName            The URL and file name to download.
'               sSaveToFile             The filename to save the file to.
'               [bOverwriteExisting]    If True overwrites the file if it existings
'Outputs     :  Returns True on success.


Function InternetGetFile(sURLFileName As String, sSaveToFile As String, Optional bOverwriteExisting As Boolean = False) As Boolean
    Dim lRet As Long
    Const S_OK As Long = 0, E_OUTOFMEMORY = &H8007000E
    Const INTERNET_OPEN_TYPE_PRECONFIG = 0, INTERNET_FLAG_EXISTING_CONNECT = &H20000000
    Const INTERNET_OPEN_TYPE_DIRECT = 1, INTERNET_OPEN_TYPE_PROXY = 3
    Const INTERNET_FLAG_RELOAD = &H80000000
    
    On Error Resume Next
    'Create an internet connection
    lRet = InternetOpen("", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    
    If bOverwriteExisting Then
        If Len(Dir$(sSaveToFile)) Then
            VBA.Kill sSaveToFile
        End If
    End If
    'Check file doesn't already exist
    If Len(Dir$(sSaveToFile)) = 0 Then
        'Download file
        lRet = URLDownloadToFile(0&, sURLFileName, sSaveToFile, 0&, 0)
        If Len(Dir$(sSaveToFile)) Then
            'File successfully downloaded
            InternetGetFile = True
        Else
            'Failed to download file
            If lRet = E_OUTOFMEMORY Then
                Debug.Print "The buffer length is invalid or there was insufficient memory to complete the operation."
            Else
                Debug.Assert False
                Debug.Print "Error occurred " & lRet & " (this is probably a proxy server error)."
            End If
            InternetGetFile = False
        End If
    End If
    On Error GoTo 0
    
End Function

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Public Function RetornaUsuarioFullName() As String
Dim Sql As String, RdoAux As rdoResultset

Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE NOMELOGIN='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    RetornaUsuarioFullName = !NomeCompleto
   .Close
End With

End Function

Public Function RetornaUsuarioFullName2(sNomeUser) As String
Dim Sql As String, RdoAux As rdoResultset

Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE NOMELOGIN='" & sNomeUser & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    RetornaUsuarioFullName2 = !NomeCompleto
   .Close
End With

End Function

Public Function RetornaUsuarioLoginName(sFullName) As String
Dim Sql As String, RdoAux As rdoResultset

Sql = "SELECT NOMELOGIN FROM USUARIO WHERE NOMECOMPLETO='" & Mask(CStr(sFullName)) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    RetornaUsuarioLoginName = !NomeLogin
   .Close
End With

End Function


Public Function RetornaUsuarioID(sLoginName) As Integer
Dim Sql As String, RdoAux As rdoResultset

Sql = "SELECT id FROM USUARIO WHERE NOMELOGIN='" & sLoginName & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        RetornaUsuarioID = !id
    Else
        RetornaUsuarioID = 0
    End If
   .Close
End With

End Function


Public Sub BuildReportMob1()

Dim sNomeArq As String, FF1 As Integer, ret As Double, Sql As String, RdoAux As rdoResultset, nContaTx As Long, nContaVs As Long
Dim nSomaTxI As Double, nSomaVsI As Double, aSomaTotalTx(0 To 3) As Double, aSomaTotalVs(0 To 3) As Double
Dim aRel() As RELMOB1, x As Integer, bFind As Boolean, nPos As Long, ax As String, nCodOld As Long, nCodNew As Long
'LISTA DE EMPRESAS COM TAXA DE LICENÇA E VIG.SANITÁRIA EM ATRASO ENTRE 2008 E 2010
'                                                        TAXA DE LIC.      VIG.SANITÁRIA
'CÓDIGO   RAZAO SOCIAL                            ANO   PC1  PC2  PC3    PC1  PC2  PC3  PC4
'100022   EMPREENDIMENTOS SOUZA E SILVA LTDA      2008   X         X
'                                                 2009                              X
Ocupado
ReDim aRel(0): nContaTx = 0: nContaVs = 0:
For x = 1 To 3
    aSomaTotalTx(x) = 0: aSomaTotalVs(x) = 0
Next

Sql = "SELECT COUNT(DISTINCT mobiliario.codigomob) AS soma FROM debitoparcela INNER JOIN mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob "
Sql = Sql & "WHERE (mobiliario.simples = 1) AND (debitoparcela.codlancamento = 6) AND (debitoparcela.anoexercicio BETWEEN 2008 AND 2010) AND (debitoparcela.statuslanc = 3) AND (debitoparcela.numparcela > 0)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nContaTx = RdoAux!soma
RdoAux.Close

Sql = "SELECT COUNT(DISTINCT mobiliario.codigomob) AS soma FROM debitoparcela INNER JOIN mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob "
Sql = Sql & "WHERE (mobiliario.simples = 1) AND (debitoparcela.codlancamento = 13) AND (debitoparcela.anoexercicio BETWEEN 2008 AND 2010) AND (debitoparcela.statuslanc = 3) AND (debitoparcela.numparcela > 0)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nContaVs = RdoAux!soma
RdoAux.Close

'carrega matriz
Sql = "SELECT mobiliario.codigomob, mobiliario.razaosocial, debitoparcela.codlancamento, debitoparcela.numparcela, debitoparcela.anoexercicio,SUM(debitotributo.ValorTributo) As soma "
Sql = Sql & "FROM debitoparcela INNER JOIN mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
Sql = Sql & "Where (mobiliario.SIMPLES = 1) And (debitoparcela.statuslanc = 3) And (debitotributo.CodTributo <> 3) GROUP BY mobiliario.codigomob, mobiliario.razaosocial, debitoparcela.codlancamento, debitoparcela.numparcela, debitoparcela.anoexercicio "
Sql = Sql & "HAVING (debitoparcela.codlancamento = 6 OR debitoparcela.codlancamento = 13) AND (debitoparcela.anoexercicio BETWEEN 2008 AND 2010) AND (debitoparcela.numparcela > 0) "
Sql = Sql & "ORDER BY mobiliario.codigomob, debitoparcela.anoexercicio, debitoparcela.numparcela"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If cGetInputState() <> 0 Then DoEvents
        bFind = False
        For x = 1 To UBound(aRel)
            If aRel(x).nCodReduz = !codigomob And aRel(x).nAno = !AnoExercicio Then
                bFind = True
                Exit For
            End If
        Next
        If bFind Then
            If !CodLancamento = 6 Then
                If !NumParcela = 1 Then
                    aRel(x).sTx1 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 2 Then
                    aRel(x).sTx2 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 3 Then
                    aRel(x).sTx3 = FillLeft(FormatNumber(!soma, 2), 11)
                End If
            ElseIf !CodLancamento = 13 Then
                If !NumParcela = 1 Then
                    aRel(x).sVs1 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 2 Then
                    aRel(x).sVs2 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 3 Then
                    aRel(x).sVs3 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 4 Then
                    aRel(x).sVs4 = FillLeft(FormatNumber(!soma, 2), 11)
                End If
            End If
        Else
            nPos = UBound(aRel) + 1
            ReDim Preserve aRel(nPos)
            aRel(nPos).nCodReduz = !codigomob
            aRel(nPos).sRazao = !RazaoSocial
            aRel(nPos).nAno = !AnoExercicio
            aRel(nPos).sTx1 = FillLeft(FormatNumber(0, 2), 11)
            aRel(nPos).sTx2 = FillLeft(FormatNumber(0, 2), 11)
            aRel(nPos).sTx3 = FillLeft(FormatNumber(0, 2), 11)
            aRel(nPos).sVs1 = FillLeft(FormatNumber(0, 2), 11)
            aRel(nPos).sVs2 = FillLeft(FormatNumber(0, 2), 11)
            aRel(nPos).sVs3 = FillLeft(FormatNumber(0, 2), 11)
            aRel(nPos).sVs4 = FillLeft(FormatNumber(0, 2), 11)
            If !CodLancamento = 6 Then
                If !NumParcela = 1 Then
                    aRel(nPos).sTx1 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 2 Then
                    aRel(nPos).sTx2 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 3 Then
                    aRel(nPos).sTx3 = FillLeft(FormatNumber(!soma, 2), 11)
                End If
            ElseIf !CodLancamento = 13 Then
                If !NumParcela = 1 Then
                    aRel(nPos).sVs1 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 2 Then
                    aRel(nPos).sVs2 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 3 Then
                    aRel(nPos).sVs3 = FillLeft(FormatNumber(!soma, 2), 11)
                ElseIf !NumParcela = 4 Then
                    aRel(nPos).sVs4 = FillLeft(FormatNumber(!soma, 2), 11)
                End If
            End If
        End If
       .MoveNext
    Loop
   .Close
End With

nCodOld = 0
sNomeArq = sPathBin & "\REPORTMOB1.TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1
Print #FF1, "***********************************************************"
Print #FF1, "LISTA DE EMPRESAS OPTANTES PELO SIMPLES NACIONAL COM"
Print #FF1, "TAXA DE LICENÇA E VIG.SANITÁRIA EM ATRASO ENTRE 2008 E 2010"
Print #FF1, "***********************************************************"
Print #FF1, FillLeft("TAXA DE LIC.", 90) & FillLeft("VIG.SANITÁRIA", 50)
ax = FillSpace("CÓDIGO", 8) & FillSpace("RAZÃO SOCIAL", 42) & FillLeft("ANO", 8) & FillLeft("PC1", 11) & FillLeft("PC2", 11) & FillLeft("PC3", 11) & FillLeft("TOTAL", 11) & FillLeft("PC1", 11) & FillLeft("PC2", 11) & FillLeft("PC3", 11) & FillLeft("PC4", 11) & FillLeft("TOTAL", 11)
Print #FF1, ax
For x = 1 To UBound(aRel)
    With aRel(x)
        nCodNew = .nCodReduz
        If nCodNew <> nCodOld Then
            ax = FillSpace(CStr(.nCodReduz), 8) & FillSpace(Left(.sRazao, 40), 42) & FillLeft(CStr(.nAno), 8) & .sTx1 & .sTx2 & .sTx3
        Else
            ax = FillSpace(" ", 50) & FillLeft(CStr(.nAno), 8) & .sTx1 & .sTx2 & .sTx3
        End If
        nSomaTxI = CDbl(.sTx1) + CDbl(.sTx2) + CDbl(.sTx3)
        ax = ax & FillLeft(FormatNumber(nSomaTxI, 2), 11)
        ax = ax & FillLeft(.sVs1, 11) & FillLeft(.sVs2, 11) & FillLeft(.sVs3, 11) & FillLeft(.sVs4, 11)
        nSomaVsI = CDbl(.sVs1) + CDbl(.sVs2) + CDbl(.sVs3) + CDbl(.sVs4)
        ax = ax & FillLeft(FormatNumber(nSomaVsI, 2), 11)
        Print #FF1, ax
        nCodOld = nCodNew
        If .nAno = 2008 Then
            aSomaTotalTx(1) = aSomaTotalTx(1) + nSomaTxI
            aSomaTotalVs(1) = aSomaTotalVs(1) + nSomaVsI
        ElseIf .nAno = 2009 Then
            aSomaTotalTx(2) = aSomaTotalTx(2) + nSomaTxI
            aSomaTotalVs(2) = aSomaTotalVs(2) + nSomaVsI
        ElseIf .nAno = 2010 Then
            aSomaTotalTx(3) = aSomaTotalTx(3) + nSomaTxI
            aSomaTotalVs(3) = aSomaTotalVs(3) + nSomaVsI
        End If
    End With
Next
aSomaTotalTx(0) = aSomaTotalTx(1) + aSomaTotalTx(2) + aSomaTotalTx(3)
aSomaTotalVs(0) = aSomaTotalVs(1) + aSomaTotalVs(2) + aSomaTotalVs(3)
Print #FF1, FillSpace(" ", 89) & "   ----------" & FillLeft(" ", 38) & FillSpace(" ", 5) & " -----------"
Print #FF1, FillLeft("Total Geral:", 85) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalTx(0), 2), 13) & FillLeft("Total Geral:", 38) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalVs(0), 2), 13)
Print #FF1, FillLeft("   |-->2008:", 85) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalTx(1), 2), 13) & FillLeft("   |-->2008:", 38) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalVs(1), 2), 13)
Print #FF1, "Total de Empresas devendo Taxa de Licença: " & Format(nContaTx, "00000") & FillLeft("   |-->2009:", 37) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalTx(2), 2), 13) & FillLeft("   |-->2009:", 38) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalVs(2), 2), 13)
Print #FF1, "Total de Empresas devendo Vigil.Sanitária: " & Format(nContaVs, "00000") & FillLeft("   |-->2010:", 37) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalTx(3), 2), 13) & FillLeft("   |-->2010:", 38) & FillSpace(" ", 4) & FillLeft(FormatNumber(aSomaTotalVs(3), 2), 13)
Print #FF1, ""
Print #FF1, "Gerado em: " & Format(Now, "dd/mm/yyyy") & " Module - BuildReportMob1 (Gestão de Tributação Municipal Integrada - GTI) Prefeitura Municipal de Jaboticabal"
Close #FF1
ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)

Liberado
MsgBox "Relatório disponível em " & sPathBin & "\REPORTMOB1.TXT"
End Sub

Public Function SetTimeFormat(ByVal TimeValue As Double)
    On Error GoTo errorhandler
    Seconds = Fix(TimeValue)
    mins = Fix(TimeValue / 60)
    secs = TimeValue - (mins * 60)
    If secs < 10 Then secs = "0" & secs
    SetTimeFormat = mins & ":" & secs
    Exit Function
errorhandler:
SetTimeFormat = "0:00"
End Function

Public Sub AlwaysOnTop(FrmID As Form, OnTop As Boolean)
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop Then
       OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
       OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub

Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      IpAddrs(i) = s
      Next
   GetIpAddrTable = IpAddrs
   End Function


Public Function ToggleScreenSaverActive(Active As Boolean) As Boolean
'To Activate Screen Saver, set active to true
'to deactivate, set active to false
Dim lActiveFlag As Long
Dim retval As Long

lActiveFlag = IIf(Active, 1, 0)
retval = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, lActiveFlag, 0, 0)

ToggleScreenSaverActive = retval > 0

End Function

Public Function ParsePath(strFullPathName As String, ReturnType As Integer, Optional StripLastBackslash) As String

Dim strTemp As String, intX As Integer, strPathName As String, strFileName As String

If IsMissing(StripLastBackslash) Then StripLastBackslash = False
If Len(strFullPathName) > 0 Then
    strTemp = ""
    intX = Len(strFullPathName)
    Do While strTemp <> "\"
        strTemp = Mid(strFullPathName, intX, 1)
        If strTemp = "\" Then
            strPathName = Left(strFullPathName, intX + StripLastBackslash)
            strFileName = Right(strFullPathName, Len(strFullPathName) - intX)
        End If
        intX = intX - 1
    Loop
    
    Select Case ReturnType
        Case vbDirectory
            ParsePath = strPathName
        Case vbNormal
            ParsePath = strFileName
        Case Else
            ParsePath = strFullPathName
    End Select
Else
    ParsePath = ""
End If

End Function

'Searching the contents of a listview

'The function below searches a listview for a specific item. The code will search the text, subitems and tags of the listview items.

'Public Enum elvSearch
'    elvSearchText = 1
'    elvSearchSub = 2
'    elvSearchTag = 4
'End Enum

'Purpose     :  Finds and selects and item in a listview
'Inputs      :  sFileName               The path and file name of the component to register.
'               lvFind                  The listview to search for the item in.
'               [eValueType]            The type of values to search:
'                                       1 = Searches the text items.
'                                       2 = Searches sub items.
'                                       4 = Searches the item tags.
'               [lSearchFor]            The type of matching required:
'                                       lvwWhole = Find whole word.
'                                       lvwPartial = Find a partial match.
'               [lIndexBeginFrom]       The item index to begin the search from, for recursive
'                                       searches. See the index property of the listitem.
'                                       property of the listitem.
'Outputs     :  N/A


Public Function ListViewFindItem(sFindItem As String, lvFind As ListView, Optional eValueType As elvSearch = elvSearchText + elvSearchSub + elvSearchTag, Optional lSearchFor As Long = lvwPartial, Optional lIndexBeginFrom As Long = 1) As ListItem
    On Error Resume Next
    
    'Try to find item
    If eValueType And elvSearchText Then
        'Search text
        Set ListViewFindItem = lvFind.FindItem(sFindItem, lvwText, lIndexBeginFrom, lSearchFor)
    End If
    If eValueType And elvSearchSub And (ListViewFindItem Is Nothing) Then
        'Search subitems
        Set ListViewFindItem = lvFind.FindItem(sFindItem, lvwText, lIndexBeginFrom, lSearchFor)
    End If
    If eValueType And elvSearchTag And (ListViewFindItem Is Nothing) Then
        'Search tags
        Set ListViewFindItem = lvFind.FindItem(sFindItem, lvwText, lIndexBeginFrom, lSearchFor)
    End If
    
    If (ListViewFindItem Is Nothing) = False Then
        'Found a matching item, display it.
        Set lvFind.SelectedItem = ListViewFindItem
    End If
    On Error GoTo 0
End Function

Public Function ProdEventoDia(nFiscal As Integer, dData As Date) As Integer
Dim Sql As String, RdoAux As rdoResultset, bAchou As Boolean, dDataIni As Date, dDataFim As Date

ProdEventoDia = 0
bAchou = False
Sql = "select * from produtividadefiscalevento where codfiscal=" & nFiscal & " order by seq"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    dData = Format(dData, "dd/mm/yyyy")
    Do Until .EOF
        dDataIni = Format(!dataini, "dd/mm/yyyy")
        dDataFim = Format(!datafim, "dd/mm/yyyy")
        If dData >= dDataIni And dData <= dDataFim Then
            bAchou = True
            ProdEventoDia = !CODEVENTO
            Exit Do
        End If
       .MoveNext
    Loop
   .Close
End With

End Function

Public Function ProdIsBoss(nFiscal) As Boolean
Dim Sql As String, RdoAux As rdoResultset, bRet As Boolean

Sql = "select codigo,chefe from produtividadefiscal where codigo=" & nFiscal
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(RdoAux!Chefe) Then
        bRet = False
    Else
        bRet = RdoAux!Chefe
    End If
   .Close
End With

ProdIsBoss = bRet
End Function

Public Function ProdIsBossLogin() As Boolean
Dim Sql As String, RdoAux As rdoResultset, bRet As Boolean

Sql = "select codigo,chefe from produtividadefiscal where nome='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If RdoAux.RowCount > 0 Then
        If IsNull(RdoAux!Chefe) Then
            bRet = False
        Else
            bRet = RdoAux!Chefe
        End If
    Else
        bRet = False
    End If
   .Close
End With

ProdIsBossLogin = bRet
End Function

Public Function RetornaNome(nCodigo As Long) As String
Dim sNome As String, Sql As String, RdoAux As rdoResultset

If nCodigo < 100000 Then
    Sql = "SELECT NOMECIDADAO AS NOME FROM VWFULLIMOVEL WHERE CODREDUZIDO=" & nCodigo
ElseIf nCodigo > 100000 And nCodigo < 300000 Then
    Sql = "SELECT RAZAOSOCIAL AS NOME FROM MOBILIARIO WHERE CODIGOMOB=" & nCodigo
Else
    Sql = "SELECT NOMECIDADAO AS NOME FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    If .RowCount > 0 Then
        sNome = SubNull(!Nome)
    Else
        sNome = "Código inválido!!"
    End If
   .Close
End With
RetornaNome = sNome

End Function

Public Function Calculo_DV10(strNumero As String) As String
'declara As variáveis
Dim intContador As Integer

Dim intNumero As Integer

Dim intTotalNumero As Integer

Dim intMultiplicador As Integer

Dim intResto As Integer

' se nao for um valor numerico sai da função
If Not IsNumeric(strNumero) Then
  Calculo_DV10 = ""
  Exit Function
End If

'inicia o multiplicador
intMultiplicador = 2

'pega cada caracter do numero a partir da direita
For intContador = Len(strNumero) To 1 Step -1

    'extrai o caracter e multiplica pelo multiplicador
    intNumero = Val(Mid(strNumero, intContador, 1)) * intMultiplicador
    
    ' se o resultado for maior que nove soma os algarismos do resultado
    If intNumero > 9 Then
      intNumero = Val(Left(intNumero, 1)) + Val(Right(intNumero, 1))
    End If
    
    'soma o resultado para totalização
    intTotalNumero = intTotalNumero + intNumero
    
    'se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
    intMultiplicador = IIf(intMultiplicador = 2, 1, 2)

Next

    Dim DezenaSuperior As Integer
    If intTotalNumero < 10 Then
        DezenaSuperior = 10
    Else
        DezenaSuperior = 10 * (Val(Left(CStr(intTotalNumero), 1)) + 1)
    End If
    intResto = DezenaSuperior - intTotalNumero

'verifica as exceções ( 0 -> DV=0 )
Select Case intResto
  Case 0
     Calculo_DV10 = "0"
  Case 10
     Calculo_DV10 = "0"
  Case Else
     Calculo_DV10 = sTr(intResto)
End Select

End Function

Public Function Calculo_DV11(strNumero As String) As String
'declara as variáveis
Dim intContador As Integer

Dim intNumero As Integer

Dim intTotalNumero As Integer

Dim intMultiplicador As Integer

Dim intResto As Integer

' se nao for um valor numerico sai da função
If Not IsNumeric(strNumero) Then
  Calculo_DV11 = ""
  Exit Function
End If

'inicia o multiplicador
intMultiplicador = 2

'pega cada caracter do numero a partir da direita
For intContador = Len(strNumero) To 1 Step -1

    'extrai o caracter e multiplica prlo multiplicador
    intNumero = Val(Mid(strNumero, intContador, 1)) * intMultiplicador
    
    'soma o resultado para totalização
    intTotalNumero = intTotalNumero + intNumero
    
    'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
    intMultiplicador = IIf(intMultiplicador < 9, intMultiplicador + 1, 2)

Next

'calcula o resto da divisao do total por 11
intResto = (intTotalNumero * 10) Mod 11

'verifica as exceções ( 0 -> DV=0    10 -> DV=X (para o BB) e retorna o DV
Select Case intResto
  Case 0
    Calculo_DV11 = "1"
  Case 10
    Calculo_DV11 = "1"
  Case Else
    Calculo_DV11 = sTr(intResto)
End Select

End Function

Public Function RetornaUsuarioFiscal() As Boolean
Dim Sql As String, RdoAux As rdoResultset, nRet As Boolean

Sql = "SELECT COUNT(*) AS contador FROM  usuario  WHERE  (fiscal = 1) AND (ativo = 1) AND nomelogin = '" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux!contador = 1 Then
    nRet = True
Else
    nRet = False
End If
RdoAux.Close
RetornaUsuarioFiscal = nRet
End Function

Public Function InSerasa(nCodigo As Long) As Boolean
Dim nRet As Boolean, Sql As String, RdoAux As rdoResultset

Sql = "select * from serasa where codigo=" & nCodigo & " and dtsaida is null"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    nRet = True
Else
    nRet = Fals5e
End If
RdoAux.Close
InSerasa = nRet

End Function

Public Sub Close_GTI_Server()
Dim target_hwnd As Long
On Error Resume Next
    ' Get the target's window handle.
    target_hwnd = FindWindow(vbNullString, "Gerenciador de Serviços do G.T.I.")
    If target_hwnd = 0 Then
       ' MsgBox "Error finding target window handle"
        Exit Sub
    End If

    ' Send the application the WM_CLOSE message.
    PostMessage target_hwnd, WM_CLOSE, 0, 0
End Sub


Public Sub GeraRefisDam(nAno As Integer)
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, nNumDoc As Long, ax As String, z As Long, nTotal As Double
nTotal = 0
Ocupado
Open sPathBin & "\RefisDAM.txt" For Output As #1
Print #1, "RELATÓRIO DO REFIS (PAGAMENTO À VISTA)"
Print #1, "REFIS - (" & nAno & ") - Impresso em " & Format(Now, "dd/mm/yyyy")
Print #1, "-------------------------------------------------"
Print #1, " "
Print #1, "Documento     Valor     Código   Dt.Pagto."
Print #1, "-------------------------------------------------"
Print #1, " "
Sql = "SELECT DISTINCT parceladocumento.numdocumento, debitoparcela.codreduzido, debitopago.datapagamento, SUM(debitopago.valorpagoreal) AS valorpago,numdocumento.valorpago AS valordoc "
Sql = Sql & "FROM numdocumento INNER JOIN parceladocumento ON numdocumento.numdocumento = parceladocumento.numdocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
Sql = Sql & "debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitopago.NumParcela And debitoparcela.CODCOMPLEMENTO = debitopago.CODCOMPLEMENTO WHERE (debitoparcela.codlancamento NOT IN (36, 65)) AND (parceladocumento.plano = 11 OR "
Sql = Sql & "parceladocumento.plano = 12) AND (debitoparcela.datavencimento < CONVERT(DATETIME, '2016-01-01 00:00:00', 102)) GROUP BY parceladocumento.numdocumento, debitoparcela.codreduzido, debitopago.datapagamento, numdocumento.valorpago "
Sql = Sql & "HAVING (debitopago.datapagamento < CONVERT(DATETIME, '2016-12-21 00:00:00', 102)) AND (numdocumento.valorpago > 0) ORDER BY debitopago.datapagamento"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !DataPagamento < CDate("08/15/2016") Then GoTo proximo
        nTotal = nTotal + !ValorPago
        nNumDoc = !NumDocumento
        ax = nNumDoc & " " & FillLeft(FormatNumber(!valordoc, 2), 10) & "     " & FillLeft(!CODREDUZIDO, 6) & "  " & Format(!DataPagamento, "dd/mm/yyyy")
        Print #1, ax
proximo:
       .MoveNext
    Loop
   .Close
End With
Print #1, ""
Print #1, "----------------------------------"
Print #1, "Total pago: R$ " & FormatNumber(nTotal, 2)
Close #1
z = Shell("NOTEPAD" & " " & sPathBin & "\RefisDAM.txt", vbNormalFocus)
Liberado
End Sub

Public Sub DocEmitido(sDataIni As String, sDataFim As String, sUser As String)
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, nNumDoc As Long, ax As String, z As Long, nTotal As Double
nTotal = 0
Ocupado
Open sPathBin & "\RefisDAM.txt" For Output As #1
Print #1, "RELATÓRIO DO REFIS (PAGAMENTO À VISTA)"
Print #1, "REFIS - (" & nAno & ") - Impresso em " & Format(Now, "dd/mm/yyyy")
Print #1, "-------------------------------------------------"
Print #1, " "
Print #1, "Documento     Valor     Código   Dt.Pagto."
Print #1, "-------------------------------------------------"
Print #1, " "

Sql = "SELECT DISTINCT parceladocumento.numdocumento, debitoparcela.codreduzido, debitopago.datapagamento, SUM(debitopago.valorpagoreal) AS valorpago,numdocumento.valorpago AS valordoc "
Sql = Sql & "FROM numdocumento INNER JOIN parceladocumento ON numdocumento.numdocumento = parceladocumento.numdocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
Sql = Sql & "debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitopago.NumParcela And debitoparcela.CODCOMPLEMENTO = debitopago.CODCOMPLEMENTO WHERE (debitoparcela.codlancamento NOT IN (36, 65)) AND (parceladocumento.plano = 11 OR "
Sql = Sql & "parceladocumento.plano = 12) AND (debitoparcela.datavencimento < CONVERT(DATETIME, '2016-01-01 00:00:00', 102)) GROUP BY parceladocumento.numdocumento, debitoparcela.codreduzido, debitopago.datapagamento, numdocumento.valorpago "
Sql = Sql & "HAVING (debitopago.datapagamento < CONVERT(DATETIME, '2016-12-21 00:00:00', 102)) AND (numdocumento.valorpago > 0) ORDER BY debitopago.datapagamento"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !DataPagamento < CDate("08/15/2016") Then GoTo proximo
        nTotal = nTotal + !ValorPago
        nNumDoc = !NumDocumento
        ax = nNumDoc & " " & FillLeft(FormatNumber(!valordoc, 2), 10) & "     " & FillLeft(!CODREDUZIDO, 6) & "  " & Format(!DataPagamento, "dd/mm/yyyy")
        Print #1, ax
proximo:
       .MoveNext
    Loop
   .Close
End With
Print #1, ""
Print #1, "----------------------------------"
Print #1, "Total pago: R$ " & FormatNumber(nTotal, 2)
Close #1
z = Shell("NOTEPAD" & " " & sPathBin & "\RefisDAM.txt", vbNormalFocus)
Liberado
End Sub

Public Function IsAtendente() As Boolean
Dim bRet As Boolean

bRet = False

If NomeDeLogin = "BRUNO.MASCARO" Or NomeDeLogin = "NAIARA.SOUZA" Or NomeDeLogin = "RAFAEL.CANOLI" Or NomeDeLogin = "FERNANDO.MEDALHA" Or NomeDeLogin = "RENATA" Or _
   NomeDeLogin = "RAFAEL.OLIVEIRA" Or NomeDeLogin = "ANDRE.FRANCA" Or NomeDeLogin = "MICHELLE.POLETTI" Or NomeDeLogin = "NATALIA.FRACASSO" Or _
   NomeDeLogin = "MICHELE.OLIVEIRA" Or NomeDeLogin = "GABRIEL.MARQUES" Or NomeDeLogin = "ISRAEL" Or NomeDeLogin = "IORIO" Or NomeDeLogin = "POLYANA.TAVARES" Or _
   NomeDeLogin = "ISRAEL.ACHE" Or NomeDeLogin = "MIRELA.ASSONI" Or NomeDeLogin = "CARLA.REMIRO" Then
    bRet = True
End If

IsAtendente = bRet

End Function

Public Function GetHexadecimal(ByVal Binary As String) As String
Dim p As Long, m As Long, dec As Long
While p < Len(Binary)
m = InStr(1, "01", Mid(Binary, p + 1, 1)) - 1
If m >= 0 Then
dec = dec * 2 + m
End If
p = p + 1
Wend
GetHexadecimal = Hex(dec)
End Function


Public Function BitId(nPos As Integer)

BitId = Val(Mid(SecId, nPos, 1))

End Function
