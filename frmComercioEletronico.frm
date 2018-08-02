VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmComercioEletronico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comércio Eletrônico"
   ClientHeight    =   8175
   ClientLeft      =   6645
   ClientTop       =   5085
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10875
   Begin SHDocVwCtl.WebBrowser WBrowser 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   10935
      ExtentX         =   19288
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   0
      Picture         =   "frmComercioEletronico.frx":0000
      Top             =   0
      Width           =   10920
   End
End
Attribute VB_Name = "frmComercioEletronico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const conSwNormal = 1

Dim sBoletoNome As String
Dim sBoletoEndereco As String
Dim sBoletoVencto As String
Dim sBoletoCpfCnpj As String
Dim sBoletoNumDoc As String
Dim sBoletoValor As String
Dim sBoletoCidade As String
Dim sBoletoUF As String
Dim sBoletoCep As String
Dim sBoletoUser As String

Public Property Let BoletoUser(sNome As String)
    sBoletoUser = sNome
End Property

Public Property Let BoletoNome(sNome As String)
    sBoletoNome = sNome
End Property

Public Property Let BoletoEndereco(sEndereco As String)
    sBoletoEndereco = sEndereco
End Property

Public Property Let BoletoVencto(sData As String)
    sBoletoVencto = sData
End Property

Public Property Let BoletoCpfCnpj(sDoc As String)
    sBoletoCpfCnpj = sDoc
End Property

Public Property Let BoletoNumDoc(nNumDoc As Long)
    sBoletoNumDoc = "287353200" & Format(nNumDoc, "00000000")
End Property

Public Property Let BoletoValor(nValor As Double)
    sBoletoValor = FormatNumber(nValor, 2)
End Property

Public Property Let BoletoCidade(sCidade As String)
    sBoletoCidade = sCidade
End Property

Public Property Let BoletoUF(sUF As String)
    sBoletoUF = sUF
End Property

Public Property Let BoletoCep(sCep As String)
    sBoletoCep = sCep
End Property


Private Sub Form_Load()

Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String, v8 As String, v9 As String, V10 As String
v1 = sBoletoNome
v2 = sBoletoEndereco
v3 = sBoletoVencto
v4 = sBoletoCpfCnpj
v5 = sBoletoNumDoc
v6 = sBoletoValor
v7 = sBoletoCidade
v8 = sBoletoUF
v9 = sBoletoCep
V10 = sBoletoUser

Centraliza Me
Me.Top = Me.Top + 1300

 ShellExecute hwnd, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Pages/boletoBB.aspx?f1=" & v1 & "&f2=" & v2 & "&f3=" & v3 & "&f4=" & v4 & "&f5=" & v5 & "&f6=" & v6 & "&f7=" & v7 & "&f8=" & v8 & "&f9=" & v9 & "&f10=" & V10, vbNullString, vbNullString, conSwNormal
 'ShellExecute hwnd, "open", "http://www.vbmania.com.br", vbNullString, vbNullString, conSwNormal
'WBrowser.Navigate "http://sistemas.jaboticabal.sp.gov.br/gti/Pages/boletoBB.aspx?f1=" & v1 & "&f2=" & v2 & "&f3=" & v3 & "&f4=" & v4 & "&f5=" & v5 & "&f6=" & v6 & "&f7=" & v7 & "&f8=" & v8 & "&f9=" & v9 & "&f10=" & V10


End Sub

