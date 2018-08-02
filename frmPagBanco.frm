VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmPagBanco 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa Automática"
   ClientHeight    =   5520
   ClientLeft      =   3465
   ClientTop       =   3705
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8805
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Layout"
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   6615
      TabIndex        =   31
      Top             =   1035
      Width           =   1995
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Novo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   33
         Top             =   270
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Antigo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   32
         Top             =   270
         Width           =   780
      End
   End
   Begin prjChameleon.chameleonButton cmdShow 
      Height          =   315
      Left            =   4785
      TabIndex        =   20
      ToolTipText     =   "Visualizar Arquivo na Origem"
      Top             =   1290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Visualizar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPagBanco.frx":0000
      PICN            =   "frmPagBanco.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstDoc 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   2595
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      ToolTipText     =   "Reativação do Arquivo"
      Top             =   3840
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Reativação"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPagBanco.frx":0176
      PICN            =   "frmPagBanco.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBaixa 
      Height          =   315
      Left            =   150
      TabIndex        =   13
      ToolTipText     =   "Efetuar Baixa"
      Top             =   3840
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "E&fetuar Baixa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPagBanco.frx":02EC
      PICN            =   "frmPagBanco.frx":0308
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   3000
      TabIndex        =   14
      ToolTipText     =   "Sair da Tela"
      Top             =   3840
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Sair"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPagBanco.frx":03A7
      PICN            =   "frmPagBanco.frx":03C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   105
      Left            =   5790
      TabIndex        =   9
      Top             =   4050
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdReg 
      Height          =   1965
      Left            =   30
      TabIndex        =   6
      Top             =   1770
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   3466
      _Version        =   393216
      Rows            =   20
      Cols            =   20
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmPagBanco.frx":0431
   End
   Begin MSFlexGridLib.MSFlexGrid grdArq 
      Height          =   765
      Left            =   2790
      TabIndex        =   1
      Top             =   120
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   12640511
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   " |Arquivos Disponíveis                      "
   End
   Begin VB.PictureBox ImgLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   2685
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
   Begin MSFlexGridLib.MSFlexGrid grdParc 
      Height          =   1155
      Left            =   60
      TabIndex        =   16
      Top             =   5640
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   2037
      _Version        =   393216
      Rows            =   1
      Cols            =   17
      FixedCols       =   0
      BackColorSel    =   12582912
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmPagBanco.frx":0560
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   1155
      Left            =   -15
      TabIndex        =   17
      Top             =   6750
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   2037
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   12582912
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Codigo        |>Vl.Lançado  |>Vl.Multa     |>Vl.Juros       |>Vl.Correção    |>Vl.Total                |>Linha "
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Remessa:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   30
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Convênio:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   29
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Geração:"
      Height          =   255
      Index           =   2
      Left            =   2355
      TabIndex        =   28
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Sequencial....:"
      Height          =   255
      Index           =   3
      Left            =   2355
      TabIndex        =   27
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Versão Layout.....:"
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   26
      Top             =   1065
      Width           =   1335
   End
   Begin VB.Label lblCR 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1440
      TabIndex        =   25
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label lblCC 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1440
      TabIndex        =   24
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label lblDG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3735
      TabIndex        =   23
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label lblNS 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3735
      TabIndex        =   22
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label lblVL 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   6180
      TabIndex        =   21
      Top             =   1065
      Width           =   675
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos não Encontrados"
      Height          =   195
      Left            =   150
      TabIndex        =   19
      Top             =   4230
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "]"
      Height          =   225
      Index           =   1
      Left            =   8610
      TabIndex        =   12
      Top             =   3990
      Width           =   45
   End
   Begin VB.Label Label5 
      Caption         =   "["
      Height          =   225
      Index           =   0
      Left            =   5730
      TabIndex        =   11
      Top             =   3990
      Width           =   45
   End
   Begin VB.Label lblPb 
      BackColor       =   &H00EEEEEE&
      Height          =   225
      Left            =   5850
      TabIndex        =   10
      Top             =   3810
      Width           =   2655
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6930
      TabIndex        =   8
      Top             =   90
      Width           =   1725
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco..............:"
      Height          =   225
      Left            =   5700
      TabIndex        =   7
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label lblRegTot 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6960
      TabIndex        =   5
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblNumReg 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6960
      TabIndex        =   4
      Top             =   390
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total........:"
      Height          =   225
      Left            =   5700
      TabIndex        =   3
      Top             =   660
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Registros:"
      Height          =   225
      Left            =   5700
      TabIndex        =   2
      Top             =   375
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   2700
      Top             =   0
      Width           =   6075
   End
End
Attribute VB_Name = "frmPagBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FebrabanA
    CodigoRegistro As String * 1
    CodigoRemessa As String * 1
    CodigoConvenio As String * 20
    NomeEmpresa As String * 20
    CodigoBanco As String * 3
    NomeBanco As String * 20
    DataGeracao As String * 8
    NumeroSeq As String * 6
    VersaoLayout As String * 2
    Filler As String * 69
End Type
Private Type FebrabanA2
    CodigoBanco As String * 3
    NumeroLote As String * 4
    TipoRegistro As String * 1
    Filler1 As String * 8
    TipoInscricao As String * 1
    NumeroInscricao As String * 15
    Agencia As String * 5
    NumeroConta As String * 10
    Filler2 As String * 5
    CodigoCedente As String * 9
    Filler3 As String * 11
    NomeEmpresa As String * 30
    NomeBanco As String * 30
    Filler4 As String * 10
    CodigoRetorno As String * 1
    DataGeracao As String * 8 'DDMMAAAA
    Filler5 As String * 6
    NumeroSeq As String * 6
    NumeroVersao As String * 3
    Filler6 As String * 74
End Type


Private Type FebrabanA3
    CodigoRegistro As String * 2
    CodigoRetorno As String * 7
    CodigoCobrança As String * 2
    NomeCobrança As String * 8
    CodigoPrefeitura As String * 11
    NomePrefeitura As String * 26
    CodigoBanco As String * 3
    NomeBanco As String * 7
    DataCredito As String * 6
    Filler1 As String * 4
    Filler2 As String * 38
    NumRegistros As String * 6
End Type

Private Type SimplesNacionalHeader
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    CodigoConvenio As String * 20
    DataGeracao As String * 8
    NumeroRemessa As String * 6
    NumeroVersao As String * 2
    Filler1 As String * 22
    Filler2 As String * 8
    CodigoBanco As String * 3
    Filler3 As String * 422
End Type
Private Type SimplesNacionalDetalhe
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    DataArrecada As String * 8
    DataVencimento As String * 8
    Filler1 As String * 12
    Filler2 As String * 37
    Cnpj As String * 14
    Filler3 As String * 11
    Esfera As String * 1
    Competencia As String * 6
    ValorPrincipal As String * 17
    ValorMulta As String * 17
    ValorJuros As String * 17
    Filler4 As String * 47
    ValorAutentica As String * 17
    NumeroAutentica As String * 23
    CodigoBanco As String * 3
    CodigoAgencia As String * 4
    Filler6 As String * 249
End Type
Private Type SimplesNacionalTrailer
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    TotalRegistro As String * 6
    ValorRegistro As String * 17
    Filler As String * 468
End Type
Private Type FebrabanG
   CodigoRegistro As String * 1
   ContaPrefeitura As String * 20
   DataPagamento As String * 8
   DataCredito As String * 8
   PreCodBarra As String * 4
   ValorRecebido As String * 11
   CodigoMunic As String * 4
   DataVencto As String * 8
   NumDocumento As String * 9
   NumParcela As String * 2
   SituacaoRetorno As String * 2
   FillerSmar As String * 4
   ValorRetornado As String * 12
   ValorTarifa As String * 7
   NumSeq As String * 8
   CodAgencia As String * 8
   FormaPagamento As String * 1
   NumAutentica As String * 23
   Filler As String * 10
End Type

Private Type FebrabanG2
   CodigoBanco As String * 3
   NumeroLote As String * 4
   TipoRegistro As String * 1
   NumSequencial As String * 5
   CodigoSegmento As String * 1
   Filler1 As String * 1
   CodigoMovimento As String * 2
   Agencia As String * 5
   NumeroConta As String * 10
   Filler2 As String * 8
   NossoNumero As String * 13
   CodigoCarteira As String * 1
   NumDocumento As String * 15
   DataVencto As String * 8 'DDMMAAAA
   ValorTitulo As String * 15
   BancoCobrador As String * 3
   AgenciaCobradora As String * 5
   UsoCedente As String * 25
   CodigoMoeda As String * 2
   TipoInscricao As String * 1 '1=CPF, 2=CNPJ
   NumeroInscricao As String * 15
   NomeSacado As String * 40
   ContaCobranca As String * 10
   ValorTarifa As String * 15
   Custas As String * 10
   Filler3 As String * 22
End Type

Private Type FebrabanG2U
    CodigoBanco As String * 3
    NumeroLote As String * 4
    TipoRegistro As String * 1
    NumeroSeq As String * 5
    CodSegmento As String * 1
    Filler1 As String * 1
    CodigoMov As String * 2
    JurosMulta As String * 15
    ValorDesconto As String * 15
    ValorAbatimento As String * 15
    ValorIOF As String * 15
    ValorPago As String * 15
    ValorCreditado As String * 15
    ValorOutrasDespesas As String * 15
    ValorOutrosCreditos As String * 15
    DataOcorrencia As String * 8 'DDMMAAAA
    DataCredito As String * 8 'DDMMAAAA
    Outros As String * 87
End Type

Private Type FebrabanG3
   CodigoRegistro As String * 1
   ContaPrefeitura As String * 11
   CodAgencia As String * 3
   NumDocumento As String * 8
   Codigo06 As String * 2
   DataPagamento As String * 6
   CodigoBanco As String * 5
   ValorTaxa As String * 13
   Filler1 As String * 26
   ValorPago As String * 13
   ValorSeiLa As String * 13
   CodigoC As String * 1
   Filler2 As String * 12
   NumSeq As String * 6
End Type

Private Type FebrabanZ
    CodigoRegistro As String * 1
    TotalRegistro As String * 6
    ValorTotal As String * 17
    Filler As String * 126
End Type

Private Type FebrabanZ2
    CodigoBanco As String * 3
    NumeroLote As String * 4
    TipoRegistro As String * 1
    QtdeRegistros As String * 6
    QtdeTitulos As String * 6
    ValorTotal As String * 17
    Outros As String * 203
End Type

Private Type FebrabanZ3
    CodigoSeiLa As String * 7
    Filler1 As String * 3
    TotalRegistro As String * 6
    ValorTotal As String * 14
    CodigoC As String * 1
    DataCredito As String * 6
    Filler2 As String * 77
    NumSeq As String * 6
End Type

Dim aFebrabanA() As FebrabanA
Dim aFebrabanG() As FebrabanG
Dim aFebrabanZ() As FebrabanZ
Dim aFebrabanA2() As FebrabanA2
Dim aFebrabanA3() As FebrabanA3
Dim aFebrabang2U() As FebrabanG2U
Dim aFebrabanG2() As FebrabanG2
Dim aFebrabanG3() As FebrabanG3
Dim aFebrabanZ2() As FebrabanZ2
Dim aFebrabanZ3() As FebrabanZ3
Dim aSimplesH() As SimplesNacionalHeader
Dim aSimplesD() As SimplesNacionalDetalhe
Dim aSimplesT() As SimplesNacionalTrailer

Private Sub cmdBaixa_Click()
Dim sFullPath As String
Dim HeaderSN As SimplesNacionalHeader
Dim RegistroSN As SimplesNacionalDetalhe
Dim FooterSN As SimplesNacionalTrailer, nValorGuia As Double, nNumDoc As Long
Dim Posicao  As Long, RdoAux2 As rdoResultset, nSeq As Integer, bAchou As Boolean, nNumParc As Integer, nCompl As Integer, nAnoCompetencia As Integer
Dim nCount As Integer, RdoAux As rdoResultset, Sql As String, nCodReduz As Long, sVencimento As String, dDataVencto As Date, RdoAux3 As rdoResultset
Dim RdoAux4 As rdoResultset
Dim qd As New rdoQuery

If grdReg.Rows = 1 Then
    MsgBox "Não existem registros a baixar.", vbExclamation, "Atenção"
    Exit Sub
End If
If Val(lblBanco.Caption) = 0 Then grdArq_Click

If Val(lblBanco.Caption) = 0 Then
    MsgBox "Código do Banco inválido.", vbExclamation, "Atenção"
    Exit Sub
End If

If Left$(grdArq.TextMatrix(grdArq.Row, 1), 2) = "BD" Then
   Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
Else
   Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If Not IsNull(!DATABAIXA) Then
           MsgBox "Ja foi efetuado Baixa neste arquivo.", vbCritical, "Atenção"
          .Close
           Exit Sub
        End If
    End If
   .Close
End With

If MsgBox("Deseje efetuar a Baixa nas Parcelas ?", vbQuestion + vbYesNo, "CONFIRMAÇÃO DE BAIXA") = vbNo Then Exit Sub


'***********************************************************************
'SE FOR ARQUIVO DO SIMPLES VERIFICA A CRIAÇÀO DE DÉBITOS E DOCUMENTOS
sFullPath = grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1)
If InStr(1, sFullPath, "DAF607", vbBinaryCompare) > 0 Then
SIMPLES:
ReDim aSimplesH(0): ReDim aSimplesD(0): ReDim aSimplesT(0)
Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, HeaderSN
    aSimplesH(0).CodigoRegistro = Trim$(HeaderSN.CodigoRegistro)
    aSimplesH(0).SeqRegistro = Trim$(HeaderSN.SeqRegistro)
    aSimplesH(0).CodigoConvenio = Trim$(HeaderSN.CodigoConvenio)
    aSimplesH(0).DataGeracao = Trim$(HeaderSN.DataGeracao)
    aSimplesH(0).NumeroRemessa = Trim$(HeaderSN.NumeroRemessa)
    aSimplesH(0).NumeroVersao = Trim$(HeaderSN.NumeroVersao)
    aSimplesH(0).Filler1 = Trim$(HeaderSN.Filler1)
    aSimplesH(0).Filler2 = Trim$(HeaderSN.Filler2)
    aSimplesH(0).CodigoBanco = Trim$(HeaderSN.CodigoBanco)
    aSimplesH(0).Filler3 = Trim$(HeaderSN.Filler3)

    Posicao = Len(HeaderSN) + 3
    nCount = 0
    Do While Not EOF(1)
         Get #1, Posicao, RegistroSN
         If RegistroSN.SeqRegistro <> "99999999" Then
              aSimplesD(nCount).CodigoRegistro = RegistroSN.CodigoRegistro
              aSimplesD(nCount).SeqRegistro = RegistroSN.SeqRegistro
              aSimplesD(nCount).DataArrecada = RegistroSN.DataArrecada
              aSimplesD(nCount).DataVencimento = RegistroSN.DataVencimento
              aSimplesD(nCount).Filler1 = RegistroSN.Filler1
              aSimplesD(nCount).Filler2 = RegistroSN.Filler2
              aSimplesD(nCount).Cnpj = RegistroSN.Cnpj
              aSimplesD(nCount).Filler3 = RegistroSN.Filler3
              aSimplesD(nCount).Esfera = RegistroSN.Esfera
              aSimplesD(nCount).Competencia = RegistroSN.Competencia
              aSimplesD(nCount).ValorPrincipal = RegistroSN.ValorPrincipal
              aSimplesD(nCount).ValorMulta = RegistroSN.ValorMulta
              aSimplesD(nCount).ValorJuros = RegistroSN.ValorJuros
              aSimplesD(nCount).Filler4 = RegistroSN.Filler4
              aSimplesD(nCount).ValorAutentica = RegistroSN.ValorAutentica
              aSimplesD(nCount).NumeroAutentica = RegistroSN.NumeroAutentica
              aSimplesD(nCount).CodigoBanco = RegistroSN.CodigoBanco
              aSimplesD(nCount).CodigoAgencia = RegistroSN.CodigoAgencia
              nAnoCompetencia = Val(Left(aSimplesD(nCount).Competencia, 4))
              nValorGuia = CDbl(aSimplesD(nCount).ValorAutentica) / 100
              'BUSCA CÓDIGO
              Sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE CONVERT(BIGINT, cnpj) = " & Val(aSimplesD(nCount).Cnpj) & " AND DATAENCERRAMENTO IS NULL ORDER BY CODIGOMOB DESC"
              Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
              With RdoAux
                  If .RowCount > 0 Then
                      nCodReduz = !CODIGOMOB
                     'BUSCA VENCIMENTO
                      sVencimento = Right(aSimplesD(nCount).DataVencimento, 2) & "/" & Mid(aSimplesD(nCount).DataVencimento, 5, 2) & "/" & Left(aSimplesD(nCount).DataVencimento, 4)
                      dDataVencto = CDate(sVencimento)
                      nNumDoc = 0
                     'BUSCA LANCAMENTO
                      Sql = "SELECT debitoparcela.codreduzido, debitoparcela.codlancamento,DEBITOPARCELA.SEQLANCAMENTO,debitoparcela.numparcela,DEBITOPARCELA.CODCOMPLEMENTO,debitoparcela.datavencimento, debitoparcela.statuslanc, debitotributo.valortributo "
                      Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
                      Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And "
                      Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND (debitoparcela.codlancamento = 5) AND (MONTH(debitoparcela.datavencimento) = " & Month(dDataVencto) & ") AND "
                      Sql = Sql & "(YEAR(debitoparcela.datavencimento) = " & Year(dDataVencto) & ") AND (debitotributo.codtributo = 13)"
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      With RdoAux2
                         'EXISTE LANCAMENTO NESTE MÊS/ANO?
                          If .RowCount > 0 Then 'SIM
                              nNumParc = !NumParcela 'CAPTURA A PARCELA
                              bAchou = False
                             'TEM ALGUM QUE NÃO ESTA PAGO?
                              Do Until .EOF
                                  If !statuslanc = 3 Then
                                      bAchou = True
                                      Exit Do
                                  End If
                                 .MoveNext
                              Loop
                             'SE ACHOU PEGA A PARCELA
                              If bAchou Then
                                 nSeq = !SeqLancamento
                                 nCompl = !CODCOMPLEMENTO '---------------> PARCELA PRONTA PARA USO
                                'BUSCA O DOCUMENTO DELA
                                 Sql = "SELECT * FROM PARCELADOCUMENTO WHERE Codreduzido = " & nCodReduz & " AND ANOEXERCICIO=" & nAnoCompetencia & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                                 Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                                 Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                 With RdoAux3
                                     If .RowCount > 0 Then
                                         nNumDoc = !NumDocumento
                                     Else
                                        'SE NÃO ENCONTRAR CRIA APENAS O DOCUMENTO
                                         GoTo Documento
                                     End If
                                    .Close
                                 End With
                                 GoTo GRID
                              Else
                                'SE NÃO ACHAR
                                .MoveFirst
                                 nCompl = 0
                                'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
                                 'Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (MONTH(datavencimento) = " & Month(dDataVencto) & ") AND "
                                 Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5)  AND "
                                 Sql = Sql & "(YEAR(datavencimento) = " & Year(dDataVencto) & ")"
                                 Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                 With RdoAux3
                                   If IsNull(!MAXIMO) Then
                                       nSeq = 0
                                   Else
                                       nSeq = !MAXIMO + 1
                                   End If
                                  .Close
                                 End With
                              End If
                          Else
                             'NÃO ACHOU LANCAMENTOS NESTE MÊS/ANO
                             'O NÚMERO DA PARCELA A SER CRIADA SERÁ O ÚLTIMO NÚMERO DE PARCELA DO ANO
                              Sql = "SELECT MAX(NUMPARCELA) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & nAnoCompetencia & ")"
                              Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                              With RdoAux3
                                If IsNull(!MAXIMO) Then
                                    nNumParc = 1
                                Else
                                    nNumParc = !MAXIMO + 1
                                End If
                               .Close
                              End With
                              nCompl = 0
                              nSeq = 0
                          End If

                         'CRIAR PARCELA DE ISS VARIAVEL NESTE MES E ANO COM O VENCIMENTO QUE VEIO DO BANCO
                          Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
                          Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
                          Sql = Sql & nNumParc & "," & nCompl & ",3,'" & Format(dDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0)"
                          cn.Execute Sql, rdExecDirect
                         'CRIAR O TRIBUTO PARA ELA (13 - iss variavel)
                          Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                          Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
                          Sql = Sql & nNumParc & "," & nCompl & "," & 13 & "," & 0 & ")"
                          cn.Execute Sql, rdExecDirect
Documento:
                         'CRIAR O DOCUMENTO PARA ELA
                          Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
                          Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                          With RdoAux3
                               nNumDoc = !MAXIMO + 1
                              .Close
                          End With
                          Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA) VALUES(" & nNumDoc & ",'"
                          Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & ")"
                          cn.Execute Sql, rdExecDirect
                         'CRIAR A PARCELADOCUMENTO
                          Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
                          Sql = Sql & nNumParc & "," & nCompl & "," & nNumDoc & ")"
                          cn.Execute Sql, rdExecDirect

                         .Close
                      End With
                  Else
                    MsgBox "O CNPJ " & aSimplesD(nCount).Cnpj & " NÃO FOI LOCALIZADO.", vbCritical, "Atenção"
                  End If
                 .Close
              End With

GRID:
             'AGORA QUE JA POSSUIMOS O DOCUMENTO A SER EFETUADA A BAIXA
             'PODEMOS POPULAR O GRID
             For x = 1 To grdReg.Rows - 1
                If Val(grdReg.TextMatrix(x, 8)) = nCodReduz And Val(grdReg.TextMatrix(x, 9)) = nNumParc Then
                    grdReg.TextMatrix(x, 8) = nNumDoc
                    Exit For
                End If
             Next


         Else
              Get #1, Posicao, FooterSN
              aSimplesT(0).CodigoRegistro = FooterSN.CodigoRegistro
              aSimplesT(0).TotalRegistro = FooterSN.TotalRegistro
              aSimplesT(0).ValorRegistro = FooterSN.ValorRegistro
              lblNumReg.Caption = Val(aSimplesT(0).TotalRegistro) - 2
              lblRegTot.Caption = "R$ " & FormatNumber(aSimplesT(0).ValorRegistro / 100, 2)
              Exit Do
         End If
         Posicao = Posicao + Len(RegistroSN) + 2
         nCount = nCount + 1
         ReDim Preserve aSimplesD(nCount)
    Loop
Close #1


End If

'************************************************************************




Sql = "DELETE FROM BAIXATMP WHERE COMPUTADOR='" & Trim$(NomeDoComputador) & "'"
cn.Execute Sql, rdExecDirect

Ocupado
Pb.Value = 0
lblPb.Caption = "Efetuando Baixa"
lblPb.Refresh
MontaResumo
Pb.Value = 0

Set qd.ActiveConnection = cn
Sql = "SELECT * From RECEITACLASSIFICAR WHERE DATARECEITA = '" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "' AND CODBANCO = " & Val(Left$(lblBanco.Caption, 3))
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
    Do Until .EOF
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
        
        qd.Sql = "{ Call spGRAVABAIXATMP(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = Trim$(NomeDoComputador) 'COMPUTADOR
        qd(1) = grdArq.TextMatrix(grdArq.Row, 1) 'ARQUIVO
        qd(2) = NomeDeLogin 'USUARIO
        qd(3) = Trim$(lblBanco.Caption) 'BANCO
        qd(4) = lblNumReg.Caption 'NUM REGS
        qd(5) = Virg2Ponto(Mid(RemovePonto(lblRegTot.Caption), 4, Len(lblRegTot.Caption) - 2)) 'TOTAL REG
        qd(6) = lblCR.Caption 'REMESSA
        qd(7) = lblCC.Caption 'CONVENIO
        qd(8) = Format(lblDG.Caption, "mm/dd/yyyy") 'DATA GERACAO
        qd(9) = lblNS.Caption 'NUM SEQ
        qd(10) = lblVL.Caption 'VERSAO LAYOUT
        qd(11) = grdReg.TextMatrix(1, 0) 'TR
        qd(12) = Trim$(grdReg.TextMatrix(1, 1)) 'CONTA PREFEITURA
        qd(13) = Format(grdReg.TextMatrix(1, 2), "mm/dd/yyyy") 'DATA PAGTO
        qd(14) = Format(!DataReceita, "mm/dd/yyyy") 'DATA CREDITO
        qd(15) = !NumDocumento & "-" & RetornaDVNumDoc(!NumDocumento)  'NUM DOC
        qd(16) = Null 'CODREDUZ
        qd(17) = Null 'EXERCICIO
        qd(18) = Null 'LANCAMENTO
        qd(19) = Null 'SEQLANCAMENTO
        qd(20) = Null 'PARCELA
        qd(21) = Null 'COMPLEMENTO
        qd(22) = Null
        qd(23) = 0 'VALOR LANCADO
        qd(24) = 0 'VALOR JUROS
        qd(25) = 0 'VALOR MULTA
        qd(26) = 0 'VALOR CORRECAO
        qd(27) = 0 'VALOR CALCULADO
        qd(28) = 0 'VALOR DIF
        qd(29) = Virg2Ponto(!ValorTotal) 'VALOR PAGO
        qd(30) = 0 'VALOR TARIFA
        qd(31) = Virg2Ponto(!ValorTotal) 'VALOR PAGO REAL
        qd(32) = "Não Existe" 'SITUACAO
        qd(33) = grdReg.TextMatrix(1, 14)
        qd(34) = 0
        qd(35) = 0
        qd(36) = Null
        qd(37) = Null
        Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)
       .MoveNext
    Loop
End With

Sql = "SELECT * FROM DEBITOCLASSIFICAR WHERE DATARECEITA = '" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "' AND CODBANCO = " & Val(Left$(lblBanco.Caption, 3))
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
    Do Until .EOF
        
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
        qd.Sql = "{ Call spGRAVABAIXATMP(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = Trim$(NomeDoComputador) 'COMPUTADOR
        qd(1) = grdArq.TextMatrix(grdArq.Row, 1) 'ARQUIVO
        qd(2) = NomeDeLogin 'USUARIO
        qd(3) = Trim$(lblBanco.Caption) 'BANCO
        qd(4) = lblNumReg.Caption 'NUM REGS
        qd(5) = Virg2Ponto(Mid(RemovePonto(lblRegTot.Caption), 4, Len(lblRegTot.Caption) - 2)) 'TOTAL REG
        qd(6) = lblCR.Caption 'REMESSA
        qd(7) = lblCC.Caption 'CONVENIO
        qd(8) = Format(lblDG.Caption, "mm/dd/yyyy") 'DATA GERACAO
        qd(9) = lblNS.Caption 'NUM SEQ
        qd(10) = lblVL.Caption 'VERSAO LAYOUT
        qd(11) = grdReg.TextMatrix(1, 0) 'TR
        qd(12) = Trim$(grdReg.TextMatrix(1, 1)) 'CONTA PREFEITURA
        qd(13) = Format(grdReg.TextMatrix(1, 2), "mm/dd/yyyy") 'DATA PAGTO
        qd(14) = Format(!DataReceita, "mm/dd/yyyy") 'DATA CREDITO
        qd(15) = !NumDocumento & "-" & RetornaDVNumDoc(!NumDocumento) 'NUM DOC
        qd(16) = !CODREDUZIDO 'CODREDUZ
        qd(17) = !AnoExercicio 'EXERCICIO
        qd(18) = !CodLancamento 'LANCAMENTO
        qd(19) = !SeqLancamento 'SEQLANCAMENTO
        qd(20) = !NumParcela 'PARCELA
        qd(21) = !CODCOMPLEMENTO 'COMPLEMENTO
        qd(22) = Null
        qd(23) = 0 'VALOR LANCADO
        qd(24) = 0 'VALOR JUROS
        qd(25) = 0 'VALOR MULTA
        qd(26) = 0 'VALOR CORRECAO
        qd(27) = 0 'VALOR CALCULADO
        qd(28) = 0 'VALOR DIF
        If Not IsNull(!VALORCLASS) Then
            qd(29) = Virg2Ponto(!VALORCLASS) 'VALOR PAGO
        Else
            qd(29) = Virg2Ponto(0) 'VALOR PAGO
        End If
        qd(30) = 0 'VALOR TARIFA
        If Not IsNull(!VALORCLASS) Then
            qd(31) = Virg2Ponto(!VALORCLASS) 'VALOR PAGO REAL
        Else
            qd(31) = Virg2Ponto(0) 'VALOR PAGO REAL
        End If
        qd(32) = "Não Existente" 'SITUACAO
        qd(33) = grdReg.TextMatrix(1, 14)
        qd(34) = 0
        qd(35) = 0
        qd(36) = Null
        qd(37) = Null
        Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)
       .MoveNext
    Loop
End With

Liberado
MsgBox "Baixa efetuada.", vbInformation, "Confirmação"
'MostraRpt
Fim:

Pb.Value = 0
lblPb.Caption = "Pronto"
lblPb.Refresh
End Sub

Private Sub MostraRpt()

'EXIBE RELATORIO
frmReport.ShowReport "BAIXATMP2", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub cmdCancel_Click()
Dim RdoAux As rdoResultset, Sql As String

If grdReg.Rows = 1 Then
    MsgBox "Não existem registros a reativar.", vbExclamation, "Atenção"
    Exit Sub
End If
If Left$(grdArq.TextMatrix(grdArq.Row, 1), 2) = "BD" Then
   Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
Else
   Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If IsNull(!DATABAIXA) Then
           MsgBox "Não foi efetuado Baixa neste arquivo.", vbCritical, "Atenção"
          .Close
           Exit Sub
        End If
    End If
   .Close
End With

sFullPath = grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1)
If InStr(1, sFullPath, "DAF607", vbBinaryCompare) > 0 Then
    MsgBox "Arquivos do simples nào podem ser reativados automaticamente.", vbInformation, "Atenção"
    Exit Sub
End If


If MsgBox("Deseja REATIVAR os pagamentos deste arquivo ?", vbQuestion + vbYesNo, "CONFIRMAÇÃO DE REATIVAÇÃO") = vbYes Then
    Ocupado
    Pb.Value = 0
    lblPb.Caption = "Reativando Parcelas"
    lblPb.Refresh
    Reativa
    Pb.Value = 0
    lblPb.Caption = "Parcelas foram Canceladas"
    lblPb.Refresh
    Liberado
End If

End Sub

Private Sub EfetuaCancelamento()

Dim x As Integer
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqLanc As Integer
Dim nNumParc As Integer
Dim nCompl As Integer
Dim nStatus As Integer
Dim dDataPag As Date
Dim dDataVencto As Date
Dim sCodAgencia As String
Dim nQtdeDup As Integer

'EFETUA O CANCELAMENTO
With grdPag
    For x = 1 To .Rows - 1
        Pb.Value = Abs(x * 100 / .Rows - 1)
        If .TextMatrix(x, 1) <> "N/A" Then
             nCodReduz = .TextMatrix(x, 1)
             nAnoExercicio = .TextMatrix(x, 3)
             nCodLanc = .TextMatrix(x, 2)
             nSeqLanc = .TextMatrix(x, 4)
             nNumParc = .TextMatrix(x, 5)
             nCompl = .TextMatrix(x, 6)
             dDataPag = CDate(.TextMatrix(x, 18))
             dDataVencto = CDate(.TextMatrix(x, 21))
             nStatus = 3 'NÃO PAGO
             'BUSCA AGENCIA
             Sql = "SELECT AGCONTAPREF FROM BANCO WHERE CODBANCO=" & Val(Left$(lblBanco.Caption, 3))
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
             With RdoAux2
                    sCodAgencia = !AGCONTAPREF
                   .Close
             End With
            'CANCELA BAIXA NA TABELA NUMDOCUMENTO
             Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=0,CODAGENCIA ='0' , VALORPAGO=0 "
             Sql = Sql & " WHERE NUMDOCUMENTO=" & Val(.TextMatrix(x, 0))
             cn.Execute Sql, rdExecDirect
            'CANCELA BAIXA NA TABELA DEBITOPARCELA
             Sql = "SELECT CODREDUZIDO,QTDEDUPLICADO FROM DEBITOPARCELA  WHERE CODREDUZIDO=" & nCodReduz & " AND "
             Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
             Sql = Sql & "CODCOMPLEMENTO=" & nCompl
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                  nQtdeDup = Val(SubNull(!QTDEDUPLICADO))
                 .Close
             End With
             Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " ,PAGTODUPLICADO=" & IIf(nQtdeDup > 1, 1, 0) & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
             Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
             Sql = Sql & "CODCOMPLEMENTO=" & nCompl
             cn.Execute Sql, rdExecDirect
            'SE FOR PARCELA UNICA EFETUA CANCELAMENTO EM TODAS AS PARCELAS AUTOMATICAMENTO
             If nNumParc = 0 Then
                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=3 ,PAGTODUPLICADO=0 WHERE CODREDUZIDO=" & nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                Sql = Sql & " AND NUMPARCELA<>0"
                cn.Execute Sql, rdExecDirect
             End If
            'SE FOR DUPLICADO REDUZ O Nº DE VEZES
             If nQtdeDup > 0 Then
                Sql = "UPDATE DEBITOPARCELA SET QTDEDUPLICADO=QTDEDUPLICADO-1 "
                Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                Sql = Sql & " AND NUMPARCELA=" & nNumParc
                cn.Execute Sql, rdExecDirect
               'REMOVE DEBITO ADICIONAL
                Sql = "DELETE FROM DEBITOADICIONAL WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND DATAPAGAMENTO='" & Format(dDataPag, "mm/dd/yyyy") & "' AND "
                Sql = Sql & "DATARECEBIMENTO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
                cn.Execute Sql, rdExecDirect
             Else
                'EFETUA CANCELAMENTO NA TABELA DEBITOTRIBUTO
                 Sql = "UPDATE DEBITOTRIBUTO SET VALORCORRECAO=0 ,VALORMULTA=0 ,VALORJUROS=0 ,DATAPAGAMENTO='" & Null & "' ,DATARECEBIMENTO='" & Null & "',VALORPAGO=0"
                 Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                 Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                 cn.Execute Sql, rdExecDirect
             
                'ATUALIZA A TABELA DEBITOPAGO
                 Sql = "SELECT * FROM DEBITOPAGO "
                 Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                 Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS NULL"
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 If RdoAux.RowCount > 1 Then
                      Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' "
                      Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                      'Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND SEQPAG=" & nSeqPag
                      Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl
                 Else
                      Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' "
                      Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                      Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS NULL"
                 End If
                 cn.Execute Sql, rdExecDirect
                 RdoAux.Close
             End If
        End If
    Next
End With

'APAGA DA TABELA ARQUIVOBAIXA
Sql = "DELETE FROM ARQUIVOBAIXA WHERE NOMEARQUIVO='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
cn.Execute Sql, rdExecDirect

End Sub


Private Sub cmdSair_Click()

Unload Me
End Sub

Private Sub GravaBaixaTmp(nNumDoc As Long)
Dim qd As New rdoQuery
Dim x As Long
Dim sStatus As String
Dim nSomaTotal As Double
Dim nContaReg As Integer
Dim bDif As Boolean, nCodReduz As Long, nNumParc As Integer

Set qd.ActiveConnection = cn

nSomaTotal = 0

nSomaTotal = CDbl(lblRegTot.Caption)
nContaReg = Val(lblNumReg.Caption)


With grdParc
    If .Rows = 1 Then
        MsgBox "Documento não encontrado: " & nNumDoc
        Exit Sub
    End If
    If grdParc.TextMatrix(1, 0) <> "N/A" Then
        For x = 1 To .Rows - 1
            bDif = IIf(.TextMatrix(x, 15) = 0, False, True)
            bDup = IIf(.TextMatrix(x, 11) = "Não", False, True)
            nLinha = .TextMatrix(x, 16)
            nNumDoc = Val(Left$(grdReg.TextMatrix(nLinha, 8), Len(grdReg.TextMatrix(nLinha, 8)) - 1))
            
'            nNumDoc = Val(Left$(grdReg.TextMatrix(x, 1), Len(grdReg.TextMatrix(x, 8)) - 1))
            If Len(CStr(Val(grdReg.TextMatrix(nLinha, 8)))) < 8 Then
              nNumDoc = Val(grdReg.TextMatrix(nLinha, 8))
            End If
            
            If bDup Then
                sStatus = "DUPLICADO"
            Else
                If bDif Then
                    sStatus = "C/DIFERENÇA"
                Else
                    sStatus = "NORMAL"
                End If
            End If
            On Error Resume Next
            RdoAux.Close
            On Error GoTo 0
            qd.Sql = "{ Call spGRAVABAIXATMP(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = Trim$(NomeDoComputador) 'COMPUTADOR
            qd(1) = grdArq.TextMatrix(grdArq.Row, 1) 'ARQUIVO
            'qd(2) = Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) 'USUARIO
            qd(2) = NomeDeLogin 'USUARIO
            qd(3) = Trim$(lblBanco.Caption) 'BANCO
            qd(4) = lblNumReg.Caption 'NUM REGS
            qd(5) = Virg2Ponto(Mid(RemovePonto(lblRegTot.Caption), 4, Len(lblRegTot.Caption) - 2)) 'TOTAL REG
            qd(6) = lblCR.Caption 'REMESSA
            qd(7) = lblCC.Caption 'CONVENIO
            qd(8) = Format(lblDG.Caption, "mm/dd/yyyy") 'DATA GERACAO
            qd(9) = lblNS.Caption 'NUM SEQ
            qd(10) = lblVL.Caption 'VERSAO LAYOUT
            qd(11) = grdReg.TextMatrix(1, 0) 'TR
            qd(12) = Trim$(grdReg.TextMatrix(1, 1)) 'CONTA PREFEITURA
            qd(13) = Format(grdReg.TextMatrix(nLinha, 2), "mm/dd/yyyy") 'DATA PAGTO
            qd(14) = Format(grdReg.TextMatrix(nLinha, 3), "mm/dd/yyyy") 'DATA CREDITO
            qd(15) = nNumDoc & "-" & RetornaDVNumDoc(CStr(nNumDoc)) 'NUM DOC
            qd(16) = Format(.TextMatrix(x, 1), "0000000")
            qd(17) = .TextMatrix(x, 0) 'EXERCICIO
            qd(18) = Val(Left$(.TextMatrix(x, 2), 3)) 'LANCAMENTO
            qd(19) = .TextMatrix(x, 3) 'SEQLANCAMENTO
            qd(20) = .TextMatrix(x, 4) 'PARCELA
            qd(21) = .TextMatrix(x, 5) 'COMPLEMENTO
            qd(22) = Format(.TextMatrix(x, 12), "mm/dd/yyyy") 'DATA VENCTO
            qd(23) = Virg2Ponto(.TextMatrix(x, 6)) 'VALOR LANCADO
            qd(24) = Virg2Ponto(.TextMatrix(x, 8)) 'VALOR JUROS
            qd(25) = Virg2Ponto(.TextMatrix(x, 7)) 'VALOR MULTA
            qd(26) = Virg2Ponto(.TextMatrix(x, 9)) 'VALOR CORRECAO
            'qd(27) = Virg2Ponto(.TextMatrix(x, 13)) 'VALOR CALCULADO
            qd(27) = 0 'VALOR CALCULADO
            qd(28) = Virg2Ponto(.TextMatrix(x, 15)) 'VALOR DIF
            qd(29) = Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) 'VALOR PAGO
            qd(30) = Virg2Ponto(.TextMatrix(x, 14)) 'VALOR TARIFA
            qd(31) = Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) 'VALOR PAGO REAL
            qd(32) = sStatus 'SITUACAO
            qd(33) = grdReg.TextMatrix(nLinha, 14)
            qd(34) = Virg2Ponto(sTr(nSomaTotal)) 'SOMA TOTAL DO BANCO
            qd(35) = nContaReg
            qd(36) = Null
            qd(37) = Null
            Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)
        Next
    Else
        nLinha = .TextMatrix(1, 16)
        nNumDoc = Val(Left$(grdReg.TextMatrix(nLinha, 8), Len(grdReg.TextMatrix(nLinha, 8)) - 1))
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
        qd.Sql = "{ Call spGRAVABAIXATMP(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = Trim$(NomeDoComputador) 'COMPUTADOR
        qd(1) = grdArq.TextMatrix(grdArq.Row, 1) 'ARQUIVO
        'qd(2) = Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) 'USUARIO
        qd(2) = NomeDeLogin 'USUARIO
        qd(3) = Trim$(lblBanco.Caption) 'BANCO
        qd(4) = lblNumReg.Caption 'NUM REGS
        qd(5) = Virg2Ponto(Mid(RemovePonto(lblRegTot.Caption), 4, Len(lblRegTot.Caption) - 2)) 'TOTAL REG
        qd(6) = lblCR.Caption 'REMESSA
        qd(7) = lblCC.Caption 'CONVENIO
        qd(8) = Format(lblDG.Caption, "mm/dd/yyyy") 'DATA GERACAO
        qd(9) = lblNS.Caption 'NUM SEQ
        qd(10) = lblVL.Caption 'VERSAO LAYOUT
        qd(11) = grdReg.TextMatrix(1, 0) 'TR
        qd(12) = Trim$(grdReg.TextMatrix(1, 1)) 'CONTA PREFEITURA
        qd(13) = Format(grdReg.TextMatrix(nLinha, 2), "mm/dd/yyyy") 'DATA PAGTO
        qd(14) = Format(grdReg.TextMatrix(nLinha, 3), "mm/dd/yyyy") 'DATA CREDITO
        qd(15) = nNumDoc & "-" & RetornaDVNumDoc(CStr(nNumDoc)) 'NUM DOC
        qd(16) = Null 'CODREDUZ
        qd(17) = Null 'EXERCICIO
        qd(18) = Null 'LANCAMENTO
        qd(19) = Null 'SEQLANCAMENTO
        qd(20) = Null 'PARCELA
        qd(21) = Null 'COMPLEMENTO
        qd(22) = Format(grdReg.TextMatrix(nLinha, 7), "mm/dd/yyyy") 'DATA VENCTO
        qd(23) = 0 'VALOR LANCADO
        qd(24) = 0 'VALOR JUROS
        qd(25) = 0 'VALOR MULTA
        qd(26) = 0 'VALOR CORRECAO
        qd(27) = 0 'VALOR CALCULADO
        qd(28) = 0 'VALOR DIF
        qd(29) = Virg2Ponto(RemovePonto(grdReg.TextMatrix(nLinha, 5))) 'VALOR PAGO
        qd(30) = 0 'VALOR TARIFA
        qd(31) = Virg2Ponto(RemovePonto(grdReg.TextMatrix(nLinha, 5))) 'VALOR PAGO REAL
        qd(32) = "Não Existe" 'SITUACAO
        qd(33) = grdReg.TextMatrix(nLinha, 14)
        qd(34) = Virg2Ponto(sTr(nSomaTotal)) 'SOMA TOTAL DO BANCO
        qd(35) = nContaReg
        qd(36) = Null
        qd(37) = Null
        Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)
    End If
End With



End Sub

Private Sub cmdShow_Click()
If grdArq.Rows > 1 Then
   x = Shell("NOTEPAD" & " " & grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1), vbNormalFocus)
End If
End Sub

Private Sub Form_Activate()

If grdArq.Rows > 1 And grdReg.Rows = 1 Then
    grdArq.Row = grdArq.Rows - 1
    LeArquivo
End If

End Sub


Private Sub Form_Load()


'ImgLogo.Picture = frmPagAutomatico.cmdBanco(frmPagAutomatico.lblAux.Caption).PictureNormal
grdArq.ColWidth(0) = 0
grdArq.Row = grdArq.Rows - 1
grdArq.Col = 0
grdArq.ColSel = 1
grdReg.Rows = 1
Me.Left = frmMdi.ScaleWidth / 2 - Me.Width / 2
Me.Top = frmMdi.ScaleHeight / 2 - Me.Height / 2 + 1800
lblPb.Caption = ""
frmMdi.AddWindow Me.Name, Me.Caption

'carrega os arquivos do banco

End Sub

Private Sub LeArquivo()

On Error GoTo Erro

Dim sFullPath As String
Dim Header As FebrabanA
Dim Header2 As FebrabanA2
Dim Header3 As FebrabanA3
Dim HeaderSN As SimplesNacionalHeader
Dim Registro As FebrabanG
Dim Registro2 As FebrabanG2
Dim Registro3 As FebrabanG3
Dim Registro2U As FebrabanG2U
Dim RegistroSN As SimplesNacionalDetalhe
Dim Footer As FebrabanZ
Dim Footer2 As FebrabanZ2
Dim Footer3 As FebrabanZ3
Dim FooterSN As SimplesNacionalTrailer, nValorGuia As Double, nNumDoc As Long, sNomeBanco As String, nQtde As Integer, nValor As Double
Dim Posicao  As Long, RdoAux2 As rdoResultset, nSeq As Integer, bAchou As Boolean, nNumParc As Integer, nCompl As Integer, nAnoCompetencia As Integer
Dim nCount As Integer, RdoAux As rdoResultset, Sql As String, nCodReduz As Long, sVencimento As String, dDataVencto As Date, RdoAux3 As rdoResultset

Pb.Value = 0
sFullPath = grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1)
grdReg.Rows = 1

If frmMdi.frTeste.Visible = True Then
    If frmMdi.frTeste.Caption = "ACESSANDO OS DADOS LOCAIS" Then
        sFullPath = grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1)
    End If
End If

If InStr(1, sFullPath, "DAF607", vbBinaryCompare) > 0 Then GoTo SIMPLES

ReDim aFebrabanA(0): ReDim aFebrabanG(0): ReDim aFebrabanZ(0)
ReDim aFebrabanA2(0): ReDim aFebrabanG2(0): ReDim aFebrabanZ2(0)
ReDim aFebrabang2U(0): ReDim aFebrabanA3(0): ReDim aFebrabanG3(0): ReDim aFebrabanZ3(0)

Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header
    aFebrabanA(0).CodigoRegistro = Trim$(Header.CodigoRegistro)
    If aFebrabanA(0).CodigoRegistro <> "A" Then
        Close #1
        If Option1(0).Value = True Then
            GoTo BANESPA3 'old
        Else
            GoTo BANESPA2 'new
        End If
    End If
    aFebrabanA(0).CodigoRemessa = Trim$(Header.CodigoRemessa)
    aFebrabanA(0).CodigoConvenio = Trim$(Header.CodigoConvenio)
    aFebrabanA(0).NomeEmpresa = Trim$(Header.NomeEmpresa)
    aFebrabanA(0).CodigoBanco = Trim$(Header.CodigoBanco)
    aFebrabanA(0).NomeBanco = Trim$(Header.NomeBanco)
    aFebrabanA(0).DataGeracao = Trim$(Header.DataGeracao)
    aFebrabanA(0).NumeroSeq = Trim$(Header.NumeroSeq)
    aFebrabanA(0).VersaoLayout = Trim$(Header.VersaoLayout)
    aFebrabanA(0).Filler = Trim$(Header.Filler)
    If Left$(aFebrabanA(0).Filler, 10) = "DEBITO AUT" Then
        MsgBox "Arquivo de débito automático deve ser baixado pela tela de Débito Automático.", vbExclamation, "Atenção"
        Close #1
        GoTo Fim
    End If
    lblBanco.Caption = aFebrabanA(0).CodigoBanco & " - " & aFebrabanA(0).NomeBanco
    lblCR.Caption = aFebrabanA(0).CodigoRegistro
    lblCC.Caption = aFebrabanA(0).CodigoConvenio
    lblDG.Caption = ConvDataSerial(aFebrabanA(0).DataGeracao)
    lblNS.Caption = aFebrabanA(0).NumeroSeq
    lblVL.Caption = aFebrabanA(0).VersaoLayout
    Posicao = Len(Header) + 3
    nCount = 0
    Do While Not EOF(1)
         Get #1, Posicao, Registro
         If Registro.CodigoRegistro <> "Z" Then
              aFebrabanG(nCount).CodigoRegistro = Registro.CodigoRegistro
              aFebrabanG(nCount).ContaPrefeitura = Registro.ContaPrefeitura
              aFebrabanG(nCount).DataPagamento = Registro.DataPagamento
              aFebrabanG(nCount).DataCredito = Registro.DataCredito
              aFebrabanG(nCount).PreCodBarra = Registro.PreCodBarra
              aFebrabanG(nCount).ValorRecebido = Registro.ValorRecebido
              aFebrabanG(nCount).CodigoMunic = Registro.CodigoMunic
              aFebrabanG(nCount).DataVencto = Registro.DataVencto
              aFebrabanG(nCount).NumDocumento = Registro.NumDocumento
              aFebrabanG(nCount).NumParcela = Registro.NumParcela
              aFebrabanG(nCount).SituacaoRetorno = Registro.SituacaoRetorno
              aFebrabanG(nCount).FillerSmar = Registro.FillerSmar
              aFebrabanG(nCount).ValorRetornado = Registro.ValorRetornado
              aFebrabanG(nCount).ValorTarifa = Registro.ValorTarifa
              aFebrabanG(nCount).NumSeq = Registro.NumSeq
              aFebrabanG(nCount).CodAgencia = Registro.CodAgencia
              aFebrabanG(nCount).FormaPagamento = Registro.FormaPagamento
              aFebrabanG(nCount).NumAutentica = Registro.NumAutentica
              aFebrabanG(nCount).Filler = Registro.Filler
              With aFebrabanG(nCount)
                    grdReg.AddItem .CodigoRegistro & Chr(9) & .ContaPrefeitura & Chr(9) & ConvDataSerial(.DataPagamento) & Chr(9) & ConvDataSerial(.DataCredito) & Chr(9) & .PreCodBarra & Chr(9) & FormatNumber(.ValorRecebido / 100, 2) & Chr(9) & Val(.CodigoMunic) & _
                    Chr(9) & ConvDataSerial(.DataVencto) & Chr(9) & .NumDocumento & Chr(9) & .NumParcela & Chr(9) & .SituacaoRetorno & Chr(9) & .FillerSmar & Chr(9) & FormatNumber(.ValorRetornado / 100, 2) & Chr(9) & FormatNumber(.ValorTarifa / 100, 2) & Chr(9) & Val(.NumSeq) & _
                    Chr(9) & Val(.CodAgencia) & Chr(9) & .FormaPagamento & Chr(9) & Trim$(.NumAutentica) & Chr(9) & .Filler & Chr(9) & aFebrabanA(0).CodigoBanco
              End With
         Else
              Get #1, Posicao, Footer
              aFebrabanZ(0).CodigoRegistro = Footer.CodigoRegistro
              aFebrabanZ(0).TotalRegistro = Footer.TotalRegistro
              aFebrabanZ(0).ValorTotal = Footer.ValorTotal
              aFebrabanZ(0).Filler = Footer.Filler
              lblNumReg.Caption = Val(aFebrabanZ(0).TotalRegistro) - 2
              lblRegTot.Caption = "R$ " & FormatNumber(aFebrabanZ(0).ValorTotal / 100, 2)
              Exit Do
         End If
         If Left$(grdArq.TextMatrix(grdArq.Row, 1), 2) = "R2" Then
            Posicao = Posicao + Len(Registro) + 3
         Else
            Posicao = Posicao + Len(Registro) + 2
         End If
         nCount = nCount + 1
         ReDim Preserve aFebrabanG(nCount)
    Loop
 Close #1
Fim:
 Exit Sub
 
BANESPA2:

Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header2
    aFebrabanA2(0).TipoRegistro = Trim$(Header2.TipoRegistro)
    aFebrabanA2(0).DataGeracao = Trim$(Header2.DataGeracao)
    
    lblCR.Caption = aFebrabanA2(0).TipoRegistro
    lblCC.Caption = "0"
    lblDG.Caption = (Left$(aFebrabanA2(0).DataGeracao, 2) & "/" & Mid$(aFebrabanA2(0).DataGeracao, 3, 2) & "/" & Right$(aFebrabanA2(0).DataGeracao, 4))
    lblNS.Caption = "0"
    lblVL.Caption = "0"
    Posicao = Len(Header2) + 3
    nCount = 0: nQtde = 0: nValor = 0
    On Error Resume Next
    Do While Not EOF(1)
         Get #1, Posicao, Registro2
         If Registro2.TipoRegistro = 1 Or Registro2.TipoRegistro = 9 Then GoTo PROXIMO
         If Registro2.TipoRegistro = 5 Then GoTo Rodape
         
              aFebrabanG2(nCount).TipoRegistro = Registro2.TipoRegistro
              aFebrabanG2(nCount).ContaCobranca = Registro2.NumeroConta
              aFebrabanG2(nCount).DataVencto = Registro2.DataVencto
              aFebrabanG2(nCount).ValorTitulo = (CDbl(FormatNumber(Registro2.ValorTitulo, 2))) / 100
              aFebrabanG2(nCount).NossoNumero = Registro2.NossoNumero
              aFebrabanG2(nCount).ValorTitulo = (CDbl(FormatNumber(Registro2.ValorTitulo, 2))) / 100
              aFebrabanG2(nCount).ValorTarifa = Registro2.ValorTarifa
              aFebrabanG2(nCount).NumSequencial = Registro2.NumSequencial
              aFebrabanG2(nCount).AgenciaCobradora = Registro2.AgenciaCobradora
              
              Posicao = Posicao + Len(Registro2) + 2
              Get #1, Posicao, Registro2U
              aFebrabang2U(nCount).DataCredito = Registro2U.DataCredito
              aFebrabang2U(nCount).ValorCreditado = (CDbl(FormatNumber(Registro2U.ValorCreditado, 2))) / 100
              With aFebrabanG2(nCount)
                    grdReg.AddItem .TipoRegistro & Chr(9) & .ContaCobranca & Chr(9) & (Left$(.DataVencto, 2) & "/" & Mid$(.DataVencto, 3, 2) & "/" & Right$(.DataVencto, 4)) & Chr(9) & (Left$(aFebrabang2U(nCount).DataCredito, 2) & "/" & Mid$(aFebrabang2U(nCount).DataCredito, 3, 2) & "/" & Right$(aFebrabang2U(nCount).DataCredito, 4)) & Chr(9) & "-" & Chr(9) & .ValorTitulo & Chr(9) & "-" & _
                    Chr(9) & "-" & Chr(9) & .NossoNumero & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & aFebrabang2U(nCount).ValorCreditado & Chr(9) & FormatNumber(.ValorTarifa / 100, 2) & Chr(9) & Val(.NumSequencial) & _
                    Chr(9) & Val(.AgenciaCobradora) & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & "-"
              End With
              nQtde = nQtde + 1
              nValor = nValor + CDbl(aFebrabang2U(nCount).ValorCreditado)
         nCount = nCount + 1
         ReDim Preserve aFebrabanG2(nCount)
         ReDim Preserve aFebrabang2U(nCount)
PROXIMO:
         Posicao = Posicao + Len(Registro2) + 2
    Loop
Rodape:
        lblNumReg.Caption = nQtde
        lblRegTot.Caption = "R$ " & FormatNumber(nValor, 2)

 Close #1
 
 Exit Sub
 
 
BANESPA3:
Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header3
    aFebrabanA3(0).CodigoRegistro = Trim$(Header3.CodigoRegistro)
    aFebrabanA3(0).CodigoRetorno = Trim$(Header3.CodigoRetorno)
    aFebrabanA3(0).CodigoCobrança = Trim$(Header3.CodigoCobrança)
    aFebrabanA3(0).NomeCobrança = Trim$(Header3.NomeCobrança)
    aFebrabanA3(0).CodigoPrefeitura = Trim$(Header3.CodigoPrefeitura)
    aFebrabanA3(0).CodigoBanco = Trim$(Header3.CodigoBanco)
    aFebrabanA3(0).NomeBanco = Trim$(Header3.NomeBanco)
    aFebrabanA3(0).DataCredito = Trim$(Header3.DataCredito)
    aFebrabanA3(0).Filler1 = Trim$(Header3.Filler1)
    aFebrabanA3(0).Filler2 = Trim$(Header3.Filler2)
    aFebrabanA3(0).NumRegistros = Trim$(Header3.NumRegistros)
    
'    lblBanco.Caption = aFebrabanA2(0).CodigoBanco & " - " & aFebrabanA2(0).NomeBanco
    lblCR.Caption = aFebrabanA3(0).CodigoRegistro
    lblCC.Caption = "0"
    lblDG.Caption = ConvDataSerial(aFebrabanA3(0).DataCredito)
    lblNS.Caption = "0"
    lblVL.Caption = "0"
    Posicao = Len(Header3) + 3
    nCount = 0
    On Error Resume Next
    Do While Not EOF(1)
         Get #1, Posicao, Registro3
         If Registro3.CodigoRegistro <> "9" Then
              aFebrabanG3(nCount).CodigoRegistro = Registro3.CodigoRegistro
              aFebrabanG3(nCount).ContaPrefeitura = Registro3.ContaPrefeitura
              aFebrabanG3(nCount).DataPagamento = Registro3.DataPagamento
              aFebrabanG3(nCount).ValorPago = (CDbl(FormatNumber(Registro3.ValorPago, 2))) / 100
              'aFebrabanG2(nCount).ValorPago = (CDbl(FormatNumber(Registro2.ValorPago, 2)) + CDbl(FormatNumber(Registro2.ValorTaxa, 2))) / 100
              aFebrabanG3(nCount).NumDocumento = Registro3.NumDocumento
              'aFebrabanG2(nCount).ValorPago = (CDbl(FormatNumber(Registro2.ValorPago, 2)) + CDbl(FormatNumber(Registro2.ValorTaxa, 2))) / 100
              aFebrabanG3(nCount).ValorPago = (CDbl(FormatNumber(Registro3.ValorPago, 2))) / 100
              aFebrabanG3(nCount).ValorTaxa = Registro3.ValorTaxa
              aFebrabanG3(nCount).NumSeq = Registro3.NumSeq
              aFebrabanG3(nCount).CodAgencia = Registro3.CodAgencia
              With aFebrabanG3(nCount)
                    grdReg.AddItem .CodigoRegistro & Chr(9) & .ContaPrefeitura & Chr(9) & ConvDataSerial(.DataPagamento) & Chr(9) & ConvDataSerial(aFebrabanA3(0).DataCredito) & Chr(9) & "-" & Chr(9) & .ValorPago & Chr(9) & "-" & _
                    Chr(9) & "-" & Chr(9) & .NumDocumento & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & .ValorPago & Chr(9) & FormatNumber(.ValorTaxa / 100, 2) & Chr(9) & Val(.NumSeq) & _
                    Chr(9) & Val(.CodAgencia) & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & "-"
              End With
         Else
              Get #1, Posicao, Footer3
              aFebrabanZ3(0).CodigoSeiLa = Footer3.CodigoSeiLa
              aFebrabanZ3(0).Filler1 = Footer3.Filler1
              aFebrabanZ3(0).TotalRegistro = Footer3.TotalRegistro
              aFebrabanZ3(0).ValorTotal = Footer3.ValorTotal
              'lblNumReg.Caption = Val(aFebrabanZ2(0).TotalRegistro) - 2
              lblNumReg.Caption = Val(aFebrabanZ3(0).TotalRegistro)
              'lblRegTot.Caption = "R$ " & CDbl(aFebrabanZ2(0).ValorTotal / 100) + (CDbl(aFebrabanG2(0).ValorTaxa / 100) * Val(aFebrabanZ2(0).TotalRegistro))
              lblRegTot.Caption = "R$ " & CDbl(aFebrabanZ3(0).ValorTotal / 100)
              Exit Do
         End If
         Posicao = Posicao + Len(Registro3) + 2
         nCount = nCount + 1
         ReDim Preserve aFebrabanG3(nCount)
    
    Loop
 Close #1
 Exit Sub
 
SIMPLES:
ReDim aSimplesH(0): ReDim aSimplesD(0): ReDim aSimplesT(0)
Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, HeaderSN
    aSimplesH(0).CodigoRegistro = Trim$(HeaderSN.CodigoRegistro)
    aSimplesH(0).SeqRegistro = Trim$(HeaderSN.SeqRegistro)
    aSimplesH(0).CodigoConvenio = Trim$(HeaderSN.CodigoConvenio)
    aSimplesH(0).DataGeracao = Trim$(HeaderSN.DataGeracao)
    aSimplesH(0).NumeroRemessa = Trim$(HeaderSN.NumeroRemessa)
    aSimplesH(0).NumeroVersao = Trim$(HeaderSN.NumeroVersao)
    aSimplesH(0).Filler1 = Trim$(HeaderSN.Filler1)
    aSimplesH(0).Filler2 = Trim$(HeaderSN.Filler2)
    aSimplesH(0).CodigoBanco = Trim$(HeaderSN.CodigoBanco)
    aSimplesH(0).Filler3 = Trim$(HeaderSN.Filler3)
    
    '** TROCA OS BANCOS PELOS BANCOS VIRTUAIS **
    If Val(aSimplesH(0).CodigoBanco) = 1 Then
        aSimplesH(0).CodigoBanco = 91: sNomeBanco = "SN - BBRA"
    ElseIf Val(aSimplesH(0).CodigoBanco) = 33 Then
        aSimplesH(0).CodigoBanco = 92: sNomeBanco = "SN - BANE"
    ElseIf Val(aSimplesH(0).CodigoBanco) = 237 Then
        aSimplesH(0).CodigoBanco = 93: sNomeBanco = "SN - BRAD"
    ElseIf Val(aSimplesH(0).CodigoBanco) = 341 Then
        aSimplesH(0).CodigoBanco = 94: sNomeBanco = "SN - ITAU"
    ElseIf Val(aSimplesH(0).CodigoBanco) = 409 Then
        aSimplesH(0).CodigoBanco = 95: sNomeBanco = "SN - UNIB"
    ElseIf Val(aSimplesH(0).CodigoBanco) = 151 Then
        aSimplesH(0).CodigoBanco = 96: sNomeBanco = "SN - NCAI"
    ElseIf Val(aSimplesH(0).CodigoBanco) = 104 Then
        aSimplesH(0).CodigoBanco = 97: sNomeBanco = "SN - CFED"
    ElseIf Val(aSimplesH(0).CodigoBanco) = 399 Then
        aSimplesH(0).CodigoBanco = 98: sNomeBanco = "SN - CFED"
    Else
        aSimplesH(0).CodigoBanco = 91: sNomeBanco = "SN - BBRA"
    End If
    aSimplesH(0).CodigoBanco = Format(aSimplesH(0).CodigoBanco, "000")
    lblBanco.Caption = aSimplesH(0).CodigoBanco & "-" & sNomeBanco
    
    '*******************************************
    
    
    lblCR.Caption = aSimplesH(0).CodigoRegistro
    lblCC.Caption = aSimplesH(0).CodigoConvenio
    lblDG.Caption = ConvDataSerial(aSimplesH(0).DataGeracao)
    lblNS.Caption = aSimplesH(0).NumeroRemessa
    lblVL.Caption = aSimplesH(0).NumeroVersao

    Posicao = Len(HeaderSN) + 3
    nCount = 0
    Do While Not EOF(1)
         Get #1, Posicao, RegistroSN
         If RegistroSN.SeqRegistro <> "99999999" Then
              aSimplesD(nCount).CodigoRegistro = RegistroSN.CodigoRegistro
              aSimplesD(nCount).SeqRegistro = RegistroSN.SeqRegistro
              aSimplesD(nCount).DataArrecada = RegistroSN.DataArrecada
              aSimplesD(nCount).DataVencimento = RegistroSN.DataVencimento
              aSimplesD(nCount).Filler1 = RegistroSN.Filler1
              aSimplesD(nCount).Filler2 = RegistroSN.Filler2
              aSimplesD(nCount).Cnpj = RegistroSN.Cnpj
              aSimplesD(nCount).Filler3 = RegistroSN.Filler3
              aSimplesD(nCount).Esfera = RegistroSN.Esfera
              aSimplesD(nCount).Competencia = RegistroSN.Competencia
              aSimplesD(nCount).ValorPrincipal = RegistroSN.ValorPrincipal
              aSimplesD(nCount).ValorMulta = RegistroSN.ValorMulta
              aSimplesD(nCount).ValorJuros = RegistroSN.ValorJuros
              aSimplesD(nCount).Filler4 = RegistroSN.Filler4
              aSimplesD(nCount).ValorAutentica = RegistroSN.ValorAutentica
              aSimplesD(nCount).NumeroAutentica = RegistroSN.NumeroAutentica
              aSimplesD(nCount).CodigoBanco = RegistroSN.CodigoBanco
              aSimplesD(nCount).CodigoAgencia = RegistroSN.CodigoAgencia
              
            '** TROCA OS BANCOS PELOS BANCOS VIRTUAIS **
            If Val(aSimplesD(nCount).CodigoBanco) = 1 Then
                aSimplesD(nCount).CodigoBanco = 91
            ElseIf Val(aSimplesD(nCount).CodigoBanco) = 33 Then
                aSimplesD(nCount).CodigoBanco = 92
            ElseIf Val(aSimplesD(nCount).CodigoBanco) = 237 Then
                aSimplesD(nCount).CodigoBanco = 93
            ElseIf Val(aSimplesD(nCount).CodigoBanco) = 341 Then
                aSimplesD(nCount).CodigoBanco = 94
            ElseIf Val(aSimplesD(nCount).CodigoBanco) = 409 Then
                aSimplesD(nCount).CodigoBanco = 95
            ElseIf Val(aSimplesD(nCount).CodigoBanco) = 151 Then
                aSimplesD(nCount).CodigoBanco = 96
            ElseIf Val(aSimplesD(nCount).CodigoBanco) = 104 Then
                aSimplesD(nCount).CodigoBanco = 97
            ElseIf Val(aSimplesD(nCount).CodigoBanco) = 399 Then
                aSimplesD(nCount).CodigoBanco = 98
            End If
            '*******************************************
              
              nAnoCompetencia = Val(Left(aSimplesD(nCount).Competencia, 4))
              nValorGuia = CDbl(aSimplesD(nCount).ValorPrincipal) / 100 + CDbl(aSimplesD(nCount).ValorMulta) / 100 + CDbl(aSimplesD(nCount).ValorJuros) / 100
              'BUSCA CÓDIGO
              Sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE CONVERT(BIGINT, cnpj) = " & Val(aSimplesD(nCount).Cnpj) & "  ORDER BY DATAABERTURA DESC"
              Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
              With RdoAux
                  If .RowCount > 0 Then
                      nCodReduz = !CODIGOMOB
'                      If nCodReduz = 108785 Then MsgBox "teste"
                     'BUSCA VENCIMENTO
                      sVencimento = Right(aSimplesD(nCount).DataVencimento, 2) & "/" & Mid(aSimplesD(nCount).DataVencimento, 5, 2) & "/" & Left(aSimplesD(nCount).DataVencimento, 4)
                      dDataVencto = CDate(sVencimento)
                      nNumDoc = 0
                     'BUSCA LANCAMENTO
                      Sql = "SELECT debitoparcela.codreduzido, debitoparcela.codlancamento,DEBITOPARCELA.SEQLANCAMENTO,debitoparcela.numparcela,DEBITOPARCELA.CODCOMPLEMENTO,debitoparcela.datavencimento, debitoparcela.statuslanc, debitotributo.valortributo "
                      Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
                      Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And "
                      Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND (debitoparcela.codlancamento = 5) AND (MONTH(debitoparcela.datavencimento) = " & Month(dDataVencto) & ") AND "
                      Sql = Sql & "(YEAR(debitoparcela.datavencimento) = " & Year(dDataVencto) & ") AND (debitotributo.codtributo = 13)"
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      With RdoAux2
                         'EXISTE LANCAMENTO NESTE MÊS/ANO?
                          If .RowCount > 0 Then 'SIM
                             'SE FOR MESMO VENCIMENTO E MESMO VALOR PULA
                              If Format(dDataVencto, "dd/mm/yyyy") = Format(!DataVencimento, "dd/mm/yyyy") Then
                                 If (CDbl(aSimplesD(nCount).ValorPrincipal / 100)) + (CDbl(aSimplesD(nCount).ValorJuros) / 100) + (CDbl(aSimplesD(nCount).ValorMulta) / 100) = !valortributo Then
                                    GoTo PROXIMOSIMPLES
                                 End If
                              End If
                          
                          
                              nNumParc = !NumParcela 'CAPTURA A PARCELA
                              bAchou = False
                             'TEM ALGUM QUE NÃO ESTA PAGO?
                              Do Until .EOF
                                  If !statuslanc = 3 Then
                                      bAchou = True
                                      Exit Do
                                  End If
                                 .MoveNext
                              Loop
                             'SE ACHOU PEGA A PARCELA
                              If bAchou Then
                                 nSeq = !SeqLancamento
                                 nCompl = !CODCOMPLEMENTO '---------------> PARCELA PRONTA PARA USO
                                'BUSCA O DOCUMENTO DELA
                                 Sql = "SELECT * FROM PARCELADOCUMENTO WHERE Codreduzido = " & nCodReduz & " AND ANOEXERCICIO=" & nAnoCompetencia & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                                 Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                                 Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                 With RdoAux3
                                     If .RowCount > 0 Then
                                         nNumDoc = !NumDocumento
                                     Else
                                        'SE NÃO ENCONTRAR CRIA APENAS O DOCUMENTO
                                         GoTo Documento
                                     End If
                                    .Close
                                 End With
                                 GoTo GRID
                              Else
                                'SE NÃO ACHAR
                                .MoveFirst
                                 nCompl = 0
                                'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
                                 Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (MONTH(datavencimento) = " & Month(dDataVencto) & ") AND "
                                 Sql = Sql & "(YEAR(datavencimento) = " & Year(dDataVencto) & ")"
                                 Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                 With RdoAux3
                                   If IsNull(!MAXIMO) Then
                                       nSeq = 0
                                   Else
                                       nSeq = !MAXIMO + 1
                                   End If
                                  .Close
                                 End With
                              End If
                          Else
                             'NÃO ACHOU LANCAMENTOS NESTE MÊS/ANO
                             'O NÚMERO DA PARCELA A SER CRIADA SERÁ O ÚLTIMO NÚMERO DE PARCELA DO ANO
                              Sql = "SELECT MAX(NUMPARCELA) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & nAnoCompetencia & ")"
                              Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                              With RdoAux3
                                If IsNull(!MAXIMO) Then
                                    nNumParc = 1
                                Else
                                    nNumParc = !MAXIMO + 1
                                End If
                               .Close
                              End With
                              nCompl = 0
                              nSeq = 0
                          End If
                        
                         'CRIAR PARCELA DE ISS VARIAVEL NESTE MES E ANO COM O VENCIMENTO QUE VEIO DO BANCO
                          Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
                          Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
                          Sql = Sql & nNumParc & "," & nCompl & ",3,'" & Format(dDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0)"
                          cn.Execute Sql, rdExecDirect
                         'CRIAR O TRIBUTO PARA ELA (13 - iss variavel)
                          Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                          Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
                          Sql = Sql & nNumParc & "," & nCompl & "," & 13 & "," & 0 & ")"
                          cn.Execute Sql, rdExecDirect
                         'CRIAR COMPLEMENTO DO SIMPLES
                          Sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR) VALUES(" & nCodReduz & ","
                          Sql = Sql & nAnoCompetencia & ",5," & nSeq & "," & nNumParc & "," & nCompl & ",'" & grdArq.TextMatrix(grdArq.Row, 1) & "','" & Format(ConvDataSerial(aSimplesD(nCount).DataArrecada), "mm/dd/yyyy") & "'," & Virg2Ponto(CStr((CDbl(aSimplesD(nCount).ValorPrincipal / 100)) + (CDbl(aSimplesD(nCount).ValorJuros) / 100) + (CDbl(aSimplesD(nCount).ValorMulta) / 100))) & ")"
                          cn.Execute Sql, rdExecDirect

Documento:
                         'CRIAR O DOCUMENTO PARA ELA
                          Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
                          Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                          With RdoAux3
                               nNumDoc = !MAXIMO + 1
                              .Close
                          End With
                          Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA) VALUES(" & nNumDoc & ",'"
                          Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & ")"
                          cn.Execute Sql, rdExecDirect
                         'CRIAR A PARCELADOCUMENTO
                          Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
                          Sql = Sql & nNumParc & "," & nCompl & "," & nNumDoc & ")"
                          cn.Execute Sql, rdExecDirect
                         
                         .Close
                      End With
                  Else
                    MsgBox "O CNPJ " & aSimplesD(nCount).Cnpj & " NÃO FOI LOCALIZADO.", vbCritical, "Atenção"
                  End If
                 .Close
              End With

GRID:
             'AGORA QUE JA POSSUIMOS O DOCUMENTO A SER EFETUADA A BAIXA
             'PODEMOS POPULAR O GRID
              With aSimplesD(nCount)
                   grdReg.AddItem .CodigoRegistro & Chr(9) & "" & Chr(9) & ConvDataSerial(.DataArrecada) & Chr(9) & ConvDataSerial(.DataArrecada) & Chr(9) & "" & Chr(9) & FormatNumber(nValorGuia, 2) & Chr(9) & "" & Chr(9) & _
                   ConvDataSerial(.DataVencimento) & Chr(9) & nCodReduz & Chr(9) & nNumParc & Chr(9) & "" & Chr(9) & "" & Chr(9) & FormatNumber(nValorGuia, 2) & Chr(9) & "0,00" & Chr(9) & .SeqRegistro & Chr(9) & .CodigoAgencia & Chr(9) & _
                   "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & .CodigoBanco
                   
              End With
         Else
              Get #1, Posicao, FooterSN
              aSimplesT(0).CodigoRegistro = FooterSN.CodigoRegistro
              aSimplesT(0).TotalRegistro = FooterSN.TotalRegistro
              aSimplesT(0).ValorRegistro = FooterSN.ValorRegistro
              lblNumReg.Caption = Val(aSimplesT(0).TotalRegistro) - 2
              lblRegTot.Caption = "R$ " & FormatNumber(aSimplesT(0).ValorRegistro / 100, 2)
              Exit Do
         End If
         nCount = nCount + 1
         ReDim Preserve aSimplesD(nCount)
PROXIMOSIMPLES:
         Posicao = Posicao + Len(RegistroSN) + 2

    Loop
Close #1

 
 
Fim2:
 Exit Sub
 
 
Erro:
 MsgBox Err.Description
Resume Next
End Sub

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Private Sub Form_Paint()
frmMdi.RemoveWindow Me.Name
End Sub

Private Sub grdArq_Click()
If grdArq.Row > 0 Then
     LeArquivo
     grdReg.SetFocus
     
End If
End Sub

Private Function CalculaJuros2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String
Dim sDataVencto As String, nDia As Integer, nMes As Integer, nAno As Integer

'SE O VENCIMENTO FOR MAIOR OU IGUAL A DATA ATUAL, NÃO EXISTE JUROS
If dDataVencto >= dDataPagto Then
    CalculaJuros2 = 0
    Exit Function
End If

'SE ESTIVER NO MESMO MES E ANO QUE A DATA ATUAL, NAO EXISTE JUROS
If Month(dDataVencto) = Month(dDataPagto) And Year(dDataVencto) = Year(dDataPagto) Then
    CalculaJuros2 = 0
    Exit Function
End If

If Not dcJuros.Exists(Year(dDataPagto)) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJuros2 = 0
   Exit Function
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

sDataVencto = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
dDataVencto = Format(sDataVencto, "dd/mm/yyyy")
nNumMes = Int(DateDiff("d", dDataVencto, dDataPagto) / 30) + 1

nValorPerc = dcJuros.Item(Year(dDataPagto))

nValorPerc = nValorPerc / 100

CalculaJuros2 = nValorDebito * nValorPerc * nNumMes
If CalculaJuros2 > 0 Then
   CalculaJuros2 = FormatNumber(CalculaJuros2, 3)
End If

End Function

Private Function CalculaMulta2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumDia As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String


If dDataVencto >= dDataPagto Then
    CalculaMulta2 = 0
    Exit Function
End If

nNumDia = Abs(DateDiff("d", dDataPagto, dDataVencto))

If nNumDia = 0 Then
   CalculaMulta2 = 0
   Exit Function
End If

Sql = "SELECT MINDIA,MAXDIA,PERCDIA FROM MULTA WHERE ANOMULTA=" & Year(dDataVencto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         If nNumDia >= !MINDIA And nNumDia <= !MAXDIA Then
             nValorPerc = !PERCDIA
             Exit Do
         ElseIf nNumDia >= !MINDIA And !MAXDIA = 0 Then
             nValorPerc = !PERCDIA
             Exit Do
         End If
        .MoveNext
    Loop
End With

nValorPerc = nValorPerc / 100
CalculaMulta2 = nValorDebito * nValorPerc
If CalculaMulta2 > 0 Then
   CalculaMulta2 = FormatNumber(CalculaMulta2, 3)
End If

End Function

Private Function CalculaCorrecao2(nValorDebito As Double, dDataBase As Date, dDataVencto As Date) As Double

Dim UfirAtual As Double
Dim UfirBase As Double

If Year(dDataVencto) > Year(Now) Then
   CalculaCorrecao2 = 0
   Exit Function
End If
UfirAtual = RetornaUFIR(Year(dDataVencto))
UfirBase = RetornaUFIR(Year(dDataBase))

CalculaCorrecao2 = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecao2 > 0 Then
   CalculaCorrecao2 = FormatNumber(CalculaCorrecao2, 2)
End If
End Function

Private Sub MontaResumo()
Dim x As Long, z As Long, y As Integer
Dim nNumDoc As Long
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqLanc As Integer
Dim nNumParc As Integer
Dim nCodTributo As Integer
Dim nCompl As Integer
Dim nStatus As Integer
Dim nValorLanc As Double
Dim nValorJuros As Double
Dim nValorMulta As Double, nValorDif As Double
Dim nValorCorrecao As Double, nValorTotal As Double
Dim nValorPago As Double, nValorPagoSTaxa As Double
Dim nValorTaxa As Double
Dim nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double
Dim bDupl As Boolean
Dim dDataPag As Date, dDataCred As Date, dDataDoc As Date
Dim dDataVencto As Date
Dim nValorPrincipal As Double
Dim nSomaTotal As Double, nSomaTotal2 As Double
Dim nSomaPrincipal As Double
Dim bDupS As Boolean, bDupN As Boolean, nSomaClass As Double, nSomaClass2 As Double
Dim bTemIssVar As Boolean, nValorSIssVar As Double
Dim nCodBanco As Integer, sCodAgencia As String, bBanespa2 As Boolean
Dim nValorPagoReal As Double, nResto As Double, nContaResto As Integer, bDebClassificar As Boolean

lstDoc.Clear
grdParc.Rows = 1
nSomaClass = 0: nSomaClass2 = 0
bDebClassificar = False
For x = 1 To grdReg.Rows - 1
    CallPb x, grdReg.Rows - 1
    grdParc.Rows = 1
'    If x > grdReg.Rows - 1 Then Exit For
    If Len(CStr(Val(grdReg.TextMatrix(x, 8)))) > 7 Then
       
       nNumDoc = Val(Left$(grdReg.TextMatrix(x, 8), Len(grdReg.TextMatrix(x, 8)) - 1))
       'If nNumDoc = 11007402 Then MsgBox "teste"
        If Val(grdReg.TextMatrix(x, 8)) > 10000000 And Val(grdReg.TextMatrix(x, 8)) < 30000000 Then
    '    If Val(grdReg.TextMatrix(x, 8)) > 10000000 Then
            nNumDoc = Val(grdReg.TextMatrix(x, 8))
            Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
            If RdoAux.RowCount = 0 Then
                nNumDoc = Val(Left$(grdReg.TextMatrix(x, 8), Len(grdReg.TextMatrix(x, 8)) - 1))
                RdoAux.Close
            End If
            
        End If
    Else
       nNumDoc = Val(grdReg.TextMatrix(x, 8))
    End If
    
    
'    If Val(grdReg.TextMatrix(x, 8)) = 784608 Then MsgBox "A"
'    nNumDoc = Val(Left$(grdReg.TextMatrix(x, 8), Len(grdReg.TextMatrix(x, 8)) - 1))
    
'    If Len(CStr(Val(grdReg.TextMatrix(x, 8)))) < 8 Then'
'      nNumDoc = Val(grdReg.TextMatrix(x, 8))
'    End If
    
    
'    If nNumDoc = 5837250 Then
'        MsgBox "AQUI"
'    End If
    
    dDataPag = Format(grdReg.TextMatrix(x, 2), "dd/mm/yyyy")
    dDataCred = Format(grdReg.TextMatrix(1, 3), "dd/mm/yyyy")
    nSomaTotal2 = CDbl(lblRegTot.Caption)
    Sql = "DELETE FROM DEBITOCLASSIFICAR WHERE DATARECEITA='" & Format(dDataCred, "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND NUMDOCUMENTO=" & nNumDoc
    cn.Execute Sql, rdExecDirect
   
    nValorPago = CDbl(grdReg.TextMatrix(x, 12))
    nResto = nValorPago
    If grdReg.TextMatrix(x, 19) <> "-" Then
       nCodBanco = grdReg.TextMatrix(x, 19)
       bBanespa2 = False
    Else
       nCodBanco = Val(Left$(lblBanco.Caption, 3))
       bBanespa2 = True
    End If
    If nCodBanco = 0 Then nCodBanco = Val(Left$(lblBanco.Caption, 3))
    sCodAgencia = grdReg.TextMatrix(x, 15)
    nSomaPrincipal = 0
    'CARREGA OS LANÇAMENTOS DO DOCUMENTO
    Sql = "SELECT PARCELADOCUMENTO.CODREDUZIDO,PARCELADOCUMENTO.ANOEXERCICIO,PARCELADOCUMENTO.CODLANCAMENTO,LANCAMENTO.DESCREDUZ,"
    Sql = Sql & "PARCELADOCUMENTO.SEQLANCAMENTO,PARCELADOCUMENTO.NUMPARCELA,PARCELADOCUMENTO.CODCOMPLEMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
    Sql = Sql & "NUMDOCUMENTO.DATADOCUMENTO,NUMDOCUMENTO.CODBANCO,NUMDOCUMENTO.VALORTAXADOC,PARCELADOCUMENTO.VALORJUROS,PARCELADOCUMENTO.VALORMULTA,"
    Sql = Sql & "PARCELADOCUMENTO.VALORCORRECAO,NUMDOCUMENTO.CODAGENCIA,NUMDOCUMENTO.VALORPAGO,DEBITOPARCELA.STATUSLANC,SITUACAOLANCAMENTO.DESCSITUACAO,"
    Sql = Sql & "DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE FROM LANCAMENTO INNER JOIN DEBITOPARCELA ON LANCAMENTO.CodLancamento = DEBITOPARCELA.CodLancamento Inner Join "
    Sql = Sql & "SITUACAOLANCAMENTO ON DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO RIGHT OUTER JOIN PARCELADOCUMENTO ON "
    Sql = Sql & "DEBITOPARCELA.CODREDUZIDO = PARCELADOCUMENTO.CODREDUZIDO AND DEBITOPARCELA.AnoExercicio = PARCELADOCUMENTO.AnoExercicio AND "
    Sql = Sql & "DEBITOPARCELA.CodLancamento = PARCELADOCUMENTO.CodLancamento AND DEBITOPARCELA.SeqLancamento = PARCELADOCUMENTO.SeqLancamento AND "
    Sql = Sql & "DEBITOPARCELA.NumParcela = PARCELADOCUMENTO.NumParcela AND DEBITOPARCELA.CODCOMPLEMENTO = PARCELADOCUMENTO.CODCOMPLEMENTO RIGHT OUTER JOIN "
    Sql = Sql & "NUMDOCUMENTO ON PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO Where PARCELADOCUMENTO.NumDocumento = " & nNumDoc
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            lstDoc.AddItem nNumDoc
            'DOCUMENTO NÃO ENCONTRADO (VERIFICAR ...........)
'             If nNumDoc = 657380 Then MsgBox "TESTE"
             nSomaClass = nSomaClass + nValorPago
             nResto = nResto - nValorPago
             Sql = "SELECT * FROM RECEITACLASSIFICAR WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATARECEITA='"
             Sql = Sql & Format(dDataCred, "mm/dd/yyyy") & "' AND NUMDOCUMENTO=" & nNumDoc
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                If .RowCount = 0 Then
                    Sql = "INSERT RECEITACLASSIFICAR (NOMEARQ,DATARECEITA,CODBANCO,NUMDOCUMENTO,VALORTOTAL) VALUES('"
                    Sql = Sql & grdArq.TextMatrix(grdArq.Row, 1) & "','" & Format(dDataCred, "mm/dd/yyyy") & "'," & nCodBanco & ","
                    Sql = Sql & nNumDoc & "," & Virg2Ponto(CStr(nValorPago)) & ")"
                    cn.Execute Sql, rdExecDirect
                End If
               .Close
             End With
             GoTo PROXIMO
        Else
                        
            'CARREGA VALOR DA TAXA DOCUMENTO
            'SE NÃO TIVER TAXADOC SINAL QUE VEIO DA SMARK ENTÃO PEGAMOS A TAXADOC DO 1º LANCAMENTO
            If IsNull(!VALORTAXADOC) Or !VALORTAXADOC = 0 Then
               Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
               Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                      nValorTaxa = FormatNumber(!valortributo, 2)
                   Else
                      nValorTaxa = 0
                   End If
               End With
            Else
               nValorTaxa = FormatNumber(!VALORTAXADOC, 2)
            End If
        End If
        nValorPagoSTaxa = nValorPago - nValorTaxa
        If Not IsNull(!DATADOCUMENTO) Then
            dDataDoc = !DATADOCUMENTO
        End If
        nContaResto = 1
        Do Until .EOF '(RDOAUX)
             If IsNull(!statuslanc) Then
                Sql = "INSERT DEBITOCLASSIFICAR (DATARECEITA,CODBANCO,NOMEARQ,NUMDOCUMENTO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO) VALUES('"
                Sql = Sql & Format(dDataCred, "mm/dd/yyyy") & "'," & nCodBanco & ",'" & grdArq.TextMatrix(grdArq.Row, 1) & "'," & nNumDoc & "," & nCodReduz & "," & nAnoExercicio & "," & nCodLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & ")"
                cn.Execute Sql, rdExecDirect
                bDebClassificar = True
                GoTo PROXIMOAUX
             End If
             nStatus = !statuslanc
             dDataVencto = !DataVencimento
             If nStatus = 1 Or nStatus = 2 Or nStatus = 7 Or nStatus = 9 Then
                bDupS = True
                bDupl = True
             Else
                bDupN = True
                bDupl = False
             End If
            'ADICIONA NO GRID PARCELA
'             grdParc.AddItem !AnoExercicio & Chr(9) & Format(!CODREDUZIDO, "000000") & Chr(9) & Format(!CodLancamento, "000") & " - " & !DESCREDUZ & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & Format(IIf(!NumParcela = 13, 0, !NumParcela), "00") & Chr(9) & _
                 !CODCOMPLEMENTO & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & _
                 "-" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & IIf(bDupl, "Sim", "Não") & Chr(9) & Format(dDataVencto, "dd/mm/yyyy") & Chr(9) & nValorPago & Chr(9) & nValorTaxa & Chr(9) & "-" & Chr(9) & x
             grdParc.AddItem !AnoExercicio & Chr(9) & Format(!CODREDUZIDO, "000000") & Chr(9) & Format(!CodLancamento, "000") & " - " & !descreduz & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & Format(!NumParcela, "00") & Chr(9) & _
                 !CODCOMPLEMENTO & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & _
                 "-" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & IIf(bDupl, "Sim", "Não") & Chr(9) & Format(dDataVencto, "dd/mm/yyyy") & Chr(9) & nValorPago & Chr(9) & nValorTaxa & Chr(9) & "-" & Chr(9) & x
            
            'PARA CADA LANCAMENTO CARREGAMOS OS TRIBUTOS
             Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO "
             Sql = Sql & "WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
             Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3 "
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                nSomaPrincipal = 0
                nSomaJuros = 0
                nSomaMulta = 0
                nSomaCorrecao = 0
                nSomaTotal = 0
                grdTrib.Rows = 1
                bTemIssVar = False
                Do Until .EOF
                   If !CodTributo = 13 Or !CodTributo = 85 Then
                     'CALCULA ISS VARIAVEL
                      Sql = "SELECT SUM(VALORTRIBUTO) AS TOTAL FROM DEBITOTRIBUTO "
                      Sql = Sql & "WHERE CODREDUZIDO = " & RdoAux!CODREDUZIDO & " AND ANOEXERCICIO = " & RdoAux!AnoExercicio & " AND CODLANCAMENTO = " & RdoAux!CodLancamento & " AND "
                      'Sql = Sql & "SEQLANCAMENTO = " & RdoAux!SEQLANCAMENTO & " AND NUMPARCELA = " & RdoAux!NUMPARCELA & " AND CODCOMPLEMENTO = " & RdoAux!CODCOMPLEMENTO & " AND CODTRIBUTO<>13 "
                      Sql = Sql & "SEQLANCAMENTO = " & RdoAux!SeqLancamento & " AND NUMPARCELA = " & RdoAux!NumParcela & " AND CODCOMPLEMENTO = " & RdoAux!CODCOMPLEMENTO & " AND CODTRIBUTO<>3 "
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      With RdoAux2
                          If Not IsNull(!Total) Then
                             nValorSIssVar = !Total
                          Else
                            nValorSIssVar = 0
                         End If
                        .Close
                      End With
                      If nValorSIssVar = 0 Then
                         nValorLanc = FormatNumber(nValorPagoSTaxa, 2)
                      Else
                     'nValorLanc = FormatNumber(nValorPago - nValorSIssVar, 2)
                         nValorLanc = FormatNumber(nValorSIssVar, 2)
                      End If
                   Else
                      nValorLanc = FormatNumber(!valortributo, 2)
                   End If
                   
                  'ADICIONA NO GRID TRIBUTO
                  If (dDataPag > dDataVencto) Then 'PAGO APOS O VENCIMENTO
                      nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
                      If nStatus = 20 Then 'JULGAMENTO
                         nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataDoc), 2)
                      Else
                         nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                      End If
                      nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                  Else
                      nValorCorrecao = 0
                      nValorJuros = 0
                      nValorMulta = 0
                  End If
                  nValorTotal = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
                  nSomaJuros = nSomaJuros + nValorJuros
                  nSomaMulta = nSomaMulta + nValorMulta
                  nSomaCorrecao = nSomaCorrecao + nValorCorrecao
                  nSomaTotal = nSomaTotal + nValorTotal
                  grdTrib.AddItem !CodTributo & Chr(9) & nValorLanc & Chr(9) & nValorMulta & Chr(9) & nValorJuros & Chr(9) & nValorCorrecao & Chr(9) & nValorTotal & Chr(9) & x
                  nSomaPrincipal = nSomaPrincipal + nValorLanc
                 .MoveNext
               Loop
            End With
            
           'ATUALIZA GRID PARCELA
            nValorDif = Round(nValorPago - (nValorTaxa + nSomaTotal), 2)
            grdParc.TextMatrix(grdParc.Rows - 1, 6) = nSomaPrincipal
            grdParc.TextMatrix(grdParc.Rows - 1, 7) = nSomaMulta
            grdParc.TextMatrix(grdParc.Rows - 1, 8) = nSomaJuros
            grdParc.TextMatrix(grdParc.Rows - 1, 9) = nSomaCorrecao
            grdParc.TextMatrix(grdParc.Rows - 1, 10) = nSomaTotal
            grdParc.TextMatrix(grdParc.Rows - 1, 15) = nValorDif
            
            'CORRIGE VALOR PAGO QUANDO + DE 1 LANCAMENTO
            If grdParc.Rows > 2 Then
                For z = 1 To grdParc.Rows - 1
                    If grdParc.TextMatrix(z, 6) <> "N/A" Then
                        nValorPrincipal = CDbl(grdParc.TextMatrix(z, 6))
                        'nValorPrincipal = CDbl(grdParc.TextMatrix(z, 6)) + CDbl(grdParc.TextMatrix(z, 14))
                        If nValorPago >= nValorPrincipal Then
                           grdParc.TextMatrix(z, 13) = FormatNumber(nValorPrincipal, 2)
                           grdParc.TextMatrix(z, 15) = FormatNumber(CDbl(grdParc.TextMatrix(z, 13)) - (CDbl(grdParc.TextMatrix(z, 6)) + CDbl(grdParc.TextMatrix(z, 14))), 2)
                        End If
                    End If
                Next
            End If
            
           'CARREGA DADOS PARA BAIXA DE PARCELA
            With grdParc
                nValorLanc = .TextMatrix(.Rows - 1, 13)
                nCodReduz = .TextMatrix(.Rows - 1, 1)
                nAnoExercicio = .TextMatrix(.Rows - 1, 0)
                nCodLanc = Val(Left$(.TextMatrix(.Rows - 1, 2), 3))
                nSeqLanc = .TextMatrix(.Rows - 1, 3)
                nNumParc = .TextMatrix(.Rows - 1, 4)
                nCompl = .TextMatrix(.Rows - 1, 5)
                If nNumParc = 0 Then
                   If CDbl(.TextMatrix(.Rows - 1, 15)) <= 0 Then
                       nStatus = 1 'UNICA SEM DIF
                   Else
                       nStatus = 9 'UNICA COM DIF
                   End If
                Else
                   If CDbl(.TextMatrix(.Rows - 1, 15)) <= 0 Then
                       nStatus = 2 'PAGO SEM DIF
                   Else
                       nStatus = 7 'PAGO COM DIF
                   End If
                End If

                If UCase$(.TextMatrix(.Rows - 1, 11)) <> "SIM" Then 'não é duplicado
                     'SE A PARCELA JÁ ESTIVER CANCELADA QUANDO FOR FEITA A BAIXA
                     'COLOCAMOS STATUS 15 (PAGO APÓS CANCELADO
                      If nStatus = 5 Or nStatus = 10 Or nStatus = 12 Or nStatus = 8 Then
                            'EFETUA BAIXA NA TABELA DEBITOPARCELA
                             Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=15  WHERE CODREDUZIDO=" & nCodReduz & " AND "
                             Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                             Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                             cn.Execute Sql, rdExecDirect
                      Else
                            'EFETUA BAIXA NA TABELA DEBITOPARCELA
                             Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
                             Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                             Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                             cn.Execute Sql, rdExecDirect
                            'SE FOR PARCELA UNICA EFETUA BAIXA EM TODAS AS PARCELAS AUTOMATICAMENTO
                            'SERA?
                             If nNumParc = 0 Then
                                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=1  WHERE CODREDUZIDO=" & nCodReduz & " AND "
                                Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                                Sql = Sql & " AND NUMPARCELA<>0"
                                cn.Execute Sql, rdExecDirect
                             Else
                                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5  WHERE CODREDUZIDO=" & nCodReduz & " AND "
                                Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                                Sql = Sql & " AND NUMPARCELA=0"
                                cn.Execute Sql, rdExecDirect
                             End If
                      End If
                     'EFETUA BAIXA NA TABELA DEBITOTRIBUTO
                      With grdTrib
                          For y = 1 To .Rows - 1
                              nCodTributo = .TextMatrix(y, 0)
                              
                              nValorCorrecao = .TextMatrix(y, 4)
                              nValorJuros = .TextMatrix(y, 3)
                              nValorMulta = .TextMatrix(y, 2)
                              If nCodTributo = 13 Then
                                  If nValorSIssVar = 0 Then
                                     Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=" & Virg2Ponto(CStr(nValorLanc)) & ",VALORCORRECAO=" & Virg2Ponto(sTr(nValorCorrecao)) & " ,VALORMULTA=" & Virg2Ponto(sTr(nValorMulta)) & " ,VALORJUROS=" & Virg2Ponto(sTr(nValorJuros))
                                     Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                                     Sql = Sql & "CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO=" & nCodTributo
                                     cn.Execute Sql, rdExecDirect
                                  End If
                              Else
                                  Sql = "UPDATE DEBITOTRIBUTO SET VALORCORRECAO=" & Virg2Ponto(sTr(nValorCorrecao)) & " ,VALORMULTA=" & Virg2Ponto(sTr(nValorMulta)) & " ,VALORJUROS=" & Virg2Ponto(sTr(nValorJuros))
                                  Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                                  Sql = Sql & "CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO=" & nCodTributo
                                  cn.Execute Sql, rdExecDirect
                              End If
                          Next
                      End With
                 End If
                
                
                
                
                'EFETUA BAIXA NA TABELA DEBITOPAGO
                 Sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND "
                 Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                 Sql = Sql & " AND NUMPARCELA=" & nNumParc
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                      If IsNull(!MAXIMO) Then
                         nSeqAdd = 0
                      Else
                         If .RowCount = 0 Then
                            nSeqAdd = 0
                        Else
                           nSeqAdd = !MAXIMO + 1
                        End If
                     End If
                    .Close
                End With
                If nContaResto = RdoAux.RowCount Then
                    nValorPagoReal = nResto
                    nResto = 0
                Else
                    If nResto >= nSomaTotal Then
                        nValorPagoReal = nSomaTotal
                        nResto = nResto - nSomaTotal
                    Else
                        nValorPagoReal = nResto
                        nResto = 0
                    End If
                End If
                nContaResto = nContaResto + 1
                nSomaClass2 = nSomaClass2 + nValorPagoReal
                Sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "SEQPAG,DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL,ARQUIVOBANCO) VALUES(" & nCodReduz & ","
                Sql = Sql & nAnoExercicio & "," & nCodLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nSeqAdd & ",'"
                Sql = Sql & Format(dDataPag, "mm/dd/yyyy") & "','" & Format(dDataCred, "mm/dd/yyyy") & "'," & Virg2Ponto(sTr(nSomaTotal)) & ","
                Sql = Sql & nCodBanco & ",'" & sCodAgencia & "'," & nNumDoc & "," & Virg2Ponto(sTr(nValorPagoReal)) & ",'" & grdArq.TextMatrix(grdArq.Row, 1) & "')"
                cn.Execute Sql, rdExecDirect
            End With
           'PROXIMO LANCAMENTO (RDOAUX)
PROXIMOAUX:
           .MoveNext
        Loop
        
       'EFETUA BAIXA NO DOCUMENTO
        Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=" & nCodBanco & " ,CODAGENCIA ='" & sCodAgencia & "' , VALORPAGO=" & Virg2Ponto(sTr(nValorPago))
        Sql = Sql & " WHERE NUMDOCUMENTO=" & nNumDoc
        cn.Execute Sql, rdExecDirect
        
        GravaBaixaTmp nNumDoc
       'GoTo FIM
    End With
    
PROXIMO:
Next

'SE TIVER DEBITO A CLASSIFICAR GRAVA NELE"
nResto = nSomaTotal2 - (nSomaClass + nSomaClass2)
If nResto > 0 Then
    Sql = "UPDATE DEBITOCLASSIFICAR SET VALORCLASS=" & Virg2Ponto(CStr(nResto)) & " WHERE DATARECEITA='" & Format(dDataCred, "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "'"
    cn.Execute Sql, rdExecDirect
End If

'EFETUA BAIXA NO ARQUIVO
Sql = "UPDATE ARQUIVOBANCO SET DATABAIXA='" & Format(Now, "mm/dd/yyyy") & " ' WHERE "
If bBanespa2 Then
   Sql = Sql & "NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(dDataCred, "mm/dd/yyyy") & "'"
Else
   Sql = Sql & "NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(dDataCred, "mm/dd/yyyy") & "'"
End If
cn.Execute Sql, rdExecDirect

Fim:
Liberado

End Sub

Private Sub LoadMatrix()
ReDim aCodBarra(0)
With aCodBarra(0)
    .PreCodBarra = Mid$(sBloco, 1, 3)
    .ValorRecebido = Mid$(sBloco, 4, 11)
    .CodigoMunic = Mid$(sBloco, 15, 4)
    .DataVencto = Mid$(sBloco, 19, 8)
    .NumDocumento = Mid$(sBloco, 27, 9)
    .NumParcela = Mid$(sBloco, 36, 2)
    .SituacaoRetorno = Mid$(sBloco, 38, 2)
    .FillerSmar = Mid$(sBloco, 40, 4)
End With

End Sub

Private Function RetornaDV2of5(sBloco As String) As Integer
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
RetornaDV2of5 = 10 - nResto

End Function

Private Sub CallPb(nVal As Long, nTot As Long)

If ((nVal * 100) / nTot) <= 100 Then
   Pb.Value = (nVal * 100) / nTot
Else
   Pb.Value = 100
End If

Me.Refresh
DoEvents

End Sub

Private Sub Reativa()
Dim nNumDoc As Long
Dim x As Integer
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqLanc As Integer
Dim nNumParc As Integer
Dim nCompl As Integer
Dim RdoAux As rdoResultset, Sql As String

With grdReg
    For x = 1 To .Rows - 1
        CallPb CLng(x), .Rows - 1
'        nNumDoc = Val(Left$(grdReg.TextMatrix(x, 8), Len(grdReg.TextMatrix(x, 8)) - 1))
'        If Len(CStr(nNumDoc)) < 7 Then
'          nNumDoc = Val(grdReg.TextMatrix(x, 8))
'        End If
        
'    nNumDoc = Val(Left$(grdReg.TextMatrix(x, 8), Len(grdReg.TextMatrix(x, 8)) - 1))
'    If Len(CStr(Val(grdReg.TextMatrix(x, 8)))) < 8 Then
'      nNumDoc = Val(grdReg.TextMatrix(x, 8))
'    End If
        
    If Len(CStr(Val(grdReg.TextMatrix(x, 8)))) > 7 Then
       
       nNumDoc = Val(Left$(grdReg.TextMatrix(x, 8), Len(grdReg.TextMatrix(x, 8)) - 1))
       'If nNumDoc = 11007402 Then MsgBox "teste"
        If Val(grdReg.TextMatrix(x, 8)) > 10000000 And Val(grdReg.TextMatrix(x, 8)) < 30000000 Then
    '    If Val(grdReg.TextMatrix(x, 8)) > 10000000 Then
            nNumDoc = Val(grdReg.TextMatrix(x, 8))
        End If
    Else
       nNumDoc = Val(grdReg.TextMatrix(x, 8))
    End If
        
        
        'If nNumDoc = 339451 Then MsgBox "TESTE"
        Sql = "SELECT PARCELADOCUMENTO.CODREDUZIDO,PARCELADOCUMENTO.ANOEXERCICIO,PARCELADOCUMENTO.CODLANCAMENTO,LANCAMENTO.DESCREDUZ,PARCELADOCUMENTO.SEQLANCAMENTO,"
        Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA,PARCELADOCUMENTO.CODCOMPLEMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,NUMDOCUMENTO.DATADOCUMENTO,NUMDOCUMENTO.CODBANCO,NUMDOCUMENTO.VALORTAXADOC,"
        Sql = Sql & "PARCELADOCUMENTO.VALORJUROS,PARCELADOCUMENTO.VALORMULTA,PARCELADOCUMENTO.VALORCORRECAO,"
        Sql = Sql & "NUMDOCUMENTO.CODAGENCIA,NUMDOCUMENTO.VALORPAGO,DEBITOPARCELA.STATUSLANC,SITUACAOLANCAMENTO.DESCSITUACAO,DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE "
        Sql = Sql & "FROM PARCELADOCUMENTO INNER JOIN DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
        Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND PARCELADOCUMENTO.NumParcela = DEBITOPARCELA.NumParcela AND "
        Sql = Sql & "PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO Inner Join LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO Inner Join SITUACAOLANCAMENTO ON "
        Sql = Sql & "DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO Inner Join NUMDOCUMENTO ON PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO Where PARCELADOCUMENTO.NumDocumento = " & nNumDoc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                GoTo PROXIMO
            End If
            Do Until .EOF
                nCodReduz = !CODREDUZIDO
'                If nCodReduz = 104295 Then MsgBox "TESTE"
                nAnoExercicio = !AnoExercicio
                nCodLanc = !CodLancamento
                nSeqLanc = !SeqLancamento
                nNumParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                'ATUALIZA A TABELA DEBITOPAGO
'                 Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' "'
'                 Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
'                 Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS NULL"
                 Sql = "DELETE FROM DEBITOPAGO "
                 Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                 Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3))
                 Sql = Sql & " AND DATAPAGAMENTO='" & Format(grdReg.TextMatrix(x, 2), "mm/dd/yyyy") & "'"
                 cn.Execute Sql, rdExecDirect
                'ATUALIZA A TABELA NUMDOCUMENTO
                 Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=0,CODAGENCIA=0,VALORPAGO=0 "
                 Sql = Sql & "WHERE NUMDOCUMENTO = " & nNumDoc
                 cn.Execute Sql, rdExecDirect
                 'SE TODOS OS REGISTROS EM DEBITOPAGO FOREM RESTITUIDOS ENTÃO ATUALIZA DÉBITOPARCELA
                 Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPAGO "
                 Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                 Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS  NULL"
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If !CONTADOR = 0 Then
                        'SE FOR ZERO SINAL QUE A PARCELA FOI TOTALMENTE RESTITUIDA
                        'ENTÃO PODEMOS ATUALIZAR O SEU STATUS PARA NÃO PAGO
                         Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=3 "
                         Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                         Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl
                         cn.Execute Sql, rdExecDirect
                     End If
                 End With
              .MoveNext
            Loop
           .Close
        End With
        
PROXIMO:
    Next
End With

Sql = "SELECT * FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    Sql = "DELETE FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL"
    cn.Execute Sql, rdExecDirect
End If

Sql = "DELETE FROM DEBITOCLASSIFICAR WHERE DATARECEITA='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "'"
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM RECEITACLASSIFICAR WHERE DATARECEITA='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "'"
cn.Execute Sql, rdExecDirect

If Left$(grdArq.TextMatrix(grdArq.Row, 1), 2) = "BD" Then
   Sql = "UPDATE ARQUIVOBANCO SET DATABAIXA=NULL WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
Else
   Sql = "UPDATE ARQUIVOBANCO SET DATABAIXA=NULL WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdReg.TextMatrix(1, 3), "mm/dd/yyyy") & "'"
End If
cn.Execute Sql, rdExecDirect

MsgBox "Todos os lançamentos descriminados e seus documentos foram reativados.", vbInformation, "INFORMAÇÃO"
grdParc.Rows = 1
'lblBanco.Caption = "0"

End Sub

Private Sub grdReg_Click()

    ' See if the user clicked row 0.
    If grdReg.MouseRow > 0 Then Exit Sub

    ' See if this is the same column.
    If grdReg.MouseCol = m_SortColumn Then
        ' This is the current sort column.
        ' Change the sort order and the column title.
        m_SortAscending = Not m_SortAscending
        If m_SortAscending Then
            grdReg.TextMatrix(0, m_SortColumn) = _
                "> " & Mid$(grdReg.TextMatrix(0, _
                    m_SortColumn), 3)
        Else
            grdReg.TextMatrix(0, m_SortColumn) = _
                "< " & Mid$(grdReg.TextMatrix(0, _
                    m_SortColumn), 3)
        End If
    Else
        ' This is a new sort column.
        ' Restore the previous sorting column's name.
        If m_SortColumn >= 0 Then
            grdReg.TextMatrix(0, m_SortColumn) = _
                Mid$(grdReg.TextMatrix(0, _
                    m_SortColumn), 3)
        End If

        ' Save the new sort column.
        m_SortColumn = grdReg.MouseCol

        ' Sort using the new column.
        m_SortAscending = True
        grdReg.TextMatrix(0, m_SortColumn) = _
            "> " & grdReg.TextMatrix(0, m_SortColumn)
    End If

    grdReg.Row = 1
    grdReg.RowSel = grdReg.Rows - 1
    grdReg.Col = m_SortColumn

    If m_SortAscending Then
        Select Case m_SortColumn
            Case 2, 5, 6
                grdReg.Sort = flexSortNumericAscending
            Case Else
                grdReg.Sort = flexSortStringAscending
        End Select
    Else
        Select Case m_SortColumn
            Case 2, 5, 6
                grdReg.Sort = flexSortNumericDescending
            Case Else
                grdReg.Sort = flexSortStringDescending
        End Select
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
LeArquivo
End Sub
