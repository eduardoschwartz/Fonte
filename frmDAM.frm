VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmDAM 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de ficha de compensação (D.A.M.)"
   ClientHeight    =   4950
   ClientLeft      =   5355
   ClientTop       =   3435
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRegistrado 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Registrado"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   6060
      TabIndex        =   31
      Top             =   4110
      Width           =   1860
   End
   Begin prjChameleon.chameleonButton btDesconto 
      Height          =   270
      Left            =   2745
      TabIndex        =   29
      ToolTipText     =   "Aplicar Desconto"
      Top             =   4545
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   476
      BTYPE           =   3
      TX              =   "Ok"
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
      MICON           =   "frmDAM.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtDesconto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1935
      TabIndex        =   28
      Text            =   "0"
      Top             =   4545
      Width           =   525
   End
   Begin MSComCtl2.UpDown UpDown 
      Height          =   285
      Left            =   2430
      TabIndex        =   27
      Top             =   4545
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtDesconto"
      BuddyDispid     =   196610
      OrigLeft        =   2880
      OrigTop         =   4500
      OrigRight       =   3135
      OrigBottom      =   4605
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox chkCobranca 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Cobrança"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   6060
      TabIndex        =   24
      Top             =   3840
      Width           =   1860
   End
   Begin VB.CheckBox chkCorrecao 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Isenção de Correção"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3900
      TabIndex        =   23
      Top             =   4110
      Width           =   1905
   End
   Begin prjChameleon.chameleonButton cmdAnistia 
      Height          =   240
      Left            =   3195
      TabIndex        =   22
      ToolTipText     =   "Vencimentos da Ansitia"
      Top             =   3240
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   14869218
      BCOLO           =   14869218
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmDAM.frx":001C
      PICN            =   "frmDAM.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskVencimento 
      Height          =   285
      Left            =   2070
      TabIndex        =   21
      Top             =   3195
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmDAM.frx":0192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin VB.CheckBox chkAnistia 
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3900
      TabIndex        =   19
      Top             =   3510
      Value           =   1  'Checked
      Width           =   285
   End
   Begin VB.CheckBox chkJulgamento 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Em Julgamento"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3900
      TabIndex        =   16
      Top             =   3840
      Width           =   1455
   End
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Left            =   2070
      TabIndex        =   15
      Top             =   4140
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmDAM.frx":01AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin VB.CheckBox chkVenctoAtual 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Calcular com data de:"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   4170
      Width           =   1905
   End
   Begin VB.CheckBox chkMulta 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Isenção Total de Juros e Multa"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   3555
   End
   Begin VB.CheckBox chkTx 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Remover Taxa de Expediente da DAM !!!"
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   3540
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   7740
      TabIndex        =   2
      ToolTipText     =   "Sair da Tela"
      Top             =   4530
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      MICON           =   "frmDAM.frx":01CA
      PICN            =   "frmDAM.frx":01E6
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
      Height          =   345
      Left            =   6120
      TabIndex        =   3
      ToolTipText     =   "Impressão do boleto bancário"
      Top             =   4530
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Imprimir DAM"
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmDAM.frx":0254
      PICN            =   "frmDAM.frx":0270
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   1230
      Left            =   270
      TabIndex        =   1
      Top             =   7230
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   2170
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDAM.frx":03CA
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDAM.frx":0466
   End
   Begin VB.Label lblDI 
      Caption         =   "N"
      Height          =   255
      Left            =   7800
      TabIndex        =   30
      Top             =   7260
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label 
      Caption         =   "Desconto Multa/Juros:"
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   135
      TabIndex        =   26
      Top             =   4545
      Width           =   1680
   End
   Begin VB.Label lblSid 
      Caption         =   "0"
      Height          =   195
      Left            =   8370
      TabIndex        =   25
      Top             =   7590
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblAnistia2 
      BackStyle       =   0  'Transparent
      Caption         =   "Isenção dos juros e multa conforme REFIS-2017 em :"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   3510
      Width           =   3525
   End
   Begin VB.Label lblAnistia3 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   8370
      TabIndex        =   18
      Top             =   3510
      Width           =   375
   End
   Begin VB.Label lblAnistia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   7695
      TabIndex        =   17
      Top             =   3510
      Width           =   600
   End
   Begin VB.Label lblValorExp2 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3780
      TabIndex        =   12
      Top             =   3195
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Vencimento......:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da DAM..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5850
      TabIndex        =   9
      Top             =   2865
      Width           =   1935
   End
   Begin VB.Label lblValorTotal 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   7845
      TabIndex        =   8
      Top             =   2865
      Width           =   1065
   End
   Begin VB.Label lblValorExp 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4725
      TabIndex        =   7
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label lblTotalLanc 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1995
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxa Expediênte...:"
      Height          =   195
      Index           =   3
      Left            =   3255
      TabIndex        =   5
      Top             =   2880
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor dos Lançamentos..:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Menu Anistia 
      Caption         =   "mnuAnistia"
      Visible         =   0   'False
      Begin VB.Menu mnuA1 
         Caption         =   "até 29/09/2017 (100%)"
      End
      Begin VB.Menu mnuA2 
         Caption         =   "até 31/10/2017 (80%)"
      End
      Begin VB.Menu mnuA3 
         Caption         =   "até 30/11/2017 (60%)"
      End
      Begin VB.Menu mnuA4 
         Caption         =   "até 22/12/2017(50%)"
      End
   End
End
Attribute VB_Name = "frmDAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type HeaderArquivo
    nCodigoBanco As String
    nLote As String
    nTipoRegistro As String
    sUsoFebraban1 As String
    nTipoInscricao As String
    nNumInscricao As String
    sCodigoConvenio As String
    nAgencia As String
    sDvAgencia As String
    nNumeroConta As String
    sDvConta As String
    sDvAgenciaConta As String
    sNomeEmpresa As String
    sNomeBanco As String
    sUsoFebraban2 As String
    nCodigoRemessa As String
    nDataGeracao As String
    nHoraGeracao As String
    nNumSeqArquivo As String
    nNumVersaoLayout As String
    nDensidade As String
    sUsoBanco As String
    sUsoEmpresa As String
    sUsoFebraban3 As String
End Type

Private Type TrailerArquivo
    nCodigoBanco As String
    nLote As String
    nTipo As String
    sUsoFebraban1 As String
    nQtdeLote As String
    nQtdeRegistro As String
    nQtdeContas As String
    sUsoFebraban2 As String
End Type

Private Type HeaderLote
    nCodigoBanco As String
    nLote As String
    nTipoRegistro As String
    sTipoOperacao As String
    nTipoServico As String
    sUsoFebraban1 As String
    nNumVersao As String
    sUsoFebraban2 As String
    nTipoInscricao As String
    nNumInscricao As String
    sCodConvenio As String
    nAgencia As String
    sDvAgencia As String
    nNumeroConta As String
    sDvConta As String
    sDvAgenciaConta As String
    sNomeEmpresa As String
    sMensagem1 As String
    sMensagem2 As String
    nNumeroRemessa As String
    sDataGeracao As String
    sDataCredito As String
    sUsoFebraban3 As String
End Type

Private Type TrailerLote
    nCodigoBanco As String
    nLote As String
    nTipo As String
    sUsoFebraban1 As String
    nQtdeRegistro As String
    sUsoFebraban2 As String
End Type

Private Type SegmentoP
    nCodigoBanco As String
    nLote As String
    nTipo As String
    nSeqReg As String
    sCodSegmento As String
    sUsoFebraban1 As String
    nCodMovimento As String
    nAgencia As String
    sDvAgencia As String
    nConta As String
    sDvConta As String
    sNossoNumero As String
    nCodCarteira As String
    nFormaCadastro As String
    sTipoDocumento As String
    nIdentificacaoEmissao As String
    sIdentificacaoDistribuicao As String
    sNumeroDocumento As String
    nDataVencimento As String
    nValorNominal As String
    nAgenciaCobranca As String
    sDvAgenciaCobranca As String
    nEspecieTitulo As String
    sAceite As String
    nDataEmissao As String
    nCodigoJuros As String
    nDataJuros As String
    nJurosMora As String
    nCodigoDesconto1 As String
    nDataDesconto1 As String
    nValorConcedido As String
    nValorIOF As String
    nValorAbatimento As String
    sIdentificaTitulo As String
    nCodigoProtesto As String
    nNumDiasProtesto As String
    nCodigoBaixa As String
    sNumDiasBaixa As String
    nCodigoMoeda As String
    nNumeroContrato As String
    sUsoLivre As String
End Type

Private Type SegmentoQ
    nCodigoBanco As String
    nLote As String
    nTipo As String
    nSeqReg As String
    sCodSegmento As String
    sUsoFebraban1 As String
    nCodMovimento As String
    nTipoInscricao As String
    nNumeroInscricao As String
    sNome As String
    sEndereco As String
    sBairro As String
    nCep As String
    nCepsufixo As String
    sCidade As String
    sUF As String
    nipoInscricaoSacado As String
    nNumeroInscricaoSacado As String
    sNomeSacado As String
    nBancoCorresponde As String
    sNossoNumeroBancoCorr As String
    sUsoFebraban2 As String
End Type

Private Type Boletos
    sNossoNumero As String
    sNumDocumento As String
    sDataVencimento As String
    sValorNominal As String
    sDataBase As String
    nTipoInscricao As String
    nNumeroInscricao As String
    sNome As String
    sEndereco As String
    sBairro As String
    sCep As String
    sSufixoCep As String
    sCidade As String
    sUF As String
End Type

Dim nNumRemessa As Long, sArquivo As String, sDataArquivo As String, aBoletos() As Boletos

Dim sLANCAMENTO As String, sTributo As String, nSomaPrincipal As Double
Dim bCorrecao As Boolean, dVencto As Date, bHonorario As Boolean, nCodigoDam As Long, sDataVenctoDAM As String, bITBI As Boolean, bRefisAtivo As Boolean, bRefisAtivoDI As Boolean
Dim dDataIni As Date, dDataFim As Date, dDataIniDI As Date, dDataFimDI As Date, nPlano As Integer

Public Property Let Honorarios(bValor As Boolean)
    bHonorario = bValor
End Property
Public Property Get Honorarios() As Boolean
    Honorarios = bHonorario
End Property

Public Property Let CodigoDAM(nCodDam As Long)
    nCodigoDam = nCodDam
End Property

Public Property Let VencimentoDAM(sDataVencto As String)
    sDataVenctoDAM = sDataVencto
End Property

Private Sub btDesconto_Click()
CarregaLista2
lblAnistia.Caption = FormatNumber(txtDesconto.Text, 2)
nPerc = 100 - CDbl(lblAnistia.Caption)
With grdTemp
    For x = 1 To grdTemp.Rows - 1
        .TextMatrix(x, 11) = FormatNumber(CDbl(.TextMatrix(x, 11)) * nPerc / 100, 2)
        .TextMatrix(x, 12) = FormatNumber(CDbl(.TextMatrix(x, 12)) * nPerc / 100, 2)
        .TextMatrix(x, 13) = FormatNumber(CDbl(.TextMatrix(x, 9)) + CDbl(.TextMatrix(x, 10)) + CDbl(.TextMatrix(x, 11)) + CDbl(.TextMatrix(x, 12)), 2)
    Next
End With
CalculaTotal
End Sub

Private Sub chkAnistia_Click()
Dim nPerc As Double, bDIS As Boolean, bDIN As Boolean
'Dim x As Integer

'With grdTemp
'    For x = 1 To .Rows - 1
'        If Val(.TextMatrix(x, 0)) <> 2013 Then
'            If Year(CDate(.TextMatrix(x, 6))) > 2010 Then
'                lblAnistia.Caption = "0,00" 'em teste
             '   Exit Sub
'            End If
'        End If
'    Next
'End With

'If chkAnistia.value = vbUnchecked Then
'    lblAnistia.Caption = "0,00"
    CarregaLista2
'Else
'    If dVencto <= CDate("30/11/2016") Then
'       lblAnistia.Caption = "90,00"
'    ElseIf dVencto <= CDate("20/12/2016") Then
'       lblAnistia.Caption = "70,00"
'    ElseIf dVencto >= CDate("21/12/2016") Then
'       If chkMulta.value = 0 Then
'          lblAnistia.Caption = "0,00"
'       Else
'          lblAnistia.Caption = "100,00"
'       End If
'       Exit Sub
'    End If
'        If lblDI.Caption = "S" Then
'            nPerc = 100
'        Else
'            nPerc = 100 - CDbl(lblAnistia.Caption)
 '       End If
'        With grdTemp
'        For x = 1 To grdTemp.Rows - 1
'            .TextMatrix(x, 11) = FormatNumber(CDbl(.TextMatrix(x, 11)) * nPerc / 100, 2)
'            .TextMatrix(x, 12) = FormatNumber(CDbl(.TextMatrix(x, 12)) * nPerc / 100, 2)
'            .TextMatrix(x, 13) = FormatNumber(CDbl(.TextMatrix(x, 9)) + CDbl(.TextMatrix(x, 10)) + CDbl(.TextMatrix(x, 11)) + CDbl(.TextMatrix(x, 12)), 2)
'        Next
 '   End With
'End If
CalculaTotal
End Sub


Private Sub chkCorrecao_Click()
Dim x As Integer, bAchou As Boolean
CarregaLista2
End Sub

Private Sub chkJulgamento_Click()
CarregaLista2
End Sub

Private Sub chkMulta_Click()

Dim x As Integer, bAchou As Boolean
bAchou = False
If chkMulta.value = 1 Then
    With grdTemp
        For x = 1 To .Rows - 2
            If .TextMatrix(x, 7) = "N" Then
                bAchou = True
                Exit For
            End If
        Next
    End With
    lblAnistia.Caption = "100,00"
Else
    If Not bAnistia Then
        lblAnistia.Caption = "0,00"
    End If
End If

CarregaLista2

End Sub

Private Sub chkTx_Click()

If chkTx.value = 1 Then
    'remover
    grdTemp.TextMatrix(grdTemp.Rows - 1, 9) = "0,00"
    grdTemp.TextMatrix(grdTemp.Rows - 1, 13) = "0,00"
    nSomaPrincipal = nSomaPrincipal - CDbl(lblValorExp.Caption)
    lblValorExp.Caption = "0,00"
    
Else
    'adicionar
    grdTemp.TextMatrix(grdTemp.Rows - 1, 9) = lblValorExp2.Caption
    grdTemp.TextMatrix(grdTemp.Rows - 1, 13) = lblValorExp2.Caption
    lblValorExp.Caption = lblValorExp2.Caption
    nSomaPrincipal = nSomaPrincipal + CDbl(lblValorExp.Caption)
End If
CalculaTotal
End Sub

Private Sub chkVenctoAtual_Click()

If Not IsDate(mskVenc.Text) And chkVenctoAtual.value = 1 Then
    MsgBox "Digite uma data válida.", vbExclamation, "Atenção"
    chkVenctoAtual.value = 0
    Exit Sub
End If

CarregaLista2
End Sub

Private Sub cmdAnistia_Click()
If chkAnistia.value = vbUnchecked Then
    MsgBox "REFIS não disponível para este débito.", vbExclamation, "Atenção!"
Else
    PopupMenu Anistia
End If
End Sub

Private Sub cmdBaixa_Click()
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer
Dim nCompl As Integer, x As Integer, aAno() As Integer, Y As Integer, bAchou As Boolean, sVento As String
Dim bFindOldYear As Boolean

'If nSeqFator = nSeqFator2 Then GoTo FIM2

EmiteBoleto
Exit Sub


If bAnistia Then
    bFindOldYear = False
    With grdTemp
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 0)) <> 2016 Then
                bFindOldYear = True
            End If
        Next
    End With
    bAchou = False
    With grdTemp
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 0)) = 2016 Then
                bAchou = True
            End If
        Next
    End With
    If bFindOldYear Then
        With grdTemp
            For x = 1 To .Rows - 1
                If Val(.TextMatrix(x, 0)) = 2016 And Val(.TextMatrix(x, 1)) <> 4 And Val(.TextMatrix(x, 1)) <> 41 And Val(.TextMatrix(x, 1)) <> 69 Then
                    MsgBox "Não é permitido emitir débitos de 2016 junto com outros anos.", vbExclamation, "Atenção"
                    lblAnistia.Caption = "0,00"
                    Exit Sub
                End If
            Next
        End With
    End If
End If


nCodReduz = nCodigoDam
If (nCodReduz < 100000 Or nCodReduz > 300000) And chkJulgamento.value = 1 Then
    MsgBox "Apuração Fiscal apenas para débitos mobiliários.", vbCritical, "Atenção"
    Liberado
    Exit Sub
End If

If Not IsDate(mskVencimento.Text) Then
    MsgBox "Data de Vencimento inválido.", vbCritical, "Atenção"
    Exit Sub
End If

'carrega os anos distintos
ReDim aAno(0)
For x = 1 To grdTemp.Rows - 1
    If Val(grdTemp.TextMatrix(x, 1)) <> 4 And Val(grdTemp.TextMatrix(x, 1)) <> 41 Then
        nAno = grdTemp.TextMatrix(x, 0)
        bAchou = False
        For Y = 1 To UBound(aAno)
            If aAno(Y) = nAno Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = nAno
        End If
    End If
Next


Ocupado
GravaDam

'MUDA O STATUS DAS PARCELAS SELECIONADAS PARA EM JULGAMENTO
If chkJulgamento.value = 1 Then
    With frmDebitoImob.grdExtrato
        For x = 1 To .Rows
            If .CellText(x, 12) = "S" Then
                nAno = Val(.CellText(x, 1))
                nLanc = Val(Left$(.CellText(x, 2), 3))
                nSeq = Val(.CellText(x, 3))
                nParc = Val(.CellText(x, 4))
                nCompl = Val(.CellText(x, 5))
               .CellText(x, 12) = ""
               .CellText(x, 6) = "20 - EM JULGAMENTO"
               .CellForeColor(.Rows, 6) = &H80FF&
                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=20"
                Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
                Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
                Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
                cn.Execute Sql, rdExecDirect
            End If
        Next
    End With
End If

Liberado
Unload Me

Exit Sub
Fim2:
'MainG2

End Sub




Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset

Sql = "select valparam from parametros where nomeparam='REFIS_INICIO'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataIni = CDate(RdoAux!valparam)

Sql = "select valparam from parametros where nomeparam='REFIS_FIM'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataFim = CDate(RdoAux!valparam)

Sql = "select valparam from parametros where nomeparam='REFISDI_INICIO'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataIniDI = CDate(RdoAux!valparam)

Sql = "select valparam from parametros where nomeparam='REFISDI_FIM'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataFimDI = CDate(RdoAux!valparam)

RdoAux.Close

nPlano = 0
Centraliza Me
Me.Top = Me.Top + 1200
Ocupado
   
'If bRefisAtivo Then
'    Select Case NomeDeLogin
'        Case "ROSE", "JOSEANE", "RENATA", "PAULA", "LUIZH", "ROSANGELA", "RODRIGOC", "LEANDRO", "DANIELAR", "SOLANGE", "RHENO.SOARES", "CARMELINO", "SCHWARTZ"
'            chkAnistia.Enabled = True
'        Case Else
'            chkAnistia.Enabled = False
'    End Select
'Else
'    cmdAnistia.Visible = False
'    chkAnistia.value = vbUnchecked
'End If
   
'lblAnistia.Visible = bAnistia
'lblAnistia2.Visible = bAnistia
'lblAnistia3.Visible = bAnistia
'chkAnistia.Visible = bAnistia
   
'If Val(frmDebitoImob.txtCod) = 523872 Then bCorrecao = False

mskVencimento.Text = sDataVenctoDAM
Sql = "DELETE FROM DAM WHERE usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

CarregaLista2

Liberado
Select Case NomeDeLogin
    Case "SCHWARTZ", "RENATA", "GLEISE", "ROSE", "RITA", "LUIZH", "RODRIGOC", "JOSEANE", "SOLANGE", "ANA"
        chkVenctoAtual.Enabled = True
        mskVenc.Enabled = True
    Case Else
        chkVenctoAtual.Enabled = False
        mskVenc.Enabled = False
End Select


If NomeDeLogin <> "RENATA" And NomeDeLogin <> "SOLANGE" And NomeDeLogin <> "ROSE" Then
    chkCobranca.Enabled = False
    chkCorrecao.Enabled = False
    chkMulta.Enabled = False
End If



If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "RENATA" And NomeDeLogin <> "SOLANGE" And NomeDeLogin <> "LUIZH" And NomeDeLogin <> "ROSE" Then
    txtDesconto.Enabled = False
    btDesconto.Enabled = False
    UpDown.Enabled = False
End If

If NomeDeLogin = "SCHWARTZ" Then
    chkRegistrado.Enabled = True
End If

End Sub

Private Sub CarregaLista2()
Dim x As Integer, Y As Integer, nCodReduz As Long, aLanc() As String, Achou As Boolean, nRowGrid As Integer, nY1 As Integer
Dim sAno As String, sLanc As String, sSeq As String, sParc As String, aAno() As Integer
Dim sComp As String, sSit As String, sVencto As String, sDA As String, RdoTrib As rdoResultset
Dim sAj As String, nValorPrincipal As Double, sDataBase As String, bDA As Boolean, dVenctoMinimo As Date
Dim nValorCorrecao As Double, nValorJuros As Double, nValorMulta As Double, nValorTotal As Double, nPerc As Double, nValorHon As Double
Dim nSomaTotal As Double, nSomaHon As Double, bJuros As Boolean, bMulta As Boolean, nCodTrib As Integer, qd As New rdoQuery, nSid As Long
Dim bDIS As Boolean, bDIN As Boolean, bFind As Boolean


Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
On Error Resume Next
ReDim aLanc(0)

If chkVenctoAtual.value = 0 Then
    If mskVencimento.ClipText = "" Then mskVencimento.Text = Right(frmMdi.Sbar.Panels(6).Text, 10)
    dVencto = Format(mskVencimento.Text, "dd/mm/yyyy")
Else
    dVencto = Format(mskVenc.Text, "dd/mm/yyyy")
End If
bITBI = False
bDIS = False: bDIN = False

grdTemp.Rows = 1
nSomaHon = 0
nSomaTotal = 0
nSomaPrincipal = 0
With frmDebitoImob.grdExtrato
    nCodReduz = nCodigoDam
    For nRowGrid = 1 To .Rows
        bJuros = False: bMulta = False
        If .CellText(nRowGrid, 12) = "S" Then
           sAno = .CellText(nRowGrid, 1)
           sLanc = Left$(.CellText(nRowGrid, 2), 3)
           
           If Val(sLanc) = 81 Then
              bDIS = True
           Else
              bDIN = True
           End If
           
           sLANCAMENTO = Right$(.CellText(nRowGrid, 2), Len(.CellText(nRowGrid, 2)) - 5)
           Achou = False
           For nY1 = 1 To UBound(aLanc)
               If aLanc(nY1) = sLANCAMENTO Then
                  Achou = True
                  Exit For
               End If
           Next
           If Not Achou Then
              ReDim Preserve aLanc(UBound(aLanc) + 1)
              aLanc(UBound(aLanc)) = sLANCAMENTO
           End If
           sSeq = .CellText(nRowGrid, 3)
           sParc = IIf(.CellText(nRowGrid, 4) = "Unica", "00", .CellText(nRowGrid, 4))
           sComp = .CellText(nRowGrid, 5)
           sSit = Left$(.CellText(nRowGrid, 6), 2)
           sVencto = .CellText(nRowGrid, 7)
'
           sDA = .CellText(nRowGrid, 8)
           sAj = .CellText(nRowGrid, 9)
           
            '***********************
                       'CARREGA O EXTRATO
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = nCodReduz
            qd(2) = Val(sAno)
            qd(3) = Val(sAno)
            qd(4) = Val(sLanc)
            qd(5) = Val(sLanc)
            qd(6) = Val(sSeq)
            qd(7) = Val(sSeq)
            qd(8) = Val(sParc)
            qd(9) = Val(sParc)
            qd(10) = Val(sComp)
            qd(11) = Val(sComp)
            qd(12) = 1
            qd(13) = 99
            qd(14) = Format(dVencto, "mm/dd/yyyy")
            qd(15) = NomeDeLogin
            Set RdoAux = qd.OpenResultset(rdOpenKeyset)
            With RdoAux
                sDataBase = !DATADEBASE
                sTributo = "": nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValorTotal = 0: bITBI = False
                Do Until .EOF
                    If !CodTributo = 84 Then bITBI = True
                    sTributo = sTributo & Format(!CodTributo, "000") & "-" & !ABREVTRIBUTO & "/ "
                    nValorPrincipal = nValorPrincipal + !ValorTributo
                    nValorJuros = nValorJuros + !ValorJuros
                    nValorCorrecao = nValorCorrecao + !ValorCorrecao
                    If MI And !CodLancamento = 5 Then
                        nValorMulta = 0
                        nValorTotal = nValorTotal + !ValorTributo + !ValorJuros + !ValorCorrecao
                    Else
                        nValorMulta = nValorMulta + !ValorMulta
                        nValorTotal = nValorTotal + !ValorTotal
                    End If
                    
                   .MoveNext
                Loop
               .Close
            End With
           
'**************************************************************
           
           If bITBI Then
                sTributo = ""
                Sql = "SELECT * FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & Val(sAno)
                Sql = Sql & " AND CODLANCAMENTO=" & Val(sLanc) & " AND SEQLANCAMENTO=" & Val(sSeq) & " AND NUMPARCELA=" & Val(sParc)
                Sql = Sql & " AND CODCOMPLEMENTO=" & Val(sComp)
                
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                        Do Until .EOF
                             If UCase$(Left(!obs, 2)) <> "LA" Then
                                sTributo = sTributo & SubNull(!obs) & " "
                             End If
                            .MoveNext
                        Loop
                    End If
                   .Close
                End With
           End If
           
           nSomaTotal = nSomaTotal + nValorTotal
           nSomaPrincipal = nSomaPrincipal + nValorPrincipal
          If sTributo <> "" Then
            sTributo = Left(sTributo, Len(sTributo) - 1)
          End If
           grdTrib.AddItem "001" & Chr(9) & sTributo
                       
           If chkMulta.value = 1 Then
                nValorTotal = nValorTotal - nValorJuros - nValorMulta
                nValorJuros = 0: nValorMulta = 0
           End If
           If chkCorrecao.value = 1 Then
                nValorTotal = nValorTotal - nValorCorrecao
                nValorCorrecao = 0
           End If
                       
                       
           grdTemp.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
           sComp & Chr(9) & sSit & Chr(9) & sVencto & Chr(9) & sDA & Chr(9) & sAj & Chr(9) & _
           FormatNumber(nValorPrincipal, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2) & Chr(9) & _
           FormatNumber(nValorMulta, 2) & Chr(9) & FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorTotal, 2)
           
        End If
    Next
    
    sLANCAMENTO = ""
    For nY1 = 1 To UBound(aLanc)
        sLANCAMENTO = sLANCAMENTO & aLanc(nY1) & "/ "
    Next
    sLANCAMENTO = Left(sLANCAMENTO, Len(sLANCAMENTO) - 2)

End With

If bDIS And Not bDIN Then
    lblDI.Caption = "S"
Else
    lblDI.Caption = "N"
End If


If lblDI.Caption = "N" Then
    If Now >= dDataIni And Now <= dDataFim Then
        bRefisAtivo = True
'        chkAnistia.value = vbChecked
    Else
        bRefisAtivo = False
 '       chkAnistia = vbUnchecked
    End If
Else
    If Now >= dDataIniDI And Now <= dDataFimDI Then
        bRefisAtivoDI = True
  '      chkAnistia = vbChecked
    Else
        bRefisAtivoDI = False
   '     chkAnistia = vbUnchecked
    End If
End If

If lblDI.Caption = "S" Then
    cmdAnistia.Visible = False
Else
    If bRefisAtivo Then
        cmdAnistia.Visible = True
    End If
End If

'chkAnistia.Visible = False
lblAnistia.Visible = True
lblAnistia2.Visible = True
lblAnistia3.Visible = True
lblAnistia.Caption = "0,00"

'****** REFIS *************
If chkAnistia.value = vbChecked Then
    If Year(Now) = 2017 Then
        If bRefisAtivoDI Then
            '******** 2017 ********
            If lblDI.Caption = "S" Then
                
                With frmDebitoImob.grdExtrato
                    bFind = False
                    For x = 1 To .Rows
                        If Val(Left$(.CellText(x, 6), 2)) = 3 And Val(.CellText(x, 2)) = 81 And .CellText(x, 12) <> "S" Then
                            sVencto = .CellText(x, 7)
                            If CDate(sVencto) < Now Then
                                bFind = True
                                Exit For
                            End If
                        End If
                    Next
                End With
               
                If Not bFind Then
                    '** DISTRITO INDUSTRIAL **
                    chkAnistia.Visible = True
                    lblAnistia.Caption = "100,00"
                    lblAnistia.Visible = True
                    lblAnistia2Visible = True
                    lblAnistia3Visible = True
                    nPerc = 100 - CDbl(lblAnistia.Caption)
                    With grdTemp
                        For nY1 = 1 To grdTemp.Rows - 1
                            If .TextMatrix(nY1, 1) = 81 Then
                                .TextMatrix(nY1, 11) = FormatNumber(CDbl(.TextMatrix(nY1, 11)) * nPerc / 100, 2)
                                .TextMatrix(nY1, 12) = FormatNumber(CDbl(.TextMatrix(nY1, 12)) * nPerc / 100, 2)
                            End If
                            .TextMatrix(nY1, 13) = FormatNumber(CDbl(.TextMatrix(nY1, 9)) + CDbl(.TextMatrix(nY1, 10)) + CDbl(.TextMatrix(nY1, 11)) + CDbl(.TextMatrix(nY1, 12)), 2)
                        Next
                    End With
                End If
            End If
        Else
            '*** OUTROS ***
            If bRefisAtivo Then
                '******** 2017 ********
                If lblDI.Caption = "N" Then
                    
                    With grdTemp
                        bFind = False
                        For x = 1 To .Rows - 1
                            If CDate(.TextMatrix(x, 6)) > CDate("31/12/2016") And .TextMatrix(x, 1) <> 41 And .TextMatrix(x, 1) <> 78 And .TextMatrix(x, 1) <> 45 Then
                                bFind = True
                                Exit For
                            End If
                        Next
                    End With
                   
                    If Not bFind Then
                        If nPlano = 0 Then
                            If dVencto <= CDate("29/09/2017") Then
                                nPlano = 16
                            ElseIf dVencto > CDate("29/09/2017") And dVencto <= CDate("31/10/2017") Then
                                nPlano = 17
                            ElseIf dVencto > CDate("31/10/2017") And dVencto <= CDate("30/11/2017") Then
                                nPlano = 18
                            ElseIf dVencto > CDate("30/11/2017") And dVencto <= CDate("22/12/2017") Then
                                nPlano = 19
                            End If
                        End If
                                                
                        Sql = "select desconto from plano where codigo=" & nPlano
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        nPerc = RdoAux2!desconto
                        RdoAux2.Close
                        
                        chkAnistia.Visible = True
                        lblAnistia.Caption = FormatNumber(nPerc, 2)
                        lblAnistia.Visible = True
                        lblAnistia2Visible = True
                        lblAnistia3Visible = True
                        'nPerc = 100 - CDbl(lblAnistia.Caption)
                        With grdTemp
                            For nY1 = 1 To grdTemp.Rows - 1
                                .TextMatrix(nY1, 11) = FormatNumber(CDbl(.TextMatrix(nY1, 11)) - CDbl(.TextMatrix(nY1, 11)) * nPerc / 100, 2)
                                .TextMatrix(nY1, 12) = FormatNumber(CDbl(.TextMatrix(nY1, 12)) - CDbl(.TextMatrix(nY1, 12)) * nPerc / 100, 2)
                                .TextMatrix(nY1, 13) = FormatNumber(CDbl(.TextMatrix(nY1, 9)) + CDbl(.TextMatrix(nY1, 10)) + CDbl(.TextMatrix(nY1, 11)) + CDbl(.TextMatrix(nY1, 12)), 2)
                            Next
                        End With
                    End If
                End If
            End If
            
        End If
        '**********************
    End If
    
End If
'**************************

'

'***** ANISTIA *****
'If bAnistia And chkAnistia.value = vbChecked Then
'
'    With grdTemp
'        For nY1 = 1 To .Rows - 1
'            If .TextMatrix(nY1, 1) = 74 Then
'                chkAnistia.value = vbUnchecked
'                lblAnistia.Caption = "0,00"
'                GoTo continua
'            End If
'            'If CDate(.TextMatrix(nY1, 6)) > CDate("31/12/2015") And .TextMatrix(nY1, 1) <> 41 And .TextMatrix(nY1, 1) <> 69 And .TextMatrix(nY1, 1) <> 5 And .TextMatrix(nY1, 1) <> 78 Then
'            If CDate(.TextMatrix(nY1, 6)) > CDate("31/12/2015") And .TextMatrix(nY1, 1) <> 41 And .TextMatrix(nY1, 1) <> 78 Then
'                chkAnistia.value = vbUnchecked
'                lblAnistia.Caption = "0,00"
'                GoTo continua
'            End If
'            If .TextMatrix(nY1, 1) = 11 Then
'                Sql = "select * from debitotributo where codreduzido=" & nCodReduz & " and anoexercicio=" & .TextMatrix(nY1, 0) & " and "
'                Sql = Sql & "codlancamento=11 and seqlancamento=" & .TextMatrix(nY1, 2) & " and numparcela=" & .TextMatrix(nY1, 3) & " and "
'                Sql = Sql & " codcomplemento=" & .TextMatrix(nY1, 4) & " and codtributo=527"
'                Set RdoTrib = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                If RdoTrib.RowCount > 0 Then
'                    RdoTrib.Close
'                    chkAnistia.value = vbUnchecked
'                    lblAnistia.Caption = "0,00"
'                    GoTo continua
'                End If
'                RdoTrib.Close
'            End If
'        Next
'    End With
'
'
'
'    If dVencto <= CDate("30/11/2016") Then
'       lblAnistia.Caption = "90,00"
'    ElseIf dVencto <= CDate("20/12/2016") Then
'       lblAnistia.Caption = "70,00"
'    ElseIf dVencto >= CDate("21/12/2016") Then
'       If chkMulta.value = 0 Then
'          lblAnistia.Caption = "0,00"
'       Else
'          lblAnistia.Caption = "100,00"
'       End If
'       GoTo continua
'     End If
'    nPerc = 100 - CDbl(lblAnistia.Caption)
'    With grdTemp
'        For nY1 = 1 To grdTemp.Rows - 1
'            If .TextMatrix(nY1, 1) <> 41 Then
'                .TextMatrix(nY1, 11) = FormatNumber(CDbl(.TextMatrix(nY1, 11)) * nPerc / 100, 2)
'                .TextMatrix(nY1, 12) = FormatNumber(CDbl(.TextMatrix(nY1, 12)) * nPerc / 100, 2)
'            End If
'            .TextMatrix(nY1, 13) = FormatNumber(CDbl(.TextMatrix(nY1, 9)) + CDbl(.TextMatrix(nY1, 10)) + CDbl(.TextMatrix(nY1, 11)) + CDbl(.TextMatrix(nY1, 12)), 2)
'        Next
'    End With
'Else
'    lblAnistia.Caption = "0,00"
'End If
'*******************
continua:


'Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now) & " AND CODLANCAMENTO = 1"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
     lblValorExp.Caption = FormatNumber(0, 2)
     lblValorExp2.Caption = FormatNumber(0, 2)
     lblTotalLanc.Caption = FormatNumber(nSomaTotal, 2)
   ' .Close
'End With


'taxa expediente
' grdTemp.AddItem Year(Now) & Chr(9) & "004" & Chr(9) & "00" & Chr(9) & "01" & Chr(9) & _
  "0" & Chr(9) & "03" & Chr(9) & Format(Now, "dd/mm/yyyy") & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & _
  lblValorExp.Caption & Chr(9) & FormatNumber(0, 2) & Chr(9) & _
  FormatNumber(0, 2) & Chr(9) & FormatNumber(0, 2) & Chr(9) & lblValorExp.Caption
'sTributo = "003" & "-" & "TAXA EXP.DOC."
'grdTrib.AddItem "003" & Chr(9) & sTributo

'HONORARIOS
nSomaHon = 0
If bHonorario Then
    For nY1 = 1 To grdTemp.Rows - 1
        nSomaHon = nSomaHon + (grdTemp.TextMatrix(nY1, 13) * 10 / 100)
    Next
    nSomaHon = FormatNumber(nSomaHon, 2)
    grdTemp.AddItem Year(Now) & Chr(9) & "041" & Chr(9) & "00" & Chr(9) & "01" & Chr(9) & _
    "0" & Chr(9) & "03" & Chr(9) & Format(Now, "dd/mm/yyyy") & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & _
    nSomaHon & Chr(9) & FormatNumber(0, 2) & Chr(9) & _
    FormatNumber(0, 2) & Chr(9) & FormatNumber(0, 2) & Chr(9) & nSomaHon
    sTributo = "041" & "-" & "HONORÁRIOS"
    grdTrib.AddItem "041" & Chr(9) & sTributo
End If

If lblValorExp.Caption > 0 Then
    nSomaPrincipal = nSomaPrincipal + CDbl(lblValorExp.Caption)
End If
CalculaTotal

End Sub

Private Sub GravaDam()

On Error GoTo Erro

Dim x As Integer
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery, RdoAux3 As rdoResultset
Dim sNumInsc As String, sValorParc As String, sData As String, sObs As String
Dim nCodReduz As Long, nNumGuia As Long
Dim sNomeResp As String
Dim sEndImovel As String
Dim nNumImovel As Integer
Dim sComplImovel As String
Dim sBairroImovel As String
Dim nCodCidade As Integer
Dim nCodBairro As Integer
Dim sCidadeEntrega As String
Dim sUFEntrega As String
Dim sCPF As String
Dim nAno As Integer
Dim nNumDoc As Long
Dim sQuadra As String
Dim sLote As String
Dim nNumParc As Integer
Dim dDataVencto As Date
Dim nCodLanc As Integer
Dim nSeq As Integer, nSeq2 As Integer
Dim nComplemento As Integer
Dim nValorTotal As Double
Dim NumBarra2 As String
Dim NumBarra2a As String
Dim NumBarra2b As String
Dim NumBarra2c As String
Dim NumBarra2d As String
Dim StrBarra2 As String
Dim nLastCod As Long
Dim nValorTaxa As Double
Dim nLanc As Integer, nParc As Integer, nPlano As Integer
Dim nComp As Integer, bMulta As Boolean
Dim nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nSid As Long

If MsgBox("Confirma Emissão da DAM ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nSid = Int(Rnd(10) * 1000000)
lblSid.Caption = nSid

'DELETA TEMPORARIO
'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "' and usuario='" & NomeDeLogin & "'"
Sql = "DELETE FROM DAM WHERE SID=" & nSid
cn.Execute Sql, rdExecDirect

'RETORNA VALOR EXPEDIENTE
'Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE CODLANCAMENTO=3 AND ANOEXPED=" & Year(Now)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'nValorTaxa = RdoAux!VALORDAM
'RdoAux.Close
nValorTaxa = 1.6

nCodReduz = nCodigoDam
sNomeResp = frmDebitoImob.lblProp.Caption

Select Case nCodReduz
    Case 1 To 99999
        Sql = "SELECT * FROM vwCnsImovel WHERE CODREDUZIDO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
             If .RowCount > 0 Then
                sNumInsc = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
                sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                nNumImovel = Val(SubNull(!Li_Num))
                sComplImovel = SubNull(!Li_Compl)
                nCodBairro = !CodBairro
                sCidadeEntrega = SubNull(!descCidade)
                sUFEntrega = SubNull(!li_uf)
                nCodCidade = !LI_CODCIDADE
                sQuadra = SubNull(!Li_Quadras)
                sLote = SubNull(!Li_Lotes)
                Sql = "SELECT CODREDUZIDO,CPF,CNPJ,RG,ORGAO FROM vwCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & nCodReduz
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If Not IsNull(!CPF) Then
                       sCPF = !CPF
                    ElseIf Not IsNull(!Cnpj) Then
                       sCPF = !Cnpj
                    ElseIf Not IsNull(!rg) Then
                       sCPF = !rg
                    Else
                       sCPF = ""
                    End If
                End With
            End If
        End With
     Case 100000 To 500000
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,CNPJ,CPF,RAZAOSOCIAL,DESCCIDADE,SIGLAUF,CODCIDADE,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO,NOMELOGR "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sNumInsc = !INSCESTADUAL
                sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
                If Trim(sEndImovel) = "" Then
                    sEndImovel = SubNull(!NomeLogr)
                End If
                nNumImovel = Val(SubNull(!Numero))
                sComplImovel = SubNull(!Complemento)
                nCodBairro = !CodBairro
                sCidadeEntrega = SubNull(!descCidade)
                sUFEntrega = SubNull(!SiglaUF)
                nCodCidade = !CodCidade
                sQuadra = "0"
                sLote = "0"
                If Not IsNull(!CPF) Then
                   sCPF = !CPF
                ElseIf Not IsNull(!Cnpj) Then
                   sCPF = !Cnpj
                Else
                    sCPF = ""
                End If
            End If
         End With
     Case 500000 To 800000
        Sql = "SELECT cidadao.codcidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, "
        Sql = Sql & "cidadao.siglauf, cidade.desccidade, bairro.descbairro, cidadao.nomelogradouro AS nomerua,"
        Sql = Sql & "Cidadao.codlogradouro , vwLOGRADOURO.AbrevTipoLog, vwLOGRADOURO.AbrevTitLog, vwLOGRADOURO.NomeLogradouro FROM bairro RIGHT OUTER JOIN "
        Sql = Sql & "cidade RIGHT OUTER JOIN cidadao ON cidade.siglauf = cidadao.siglauf AND cidade.codcidade = cidadao.codcidade LEFT OUTER JOIN "
        Sql = Sql & "vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO ON bairro.siglauf = cidadao.siglauf AND bairro.codcidade = Cidadao.codcidade And bairro.codbairro = Cidadao.codbairro "
        Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If Not IsNull(!NomeLogradouro) Then
                    sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
                Else
                    sEndImovel = SubNull(!nomerua)
                End If
                nNumImovel = Val(SubNull(!NUMIMOVEL))
                sComplImovel = SubNull(!Complemento)
                nCodBairro = Val(SubNull(!CodBairro))
                nCodCidade = Val(SubNull(!CodCidade))
                sCidadeEntrega = SubNull(!descCidade)
                sUFEntrega = SubNull(!SiglaUF)
                If Not IsNull(!CPF) And Trim$(SubNull(!CPF)) <> "" Then
                   sCPF = !CPF
                ElseIf Not IsNull(!Cnpj) And Trim$(SubNull(!Cnpj)) <> "" Then
                   sCPF = !Cnpj
                ElseIf Not IsNull(!rg) Then
                   sCPF = !rg
                Else
                   sCPF = ""
                End If
             Else
                sCPF = ""
             End If
             If sCidadeEntrega = "" Then
                sCidadeEntrega = SubNull(!descCidade)
             End If
             If nCodBairro = 0 Then
                sBairroImovel = SubNull(!DescBairro)
                
                GoTo FIMBAIRRO
             End If
        End With
End Select


If nCodCidade = 0 Then GoTo FIMBAIRRO
Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & sUFEntrega & "' AND CODCIDADE=" & nCodCidade & " AND CODBAIRRO=" & nCodBairro
Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux3
    If .RowCount > 0 Then
         sBairroImovel = !DescBairro
    Else
         sBairroImovel = ""
    End If
   .Close
End With
FIMBAIRRO:
'TOTAL
nValorTotal = CDbl(lblTotalLanc.Caption)


'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
   nLastCod = 1
Else
   nLastCod = RdoAux!maximo + 1
End If
RdoAux.Close
'If dVencto = "00:00:00" Then dVencto = Format(Now, "dd/mm/yyyy")
dVencto = sDataVenctoDAM
'GERAÇÃO DOS DÉBITOS
With grdTemp
   'GRAVA NUMDOCUMENTO
    If chkMulta.value = 1 Then
       bMulta = True
    Else
       If Val(lblAnistia.Caption) > 0 Then
           bMulta = True
       Else
           bMulta = False
       End If
    End If
       
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO,TIPODOC,emissor) VALUES("
    If chkJulgamento.value = 0 Then
        Sql = Sql & nLastCod & ",'" & Format(dVencto, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ",1,'" & NomeDeLogin & " (DAM)" & "')"
    Else
        Sql = Sql & nLastCod & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ",1,'" & NomeDeLogin & " (DAM)" & "')"
    End If
    cn.Execute Sql, rdExecDirect
    For x = 1 To .Rows - 2
        nAno = Val(.TextMatrix(x, 0))
        nLanc = Val(.TextMatrix(x, 1))
        nSeq = Val(.TextMatrix(x, 2))
        nParc = Val(.TextMatrix(x, 3))
        nComp = Val(.TextMatrix(x, 4))
        nValorJuros = FormatNumber(CDbl(grdTemp.TextMatrix(x, 12)), 2)
        nValorMulta = FormatNumber(CDbl(.TextMatrix(x, 11)), 2)
        nValorCorrecao = FormatNumber(CDbl(.TextMatrix(x, 10)), 2)
       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,VALORJUROS,VALORMULTA,VALORCORRECAO,PLANO) VALUES(" & nCodReduz & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nComp & "," & nLastCod & ","
        Sql = Sql & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & "," & Virg2Ponto(CStr(nValorCorrecao)) & ",10," & ")"
        cn.Execute Sql, rdExecDirect
        If Val(lblAnistia.Caption) > 0 And bAnistia Then
            'GRAVA OBS PARCELA
             Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
             Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
             Sql = Sql & " AND CODCOMPLEMENTO=" & nComp
             Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux
                 If IsNull(!maximo) Then
                     nSeq2 = 1
                 Else
                     nSeq2 = !maximo + 1
                 End If
                .Close
             End With
             sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
             sObs = "Lancamento incluido na DAM número " & nLastCod & " com " & lblAnistia.Caption & "% de desconto em multa e juros conforme REFIS-2016"
'             Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & nCodReduz & "," & nAno & ","
'             Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nComp & "," & nSeq2 & ",'" & sObs & "','" & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
             Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduz & "," & nAno & ","
             Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nComp & "," & nSeq2 & ",'" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "')"
             cn.Execute Sql, rdExecDirect
        End If
    Next
End With

'CRIA VINCULO COM ISSELETRONICO SE HOUVER
If bAnistia Then
    If UBound(aDocDAM) > 0 Then
        For x = 1 To UBound(aDocDAM)
            Sql = "INSERT DAMISS(DOCDAM,DOCISS) VALUES(" & nLastCod & "," & aDocDAM(x) & ")"
            cn.Execute Sql, rdExecDirect
        Next
        ReDim aDocDAM(0)
    End If
End If

'ATUALIZA DEBITOTRIBUTO
With grdTemp
    For x = 1 To .Rows - 1
        nAno = Val(.TextMatrix(x, 0))
        nLanc = Val(.TextMatrix(x, 1))
        nSeq = Val(.TextMatrix(x, 2))
        nParc = Val(.TextMatrix(x, 3))
        nComp = Val(.TextMatrix(x, 4))
        If nLanc <> 4 Then
            Sql = "UPDATE DEBITOTRIBUTO SET VALORCORRECAO=" & Virg2Ponto(RemovePonto(.TextMatrix(x, 10))) & ","
            Sql = Sql & "VALORMULTA=" & Virg2Ponto(RemovePonto(.TextMatrix(x, 11))) & ","
            Sql = Sql & "VALORJUROS=" & Virg2Ponto(RemovePonto(.TextMatrix(x, 12))) & " WHERE CODREDUZIDO=" & nCodReduz
            Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
            Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
            Sql = Sql & " AND CODCOMPLEMENTO=" & nComp
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With

'DELETA TEMPORARIO
'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "' and usuario='" & NomeDeLogin & "'"
Sql = "DELETE FROM DAM WHERE SID=" & nSid
cn.Execute Sql, rdExecDirect

Set qd.ActiveConnection = cn

'GRAVA TEMPORARIO
With grdTemp
    For x = 1 To .Rows - 1
        nAno = Year(dVencto)
        nCodLanc = 1
        nSeq = 0
        nNumParc = 1
        nComplemento = 0
        dDataVencto = Format(dVencto, "dd/mm/yyyy")
        sValorParc = FormatNumber(lblValorTotal.Caption, 2)
        nNumDoc = nLastCod
        'NumBarra2 = Gera2of5Cod(sValorParc, dDataVencto, nNumDoc, nNumParc, nCodLanc, nSeq, nComplemento)
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        
        Sql = "INSERT DAM(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,CPF,ENDERECO,NUMERO,COMPLEMENTO,"
        Sql = Sql & "BAIRRO,CIDADE,UF,QUADRA,LOTE,FULLLANC,FULLTRIB,NUMDAM,ANOEXERC,LANC,NUMSEQ,NUMPARCELA,COMP,DATAVENCTO,"
        Sql = Sql & "SIT,AJ,DA,PRINCIPAL,CORRECAO,MULTA,JUROS,TOTAL,STRBARRA2,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
        Sql = Sql & "VALORDAM,VALORPRINCDAM,CODTRIBUTO,USUARIO,SID) VALUES('" & NomeDoComputador & "'," & x & ",'" & sNumInsc & "','"
        Sql = Sql & Format(nCodReduz, "000000") & "','" & "DAM" & "','" & Mask(Left$(sNomeResp, 40)) & "','" & Left(sCPF, 20) & "','" & Left$(sEndImovel, 40) & "',"
        Sql = Sql & nNumImovel & ",'" & Left$(sComplImovel, 30) & "','" & Left$(sBairroImovel, 25) & "','" & sCidadeEntrega & "','" & sUFEntrega & "','"
        Sql = Sql & Left(Mask(sQuadra), 15) & "','" & Left$(Mask(sLote), 10) & "','" & sLANCAMENTO & "','" & Left$(Mask(grdTrib.TextMatrix(x, 1)), 2000) & "','"
        Sql = Sql & CStr(nLastCod) & CStr(RetornaDVNumDoc(nLastCod)) & "','" & .TextMatrix(x, 0) & "','" & .TextMatrix(x, 1) & "','"
        Sql = Sql & .TextMatrix(x, 2) & "','" & .TextMatrix(x, 3) & "','" & .TextMatrix(x, 4) & "','" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','"
        Sql = Sql & .TextMatrix(x, 5) & "','" & .TextMatrix(x, 7) & "','" & .TextMatrix(x, 8) & "'," & Virg2Ponto(RemovePonto(.TextMatrix(x, 9))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(.TextMatrix(x, 10))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 11))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 12))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) & ",'"
        Sql = Sql & Mask(StrBarra2) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "'," & Virg2Ponto(RemovePonto(sValorParc)) & ","
        Sql = Sql & Virg2Ponto(CStr(Format(nSomaPrincipal, "#0.00"))) & "," & Val(Left$(Mask(grdTrib.TextMatrix(x, 1)), 3)) & ",'" & NomeDeLogin & "'," & nSid & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

modLg "Emissão de DAM nº " & CStr(nLastCod)
nNumGuia = nNumDoc
nNumDoc = lblSid.Caption
'EXIBE RELATORIO
If bHonorario Then
    frmReport.ShowReport "DAMHONORARIO", frmMdi.hwnd, Me.hwnd, nNumDoc, nNumGuia
Else
    If frmMdi.frTeste.Visible = True Then
        frmReport.ShowReport "DAMTMP", frmMdi.hwnd, Me.hwnd, nNumDoc, nNumGuia
    Else
        frmReport.ShowReport "DAM", frmMdi.hwnd, Me.hwnd, nNumDoc, nNumGuia
    End If
End If

'DELETA TEMPORARIO
'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "' and usuario='" & NomeDeLogin & "'"
Sql = "DELETE FROM DAM WHERE SID=" & nSid
cn.Execute Sql, rdExecDirect

Exit Sub

Erro:
For Y = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(Y).Description
Next
Resume Next

End Sub

Private Sub CalculaTotal()
Dim nTotal As Double, x As Integer

For x = 1 To grdTemp.Rows - 1
    nTotal = nTotal + grdTemp.TextMatrix(x, 13)
Next

lblValorTotal.Caption = FormatNumber(nTotal, 2)
End Sub

Private Function CalculaCorrecaoDAM(nValorDebito As Double, dDataBase As Date) As Double

Dim RdoAux As rdoResultset, Sql As String
Dim UfirAtual As Double
Dim UfirBase As Double, dDataVencto As Date

If chkVenctoAtual.value = 0 Then
    dDataVencto = Format(mskVencimento.Text)
Else
    dDataVencto = Format(mskVenc.Text)
End If

If Year(dDataBase) > Year(dDataVencto) Then
    CalculaCorrecaoDAM = 0
    Exit Function
End If

UfirAtual = RetornaUFIR(Year(dDataVencto))
If UfirAtual = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano atual.", vbCritical, "Alerta !!!"
    CalculaCorrecaoDAM = 0
    Exit Function
End If

UfirBase = RetornaUFIR(Year(dDataBase))
If UfirBase = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano base.", vbCritical, "Alerta !!!"
    CalculaCorrecaoDAM = 0
    Exit Function
End If

CalculaCorrecaoDAM = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecaoDAM > 0 Then
   CalculaCorrecaoDAM = FormatNumber(CalculaCorrecaoDAM, 2)
End If
End Function

Private Function CalculaJurosDAM(nValorDebito As Double, dDataVencto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim sDataVencto As String, nDia As Integer, nMes As Integer, nAno As Integer

'If dDataNow = "00:00:00" Then
 dDataNow = Now
'End If

'SE O VENCIMENTO FOR MAIOR OU IGUAL A DATA ATUAL, NÃO EXISTE JUROS
If dDataVencto >= dDataNow Then
    CalculaJurosDAM = 0
    Exit Function
End If

'SE ESTIVER NO MESMO MES E ANO QUE A DATA ATUAL, NAO EXISTE JUROS
If Month(dDataVencto) = Month(dDataNow) And Year(dDataVencto) = Year(dDataNow) Then
    CalculaJurosDAM = 0
    Exit Function
End If

If Not dcJuros.Exists(Year(dDataNow)) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJurosDAM = 0
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
nNumMes = Int(DateDiff("d", dDataVencto, dDataNow) / 30) + 1


'If chkVenctoAtual.Value = 0 Then
'    dDataNow = Now
'Else
'    dDataNow = Format(mskVenc.text, "dd/mm/yyyy")
'End If

'If dDataVencto >= dDataNow Then
'    CalculaJurosDAM = 0
'    Exit Function
'End If
'nNumMes = Int((DateDiff("d", dDataVencto, dDataNow)) / 30)

If Not dcJuros.Exists(Year(dDataNow)) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJurosDAM = 0
   Exit Function
End If
nValorPerc = dcJuros.Item(Year(dDataNow))

nValorPerc = nValorPerc / 100

CalculaJurosDAM = nValorDebito * nValorPerc * nNumMes
If CalculaJurosDAM > 0 Then
   CalculaJurosDAM = FormatNumber(CalculaJurosDAM, 3)
End If

End Function

Private Sub mnuA1_Click()
nPlano = 16
mskVencimento.Text = "29/09/2017"
mskVencimento_LostFocus
End Sub

Private Sub mnuA2_Click()
nPlano = 17
mskVencimento.Text = "31/10/2017"
mskVencimento_LostFocus
End Sub
Private Sub mnuA3_Click()
mskVencimento.Text = "30/11/2017"
nPlano = 18
mskVencimento_LostFocus
End Sub
Private Sub mnuA4_Click()
nPlano = 19
mskVencimento.Text = "22/12/2017"
mskVencimento_LostFocus
End Sub

Private Sub mskVencimento_GotFocus()
mskVencimento.SelStart = 0
mskVencimento.SelLength = Len(mskVencimento.Text)
End Sub

Private Sub mskVencimento_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    mskVencimento_LostFocus
Else
    nPlano = 0
End If
End Sub

Private Sub mskVencimento_LostFocus()
Dim bValid As Boolean
bValid = False

If Not IsDate(mskVencimento.Text) Then
   MsgBox "Data inválida.", vbCritical, "Atenção"
   LimpaMascara mskVencimento
   bValid = False
Else
   If CDate(mskVencimento.Text) < Format(Now, "dd/mm/yyyy") Then
      MsgBox "Data de vencimento não pode ser retroativa.", vbCritical, "Atenção"
      LimpaMascara mskVencimento
      bValid = False
   Else
      bValid = True
   End If
End If

If bValid Then
    dDataVencto = Format(mskVencimento.Text)
    CarregaLista2
End If
End Sub

Private Sub EmiteBoleto()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nPos As Integer, sDataDam As String, sDataVencto As String
Dim nCodReduz As Long, sInsc As String, sNome As String, sDoc As String, sEnd As String, nNum As Integer, nValorDoc As Double
Dim sCompl As String, sBairro As String, sCidade As String, sUF As String, sQuadras As String, sLotes As String, sCep As String
Dim sUsuario As String, nNumDoc As Long, bMulta As Boolean, nValorTaxa As Double, sNumDoc As String, bGerado As Boolean
Dim sLanc As String, sFullTrib As String, nAno As Integer, nSeq As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double
Dim nSeq2 As Integer, sAj As String, sDA As String, nValorPrincipal As Double, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long
Dim nSid As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date, bAchou As Boolean
Dim sBarra As String, sDigitavel2 As String, nValorDam As Double, nValorPrincDam As Double, nNumGuia As Long, sTipoEnd As String


bAchou = False
With grdTemp
'    For x = 1 To .Rows - 1
'        If Val(.TextMatrix(x, 3)) = 0 Then
'            sVencto = .TextMatrix(x, 6)
'            bAchou = True
'            Exit For
'        End If
'    Next
End With
'If bAchou Then
'    If CDate(mskVencimento.Text) > CDate(sVencto) Then
'        MsgBox "Vencimento da única não pode ser prorrogado.", vbCritical, "Atenção"
'        Liberado
'        Exit Sub
'    End If
'End If


'nPlano = 0

If bRefisAtivoDI Then
    bAchou = False
    With grdTemp
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 0)) <> Year(Now) Then
                bAchou = True
                Exit For
            End If
        Next
    End With
    If bAchou Then
        With grdTemp
            For x = 1 To .Rows - 1
                If Val(.TextMatrix(x, 0)) = Year(Now) And Val(.TextMatrix(x, 1)) <> 4 And Val(.TextMatrix(x, 1)) <> 41 And Val(.TextMatrix(x, 1)) <> 69 Then
                    MsgBox "Não é permitido emitir débitos de " & Year(Now) & " junto com outros anos.", vbExclamation, "Atenção"
                    lblAnistia.Caption = "0,00"
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If lblDI.Caption = "S" Then
        bAchou = False
        With grdTemp
            For x = 1 To .Rows - 1
                If Year(.TextMatrix(x, 6)) >= Year(Now) Then
                    bAchou = True
                    Exit For
                End If
            Next
        End With
        If bAchou Then
            nPlano = 0
        Else
            nPlano = 23
        End If
    End If
End If


If bRefisAtivo And chkAnistia.value = vbChecked Then
'    If dVencto <= CDate("30/11/2016") Then
'       nPlano = 11
'    ElseIf dVencto <= CDate("20/12/2016") Then
'       nPlano = 12
'    Else
'        nPlano = 0
 '   End If
    
    For x = 1 To grdTemp.Rows - 1
        If Val(grdTemp.TextMatrix(x, 0)) = 2016 And CodLancamento <> 41 And CodLancamento <> 4 Then
            nPlano = 0
            Exit For
        End If
    Next
    
    If bITBI Then nPlano = 0
    bAchou = False
    With grdTemp
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 0)) <> 2017 Then
                bAchou = True
                Exit For
            End If
        Next
    End With
    If bAchou Then
        With grdTemp
            For x = 1 To .Rows - 1
                If Val(.TextMatrix(x, 0)) = 2017 And Val(.TextMatrix(x, 1)) <> 4 And Val(.TextMatrix(x, 1)) <> 41 And Val(.TextMatrix(x, 1)) <> 69 And Val(.TextMatrix(x, 1)) <> 67 And Val(.TextMatrix(x, 1)) <> 78 Then
                    MsgBox "Não é permitido emitir débitos de 2017 junto com outros anos." & vbCrLf & "Desmarque a opção REFIS.", vbExclamation, "Atenção"
                    lblAnistia.Caption = "0,00"
                     Exit Sub
                End If
            Next
        End With
    End If
End If


If MsgBox("Confirma Emissão da DAM ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boleto where sid=" & nSid
cn.Execute Sql, rdExecDirect

'RETORNA VALOR EXPEDIENTE
'Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE CODLANCAMENTO=3 AND ANOEXPED=" & Year(Now)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValorTaxa = 0
'RdoAux.Close

bMulta = False
sDoc = ""
nPos = 0
nValorDoc = 0
nValorDam = 0
nValorPrincDam = 0
nCodReduz = nCodigoDam
sUsuario = NomeDeLogin
sDataDam = mskVencimento.Text
sDoc = ""
Select Case nCodReduz
    
    Case 1 To 99999
        Sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = !Inscricao
            sNome = !nomecidadao
            sDoc = Format(SubNull(!CPF), "00000000000")
            If sDoc = "" Then
                sDoc = Format(SubNull(!Cnpj), "00000000000000")
            End If
            sEnd = SubNull(!Logradouro)
            nNum = Val(SubNull(!Li_Num))
            sCompl = Left(SubNull(!Li_Compl), 30)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!li_uf)
            sQuadras = Left(SubNull(!Li_Quadras), 15)
            sLotes = Left(SubNull(!Li_Lotes), 10)
            sCep = CStr(RetornaCEP(!CodLogr, !Li_Num))
           .Close
        End With
    Case 100000 To 300000
        Sql = "SELECT * FROM vwFULLEMPRESA3 WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = SubNull(!INSCESTADUAL)
            sNome = !RazaoSocial
            sDoc = SubNull(!CPF)
            If sDoc = "" Then
                sDoc = SubNull(!Cnpj)
            End If
            sEnd = !Logradouro
            nNum = !Numero
            sCompl = SubNull(!Complemento)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!SiglaUF)
            sQuadras = ""
            sLotes = ""
            If !CodCidade = 413 Then
                sCep = CStr(RetornaCEP(!CodLogradouro, !Numero))
            Else
                sCep = SubNull(!Cep)
            End If
         End With
     Case 500000 To 800000
        sTipoEnd = "R"
        Sql = "select * from cidadao where codcidadao=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If SubNull(RdoAux2!etiqueta) = "N" And SubNull(RdoAux2!etiqueta2) = "S" Then
                sTipoEnd = "C"
            End If
            RdoAux2.Close
        End If
        
        If sTipoEnd = "R" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
            Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
            Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
            Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        Else
            Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
            Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
            Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
            Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        On Error Resume Next
        With RdoAux2
            If .RowCount > 0 Then
                 sCodReduz = !CodCidadao
                 sNome = !nomecidadao
                 If Val(SubNull(!FCodLogradouro)) > 0 Then
                     Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                     Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                     Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                     Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                     With RdoS
                         If .RowCount > 0 Then
                            sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                         Else
                            sEnd = ""
                         End If
                        .Close
                     End With
                 Else
                    sEnd = SubNull(!FNomeLogradouro)
                 End If
                 nNum = Val(SubNull(RdoAux2!fNUMIMOVEL))
                  
                 Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                 Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                 If RdoS.RowCount > 0 Then
                     sCidade = RdoS!descCidade
                 Else
                      sCidade = ""
                 End If
                 If Not IsNull(!CodBairro) Then
                     Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                     If .RowCount > 0 Then
                         sBairro = RdoS!DescBairro
                     Else
                         sBairro = ""
                     End If
                 Else
                     sBairro = ""
                 End If
                 sUF = SubNull(!fsiglauf)
                 sCompl = SubNull(!fcomplemento)
            Else
                sEnd = ""
                sBairro = ""
                sCidade = ""
                sUF = ""
                sCompl = ""
            End If
            sDoc = SubNull(!CPF)
            If sDoc = "" Then
                sDoc = SubNull(!Cnpj)
            End If
            If !fCodCidade = 413 Then
                sCep = CStr(RetornaCEP(!FCodLogradouro, !fNUMIMOVEL))
            Else
                sCep = SubNull(!FCEP)
            End If
           .Close
        End With
     
End Select

If sCep = "" Then
    MsgBox "Contribuinte não possui CEP válido!", vbCritical, "ERRO"
    Exit Sub
End If


If sDoc = "" Then
    MsgBox "Contribuinte não possui CPF/CNPJ válido!", vbCritical, "ERRO"
    Exit Sub
End If


If chkMulta.value = 1 Then
   bMulta = True
Else
   If Val(lblAnistia.Caption) > 0 Then
       bMulta = True
   Else
       bMulta = False
   End If
End If

With grdTemp
   'GRAVA NUMDOCUMENTO
    If chkMulta.value = 1 Then
       bMulta = True
    Else
       If Val(lblAnistia.Caption) > 0 Then
           bMulta = True
       Else
           bMulta = False
       End If
    End If
    
    If Not bHonorario Then
        'grava documento
        Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!maximo) Then
           nNumDoc = 0
        Else
           nNumDoc = RdoAux!maximo + 1
        End If
        RdoAux.Close
       
        If bComercioEletronico Then
            Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO,TIPODOC,emissor,registrado) VALUES("
            If chkJulgamento.value = 0 Then
                'Sql = Sql & nNumDoc & ",'" & Format(sDataDam, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ",1,'" & NomeDeLogin & " (DAM.REG)" & "',1)"
                Sql = Sql & nNumDoc & ",'" & Format(Now, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ",1,'" & NomeDeLogin & " (DAM.REG)" & "',1)"
            Else
                Sql = Sql & nNumDoc & ",'" & Format(Now, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ",1,'" & NomeDeLogin & " (DAM.REG)" & "',1)"
            End If
        Else
            Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO,TIPODOC,emissor,registrado) VALUES("
            If chkJulgamento.value = 0 Then
                Sql = Sql & nNumDoc & ",'" & Format(Now, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ",1,'" & NomeDeLogin & " (DAM)" & "',0)"
            Else
                Sql = Sql & nNumDoc & ",'" & Format(Now, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ",1,'" & NomeDeLogin & " (DAM)" & "',0)"
            End If
        End If
        cn.Execute Sql, rdExecDirect
    Else
        nNumDoc = 0
    End If
    
    '*******
 '   nNumDoc = 14311313
    '*******
    
 '   If NomeDeLogin = "SCHWARTZ" Then
'        nNumDoc = 15813326
  '  End If
    
    sNumDoc = CStr(nNumDoc)
    'sNumDoc = CStr(nNumDoc) & "-" & RetornaDVNumDoc(nNumDoc)
    sNumDoc2 = CStr(nNumDoc)
    'sNumDoc2 = CStr(nNumDoc) & RetornaDVNumDoc(nNumDoc)
    'sNumDoc3 = CStr(nNumDoc) & Modulo11(nNumDoc)
    
    For x = 1 To .Rows - 1
        nAno = Val(.TextMatrix(x, 0))
        nLanc = Val(.TextMatrix(x, 1))
        nSeq = Val(.TextMatrix(x, 2))
        nParc = Val(.TextMatrix(x, 3))
        nCompl = Val(.TextMatrix(x, 4))
        sDataVencto = .TextMatrix(x, 6)
        sDA = .TextMatrix(x, 7)
        sAj = .TextMatrix(x, 8)
        nValorPrincipal = FormatNumber(CDbl(.TextMatrix(x, 9)), 2)
        nValorPrincDam = nValorPrincDam + nValorPrincipal
        nValorJuros = FormatNumber(CDbl(grdTemp.TextMatrix(x, 12)), 2)
        nValorMulta = FormatNumber(CDbl(.TextMatrix(x, 11)), 2)
        nValorCorrecao = FormatNumber(CDbl(.TextMatrix(x, 10)), 2)
        nValorTotal = FormatNumber(CDbl(.TextMatrix(x, 13)), 2)
        nValorDoc = nValorDoc + nValorTotal
        sFullTrib = Left$(Mask(grdTrib.TextMatrix(x, 1)), 2000)
        sFullTrib = Left(sFullTrib, Len(sFullTrib))
       'GRAVA PARCELADOCUMENTO
        If Not bHonorario Then
            Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,VALORJUROS,VALORMULTA,VALORCORRECAO,PLANO) VALUES(" & nCodReduz & ","
            Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumDoc & ","
            Sql = Sql & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & "," & Virg2Ponto(CStr(nValorCorrecao)) & "," & nPlano & ")"
            cn.Execute Sql, rdExecDirect
        End If
        If Val(lblAnistia.Caption) > 0 And bAnistia Then
            'GRAVA OBS PARCELA
             Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
             Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
             Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
             Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux
                 If IsNull(!maximo) Then
                     nSeq2 = 1
                 Else
                     nSeq2 = !maximo + 1
                 End If
                .Close
             End With
             sObs = "Lancamento incluido na DAM número " & nLastCod + 1 & " com " & lblAnistia.Caption & "% de desconto em multa e juros conforme REFIS-2016"
'             Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & nCodReduz & "," & nAno & ","
'             Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nSeq2 & ",'" & sObs & "','" & NomeDeLogin & "','" & Format(Now, sDataFormat) & "')"
             Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduz & "," & nAno & ","
             Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nSeq2 & ",'" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, sDataFormat) & "')"
             cn.Execute Sql, rdExecDirect
        End If
            
'        If Not bComercioEletronico Then
            Sql = "insert boleto(usuario,computer,sid,seq,inscricao,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,quadra,lote,numdoc,nomefunc,datadam,fulllanc,fulltrib,"
            Sql = Sql & "anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datavencto,aj,da,principal,juros,multa,correcao,total,numdoc2,valordam) values('"
            Sql = Sql & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nPos & ",'" & sInsc & "'," & nCodReduz & ",'" & Left(Mask(sNome), 40) & "','" & sDoc & "','"
            Sql = Sql & Left(Mask(sEnd), 40) & "'," & nNum & ",'" & Left(Mask(sCompl), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Mask(sCidade) & "','" & sUF & "','" & Mask(sQuadras) & "','"
            Sql = Sql & Mask(sLotes) & "','" & sNumDoc & "','" & NomeDeLogin & "','" & Format(sDataDam, sDataFormat) & "','" & sLANCAMENTO & "','" & sFullTrib & "'," & nAno & ","
            Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",'" & Format(sDataVencto, sDataFormat) & "','" & sAj & "','" & sDA & "'," & Virg2Ponto(Format(nValorPrincipal, "#0.00")) & ","
            Sql = Sql & Virg2Ponto(Format(nValorJuros, "#0.00")) & "," & Virg2Ponto(Format(nValorMulta, "#0.00")) & "," & Virg2Ponto(Format(nValorCorrecao, "#0.00")) & "," & Virg2Ponto(Format(nValorTotal, "#0.00")) & ",'" & sNumDoc2
            Sql = Sql & "'," & Virg2Ponto(RemovePonto(lblValorTotal.Caption)) & ")"
            cn.Execute Sql, rdExecDirect
 '       End If
        nPos = nPos + 1
    
    Next

End With

'CRIA VINCULO COM ISSELETRONICO SE HOUVER
If bAnistia Then
    If Not bHonorario Then
        If UBound(aDocDAM) > 0 Then
            For x = 1 To UBound(aDocDAM)
                Sql = "INSERT DAMISS(DOCDAM,DOCISS) VALUES(" & nNumDoc & "," & aDocDAM(x) & ")"
                cn.Execute Sql, rdExecDirect
            Next
            ReDim aDocDAM(0)
        End If
    End If
End If


'*****'
'nValorDoc = 112.17
'sDataDam = "30/01/2015"
'******'


'bBoleto = True
'If NomeDeLogin = "SCHWARTZ" Then
    bBoleto = False
'End If

Sql = "update numdocumento set valorguia=" & Virg2Ponto(CStr(nValorDoc)) & " where numdocumento=" & nNumDoc
cn.Execute Sql, rdExecDirect

If bComercioEletronico Then
    GoTo ComercioEletronico
End If

'**** GERADOR DE CÓDIGO DE BARRAS ********
If chkRegistrado.value = vbChecked Then
    sNossoNumero = "2873532"
'    sDigitavel = "001900000"
'    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'    sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
'    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'    sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
'    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'    sDigitavel = sDigitavel & sDv
    
    dDataBase = "07/10/1997"
    nFatorVencto = CDate(sDataDam) - dDataBase
    sQuintoGrupo = Format(nFatorVencto, "0000")
    sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
    sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "000000287353200"
    sBarra = sBarra & sNumDoc & "17"
'    sDv = Val(Calculo_DV11(sBarra))
'    sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
    Dim sCampo1 As String, sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String
    
    sCampo1 = "0019" & Mid(sBarra, 20, 5)
    sDigitavel = sCampo1 & Val(Calculo_DV10(sCampo1))
    sCampo2 = Mid(sBarra, 24, 10)
    sDigitavel = sDigitavel & sCampo2 & Val(Calculo_DV10(sCampo2))
    sCampo3 = Mid(sBarra, 34, 10)
    sDigitavel = sDigitavel & sCampo3 & Val(Calculo_DV10(sCampo3))
    sCampo5 = Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
    
    'sCampo4 = Val(Calculo_DV11(sDigitavel & sCampo5))
    sCampo4 = Val(Calculo_DV11(sBarra))
    sDigitavel = sDigitavel & sCampo4 & sCampo5
    sBarra = Left(sBarra, 4) & sCampo4 & Mid(sBarra, 5, Len(sBarra) - 4)
    'sDigitavel = sDigitavel & sDv & sQuintoGrupo
    
    sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
    sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
    
    sBarra = Gera2of5Str(sBarra)
    Sql = "update boleto set digitavel='" & sDigitavel2 & "',codbarra='" & Mask(sBarra) & "',valorprincdam=" & Virg2Ponto(RemovePonto(Format(nValorPrincDam, "#0.00"))) & " where sid=" & nSid
    cn.Execute Sql, rdExecDirect
    nNumGuia = nNumDoc
    frmReport.ShowReport2 "BOLETODAM_V3", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
    GeraArquivo 0, nNumDoc, sDataDam
Else
    Dim sValor As String, dDataVencto As Date, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
    sValor = nValorDoc
    dDataVencto = CDate(sDataDam)
   ' nNumDoc = Val(sNumDoc2)
    nNumGuia = nNumDoc
    NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumDoc, nCodReduz)
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)

    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
    Sql = "update boleto set numbarra2a='" & NumBarra2a & "',numbarra2b='" & NumBarra2b & "',numbarra2c='" & NumBarra2c & "',numbarra2d='" & NumBarra2d & "',codbarra='" & Mask(StrBarra2) & "' where sid=" & nSid
    cn.Execute Sql, rdExecDirect
    
    nNumGuia = nNumDoc
    If frmMdi.frTeste.Visible = False Then
        frmReport.ShowReport2 "BOLETODAM_V4", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
    Else
        frmReport.ShowReport2 "BOLETODAM_v4TMP", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
    End If
End If


Sql = "delete from boleto where sid=" & nSid
cn.Execute Sql, rdExecDirect

Exit Sub


ComercioEletronico:
frmReport.ShowReport2 "BOLETODAM_V5", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
Sql = "delete from boleto where sid=" & nSid
cn.Execute Sql, rdExecDirect

'frmComercioEletronico.BoletoUser = NomeDeLogin & "-Dam"
'frmComercioEletronico.BoletoNome = sNome
'frmComercioEletronico.BoletoCidade = Left(sCidade, 18)
'frmComercioEletronico.BoletoCep = sCep
'frmComercioEletronico.BoletoCpfCnpj = sDoc
'frmComercioEletronico.BoletoEndereco = Left(sEnd & ", " & nNum & IIf(sCompl <> "", " " & sCompl, "") & " - " & sBairro, 60)
'frmComercioEletronico.BoletoNumDoc = nNumDoc
'frmComercioEletronico.BoletoUF = sUF
'frmComercioEletronico.BoletoValor = nValorDoc
'frmComercioEletronico.BoletoVencto = sDataVenctoDAM
'frmComercioEletronico.show 1
'Exit Sub
Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String, v8 As String, v9 As String, V10 As String
v1 = sNome
v2 = Left(sEnd & ", " & nNum & IIf(sCompl <> "", " " & sCompl, "") & " - " & sBairro, 60)
v3 = sDataDam
v4 = sDoc
v5 = "287353200" & Format(nNumDoc, "00000000")
Dim sValorDoc As String
sValorDoc = FormatNumber(nValorDoc, 2)
sValorDoc = RetornaNumero(sValorDoc)
v6 = sValorDoc
v7 = Left(sCidade, 18)
v8 = sUF
v9 = sCep
V10 = NomeDeLogin & "-Dam"
If Trim(sCep) = "" Or Trim(sCep) = "-" Then
    v9 = "14870-000"
End If
ShellExecute hwnd, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Pages/boletoBB.aspx?f1=" & v1 & "&f2=" & v2 & "&f3=" & v3 & "&f4=" & v4 & "&f5=" & v5 & "&f6=" & v6 & "&f7=" & v7 & "&f8=" & v8 & "&f9=" & v9 & "&f10=" & V10, vbNullString, vbNullString, conSwNormal
Unload Me

End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
Tweak txtDesconto, KeyAscii, IntegerPositive
End Sub

Private Sub txtDesconto_LostFocus()
If Val(txtDesconto.Text) = 0 Then txtDesconto.Text = "0"
If Val(txtDesconto.Text) > 100 Then txtDesconto.Text = "100"
End Sub

Private Sub GeraArquivo(nTipo As Integer, nNumDoc As Long, sDataVencto As String)

Dim RdoAux As rdoResultset, Sql As String, nPosReg As Long, nContador As Long, nInicio As Integer
Dim aHeaderArquivo() As HeaderArquivo, FF1 As Integer, sHeaderArquivo As String
Dim aTrailerArquivo() As TrailerArquivo, sTrailerArquivo As String, nQtdeRegistroArquivo As Long, nQtdeRegistroLote As Long, aHeaderLote() As HeaderLote, sHeaderLote As String
Dim aTrailerLote() As TrailerLote, sTrailerLote As String, aSegmentoP() As SegmentoP, sSegmentoP As String, aSegmentoQ() As SegmentoQ, sSegmentoQ As String

sArquivo = "C:\tmp\remessaiptu.txt"
FF1 = FreeFile()
Open sArquivo For Output As FF1

'*********************************
'******** Header Arquivo *********
'*********************************
ReDim aHeaderArquivo(1)
With aHeaderArquivo(1)
    .nCodigoBanco = "001"
    .nLote = "0000"
    .nTipoRegistro = "0"
    .sUsoFebraban1 = FillSpace(" ", 9)
    .nTipoInscricao = "2"
    .nNumInscricao = "50387844000105"
    If nTipo = 0 Then
        .sCodigoConvenio = FillLeft("2873532", 9) & "001417019  " 'IPTU/ISS/TXLIC
        .nNumeroConta = FillLeft("74000", 12)
        .sDvConta = "4 "
    Else
        .sCodigoConvenio = FillLeft("2950230", 9) & "001417019  " 'CIP
        .nNumeroConta = FillLeft("34692", 12)
        .sDvConta = "6 "
    End If
    .nAgencia = "00269"
    .sDvAgencia = "0"
    .sNomeEmpresa = FillSpace("PREFEITURA MUN. DE JABOTICABAL", 30)
    .sNomeBanco = FillSpace("BANCO DO BRASIL S.A.", 30)
    .sUsoFebraban2 = FillSpace(" ", 10)
    .nCodigoRemessa = "1"
    '.nDataGeracao = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
    .nDataGeracao = "10012017"
    .nHoraGeracao = Format(Now, "hhmmss")
    .nNumSeqArquivo = "000000"
    .nNumVersaoLayout = "000"
    .nDensidade = "00000"
    .sUsoBanco = FillSpace(" ", 20)
    .sUsoEmpresa = FillSpace(" ", 20)
    .sUsoFebraban3 = FillSpace(" ", 29)
    
    sHeaderArquivo = .nCodigoBanco & .nLote & .nTipoRegistro & .sUsoFebraban1 & .nTipoInscricao & .nNumInscricao & .sCodigoConvenio & .nAgencia & .sDvAgencia & .nNumeroConta & .sDvConta & .sNomeEmpresa
    sHeaderArquivo = sHeaderArquivo & .sNomeBanco & .sUsoFebraban2 & .nCodigoRemessa & .nDataGeracao & .nHoraGeracao & .nNumSeqArquivo & .nNumVersaoLayout & .nDensidade & .sUsoBanco & .sUsoEmpresa & .sUsoFebraban3
End With

Print #FF1, sHeaderArquivo
'****** Fim Header Arquivo ************


'*********************************
'******** Header do Lote *********
'*********************************
ReDim aHeaderLote(1)
With aHeaderLote(1)
    .nCodigoBanco = "001"
    .nLote = "0001"
    .nTipoRegistro = "1"
    .sTipoOperacao = "R"
    .nTipoServico = "01"
    .sUsoFebraban1 = "  "
    .nNumVersao = "000"
    .sUsoFebraban2 = " "
    .nTipoInscricao = "2"
    .nNumInscricao = "050387844000105"
    If nTipo = 0 Then
        .sCodConvenio = FillLeft("2873532", 9) & "001417019  " 'IPTU/ISS/TXLIC
        .nNumeroConta = FillLeft("74000", 12)
        .sDvConta = "4 "
    Else
        .sCodConvenio = FillLeft("2950230", 9) & "001417019  " 'CIP
        .nNumeroConta = FillLeft("34692", 12)
        .sDvConta = "6 "
    End If
    .nAgencia = "00269"
    .sDvAgencia = "0"
    .sNomeEmpresa = FillSpace("PREFEITURA MUN. DE JABOTICABAL", 30)
    .sMensagem1 = FillSpace(" ", 40)
    .sMensagem2 = FillSpace(" ", 40)
    .nNumeroRemessa = FillLeft(CStr(nNumRemessa), 8)
    '.sDataGeracao = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
    .sDataGeracao = "10012017"
    .sDataCredito = "00000000"
    .sUsoFebraban3 = FillSpace(" ", 33)
    
    sHeaderLote = .nCodigoBanco & .nLote & .nTipoRegistro & .sTipoOperacao & .nTipoServico & .sUsoFebraban1 & .nNumVersao & .sUsoFebraban2 & .nTipoInscricao & .nNumInscricao
    sHeaderLote = sHeaderLote & .sCodConvenio & .nAgencia & .sDvAgencia & .nNumeroConta & .sDvConta & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & .nNumeroRemessa & .sDataGeracao & .sDataCredito & .sUsoFebraban3
End With

Print #FF1, sHeaderLote
'****** Fim Header do Lote********


'*********************************
'******** Segmento P e Q *********
'*********************************

nContador = 1
'nInicio = 49999
For nPosReg = 1 To 1
    ReDim aSegmentoP(1): ReDim aSegmentoQ(1)
    With aSegmentoP(1)
    
        .nCodigoBanco = "001"
        .nLote = "0001"
        .nTipo = "3"
        .nSeqReg = FillLeft(CStr(nContador), 5)
        .sCodSegmento = "P"
        .sUsoFebraban1 = " "
        .nCodMovimento = "01"
        .nAgencia = "00269"
        .sDvAgencia = "0"
        If nTipo = 0 Or nTipo = 1 Or nTipo = 2 Then
            .nConta = FillLeft("74000", 12)
            .sDvConta = "4 "
        Else
            .nConta = FillLeft("34692", 12)
            .sDvConta = "6 "
        End If
        .sNossoNumero = FillSpace("287353200" & CStr(nNumDoc), 20)
        
        
        .nCodCarteira = "7"
        .nFormaCadastro = "1"
        .sTipoDocumento = "1"
        .nIdentificacaoEmissao = "2"
        .sIdentificacaoDistribuicao = "2"
        .sNumeroDocumento = FillSpace(CStr(nNumDoc), 15)
        .nDataVencimento = Format(RetornaNumero(sDataVencto), "000000")
        .nValorNominal = FillLeft(RetornaNumero(CStr(lblValorTotal.Caption * 100)), 15)
        .nAgenciaCobranca = "00000"
        .sDvAgenciaCobranca = "0"
        .nEspecieTitulo = "01"
        .sAceite = "N"
        .nDataEmissao = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
        .nCodigoJuros = "0"
        .nDataJuros = FillLeft("0", 8)
        .nJurosMora = FillLeft("0", 15)
        .nCodigoDesconto1 = "0"
        .nDataDesconto1 = FillLeft("0", 8)
        .nValorConcedido = FillLeft("0", 15)
        .nValorIOF = FillLeft("0", 15)
        .nValorAbatimento = FillLeft("0", 15)
        .sIdentificaTitulo = FillSpace(CStr(nNumDoc), 25)
        .nCodigoProtesto = "3"
        .nNumDiasProtesto = "00"
        .nCodigoBaixa = "0"
        .sNumDiasBaixa = "000"
        .nCodigoMoeda = "09"
        .nNumeroContrato = FillLeft("19663033", 10)
        .sUsoLivre = " "
        
        sSegmentoP = .nCodigoBanco & .nLote & .nTipo & .nSeqReg & .sCodSegmento & .sUsoFebraban1 & .nCodMovimento & .nAgencia & .sDvAgencia & .nConta & .sDvConta & .sNossoNumero
        sSegmentoP = sSegmentoP & .nCodCarteira & .nFormaCadastro & .sTipoDocumento & .nIdentificacaoEmissao & .sIdentificacaoDistribuicao & .sNumeroDocumento & .nDataVencimento
        sSegmentoP = sSegmentoP & .nValorNominal & .nAgenciaCobranca & .sDvAgenciaCobranca & .nEspecieTitulo & .sAceite & .nDataEmissao & .nCodigoJuros & .nDataJuros & .nJurosMora
        sSegmentoP = sSegmentoP & .nCodigoDesconto1 & .nDataDesconto1 & .nValorConcedido & .nValorIOF & .nValorAbatimento & .sIdentificaTitulo & .nCodigoProtesto & .nNumDiasProtesto & .nCodigoBaixa
        sSegmentoP = sSegmentoP & .sNumDiasBaixa & .nCodigoMoeda & .nNumeroContrato & .sUsoLivre
    End With
    
    nContador = nContador + 1
    With aSegmentoQ(1)
        .nCodigoBanco = "001"
        .nLote = "0001"
        .nTipo = "3"
        .nSeqReg = FillLeft(CStr(nContador), 5)
        .sCodSegmento = "Q"
        .sUsoFebraban1 = " "
        .nCodMovimento = "01"
        .nTipoInscricao = 2
        .nNumeroInscricao = FillLeft("01922044000131", 15)
        .sNome = FillSpace(frmDebitoImob.lblProp.Caption, 40)
        .sEndereco = FillSpace("Av. Major Novaes 519", 40)
        .sBairro = FillSpace("Centro", 15)
        'If Len(aBoletos(nPosReg).sCep) < 5 Then
            .nCep = "14870"
            .nCepsufixo = "080"
        'Else
        '    .nCep = ""
        '    .nCepsufixo = ""
        'End If
        
        .sCidade = FillSpace("Jaboticabal", 15)
        .sUF = "SP"
        .nipoInscricaoSacado = "2"
        .nNumeroInscricaoSacado = FillLeft("01922044000131", 15)
        .sNomeSacado = FillSpace(frmDebitoImob.lblProp.Caption, 40)
        .nBancoCorresponde = "000"
        .sNossoNumeroBancoCorr = FillSpace(" ", 20)
        .sUsoFebraban2 = FillSpace(" ", 8)
        
        sSegmentoQ = .nCodigoBanco & .nLote & .nTipo & .nSeqReg & .sCodSegmento & .sUsoFebraban1 & .nCodMovimento & .nTipoInscricao & .nNumeroInscricao & .sNome & .sEndereco & .sBairro
        sSegmentoQ = sSegmentoQ & .nCep & .nCepsufixo & .sCidade & .sUF & .nipoInscricaoSacado & .nNumeroInscricaoSacado & .sNomeSacado & .nBancoCorresponde & .sNossoNumeroBancoCorr & .sUsoFebraban2
    End With
    nContador = nContador + 1
    Print #FF1, sSegmentoP
    Print #FF1, sSegmentoQ
    DoEvents
'    If nContador > 99995 Then
'        Exit For
'    End If
Next

'****** Fim do Segmento P *******


'***************************************
'********** Trailer do Lote ************
'***************************************
nQtdeRegistroLote = 3
ReDim aTrailerLote(1)
With aTrailerLote(1)
    .nCodigoBanco = "001"
    .nLote = "0001"
    .nTipo = "5"
    .sUsoFebraban1 = FillSpace(" ", 9)
    .nQtdeRegistro = FillLeft(CStr(nQtdeRegistroLote), 6)
    .sUsoFebraban2 = FillSpace(" ", 217)
    
    sTrailerLote = .nCodigoBanco & .nLote & .nTipo & .sUsoFebraban1 & .nQtdeRegistro & .sUsoFebraban2
End With

Print #FF1, sTrailerLote
'****** Fim Trailer do Lote ************


'***************************************
'********** Trailer Arquivo ************
'***************************************
nQtdeRegistroArquivo = nQtdeRegistroLote + 2
ReDim aTrailerArquivo(1)
With aTrailerArquivo(1)
    .nCodigoBanco = "001"
    .nLote = "9999"
    .nTipo = "9"
    .sUsoFebraban1 = FillSpace(" ", 9)
    .nQtdeLote = FillLeft("1", 6)
    .nQtdeRegistro = FillLeft(CStr(nQtdeRegistroArquivo), 6)
    .nQtdeContas = FillLeft("0", 6)
    .sUsoFebraban2 = FillSpace(" ", 205)
    
    sTrailerArquivo = .nCodigoBanco & .nLote & .nTipo & .sUsoFebraban1 & .nQtdeLote & .nQtdeRegistro & .nQtdeContas & .sUsoFebraban2
End With

Print #FF1, sTrailerArquivo
'****** Fim Trailer Arquivo ************


'**********************
'****** Finaliza ******
'**********************

Close #FF1
'PBar.value = 0
Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='COBRANCA'"
cn.Execute Sql, rdExecDirect

ret = Shell("C:\Program Files\PSPad editor\pspad.exe" & " " & sArquivo, vbNormalFocus)

'AtualizaRemessa

End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Format(sTexto, String(nTamanho, "0"))

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

