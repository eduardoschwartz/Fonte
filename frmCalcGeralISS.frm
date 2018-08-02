VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCalcGeralISS 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo Geral de ISS"
   ClientHeight    =   2805
   ClientLeft      =   2325
   ClientTop       =   3765
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   6540
   Begin esMaskEdit.esMaskedEdit esMaskedEdit1 
      Height          =   30
      Left            =   1590
      TabIndex        =   28
      Top             =   300
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      MouseIcon       =   "frmCalcGeralISS.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelText         =   ""
      HideSelection   =   -1  'True
   End
   Begin VB.Timer Timer1 
      Left            =   5550
      Top             =   2160
   End
   Begin MSComctlLib.ProgressBar PbF 
      Height          =   240
      Left            =   2175
      TabIndex        =   8
      Top             =   570
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5355
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   1800
      Width           =   1125
      _ExtentX        =   1984
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCalcGeralISS.frx":001C
      PICN            =   "frmCalcGeralISS.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdHelp 
      Height          =   315
      Left            =   5355
      TabIndex        =   3
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   1410
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Ajuda"
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
      MICON           =   "frmCalcGeralISS.frx":00A6
      PICN            =   "frmCalcGeralISS.frx":00C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   315
      Left            =   5355
      TabIndex        =   2
      ToolTipText     =   "Cancelar Edição"
      Top             =   1020
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmCalcGeralISS.frx":021C
      PICN            =   "frmCalcGeralISS.frx":0238
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCalculo 
      Height          =   315
      Left            =   5355
      TabIndex        =   1
      ToolTipText     =   "Cancelar Edição"
      Top             =   630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Calcular"
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
      MICON           =   "frmCalcGeralISS.frx":0392
      PICN            =   "frmCalcGeralISS.frx":03AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar PbV 
      Height          =   240
      Left            =   2175
      TabIndex        =   9
      Top             =   900
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PbE 
      Height          =   240
      Left            =   2175
      TabIndex        =   10
      Top             =   1245
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PbL 
      Height          =   240
      Left            =   2175
      TabIndex        =   17
      Top             =   1575
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PbS 
      Height          =   240
      Left            =   2175
      TabIndex        =   18
      Top             =   1920
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin esMaskEdit.esMaskedEdit mskDataBase 
      Height          =   285
      Left            =   3555
      TabIndex        =   27
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmCalcGeralISS.frx":044D
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
   Begin prjChameleon.chameleonButton cmdVS 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   270
      TabIndex        =   30
      ToolTipText     =   "Cancelar Edição"
      Top             =   3150
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Atualizar Vig.Sanitaria"
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
      MICON           =   "frmCalcGeralISS.frx":0469
      PICN            =   "frmCalcGeralISS.frx":0485
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblWait 
      BackStyle       =   0  'Transparent
      Caption         =   "CÁLCULO EM ANDAMENTO AGUARDE........."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   180
      TabIndex        =   29
      Top             =   2370
      Width           =   6165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Base...:"
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   1
      Left            =   2520
      TabIndex        =   0
      Top             =   165
      Width           =   1065
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      Caption         =   "2011"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   1455
      TabIndex        =   26
      Top             =   135
      Width           =   585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano de Cálculo:"
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   135
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxa de Licença.:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   1575
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vigil.Sanitária.......:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   1320
   End
   Begin VB.Label lblPL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1590
      TabIndex        =   22
      Top             =   1590
      Width           =   390
   End
   Begin VB.Label lblPS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1590
      TabIndex        =   21
      Top             =   1935
      Width           =   390
   End
   Begin VB.Label lblTotL 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   4545
      TabIndex        =   20
      Top             =   1590
      Width           =   720
   End
   Begin VB.Label lblTotS 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   4545
      TabIndex        =   19
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lblTotE 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   4545
      TabIndex        =   16
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label lblTotV 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   4545
      TabIndex        =   15
      Top             =   915
      Width           =   720
   End
   Begin VB.Label lblTotF 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   4545
      TabIndex        =   14
      Top             =   585
      Width           =   720
   End
   Begin VB.Label lblPE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1590
      TabIndex        =   13
      Top             =   1260
      Width           =   390
   End
   Begin VB.Label lblPV 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1590
      TabIndex        =   12
      Top             =   915
      Width           =   390
   End
   Begin VB.Label lblPF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1590
      TabIndex        =   11
      Top             =   585
      Width           =   390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ISS Estimado......:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1245
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ISS Variável........:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   900
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ISS Fixo..............:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   570
      Width           =   1320
   End
End
Attribute VB_Name = "frmCalcGeralISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const COD_LANCISSFIXO = 6 '--> ALTERADO DE 2 PARA 6 POR CAUSA DA SEPARAÇÃO DOS LANCAMENTOS ISS FIXO/TLL
Private Const COD_LANCISSESTIMADO = 3
Private Const COD_LANCISSVARIAVEL = 5
Private Const COD_LANCTAXALICENCA = 6
Private Const COD_LANCVIGSANITARIA = 13
Private Const COD_TRIBISSFIXO = 11
Private Const COD_TRIBISSESTIMADO = 12
Private Const COD_TRIBISSVARIAVEL = 13
Private Const COD_TRIBTAXALICENCA = 14
Private Const COD_TRIBVIGSANITARIA = 25
Private Const COD_TRIBALVARA = 23
Private Const COD_TRIBVISTORIA = 24
Private Const COD_TRIBprotocolo = 28
Private Const COD_TRIBTXEXPDOC = 3 'AUTENTICAÇÃO BANCÁRIA
Private Const COD_TRIBTXEXPED = 10 'EXPEDIENTE FUNCIONÁRIO
Private Const STATUS_NAOPAGO = 3
Private Const MOEDA_REAL = 1

Private Type VS
    nDivisao As Integer
    nGrupo As Integer
    nClasse As Integer
    nSubClasse As Integer
    nCriterio As Integer
    nValor As Double
End Type

Private Type TAXALICENCA
    nValorAliq As Double
    nArea As Double
End Type

Dim aValorAliquotaTxL() As TAXALICENCA

'POSIÇÕES DOS CURSORES
Dim nPosF As Long, nTotalF As Long
Dim nPosE As Long, nTotalE As Long
Dim nPosV As Long, nTotalV As Long
Dim nPosL As Long, nTotalL As Long
Dim nPosS As Long, nTotalS As Long

Private Type Parcela
    CODREDUZIDO         As String
    AnoExercicio        As String
    CodLancamento       As String
    SeqLancamento       As String
    NumParcela          As String
    CODCOMPLEMENTO      As String
End Type
Private Type TRIBUTO
    sCodReduz As String
    sAno As String
    sCodLanc As String
    sSeq As String
    sNumParc As String
    sCompl As String
    sCodTrib As String
End Type


Dim Sql As String, RdoAux As rdoResultset, RdoEmp As rdoResultset, RdoAux2 As rdoResultset
Dim RdoVig As rdoResultset, RdoAliq As rdoResultset
Private Sub cmdCalculo_Click()

On Error GoTo Erro
Dim nSeqLanc As Integer, bTemTL As Boolean

'GoTo paramparc
'MsgBox "EM MANUTENÇÃO"
'Exit Sub

'PARAMETROS DAS PARCELAS
Dim aParcF() As Date, bUnicaF As Boolean, nDescUnicaF As Double
Dim aParcV() As Date, bUnicaV As Boolean, nDescUnicaV As Double
Dim aParcS() As Date, bUnicaS As Boolean, nDescUnicaS As Double
'PARAMETROS PARA TRIBUTOS
Dim nValorAlvara As Double
Dim nValorVistoria As Double
Dim nValorprotocolo As Double
Dim nValorExpediente As Double
'PARAMETROS TAXA EXPEDIÇÃO DE DOCUMENTO
Dim nValorExpDocParcF As Double, nValorExpDocUnicaF As Double
Dim nValorExpDocParcE As Double, nValorExpDocUnicaE As Double
Dim nValorExpDocParcV As Double, nValorExpDocUnicaV As Double
'OUTRAS VARIAVEIS
Dim x As Integer
Dim nCodEmpresa As Long
Dim nCodAtividade As Long
Dim nArea As Double
Dim nQtdeProfTL As Integer, nQtdeProfISS As Integer, nQtdeProfVS As Integer
Dim nValorAliquotaISS As Double
Dim nValorEstimado As Double, nValorAliquotaVS As Double
Dim bVistoria As Boolean
Dim nUfirAtual As Double
Dim sDataBase As String
Dim nAnoExercicio As Integer
Dim nNumParcela As Integer
Dim nValorTotal As Double, nValorParcela As Double, nValorUnica As Double
Dim nValorTxLic As Double
Dim nLastDoc As Long, bAchou As Boolean
Dim nTipoIss As Integer, aLancAno() As Long, nTeste As Integer, bTaxaLic As Boolean
ReDim aLancAno(0)
Dim Rdo99 As rdoResultset

' VALIDAÇÃO DATA BASE
If Not IsDate(mskDataBase.Text) Then
   MsgBox "Data Base inválida.", vbCritical, "Atenção"
   mskDataBase.SetFocus
   Exit Sub
End If

If frmMdi.frTeste.Visible = False Then
    MsgBox "Calculo geral apenas para base de testes."
'    Exit Sub
End If

If Year(CDate(mskDataBase.Text)) < Year(Now) Then
   MsgBox "O Cálculo Geral não pode ser efetuado para anos anteriores.", vbCritical, "Atenção"
   mskDataBase.SetFocus
   Exit Sub
End If

'CONFIRMAÇÃO
If MsgBox("Deseja efetuar o Cálculo Geral de ISS para " & lblAno.Caption & "?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
    Exit Sub
Else
    Ocupado
    Timer1.Interval = 700
    lblWait.Visible = True
    nAnoExercicio = Val(lblAno.Caption)
End If

'ZERA CONTADORES
PbE.Value = 0
PbF.Value = 0
PbL.Value = 0
PbS.Value = 0
PbV.Value = 0
nPosF = 0
nPosE = 0
nPosV = 0
nPosL = 0
nPosS = 0
nTotalF = Val(lblTotF.Caption)
nTotalE = Val(lblTotE.Caption)
nTotalV = Val(lblTotV.Caption)
nTotalL = Val(lblTotL.Caption)
nTotalS = Val(lblTotS.Caption)

'********************************
' APAGANDO AS TABELAS
'********************************
cn.QueryTimeout = 0
lblWait.Caption = "PESQUISANDO AGUARDE........"
lblWait.Refresh
'TABELA LASERISS
lblWait.Caption = "APAGANDO TABELA LASER........"
lblWait.Refresh
Sql = "TRUNCATE TABLE LASERISS"
cn.Execute Sql, rdExecDirect
If cGetInputState() <> 0 Then DoEvents


paramparc:
'********************************
' PARAMETROS DAS PARCELAS
'********************************
lblWait.Caption = "CARREGANDO PARÂMETROS AGUARDE...."
lblWait.Refresh

'PARCELAS PARA ISS FIXO E TLL
Sql = "SELECT QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,"
Sql = Sql & "VENC01,VENC02,VENC03,VENC04,VENC05,VENC06,VENC07,VENC08,"
Sql = Sql & "VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & nAnoExercicio & " AND CODTIPO=2"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    bUnicaF = IIf(!PARCELAUNICA = "S", True, False)
    nDescUnicaF = FormatNumber(!DESCONTOUNICA / 100, 2)
    ReDim aParcF(!qtdeparcela)
    Do Until .EOF
       If bUnicaF Then
          If Not IsNull(!VENCUNICA) Then aParcF(0) = Format(!VENCUNICA, "dd/mm/yyyy")
       End If
       If Not IsNull(!VENC01) Then aParcF(1) = Format(!VENC01, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC02) Then aParcF(2) = Format(!VENC02, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC03) Then aParcF(3) = Format(!VENC03, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC04) Then aParcF(4) = Format(!VENC04, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC05) Then aParcF(5) = Format(!VENC05, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC06) Then aParcF(6) = Format(!VENC06, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC07) Then aParcF(7) = Format(!VENC07, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC08) Then aParcF(8) = Format(!VENC08, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC09) Then aParcF(9) = Format(!VENC09, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC10) Then aParcF(10) = Format(!VENC10, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC11) Then aParcF(11) = Format(!VENC11, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC12) Then aParcF(12) = Format(!VENC12, "dd/mm/yyyy") Else Exit Do
       x = x + 1
      .MoveNext
    Loop
End With

'PARCELAS PARA ISS ESTIMADO E VARIÁVEL
Sql = "SELECT QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,"
Sql = Sql & "VENC01,VENC02,VENC03,VENC04,VENC05,VENC06,VENC07,VENC08,"
Sql = Sql & "VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & nAnoExercicio & " AND CODTIPO=3"
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
    bUnicaV = IIf(!PARCELAUNICA = "S", True, False)
    nDescUnicaV = FormatNumber(!DESCONTOUNICA / 100, 2)
    ReDim aParcV(!qtdeparcela)
    Do Until .EOF
       If bUnicaV Then
          If Not IsNull(!VENCUNICA) Then aParcV(0) = Format(!VENCUNICA, "dd/mm/yyyy")
       End If
       If Not IsNull(!VENC01) Then aParcV(1) = Format(!VENC01, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC02) Then aParcV(2) = Format(!VENC02, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC03) Then aParcV(3) = Format(!VENC03, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC04) Then aParcV(4) = Format(!VENC04, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC05) Then aParcV(5) = Format(!VENC05, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC06) Then aParcV(6) = Format(!VENC06, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC07) Then aParcV(7) = Format(!VENC07, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC08) Then aParcV(8) = Format(!VENC08, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC09) Then aParcV(9) = Format(!VENC09, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC10) Then aParcV(10) = Format(!VENC10, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC11) Then aParcV(11) = Format(!VENC11, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC12) Then aParcV(12) = Format(!VENC12, "dd/mm/yyyy") Else Exit Do
       x = x + 1
      .MoveNext
    Loop
    .Close
End With

'PARCELAS PARA VIGILÂNCIA SANITÁRIA
Sql = "SELECT QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,"
Sql = Sql & "VENC01,VENC02,VENC03,VENC04,VENC05,VENC06,VENC07,VENC08,"
Sql = Sql & "VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & nAnoExercicio & " AND CODTIPO=5"
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
    bUnicaS = IIf(!PARCELAUNICA = "S", True, False)
    nDescUnicaS = FormatNumber(!DESCONTOUNICA / 100, 2)
    ReDim aParcS(!qtdeparcela)
    Do Until .EOF
       If bUnicaS Then
          If Not IsNull(!VENCUNICA) Then aParcS(0) = Format(!VENCUNICA, "dd/mm/yyyy")
       End If
       If Not IsNull(!VENC01) Then aParcS(1) = Format(!VENC01, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC02) Then aParcS(2) = Format(!VENC02, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC03) Then aParcS(3) = Format(!VENC03, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC04) Then aParcS(4) = Format(!VENC04, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC05) Then aParcS(5) = Format(!VENC05, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC06) Then aParcS(6) = Format(!VENC06, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC07) Then aParcS(7) = Format(!VENC07, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC08) Then aParcS(8) = Format(!VENC08, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC09) Then aParcS(9) = Format(!VENC09, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC10) Then aParcS(10) = Format(!VENC10, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC11) Then aParcS(11) = Format(!VENC11, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC12) Then aParcS(12) = Format(!VENC12, "dd/mm/yyyy") Else Exit Do
       x = x + 1
      .MoveNext
    Loop
   .Close
End With
'GoTo ISSFIXOTLL
'********************************
' PARAMETROS DOS TRIBUTOS
'********************************
'ALVARA
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAnoExercicio & " AND CODTRIBUTO=" & COD_TRIBALVARA
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
    'nValorAlvara = FormatNumber(!VALORALIQ, 2)
    nValorAlvara = 0
   .Close
End With
'VISTORIA
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAnoExercicio & " AND CODTRIBUTO=" & COD_TRIBVISTORIA
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
    nValorVistoria = FormatNumber(!VALORALIQ, 2)
   .Close
End With
'protocolo
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAnoExercicio & " AND CODTRIBUTO=" & COD_TRIBprotocolo
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
    'nValorprotocolo = FormatNumber(!VALORALIQ, 2)
    nValorprotocolo = 0
   .Close
End With
'UFIR ATUAL
Sql = "SELECT VALORUFIR FROM UFIR WHERE ANOUFIR=" & nAnoExercicio
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
    nUfirAtual = FormatNumber(!VALORUFIR, 4)
   .Close
End With
'DATABASE
sDataBase = mskDataBase.Text
'********************************
' TAXA DE EXPEDIÇÃO DE DOCUMENTO
'********************************
Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoExercicio
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
   ' Do Until .EOF
'        Select Case !CodLancamento
 '           Case 6, 14 'FIXO
                nValorExpDocParcF = FormatNumber(!VALORPARCELA, 2)
                nValorExpDocUnicaF = FormatNumber(!VALORUNICA, 2)
  '          Case 3 'ESTIMADO
                nValorExpDocParcE = FormatNumber(!VALORPARCELA, 2)
                nValorExpDocUnicaE = FormatNumber(!VALORUNICA, 2)
   '         Case 5 'VARIAVEL
                nValorExpDocParcV = FormatNumber(!VALORPARCELA, 2)
                nValorExpDocUnicaV = FormatNumber(!VALORUNICA, 2)
   '     End Select
    '   .MoveNext
    'Loop
   .Close
End With
'ULTIMO Nº DE DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdo99
    nLastDoc = !ULTIMO + 300
   .Close
End With

'********************************
' ###### EFETUANDO CÁLCULO ######
'********************************

lblWait.Caption = "CÁLCULO EM ANDAMENTO AGUARDE....."
lblWait.Refresh

Open sPathBin & "\DEBITOPARCELA.TXT" For Output As #1
Open sPathBin & "\DEBITOTRIBUTO.TXT" For Output As #2
Open sPathBin & "\PARCELADOCUMENTO.TXT" For Output As #3
Open sPathBin & "\NUMDOCUMENTO.TXT" For Output As #4

Sql = "SELECT DISTINCT MOBILIARIO.CODIGOMOB,MOBILIARIOATIVIDADEISS.CODTRIBUTO FROM MOBILIARIOATIVIDADEISS FULL OUTER JOIN MOBILIARIO ON "
Sql = Sql & "MOBILIARIOATIVIDADEISS.CODMOBILIARIO = MOBILIARIO.CODIGOMOB WHERE DATAENCERRAMENTO Is Null AND MOBILIARIO.CODATIVIDADE>0 AND MOBILIARIO.CODIGOMOB<500000"
'Sql = Sql & "and MOBILIARIO.CODIGOMOB=111223"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       
'        If !CODIGOMOB = 115529 Or !CODIGOMOB = 11528 Or !CODIGOMOB = 115527 Or !CODIGOMOB = 115526 Or !CODIGOMOB = 115508 Or !CODIGOMOB = 115518 Then
'            GoTo proximo
'        End If
       
       
       
       
       'SUSPENÇÃO
        Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & !CODIGOMOB
        Sql = Sql & " ORDER BY DATAEVENTO DESC"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If !CODTIPOEVENTO = 2 Then
                    GoTo PROXIMO
                End If
            End If
           .Close
        End With
        
        nCodEmpresa = !CODIGOMOB
'GoTo VIGSANIT
        
        'GRAVA NA TABELA LASERISS
        Sql = "INSERT LASERISS(CODIGOMOB) VALUES(" & !CODIGOMOB & ")"
'        cn.Execute Sql, rdExecDirect
        If IsNull(!CodTributo) Then
            nTipoIss = 0
        Else
            nTipoIss = !CodTributo
        End If
        
       'CARREGA VALOR ATIVIDADE
        nValorAliquotaISS = 0
        nValorEstimado = 0
        nQtdeProfISS = 0
        Sql = "SELECT MOBILIARIOATIVIDADEISS.CODMOBILIARIO,MOBILIARIOATIVIDADEISS.CODTRIBUTO,MOBILIARIOATIVIDADEISS.CODATIVIDADE,"
        Sql = Sql & "ATIVIDADEISS.DESCATIVIDADE,MOBILIARIOATIVIDADEISS.SEQ,MOBILIARIOATIVIDADEISS.QTDEISS,MOBILIARIOATIVIDADEISS.VALORISS,"
        Sql = Sql & "TRIBUTO.DESCTRIBUTO , TABELAISS.ALIQUOTA FROM MOBILIARIOATIVIDADEISS INNER JOIN ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE Inner Join "
        Sql = Sql & "TRIBUTO ON MOBILIARIOATIVIDADEISS.CODTRIBUTO = TRIBUTO.CODTRIBUTO Inner Join TABELAISS ON MOBILIARIOATIVIDADEISS.CODTRIBUTO = TABELAISS.TIPOISS AND MOBILIARIOATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV "
        Sql = Sql & "Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO =" & !CODIGOMOB
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                Do Until .EOF
                    nQtdeProfISS = nQtdeProfISS + IIf(!QTDEISS = 0, 1, !QTDEISS)
                    'nValorAliquotaISS = (RetornaAliquotaISS(nCodEmpresa, Format(Now, "dd/mm/yyyy")) * IIf(!QTDEISS = 0, 1, !QTDEISS) * nUfirAtual)
                    'nValorAliquotaISS = (   !Aliquota * IIf(!QTDEISS = 0, 1, !QTDEISS) * nUfirAtual)
                    nValorEstimado = nValorEstimado + FormatNumber(!valoriss * nValorAliquotaISS * IIf(!QTDEISS = 0, 1, !QTDEISS), 2)
                   .MoveNext
                Loop
            End If
           .Close
        End With
       'CARREGA DADOS DA EMPRESA
        Sql = "SELECT CODIGOMOB,CODATIVIDADE,AREATL,VISTORIA,QTDEPROF FROM  MOBILIARIO WHERE CODIGOMOB=" & nCodEmpresa
        Set RdoEmp = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoEmp
             If RdoEmp.RowCount = 0 Then GoTo PROXIMO
             nQtdeProfTL = IIf(!QTDEPROF = 0, 1, !QTDEPROF)
'             nCodEmpresa = !CODIGOMOB
             nCodAtividade = !CODATIVIDADE
             nArea = FormatNumber(IIf(IsNull(!AREATL), 0, !AREATL), 2)
             bVistoria = IIf(!VISTORIA = 1, True, False)
             If nTipoIss = COD_TRIBISSFIXO Or nTipoIss = 0 Then
                '********************************
                ' CÁLCULO DE ISS FIXO
                '********************************
                If nTipoIss = COD_TRIBISSFIXO Then
                   nPosF = nPosF + 1
                End If
                'GERA DEBITOS PARCELAS
                nPosL = nPosL + 1
                If nPosL Mod 50 = 0 Then
                   CallPb
                End If
                 
                For nNumParcela = 0 To UBound(aParcF)
                    If nNumParcela = 0 And Not bUnicaF Then GoTo PROXIMOF
                                        
                    'VERIFICA SE TEM TAXA DE LICENÇA
                    ReDim aValorAliquotaTxL(0)
                    Sql = "SELECT MOBILIARIO.CODATIVIDADE, QTDEPROF,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
                    Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodEmpresa
                    Set RdoAliq = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAliq
                        Select Case !CODIGOALIQ
                            Case 1
                                aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                            Case 2
                                aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                            Case 3
                                aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                        End Select
                        bTemTL = True
                        If IsNull(!AREATL) Or !AREATL = 0 Then
                           If aValorAliquotaTxL(0).nValorAliq < 14 Then
                              bTemTL = False
                           End If
                        End If
                        aValorAliquotaTxL(0).nArea = IIf(IsNull(!AREATL), 0, !AREATL)
                        If aValorAliquotaTxL(0).nArea = 0 Then aValorAliquotaTxL(0).nArea = 1
                        
                        'TABELA DE LIMITANTES DE ÁREA
                        If aValorAliquotaTxL(0).nArea > 27000 And aValorAliquotaTxL(0).nValorAliq = 0.29 Then
                           aValorAliquotaTxL(0).nArea = 27000
                        ElseIf aValorAliquotaTxL(0).nArea > 9000 And (aValorAliquotaTxL(0).nValorAliq = 0.58 Or aValorAliquotaTxL(0).nValorAliq = 0.36) Then
                           aValorAliquotaTxL(0).nArea = 9000
                        ElseIf aValorAliquotaTxL(0).nArea > 6000 And (aValorAliquotaTxL(0).nValorAliq = 0.72 Or aValorAliquotaTxL(0).nValorAliq = 0.86) Then
                           aValorAliquotaTxL(0).nArea = 6000
                        End If
                        
                        Sql = "SELECT MOBILIARIOATIVIDADETL.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE,MOBILIARIOATIVIDADETL.CODIGOALIQ,"
                        Sql = Sql & "ATIVIDADE.VALORALIQ1, ATIVIDADE.VALORALIQ2,ATIVIDADE.VALORALIQ3, MOBILIARIO.AREATL,MOBILIARIO.QTDEPROF "
                        Sql = Sql & "FROM ATIVIDADE INNER JOIN MOBILIARIOATIVIDADETL ON ATIVIDADE.CODATIVIDADE = MOBILIARIOATIVIDADETL.CODATIVIDADE "
                        Sql = Sql & "Inner Join MOBILIARIO ON MOBILIARIOATIVIDADETL.CODIGOMOB = MOBILIARIO.CODIGOMOB "
                        Sql = Sql & "where MOBILIARIOATIVIDADETL.CODIGOMOB=" & nCodEmpresa
                        Set Rdo99 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With Rdo99
                            Do Until .EOF
                                ReDim Preserve aValorAliquotaTxL(UBound(aValorAliquotaTxL) + 1)
                                Select Case !CODIGOALIQ
                                    Case 1
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                                    Case 2
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                                    Case 3
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                                End Select
                                aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nArea = !AREATL
                               .MoveNext
                            Loop
                           .Close
                        End With
                        bTaxaLic = True
                        nValorTxLic = 0
                        If nTipoIss = COD_TRIBISSFIXO Then
                           For x = 0 To UBound(aValorAliquotaTxL)
                               nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                           Next
                        Else
                           If aValorAliquotaTxL(0).nValorAliq = 0 Then
                              bTaxaLic = False
                           End If
                           For x = 0 To UBound(aValorAliquotaTxL)
                               If aValorAliquotaTxL(0).nValorAliq <= 14 Then
                                  If aValorAliquotaTxL(0).nArea = 0 Then
                                      bTaxaLic = False
                                  End If
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * aValorAliquotaTxL(0).nArea * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               Else
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               End If
                           Next
                        End If

                        'TAXA DE LICENÇA
                         If bTemTL Then
                            If nValorTxLic > 0 And bTaxaLic Then
                             ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                             ax = ax & COD_TRIBTAXALICENCA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorTxLic - (nValorTxLic * nDescUnicaF), 2), Round(nValorTxLic / UBound(aParcF), 2))) & ","
                             ax = ax & 0 & "," & 0 & "," & 0
 '                            Print #2, ax
                            End If
                         End If
                       .Close
                    End With
                    'GRAVA NA TABELA DEBITOPARCELA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & STATUS_NAOPAGO & "," & Format(aParcF(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                    ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                    ax = ax & Null & "," & 0
                    Print #1, ax
                    'GRAVA NA TABELA NUMDOCUMENTO
                    nLastDoc = nLastDoc + 1
                    ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaF, nValorExpDocParcF))
                    Print #4, ax
                    'GRAVA NA TABELA PARCELADOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & ","
                    ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                    Print #3, ax
                    'CALCULA VALOR PARCELA ISS FIXO
                    'nValorTotal = nValorAliquotaISS * nUfirAtual * nQtdeProfISS
                    nValorTotal = nValorAliquotaISS
                    nValorParcela = nValorTotal / UBound(aParcF)
                    If bUnicaF Then
                        nValorUnica = nValorTotal - (nValorTotal * nDescUnicaF)
                    End If
                    'GERA DEBITOS TRIBUTO
                    'ISS ESTIMADO/VARIAVEL
                    If nTipoIss = COD_TRIBISSFIXO Then
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBISSFIXO & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                        Sql = "INSERT LASERISS(CODIGOMOB,NUMDOCUMENTO,NUMPARCELA,VALORPARCELA) VALUES(" & !CODIGOMOB & "," & nLastDoc & "," & nNumParcela & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ")"
                        cn.Execute Sql, rdExecDirect
                    End If
                    'ALVARA
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBALVARA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorAlvara, 2), Round(nValorAlvara / UBound(aParcF), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                    'protocolo
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBprotocolo & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorprotocolo, 2), Round(nValorprotocolo / UBound(aParcF), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                    'EXPEDIENTE FOI REMOVIDO CONFORME SOLICITAÇÃO EDUARDO 28/12/04
                    'VISTORIA
                    If bVistoria And bTaxaLic Then
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBVISTORIA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorVistoria, 2), Round(nValorVistoria / UBound(aParcF), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                    End If
                    'TAXA DE EXPEDIENTE DOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocParcF)) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
   '                 Print #2, ax
PROXIMOF:
                 Next
                .Close
             Else
                 '************************************
                 ' CÁLCULO DE ISS ESTIMADO E VARIÁVEL
                 '************************************
'                 GoTo PROXIMO
                 'ATUALIZA GAUGE
                 Select Case nTipoIss
                    Case COD_TRIBISSESTIMADO
                        nPosE = nPosE + 1
                    Case COD_TRIBISSVARIAVEL
                        GoTo TAXALIC
                        nPosV = nPosV + 1
                 End Select
                 
                 
                 'GERA LANÇAMENTOS PARCELAS ISS
                 For nNumParcela = 0 To UBound(aParcV)
                    If nNumParcela = 0 And Not bUnicaV Then GoTo PROXIMOEV
                    'GRAVA NA TABELA DEBITOPARCELA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & STATUS_NAOPAGO & "," & Format(aParcV(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                    ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                    ax = ax & Null & "," & 0
                    Print #1, ax
                    'GRAVA NA TABELA NUMDOCUMENTO
                    nLastDoc = nLastDoc + 1
                    ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & ","
                    If nTipoIss = 12 Then
                       ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaE, nValorExpDocParcE))
                    Else
                       ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaV, nValorExpDocParcV))
                    End If
                    Print #4, ax
                    'GRAVA NA TABELA PARCELADOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & ","
                    ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                    Print #3, ax
                    If nTipoIss = COD_TRIBISSESTIMADO Then
                        'CALCULA VALOR PARCELA ISS ESTIMADO
                        nValorTotal = nValorEstimado
                        nValorParcela = nValorTotal
                        If bUnicaV Then
                            nValorUnica = nValorTotal - (nValorTotal * nDescUnicaV)
                        End If
                    Else
                        'CALCULA VALOR PARCELA ISS VARIAVEL
                        nValorParcela = 0
                        nValorUnica = 0
                    End If
                    'GERA DEBITOS TRIBUTO
                    'ISS VARIAVEL/ESTIMADO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_TRIBISSESTIMADO, COD_TRIBISSVARIAVEL) & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
                    'TAXA DE EXPEDIÇÃO DOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocParcE)) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
PROXIMOEV:
                 Next
                 'GERA LANÇAMENTOS TAXA LICENÇA
TAXALIC:
                 nPosL = nPosL + 1
                 If nPosL Mod 20 = 0 Then
                    CallPb
                 End If
'*****
                'VERIFICA SE TEM AREA E ALIQUOTA SENÃO NÃO CALCULA
                ReDim aValorAliquotaTxL(0)
                Sql = "SELECT MOBILIARIO.CODATIVIDADE, QTDEPROF,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
                Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodEmpresa
                Set RdoAliq = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAliq
                    Select Case !CODIGOALIQ
                        Case 1
                            aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                        Case 2
                            aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                        Case 3
                            aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                    End Select
                    aValorAliquotaTxL(0).nArea = IIf(IsNull(!AREATL), 0, !AREATL)
                End With
                If aValorAliquotaTxL(0).nValorAliq = 0 And aValorAliquotaTxL(0).nArea = 0 Then
                    GoTo VIGSANIT
                End If
'*****
               
                For nNumParcela = 0 To UBound(aParcF)
                   'GRAVA NA TABELA DEBITOPARCELA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & STATUS_NAOPAGO & "," & Format(aParcF(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                    ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                    ax = ax & Null & "," & 0
'                    Print #1, ax
                    'GRAVA NA TABELA NUMDOCUMENTO
                    nLastDoc = nLastDoc + 1
                    ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & ","
                    ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaF, nValorExpDocParcF))
'                    Print #4, ax
                    'GRAVA NA TABELA PARCELADOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & 0 & ","
                    ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
'                    Print #3, ax
                    ReDim aValorAliquotaTxL(0)
                    Sql = "SELECT MOBILIARIO.CODATIVIDADE,QTDEPROF, DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
                    Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodEmpresa
                    Set RdoAliq = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAliq
                        Select Case !CODIGOALIQ
                            Case 1
                                aValorAliquotaTxL(0).nValorAliq = CDbl(!VALORALIQ1)
                            Case 2
                                aValorAliquotaTxL(0).nValorAliq = !VALORALIQ2
                            Case 3
                                aValorAliquotaTxL(0).nValorAliq = !VALORALIQ3
                        End Select
                        aValorAliquotaTxL(0).nArea = IIf(IsNull(!AREATL), 0, !AREATL)
                        nQtdeProfTL = Val(SubNull(!QTDEPROF))
                        If nQtdeProfTL = 0 Then nQtdeProfTL = 1

                        Sql = "SELECT MOBILIARIOATIVIDADETL.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE,MOBILIARIOATIVIDADETL.CODIGOALIQ,"
                        Sql = Sql & "ATIVIDADE.VALORALIQ1, ATIVIDADE.VALORALIQ2,ATIVIDADE.VALORALIQ3, MOBILIARIO.AREATL,MOBILIARIO.QTDEPROF "
                        Sql = Sql & "FROM ATIVIDADE INNER JOIN MOBILIARIOATIVIDADETL ON ATIVIDADE.CODATIVIDADE = MOBILIARIOATIVIDADETL.CODATIVIDADE "
                        Sql = Sql & "Inner Join MOBILIARIO ON MOBILIARIOATIVIDADETL.CODIGOMOB = MOBILIARIO.CODIGOMOB "
                        Sql = Sql & "where MOBILIARIOATIVIDADETL.CODIGOMOB=" & nCodEmpresa
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux
                            Do Until .EOF
                                ReDim Preserve aValorAliquotaTxL(UBound(aValorAliquotaTxL) + 1)
                                Select Case !CODIGOALIQ
                                    Case 1
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                                    Case 2
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                                    Case 3
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                                End Select
                                aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nArea = !AREATL
                               .MoveNext
                            Loop
                        End With

                        nValorTxLic = 0
                        If nTipoIss = COD_TRIBISSFIXO Then
                           For x = 0 To UBound(aValorAliquotaTxL)
                               nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                           Next
                        Else
                           For x = 0 To UBound(aValorAliquotaTxL)
                               If aValorAliquotaTxL(x).nValorAliq <= 14 Then
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * aValorAliquotaTxL(x).nArea * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               Else
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               End If
                           Next
                        End If
                       .Close
                    End With
                    If nTipoIss = COD_TRIBISSESTIMADO Then
                        'CALCULA VALOR TX.LICENÇA ISS ESTIMADO
                        nValorUnica = nValorTxLic
                        nValorParcela = nValorUnica / UBound(aParcF)
                    Else
                        'CALCULA VALOR PARCELA ISS VARIAVEL
                        nValorParcela = 0
                        nValorUnica = 0
                    End If
                   'GERA DEBITOS TRIBUTO
                   'ISS VARIAVEL/ESTIMADO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTAXALICENCA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorTxLic - (nValorTxLic * nDescUnicaF), 2), Round(nValorTxLic / UBound(aParcF), 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
'                    Print #2, ax
                   'ALVARA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBALVARA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorAlvara, 2), Round(nValorAlvara / UBound(aParcF), 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
'                    Print #2, ax
                   'protocolo
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBprotocolo & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorprotocolo, 2), Round(nValorprotocolo / UBound(aParcF), 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
'                    Print #2, ax
                   'VISTORIA
                    If bVistoria Then
                       ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                       ax = ax & COD_TRIBVISTORIA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorVistoria, 2), Round(nValorVistoria / UBound(aParcF), 2))) & ","
                       ax = ax & 0 & "," & 0 & "," & 0
'                       Print #2, ax
                    End If
                   'TAXA DE EXPEDIENTE DOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocUnicaF)) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
'                    Print #2, ax
PROXIMOTL:
                 Next
                .Close
             End If

'***

VIGSANIT:
            '************************************
            ' CÁLCULO DE VIGILÂNCIA SANITÁRIA
            '************************************
            GoTo PROXIMO
            Sql = "SELECT * FROM MOBILIARIOATIVIDADEVS2 WHERE CODMOBILIARIO=" & nCodEmpresa
'            Sql = "SELECT MOBILIARIOATIVIDADEVS.CODMOBILIARIO,VIGSANITARIA.VALORALIQ,MOBILIARIOATIVIDADEVS.QTDE FROM MOBILIARIOATIVIDADEVS INNER JOIN "
'            Sql = Sql & "VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT AND MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT "
'            Sql = Sql & "Inner Join MOBILIARIO ON MOBILIARIOATIVIDADEVS.CODMOBILIARIO = MOBILIARIO.CODIGOMOB "
'            Sql = Sql & "WHERE CODMOBILIARIO = " & nCodEmpresa & " AND  DATAENCERRAMENTO IS NULL"
            Set RdoVig = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoVig
                If .RowCount > 0 Then
                    nValorAliquotaVS = 0
                    nQtdeProfVS = 0
                    nPosS = nPosS + 1
                    Do Until .EOF
                       nQtdeProfVS = IIf(!QTDE = 0, 1, !QTDE)
                       'nValorAliquotaVS = nValorAliquotaVS + (!VALORALIQ * nUfirAtual * nQtdeProfVS)
                       nValorAliquotaVS = nValorAliquotaVS + (!VALOR * nQtdeProfVS)
                      .MoveNext
                    Loop
                    nValorAliquotaVS = FormatNumber(nValorAliquotaVS, 2)
                    nSeqLanc = 0
                    For nNumParcela = 0 To UBound(aParcS)
                       If nNumParcela = 0 And Not bUnicaS Then GoTo PROXIMOVS
                       'GRAVA NA TABELA DEBITOPARCELA
                         ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                         ax = ax & STATUS_NAOPAGO & "," & Format(aParcS(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                         ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                         ax = ax & Null & "," & 0
                         Print #1, ax
                       'GRAVA NA TABELA NUMDOCUMENTO
                        nLastDoc = nLastDoc + 1
                        ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & ","
                        ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaF, nValorExpDocParcF))
                        Print #4, ax
                       'GRAVA NA TABELA PARCELADOCUMENTO
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & ","
                        ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                        Print #3, ax
                       'GRAVA NA TABELA DEBITOTRIBUTO
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBVIGSANITARIA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorAliquotaVS - (nValorAliquotaVS * nDescUnicaS), 2), Round(nValorAliquotaVS / UBound(aParcS), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                        'TAXA DE EXPEDIENTE DOCUMENTO
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocParcF)) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
'                        Print #2, ax
PROXIMOVS:
                    Next
                End If
               .Close
            End With
'****
        End With
        
PROXIMO:
       .MoveNext
    Loop
End With
Close 4#
Close 3#
Close 2#
Close 1#

'MsgBox nTeste
'Exit Sub
Dim sCodReduz As String, sAno As String, sSeq As String
Dim sNumParc As String, sCompl As String, sDataVencto As String

Dim aParcela() As Parcela
ReDim aParcela(0)
Open sPathBin & "\DEBITOPARCELA.TXT" For Input As #1
Open sPathBin & "\DEBITOPARCELATMP.TXT" For Output As #2

lblWait.Caption = "REMOVENDO DUPLICADOS 1..."
lblWait.Refresh

If cGetInputState() <> 0 Then DoEvents

Do While Not EOF(1)
     Achou = False
     Input #1, sCodReduz, sAno, sCodLanc, sSeq, sNumParc, sCompl, sStatus, sDataVencto, sDataBase, sCodMoeda, sNumLivro, sPagLivro, sNumCertidao, sDataInsc, sDataAjuiza, sValorJuros
     For x = 0 To UBound(aParcela)
        If Val(aParcela(x).CODREDUZIDO) = Val(sCodReduz) And Val(aParcela(x).AnoExercicio) = Val(sAno) And Val(aParcela(x).CodLancamento) = Val(sCodLanc) And Val(aParcela(x).SeqLancamento) = Val(sSeq) And Val(aParcela(x).NumParcela) = Val(sNumParc) And Val(aParcela(x).CODCOMPLEMENTO) = Val(sCompl) Then
           Achou = True
        End If
     Next
     If Achou Then
'        MsgBox "duplicado"
     Else
        If UBound(aParcela) > 0 Then
           If sCodReduz <> aParcela(1).CODREDUZIDO Then
              ReDim aParcela(0)
           End If
        End If
        ReDim Preserve aParcela(UBound(aParcela) + 1)
        aParcela(UBound(aParcela)).CODREDUZIDO = sCodReduz
        aParcela(UBound(aParcela)).AnoExercicio = sAno
        aParcela(UBound(aParcela)).CodLancamento = sCodLanc
        aParcela(UBound(aParcela)).SeqLancamento = sSeq
        aParcela(UBound(aParcela)).NumParcela = sNumParc
        aParcela(UBound(aParcela)).CODCOMPLEMENTO = sCompl
        ax = sCodReduz & "," & sAno & "," & sCodLanc & "," & sSeq & "," & sNumParc & "," & sCompl & "," & sStatus & "," & sDataVencto & "," & sDataBase & "," & sCodMoeda & "," & sNumLivro & "," & sPagLivro & "," & sNumCertidao & "," & sDataInsc & "," & sDataAjuiza & "," & sValorJuros
        Print #2, ax
     End If
Loop
Close #2
Close #1

Kill sPathBin & "\DEBITOPARCELA.TXT"
Name sPathBin & "\DEBITOPARCELATMP.TXT" As sPathBin & "\DEBITOPARCELA.TXT"

lblWait.Caption = "REMOVENDO DUPLICADOS 2..."
lblWait.Refresh
If cGetInputState() <> 0 Then DoEvents

Dim sCodTrib As String, sValorTrib As String, sValorMulta As String, sValorCorrecao As String
Dim aTrib() As TRIBUTO
ReDim aTrib(0)
Open sPathBin & "\DEBITOTRIBUTO.TXT" For Input As #1
Open sPathBin & "\DEBITOTRIBUTOTMP.TXT" For Output As #2

Do While Not EOF(1)
     Achou = False
     Input #1, sCodReduz, sAno, sCodLanc, sSeq, sNumParc, sCompl, sCodTrib, sValorTrib, sValorCorrecao, sValorMulta, sValorJuros
     For x = 0 To UBound(aTrib)
        If aTrib(x).sCodReduz = sCodReduz And aTrib(x).sAno = sAno And aTrib(x).sCodLanc = sCodLanc And aTrib(x).sSeq = sSeq And aTrib(x).sNumParc = sNumParc And aTrib(x).sCompl = sCompl And aTrib(x).sCodTrib = sCodTrib Then
           Achou = True
        End If
     Next
     If Achou Then
'        MsgBox "duplicado"
     Else
        If UBound(aTrib) > 0 Then
           If sCodReduz <> aTrib(1).sCodReduz Then
              ReDim aTrib(0)
           End If
        End If
        ReDim Preserve aTrib(UBound(aTrib) + 1)
        aTrib(UBound(aTrib)).sCodReduz = sCodReduz
        aTrib(UBound(aTrib)).sAno = sAno
        aTrib(UBound(aTrib)).sCodLanc = sCodLanc
        aTrib(UBound(aTrib)).sSeq = sSeq
        aTrib(UBound(aTrib)).sNumParc = sNumParc
        aTrib(UBound(aTrib)).sCompl = sCompl
        aTrib(UBound(aTrib)).sCodTrib = sCodTrib
        ax = sCodReduz & "," & sAno & "," & sCodLanc & "," & sSeq & "," & sNumParc & "," & sCompl & "," & sCodTrib & "," & sValorTrib & "," & sValorCorrecao & "," & sValorMulta & "," & sValorJuros
        Print #2, ax
     End If
Loop
Close #2
Close #1

Kill sPathBin & "\DEBITOTRIBUTO.TXT"
Name sPathBin & "\DEBITOTRIBUTOTMP.TXT" As sPathBin & "\DEBITOTRIBUTO.TXT"

lblWait.Caption = "REMOVENDO DUPLICADOS 3..."
lblWait.Refresh
If cGetInputState() <> 0 Then DoEvents

Dim sNumDocumento As String, sCodBanco As String
ReDim aParcela(0)
Open sPathBin & "\PARCELADOCUMENTO.TXT" For Input As #1
Open sPathBin & "\PARCELADOCUMENTOTMP.TXT" For Output As #2

Do While Not EOF(1)
     Achou = False
     Input #1, sCodReduz, sAno, sCodLanc, sSeq, sNumParc, sCompl, sNumDocumento, sValorJuros, sCodBanco
     For x = 0 To UBound(aParcela)
        If aParcela(x).CODREDUZIDO = sCodReduz And aParcela(x).AnoExercicio = sAno And aParcela(x).CodLancamento = sCodLanc And aParcela(x).SeqLancamento = sSeq And aParcela(x).NumParcela = sNumParc And aParcela(x).CODCOMPLEMENTO = sCompl Then
           Achou = True
        End If
     Next
     If Achou Then
'        MsgBox "duplicado"
     Else
        If UBound(aParcela) > 0 Then
           If sCodReduz <> aParcela(1).CODREDUZIDO Then
              ReDim aParcela(0)
           End If
        End If
        ReDim Preserve aParcela(UBound(aParcela) + 1)
        aParcela(UBound(aParcela)).CODREDUZIDO = sCodReduz
        aParcela(UBound(aParcela)).AnoExercicio = sAno
        aParcela(UBound(aParcela)).CodLancamento = sCodLanc
        aParcela(UBound(aParcela)).SeqLancamento = sSeq
        aParcela(UBound(aParcela)).NumParcela = sNumParc
        aParcela(UBound(aParcela)).CODCOMPLEMENTO = sCompl
        ax = sCodReduz & "," & sAno & "," & sCodLanc & "," & sSeq & "," & sNumParc & "," & sCompl & "," & sNumDocumento & "," & sValorJuros & "," & sCodBanco
        Print #2, ax
     End If
Loop
Close #2
Close #1

Kill sPathBin & "\PARCELADOCUMENTO.TXT"
Name sPathBin & "\PARCELADOCUMENTOTMP.TXT" As sPathBin & "\PARCELADOCUMENTO.TXT"


ISSFIXOTLL:
lblWait.Caption = "GERANDO ISS\TLL..."
lblWait.Refresh
If cGetInputState() <> 0 Then DoEvents

Open sPathBin & "\DEBITOPARCELA2.TXT" For Output As #1
Open sPathBin & "\DEBITOTRIBUTO2.TXT" For Output As #2
Open sPathBin & "\PARCELADOCUMENTO2.TXT" For Output As #3

Sql = "SELECT * FROM LASERISS ORDER BY CODIGOMOB,NUMPARCELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       'GRAVA NA TABELA DEBITOPARCELA
        ax = !CODIGOMOB & "," & nAnoExercicio & "," & 14 & "," & 0 & "," & !NumParcela & "," & 0 & ","
        ax = ax & STATUS_NAOPAGO & "," & Format(aParcS(!NumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
        ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
        ax = ax & Null & "," & 0
        Print #1, ax
       'GRAVA NA TABELA DEBITOTRIBUTO
        ax = !CODIGOMOB & "," & nAnoExercicio & "," & 14 & "," & 0 & "," & !NumParcela & "," & 0 & ","
        ax = ax & 11 & "," & Virg2Ponto(!VALORPARCELA) & ","
        ax = ax & 0 & "," & 0 & "," & 0
        Print #2, ax
       'GRAVA NA TABELA DEBITOTRIBUTO EXP DOC
        ax = !CODIGOMOB & "," & nAnoExercicio & "," & 14 & "," & 0 & "," & !NumParcela & "," & 0 & ","
        ax = ax & 3 & "," & Virg2Ponto(CStr(nValorExpDocParcF)) & ","
        ax = ax & 0 & "," & 0 & "," & 0
 '       Print #2, ax
       'GRAVA NA TABELA PARCELADOCUMENTO
        ax = !CODIGOMOB & "," & nAnoExercicio & "," & 14 & "," & 0 & ","
        ax = ax & !NumParcela & "," & 0 & "," & !NumDocumento & "," & "0" & "," & "0"
        Print #3, ax
        
        
       .MoveNext
    Loop
   .Close
End With


Close 3#
Close 2#
Close 1#




Liberado
MsgBox "fim"

fim:
'********************************
' FINAL DO CÁLCULO
'********************************
Timer1.Interval = 0
lblWait.Caption = "CÁLCULO EFETUADO COM SUCESSO....."
lblWait.Refresh

Liberado

Exit Sub
Erro:
For x = 0 To rdoErrors.Count - 1
   MsgBox rdoErrors(x).Description

Next
Resume Next
Liberado
lblWait.Caption = "CÁLCULO EFETUADO COM ERRO......."
lblWait.Refresh

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub


Private Sub cmdVS_Click()
Dim aVS() As VS, Sql As String, RdoAux As rdoResultset, bAchou As Boolean, x As Integer
ReDim aVS(0)
Exit Sub
Sql = "SELECT * FROM CNAECRITERIO ORDER BY DIVISAO,GRUPO,CLASSE,SUBCLASSE,CRITERIO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aVS(UBound(aVS) + 1)
        aVS(UBound(aVS)).nDivisao = !divisao
        aVS(UBound(aVS)).nGrupo = !grupo
        aVS(UBound(aVS)).nClasse = !classe
        aVS(UBound(aVS)).nSubClasse = !subclasse
        aVS(UBound(aVS)).nCriterio = !CRITERIO
        aVS(UBound(aVS)).nValor = FormatNumber(!VALOR, 2)
       .MoveNext
    Loop
   .Close
End With


For x = 1 To UBound(aVS) - 1
    With aVS(x)
        Sql = "UPDATE MOBILIARIOATIVIDADEVS2 SET VALOR=" & Virg2Ponto(CStr(.nValor)) & " WHERE "
        Sql = Sql & "DIVISAO=" & .nDivisao & " AND GRUPO=" & .nGrupo & " AND CLASSE=" & .nClasse & " AND "
        Sql = Sql & "SUBCLASSE=" & .nSubClasse & " AND CRITERIO=" & .nCriterio
        cn.Execute Sql, rdExecDirect
'        MsgBox cn.RowsAffected
    End With
Next

MsgBox "FIM"

End Sub

Private Sub Form_Load()
lblWait.Caption = "Pronto para Calcular...."
Timer1.Interval = 0
Centraliza Me
'mskDataBase.text = Mid$(frmMdi.Sbar.Panels(6).text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).text, 4)
mskDataBase.Text = "01/01/2011"
CarregaTotal
End Sub

Private Sub CarregaTotal()

Sql = "SELECT COUNT(CODMOBILIARIO) AS TOTALF FROM MOBILIARIO INNER JOIN MOBILIARIOATIVIDADEISS ON MOBILIARIO.CODIGOMOB = MOBILIARIOATIVIDADEISS.CODMOBILIARIO WHERE CODTRIBUTO = 11 AND DATAENCERRAMENTO IS NULL AND MOBILIARIO.CODATIVIDADE>0; " & _
      "SELECT COUNT(CODMOBILIARIO) AS TOTALE FROM MOBILIARIO INNER JOIN MOBILIARIOATIVIDADEISS ON MOBILIARIO.CODIGOMOB = MOBILIARIOATIVIDADEISS.CODMOBILIARIO WHERE CODTRIBUTO = 12 AND DATAENCERRAMENTO IS NULL AND MOBILIARIO.CODATIVIDADE>0; " & _
      "SELECT COUNT(CODMOBILIARIO) AS TOTALV FROM MOBILIARIO INNER JOIN MOBILIARIOATIVIDADEISS ON MOBILIARIO.CODIGOMOB = MOBILIARIOATIVIDADEISS.CODMOBILIARIO WHERE CODTRIBUTO = 13 AND DATAENCERRAMENTO IS NULL AND MOBILIARIO.CODATIVIDADE>0; " & _
      "SELECT COUNT(*) AS TOTALL FROM MOBILIARIO WHERE DATAENCERRAMENTO IS NULL AND CODATIVIDADE>0; " & _
      "SELECT COUNT(DISTINCT CODMOBILIARIO) AS TOTALS FROM MOBILIARIOATIVIDADEVS INNER JOIN MOBILIARIO ON MOBILIARIOATIVIDADEVS.CODMOBILIARIO = MOBILIARIO.CODIGOMOB WHERE DATAENCERRAMENTO IS NULL AND MOBILIARIO.CODATIVIDADE>0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblTotF.Caption = !TOTALF
   .MoreResults
    lblTotE.Caption = !TOTALE
   .MoreResults
    lblTotV.Caption = !TOTALV
   .MoreResults
    lblTotL.Caption = !TOTALL
   .MoreResults
    lblTotS.Caption = !TOTALS
   .Close
End With

End Sub

Private Sub Timer1_Timer()
If lblWait.ForeColor = &HC0& Then
     lblWait.ForeColor = vbBlack
Else
     lblWait.ForeColor = &HC0&
End If

End Sub

Private Sub CallPb()
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotalF) <= 100 Then
   PbF.Value = (nPosF * 100) / nTotalF
   lblPF.Caption = sTr(Int(PbF.Value)) & " %"
   lblPF.Refresh
Else
   PbF.Value = 100
End If

If ((nPosL * 100) / nTotalL) <= 100 Then
   PbL.Value = (nPosL * 100) / nTotalL
   lblPL.Caption = sTr(Int(PbL.Value)) & " %"
   lblPL.Refresh
Else
   PbL.Value = 100
End If

If ((nPosE * 100) / nTotalE) <= 100 Then
   PbE.Value = (nPosE * 100) / nTotalE
   lblPE.Caption = sTr(Int(PbE.Value)) & " %"
   lblPE.Refresh
Else
   PbE.Value = 100
End If

If ((nPosV * 100) / nTotalV) <= 100 Then
   PbV.Value = (nPosV * 100) / nTotalV
   lblPV.Caption = sTr(Int(PbV.Value)) & " %"
   lblPV.Refresh
Else
   PbV.Value = 100
End If

If ((nPosS * 100) / nTotalS) <= 100 Then
   PbS.Value = (nPosS * 100) / nTotalS
   lblPS.Caption = sTr(Int(PbS.Value)) & " %"
   lblPS.Refresh
Else
   PbS.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub Calculo()
Dim aParcF() As Date, bUnicaF As Boolean, nDescUnicaF As Double
Dim aParcV() As Date, bUnicaV As Boolean, nDescUnicaV As Double
Dim aParcS() As Date, bUnicaS As Boolean, nDescUnicaS As Double
Dim nAnoExercicio As Integer, bVistoria As Boolean, nUfirAtual As Double, sDataBase As String
Dim nValorAlvara As Double, nValorVistoria As Double, nValorprotocolo As Double
Dim nValorExpDocParc As Double, nValorExpDocUnica As Double
Dim nLastDoc As Long, bAchou As Boolean

nAnoExercicio = Val(lblAno.Caption)

'********************************
' PARAMETROS DAS PARCELAS
'********************************

'PARCELAS PARA ISS FIXO E TLL
Sql = "SELECT QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,"
Sql = Sql & "VENC01,VENC02,VENC03,VENC04,VENC05,VENC06,VENC07,VENC08,"
Sql = Sql & "VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & nAnoExercicio & " AND CODTIPO=2"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    bUnicaF = IIf(!PARCELAUNICA = "S", True, False)
    nDescUnicaF = FormatNumber(!DESCONTOUNICA / 100, 2)
    ReDim aParcF(!qtdeparcela)
    Do Until .EOF
       If bUnicaF Then
          If Not IsNull(!VENCUNICA) Then aParcF(0) = Format(!VENCUNICA, "dd/mm/yyyy")
       End If
       If Not IsNull(!VENC01) Then aParcF(1) = Format(!VENC01, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC02) Then aParcF(2) = Format(!VENC02, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC03) Then aParcF(3) = Format(!VENC03, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC04) Then aParcF(4) = Format(!VENC04, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC05) Then aParcF(5) = Format(!VENC05, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC06) Then aParcF(6) = Format(!VENC06, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC07) Then aParcF(7) = Format(!VENC07, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC08) Then aParcF(8) = Format(!VENC08, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC09) Then aParcF(9) = Format(!VENC09, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC10) Then aParcF(10) = Format(!VENC10, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC11) Then aParcF(11) = Format(!VENC11, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC12) Then aParcF(12) = Format(!VENC12, "dd/mm/yyyy") Else Exit Do
       x = x + 1
      .MoveNext
    Loop
   .Close
End With

'PARCELAS PARA ISS ESTIMADO
Sql = "SELECT QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,"
Sql = Sql & "VENC01,VENC02,VENC03,VENC04,VENC05,VENC06,VENC07,VENC08,"
Sql = Sql & "VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & nAnoExercicio & " AND CODTIPO=3"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    bUnicaV = IIf(!PARCELAUNICA = "S", True, False)
    nDescUnicaV = FormatNumber(!DESCONTOUNICA / 100, 2)
    ReDim aParcV(!qtdeparcela)
    Do Until .EOF
       If bUnicaV Then
          If Not IsNull(!VENCUNICA) Then aParcV(0) = Format(!VENCUNICA, "dd/mm/yyyy")
       End If
       If Not IsNull(!VENC01) Then aParcV(1) = Format(!VENC01, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC02) Then aParcV(2) = Format(!VENC02, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC03) Then aParcV(3) = Format(!VENC03, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC04) Then aParcV(4) = Format(!VENC04, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC05) Then aParcV(5) = Format(!VENC05, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC06) Then aParcV(6) = Format(!VENC06, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC07) Then aParcV(7) = Format(!VENC07, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC08) Then aParcV(8) = Format(!VENC08, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC09) Then aParcV(9) = Format(!VENC09, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC10) Then aParcV(10) = Format(!VENC10, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC11) Then aParcV(11) = Format(!VENC11, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC12) Then aParcV(12) = Format(!VENC12, "dd/mm/yyyy") Else Exit Do
       x = x + 1
      .MoveNext
    Loop
   .Close
End With

'PARCELAS PARA VIGILÂNCIA SANITÁRIA
Sql = "SELECT QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,"
Sql = Sql & "VENC01,VENC02,VENC03,VENC04,VENC05,VENC06,VENC07,VENC08,"
Sql = Sql & "VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & nAnoExercicio & " AND CODTIPO=5"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    bUnicaS = IIf(!PARCELAUNICA = "S", True, False)
    nDescUnicaS = FormatNumber(!DESCONTOUNICA / 100, 2)
    ReDim aParcS(!qtdeparcela)
    Do Until .EOF
       If bUnicaS Then
          If Not IsNull(!VENCUNICA) Then aParcS(0) = Format(!VENCUNICA, "dd/mm/yyyy")
       End If
       If Not IsNull(!VENC01) Then aParcS(1) = Format(!VENC01, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC02) Then aParcS(2) = Format(!VENC02, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC03) Then aParcS(3) = Format(!VENC03, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC04) Then aParcS(4) = Format(!VENC04, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC05) Then aParcS(5) = Format(!VENC05, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC06) Then aParcS(6) = Format(!VENC06, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC07) Then aParcS(7) = Format(!VENC07, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC08) Then aParcS(8) = Format(!VENC08, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC09) Then aParcS(9) = Format(!VENC09, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC10) Then aParcS(10) = Format(!VENC10, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC11) Then aParcS(11) = Format(!VENC11, "dd/mm/yyyy") Else Exit Do
       If Not IsNull(!VENC12) Then aParcS(12) = Format(!VENC12, "dd/mm/yyyy") Else Exit Do
       x = x + 1
      .MoveNext
    Loop
   .Close
End With

'********************************
' PARAMETROS DOS TRIBUTOS
'********************************
'ALVARA
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAnoExercicio & " AND CODTRIBUTO=" & COD_TRIBALVARA
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorAlvara = FormatNumber(!VALORALIQ, 2)
   .Close
End With
'VISTORIA
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAnoExercicio & " AND CODTRIBUTO=" & COD_TRIBVISTORIA
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorVistoria = FormatNumber(!VALORALIQ, 2)
   .Close
End With
'protocolo
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAnoExercicio & " AND CODTRIBUTO=" & COD_TRIBprotocolo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorprotocolo = FormatNumber(!VALORALIQ, 2)
   .Close
End With
'UFIR ATUAL
Sql = "SELECT VALORUFIR FROM UFIR WHERE ANOUFIR=" & nAnoExercicio
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nUfirAtual = FormatNumber(!VALORUFIR, 4)
   .Close
End With
'TAXA EXPEDIENTE
Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoExercicio
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
    nValorExpDocUnica = FormatNumber(!VALORUNICA, 2)
   .Close
End With
'DATABASE
sDataBase = mskDataBase.Text

'ULTIMO Nº DE DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nLastDoc = !ULTIMO + 1000
   .Close
End With

Sql = "TRUNCATE TABLE LASERISS"
cn.Execute Sql, rdExecDirect

'********************************
' ###### EFETUANDO CÁLCULO ######
'********************************

'######## ISS ESTIMADO #######

Open sPathBin & "\DEBITOPARCELAE.TXT" For Output As #1
Open sPathBin & "\DEBITOTRIBUTOE.TXT" For Output As #2
Open sPathBin & "\PARCELADOCUMENTOE.TXT" For Output As #3
Open sPathBin & "\NUMDOCUMENTOE.TXT" For Output As #4

Sql = "SELECT DISTINCT MOBILIARIO.CODIGOMOB,MOBILIARIO.ISENTOTAXA,MOBILIARIOATIVIDADEISS.CODTRIBUTO FROM MOBILIARIOATIVIDADEISS FULL OUTER JOIN MOBILIARIO ON "
Sql = Sql & "MOBILIARIOATIVIDADEISS.CODMOBILIARIO = MOBILIARIO.CODIGOMOB WHERE DATAENCERRAMENTO Is Null AND MOBILIARIO.CODATIVIDADE>0 "
Sql = Sql & "and MOBILIARIO.CODIGOMOB=107893"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
'        If !CODIGOMOB = 115529 Or !CODIGOMOB = 11528 Or !CODIGOMOB = 115527 Or !CODIGOMOB = 115526 Or !CODIGOMOB = 115508 Or !CODIGOMOB = 115518 Or !CODIGOMOB = 115233 Then
'            GoTo PROXIMO
'        End If
        
       'SUSPENÇÃO
        Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & !CODIGOMOB
        Sql = Sql & " ORDER BY DATAEVENTO DESC"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If !CODTIPOEVENTO = 2 Then
                    GoTo PROXIMO
                End If
            End If
           .Close
        End With
        
        
        'SE FOR ISENTA VAI PARA VIG.SANITARIA
        If Val(SubNull(!ISENTOTAXA)) = 1 Then
            GoTo VIGSANIT
        End If
        
        
        '____>
        GoTo VIGSANIT
        
        
       'CARREGA VALOR ATIVIDADE
        nValorAliquotaISS = 0
        nValorEstimado = 0
        nQtdeProfISS = 0
        Sql = "SELECT MOBILIARIOATIVIDADEISS.CODMOBILIARIO,MOBILIARIOATIVIDADEISS.CODTRIBUTO,MOBILIARIOATIVIDADEISS.CODATIVIDADE,"
        Sql = Sql & "ATIVIDADEISS.DESCATIVIDADE,MOBILIARIOATIVIDADEISS.SEQ,MOBILIARIOATIVIDADEISS.QTDEISS,MOBILIARIOATIVIDADEISS.VALORISS,"
        Sql = Sql & "TRIBUTO.DESCTRIBUTO , TABELAISS.ALIQUOTA FROM MOBILIARIOATIVIDADEISS INNER JOIN ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE Inner Join "
        Sql = Sql & "TRIBUTO ON MOBILIARIOATIVIDADEISS.CODTRIBUTO = TRIBUTO.CODTRIBUTO Inner Join TABELAISS ON MOBILIARIOATIVIDADEISS.CODTRIBUTO = TABELAISS.TIPOISS AND MOBILIARIOATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV "
        Sql = Sql & "Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO =" & !CODIGOMOB
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                Do Until .EOF
                    nQtdeProfISS = nQtdeProfISS + IIf(!QTDEISS = 0, 1, !QTDEISS)
                    nValorAliquotaISS = (!Aliquota * IIf(!QTDEISS = 0, 1, !QTDEISS) * nUfirAtual)
                    nValorEstimado = nValorEstimado + FormatNumber(!valoriss * nValorAliquotaISS * IIf(!QTDEISS = 0, 1, !QTDEISS), 2)
                   .MoveNext
                Loop
            End If
           .Close
        End With
       'CARREGA DADOS DA EMPRESA
        Sql = "SELECT CODIGOMOB,CODATIVIDADE,AREATL,VISTORIA,QTDEPROF FROM  MOBILIARIO WHERE CODIGOMOB=" & !CODIGOMOB
        Set RdoEmp = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoEmp
             If RdoEmp.RowCount = 0 Then GoTo PROXIMO
             nQtdeProfTL = IIf(!QTDEPROF = 0, 1, !QTDEPROF)
             nCodEmpresa = !CODIGOMOB
             nCodAtividade = !CODATIVIDADE
             nArea = FormatNumber(IIf(IsNull(!AREATL), 0, !AREATL), 2)
             bVistoria = IIf(!VISTORIA = 1, True, False)
             If nTipoIss = COD_TRIBISSFIXO Or nTipoIss = 0 Then
                '********************************
                ' CÁLCULO DE ISS FIXO
                '********************************
                If nTipoIss = COD_TRIBISSFIXO Then
                   nPosF = nPosF + 1
                End If
                'GERA DEBITOS PARCELAS
                nPosL = nPosL + 1
                If nPosL Mod 50 = 0 Then
                   CallPb
                End If
                 
                For nNumParcela = 0 To UBound(aParcF)
                    If nNumParcela = 0 And Not bUnicaF Then GoTo PROXIMOF
                                        
                    'VERIFICA SE TEM TAXA DE LICENÇA
                    ReDim aValorAliquotaTxL(0)
                    Sql = "SELECT MOBILIARIO.CODATIVIDADE, QTDEPROF,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
                    Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodEmpresa
                    Set RdoAliq = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAliq
                        Select Case !CODIGOALIQ
                            Case 1
                                aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                            Case 2
                                aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                            Case 3
                                aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                        End Select
                        bTemTL = True
                        If IsNull(!AREATL) Or !AREATL = 0 Then
                           If aValorAliquotaTxL(0).nValorAliq < 14 Then
                              bTemTL = False
                           End If
                        End If
                        aValorAliquotaTxL(0).nArea = IIf(IsNull(!AREATL), 0, !AREATL)
                        If aValorAliquotaTxL(0).nArea = 0 Then aValorAliquotaTxL(0).nArea = 1
                        
                        'TABELA DE LIMITANTES DE ÁREA
                        If aValorAliquotaTxL(0).nArea > 27000 And aValorAliquotaTxL(0).nValorAliq = 0.29 Then
                           aValorAliquotaTxL(0).nArea = 27000
                        ElseIf aValorAliquotaTxL(0).nArea > 9000 And (aValorAliquotaTxL(0).nValorAliq = 0.58 Or aValorAliquotaTxL(0).nValorAliq = 0.36) Then
                           aValorAliquotaTxL(0).nArea = 9000
                        ElseIf aValorAliquotaTxL(0).nArea > 6000 And (aValorAliquotaTxL(0).nValorAliq = 0.72 Or aValorAliquotaTxL(0).nValorAliq = 0.86) Then
                           aValorAliquotaTxL(0).nArea = 6000
                        End If
                        
                        Sql = "SELECT MOBILIARIOATIVIDADETL.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE,MOBILIARIOATIVIDADETL.CODIGOALIQ,"
                        Sql = Sql & "ATIVIDADE.VALORALIQ1, ATIVIDADE.VALORALIQ2,ATIVIDADE.VALORALIQ3, MOBILIARIO.AREATL,MOBILIARIO.QTDEPROF "
                        Sql = Sql & "FROM ATIVIDADE INNER JOIN MOBILIARIOATIVIDADETL ON ATIVIDADE.CODATIVIDADE = MOBILIARIOATIVIDADETL.CODATIVIDADE "
                        Sql = Sql & "Inner Join MOBILIARIO ON MOBILIARIOATIVIDADETL.CODIGOMOB = MOBILIARIO.CODIGOMOB "
                        Sql = Sql & "where MOBILIARIOATIVIDADETL.CODIGOMOB=" & nCodEmpresa
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux
                            Do Until .EOF
                                ReDim Preserve aValorAliquotaTxL(UBound(aValorAliquotaTxL) + 1)
                                Select Case !CODIGOALIQ
                                    Case 1
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                                    Case 2
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                                    Case 3
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                                End Select
                                aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nArea = !AREATL
                               .MoveNext
                            Loop
                           .Close
                        End With
                        bTaxaLic = True
                        nValorTxLic = 0
                        If nTipoIss = COD_TRIBISSFIXO Then
                           For x = 0 To UBound(aValorAliquotaTxL)
                               nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                           Next
                        Else
                           If aValorAliquotaTxL(0).nValorAliq = 0 Then
                              bTaxaLic = False
                           End If
                           For x = 0 To UBound(aValorAliquotaTxL)
                               If aValorAliquotaTxL(0).nValorAliq <= 14 Then
                                  If aValorAliquotaTxL(0).nArea = 0 Then
                                      bTaxaLic = False
                                  End If
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * aValorAliquotaTxL(0).nArea * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               Else
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               End If
                           Next
                        End If

                        'TAXA DE LICENÇA
                         If bTemTL Then
                            If nValorTxLic > 0 And bTaxaLic Then
                             ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                             ax = ax & COD_TRIBTAXALICENCA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorTxLic - (nValorTxLic * nDescUnicaF), 2), Round(nValorTxLic / UBound(aParcF), 2))) & ","
                             ax = ax & 0 & "," & 0 & "," & 0
                             Print #2, ax
                            End If
                         End If
                       .Close
                    End With
                    'GRAVA NA TABELA DEBITOPARCELA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & STATUS_NAOPAGO & "," & Format(aParcF(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                    ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                    ax = ax & Null & "," & 0
                    Print #1, ax
                    'GRAVA NA TABELA NUMDOCUMENTO
                    nLastDoc = nLastDoc + 1
                    ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaF, nValorExpDocParcF))
                    Print #4, ax
                    'GRAVA NA TABELA PARCELADOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & ","
                    ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                    Print #3, ax
                    'CALCULA VALOR PARCELA ISS FIXO
                    'nValorTotal = nValorAliquotaISS * nUfirAtual * nQtdeProfISS
                    nValorTotal = nValorAliquotaISS
                    nValorParcela = nValorTotal / UBound(aParcF)
                    If bUnicaF Then
                        nValorUnica = nValorTotal - (nValorTotal * nDescUnicaF)
                    End If
                    'GERA DEBITOS TRIBUTO
                    'ISS ESTIMADO/VARIAVEL
                    If nTipoIss = COD_TRIBISSFIXO Then
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBISSFIXO & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                    End If
                    'ALVARA
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBALVARA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorAlvara, 2), Round(nValorAlvara / UBound(aParcF), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                    'protocolo
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBprotocolo & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorprotocolo, 2), Round(nValorprotocolo / UBound(aParcF), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                    'EXPEDIENTE FOI REMOVIDO CONFORME SOLICITAÇÃO EDUARDO 28/12/04
                    'VISTORIA
                    If bVistoria And bTaxaLic Then
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBVISTORIA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorVistoria, 2), Round(nValorVistoria / UBound(aParcF), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                    End If
                    'TAXA DE EXPEDIENTE DOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCISSFIXO & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocParcF)) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
PROXIMOF:
                 Next
                .Close
             Else
                 '************************************
                 ' CÁLCULO DE ISS ESTIMADO E VARIÁVEL
                 '************************************
'                 GoTo PROXIMO
                 'ATUALIZA GAUGE
                 Select Case nTipoIss
                    Case COD_TRIBISSESTIMADO
                        nPosE = nPosE + 1
                    Case COD_TRIBISSVARIAVEL
                        GoTo TAXALIC
                        nPosV = nPosV + 1
                 End Select
                 
                 
                 'GERA LANÇAMENTOS PARCELAS ISS
                 For nNumParcela = 0 To UBound(aParcV)
                    If nNumParcela = 0 And Not bUnicaV Then GoTo PROXIMOEV
                    'GRAVA NA TABELA DEBITOPARCELA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & STATUS_NAOPAGO & "," & Format(aParcV(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                    ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                    ax = ax & Null & "," & 0
                    Print #1, ax
                    'GRAVA NA TABELA NUMDOCUMENTO
                    nLastDoc = nLastDoc + 1
                    ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & ","
                    If nTipoIss = 12 Then
                       ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaE, nValorExpDocParcE))
                    Else
                       ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaV, nValorExpDocParcV))
                    End If
                    Print #4, ax
                    'GRAVA NA TABELA PARCELADOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & ","
                    ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                    Print #3, ax
                    If nTipoIss = COD_TRIBISSESTIMADO Then
                        'CALCULA VALOR PARCELA ISS ESTIMADO
                        nValorTotal = nValorEstimado
                        nValorParcela = nValorTotal
                        If bUnicaV Then
                            nValorUnica = nValorTotal - (nValorTotal * nDescUnicaV)
                        End If
                    Else
                        'CALCULA VALOR PARCELA ISS VARIAVEL
                        nValorParcela = 0
                        nValorUnica = 0
                    End If
                    'GERA DEBITOS TRIBUTO
                    'ISS VARIAVEL/ESTIMADO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_TRIBISSESTIMADO, COD_TRIBISSVARIAVEL) & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
                    'TAXA DE EXPEDIÇÃO DOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & IIf(nTipoIss = COD_TRIBISSESTIMADO, COD_LANCISSESTIMADO, COD_LANCISSVARIAVEL) & "," & 0 & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocParcE)) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
PROXIMOEV:
                 Next
                 'GERA LANÇAMENTOS TAXA LICENÇA
TAXALIC:
                 nPosL = nPosL + 1
                 If nPosL Mod 20 = 0 Then
                    CallPb
                 End If
'*****
                'VERIFICA SE TEM AREA E ALIQUOTA SENÃO NÃO CALCULA
                ReDim aValorAliquotaTxL(0)
                Sql = "SELECT MOBILIARIO.CODATIVIDADE, QTDEPROF,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
                Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodEmpresa
                Set RdoAliq = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAliq
                    Select Case !CODIGOALIQ
                        Case 1
                            aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                        Case 2
                            aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                        Case 3
                            aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                    End Select
                    aValorAliquotaTxL(0).nArea = IIf(IsNull(!AREATL), 0, !AREATL)
                End With
                If aValorAliquotaTxL(0).nValorAliq = 0 And aValorAliquotaTxL(0).nArea = 0 Then
                    GoTo VIGSANIT
                End If
'*****
               
                For nNumParcela = 0 To UBound(aParcF)
                   'GRAVA NA TABELA DEBITOPARCELA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & STATUS_NAOPAGO & "," & Format(aParcF(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                    ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                    ax = ax & Null & "," & 0
                    Print #1, ax
                    'GRAVA NA TABELA NUMDOCUMENTO
                    nLastDoc = nLastDoc + 1
                    ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & ","
                    ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaF, nValorExpDocParcF))
                    Print #4, ax
                    'GRAVA NA TABELA PARCELADOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & 0 & ","
                    ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                    Print #3, ax
                    ReDim aValorAliquotaTxL(0)
                    Sql = "SELECT MOBILIARIO.CODATIVIDADE,QTDEPROF, DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
                    Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodEmpresa
                    Set RdoAliq = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAliq
                        Select Case !CODIGOALIQ
                            Case 1
                                aValorAliquotaTxL(0).nValorAliq = CDbl(!VALORALIQ1)
                            Case 2
                                aValorAliquotaTxL(0).nValorAliq = !VALORALIQ2
                            Case 3
                                aValorAliquotaTxL(0).nValorAliq = !VALORALIQ3
                        End Select
                        aValorAliquotaTxL(0).nArea = IIf(IsNull(!AREATL), 0, !AREATL)
                        nQtdeProfTL = Val(SubNull(!QTDEPROF))
                        If nQtdeProfTL = 0 Then nQtdeProfTL = 1

                        Sql = "SELECT MOBILIARIOATIVIDADETL.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE,MOBILIARIOATIVIDADETL.CODIGOALIQ,"
                        Sql = Sql & "ATIVIDADE.VALORALIQ1, ATIVIDADE.VALORALIQ2,ATIVIDADE.VALORALIQ3, MOBILIARIO.AREATL,MOBILIARIO.QTDEPROF "
                        Sql = Sql & "FROM ATIVIDADE INNER JOIN MOBILIARIOATIVIDADETL ON ATIVIDADE.CODATIVIDADE = MOBILIARIOATIVIDADETL.CODATIVIDADE "
                        Sql = Sql & "Inner Join MOBILIARIO ON MOBILIARIOATIVIDADETL.CODIGOMOB = MOBILIARIO.CODIGOMOB "
                        Sql = Sql & "where MOBILIARIOATIVIDADETL.CODIGOMOB=" & nCodEmpresa
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux
                            Do Until .EOF
                                ReDim Preserve aValorAliquotaTxL(UBound(aValorAliquotaTxL) + 1)
                                Select Case !CODIGOALIQ
                                    Case 1
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                                    Case 2
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                                    Case 3
                                        aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                                End Select
                                aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nArea = !AREATL
                               .MoveNext
                            Loop
                        End With

                        nValorTxLic = 0
                        If nTipoIss = COD_TRIBISSFIXO Then
                           For x = 0 To UBound(aValorAliquotaTxL)
                               nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                           Next
                        Else
                           For x = 0 To UBound(aValorAliquotaTxL)
                               If aValorAliquotaTxL(x).nValorAliq <= 14 Then
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * aValorAliquotaTxL(x).nArea * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               Else
                                  nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAnoExercicio) * nQtdeProfTL)
                               End If
                           Next
                        End If
                       .Close
                    End With
                    If nTipoIss = COD_TRIBISSESTIMADO Then
                        'CALCULA VALOR TX.LICENÇA ISS ESTIMADO
                        nValorUnica = nValorTxLic
                        nValorParcela = nValorUnica / UBound(aParcF)
                    Else
                        'CALCULA VALOR PARCELA ISS VARIAVEL
                        nValorParcela = 0
                        nValorUnica = 0
                    End If
                   'GERA DEBITOS TRIBUTO
                   'ISS VARIAVEL/ESTIMADO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTAXALICENCA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorTxLic - (nValorTxLic * nDescUnicaF), 2), Round(nValorTxLic / UBound(aParcF), 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
                   'ALVARA
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBALVARA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorAlvara, 2), Round(nValorAlvara / UBound(aParcF), 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
                   'protocolo
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBprotocolo & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorprotocolo, 2), Round(nValorprotocolo / UBound(aParcF), 2))) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
                   'VISTORIA
                    If bVistoria Then
                       ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                       ax = ax & COD_TRIBVISTORIA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorVistoria, 2), Round(nValorVistoria / UBound(aParcF), 2))) & ","
                       ax = ax & 0 & "," & 0 & "," & 0
                       Print #2, ax
                    End If
                   'TAXA DE EXPEDIENTE DOCUMENTO
                    ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCTAXALICENCA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                    ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocUnicaF)) & ","
                    ax = ax & 0 & "," & 0 & "," & 0
                    Print #2, ax
PROXIMOTL:
                 Next
                .Close
             End If

'***

VIGSANIT:
            '************************************
            ' CÁLCULO DE VIGILÂNCIA SANITÁRIA
            '************************************
            'GoTo proximo
            Sql = "SELECT * FROM MOBILIARIOATIVIDADEVS2 WHERE CODMOBILIARIO=" & nCodEmpresa
'            Sql = "SELECT MOBILIARIOATIVIDADEVS.CODMOBILIARIO,VIGSANITARIA.VALORALIQ,MOBILIARIOATIVIDADEVS.QTDE FROM MOBILIARIOATIVIDADEVS INNER JOIN "
'            Sql = Sql & "VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT AND MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT "
'            Sql = Sql & "Inner Join MOBILIARIO ON MOBILIARIOATIVIDADEVS.CODMOBILIARIO = MOBILIARIO.CODIGOMOB "
'            Sql = Sql & "WHERE CODMOBILIARIO = " & nCodEmpresa & " AND  DATAENCERRAMENTO IS NULL"
            Set RdoVig = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoVig
                If .RowCount > 0 Then
                    nValorAliquotaVS = 0
                    nQtdeProfVS = 0
                    nPosS = nPosS + 1
                    Do Until .EOF
                       nQtdeProfVS = IIf(!QTDE = 0, 1, !QTDE)
                       nValorAliquotaVS = nValorAliquotaVS + (!VALOR * nUfirAtual * nQtdeProfVS)
                       'nValorAliquotaVS = nValorAliquotaVS + (!VALORALIQ * nUfirAtual * nQtdeProfVS)
                      .MoveNext
                    Loop
                    nValorAliquotaVS = FormatNumber(nValorAliquotaVS, 2)
                    nSeqLanc = 0
                    For nNumParcela = 0 To UBound(aParcS)
                       If nNumParcela = 0 And Not bUnicaS Then
                            GoTo PROXIMOVS
                       End If
                       'GRAVA NA TABELA DEBITOPARCELA
                         ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                         ax = ax & STATUS_NAOPAGO & "," & Format(aParcS(nNumParcela), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
                         ax = ax & MOEDA_REAL & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                         ax = ax & Null & "," & 0
                         Print #1, ax
                       'GRAVA NA TABELA NUMDOCUMENTO
                        nLastDoc = nLastDoc + 1
                        ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & ","
                        ax = ax & Virg2Ponto(IIf(nNumParcela = 0, nValorExpDocUnicaF, nValorExpDocParcF))
                        Print #4, ax
                       'GRAVA NA TABELA PARCELADOCUMENTO
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & ","
                        ax = ax & nNumParcela & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                        Print #3, ax
                       'GRAVA NA TABELA DEBITOTRIBUTO
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBVIGSANITARIA & "," & Virg2Ponto(IIf(nNumParcela = 0, Round(nValorAliquotaVS - (nValorAliquotaVS * nDescUnicaS), 2), Round(nValorAliquotaVS / UBound(aParcS), 2))) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
                        'TAXA DE EXPEDIENTE DOCUMENTO
                        ax = nCodEmpresa & "," & nAnoExercicio & "," & COD_LANCVIGSANITARIA & "," & nSeqLanc & "," & nNumParcela & "," & 0 & ","
                        ax = ax & COD_TRIBTXEXPDOC & "," & Virg2Ponto(CStr(nValorExpDocParcF)) & ","
                        ax = ax & 0 & "," & 0 & "," & 0
                        Print #2, ax
PROXIMOVS:
                    Next
                Else
                    MsgBox "teste"
                End If
               .Close
            End With
'****
        End With
        
PROXIMO:
       .MoveNext
    Loop
End With


Close 4#
Close 3#
Close 2#
Close 1#
MsgBox "FIM"
fim:
Exit Sub
Erro:
For x = 0 To rdoErrors.Count - 1
   MsgBox rdoErrors(x).Description
Next

End Sub
