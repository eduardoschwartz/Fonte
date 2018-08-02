VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCobranca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Arquivo Bancário"
   ClientHeight    =   2400
   ClientLeft      =   11130
   ClientTop       =   6060
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   5895
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2310
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1260
      Width           =   3435
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   1305
      Left            =   60
      Top             =   960
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   2302
      FrameColor      =   12829635
      Style           =   0
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.ComboBox cmbLista 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCobranca.frx":0000
         Left            =   210
         List            =   "frmCobranca.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1875
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Outros"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   540
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Arrecadação diária"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRADA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   390
         Width           =   2955
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COBRANÇA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   60
         Width           =   2955
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   0
         Picture         =   "frmCobranca.frx":0050
         Top             =   0
         Width           =   2655
      End
   End
   Begin prjChameleon.chameleonButton btExec 
      Default         =   -1  'True
      Height          =   345
      Left            =   4260
      TabIndex        =   5
      ToolTipText     =   "Executar cálculo"
      Top             =   1950
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Gerar Arquivo"
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
      MICON           =   "frmCobranca.frx":0D5D
      PICN            =   "frmCobranca.frx":0D79
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   2310
      TabIndex        =   10
      Top             =   1620
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.Label lblNumRemessa 
      Caption         =   "Label4"
      Height          =   285
      Left            =   570
      TabIndex        =   9
      Top             =   2790
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do próximo arquivo..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2340
      TabIndex        =   6
      Top             =   990
      Width           =   2205
   End
End
Attribute VB_Name = "frmCobranca"
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

Private Sub Form_Load()

Centraliza Me
If NomeDeLogin <> "SCHWARTZ" Then
 Opt(1).Enabled = False
End If

AtualizaRemessa

cmbLista.ListIndex = 0
cmbLista.Enabled = False
cmbLista.BackColor = Me.BackColor

End Sub

Private Sub AtualizaRemessa()
Dim RdoAux As rdoResultset, Sql As String

ReDim aBoletos(0)
Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='COBRANCA'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nNumRemessa = RdoAux!valparam + 1
lblNumRemessa.Caption = Format(RdoAux!valparam + 1, "00000000")
RdoAux.Close

sDataArquivo = Format(Now, "dd") & Format(Now, "mm") & Right(Format(Now, "yyyy"), 2)
sArquivo = "\\192.168.200.130\ATUALIZAGTI\COBRANCA\COB_" & sDataArquivo & "_" & lblNumRemessa.Caption & ".TXT"
txtFile.Text = sArquivo
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

Private Sub Opt_Click(Index As Integer)
If Opt(0).value = True Then
    cmbLista.Enabled = False
    cmbLista.BackColor = Me.BackColor
Else
    cmbLista.Enabled = True
    cmbLista.BackColor = Branco
End If

End Sub

Private Sub btExec_Click()
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

If MsgBox("Deseja gerar este arquivo de cobrança?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") = vbNo Then Exit Sub

CarregaMatrizBoletos
If UBound(aBoletos) = 0 Then
    MsgBox "Não foi localizado nenhum boleto a ser enviado.", vbCritical, "Atenção"
Else
    GeraArquivo cmbLista.ListIndex
End If

End Sub

Private Sub CarregaMatrizBoletos()
Dim RdoAux As rdoResultset, Sql As String, nTipoDoc As Integer, sCPFCNPJ As String, sCep As String, nTotBar As Long, nNumDoc As Long, nValorGuia As Double, RdoAux2 As rdoResultset
ReDim aBoletos(0)
DoEvents
PBar.value = 0
If Opt(0).value = True Then
Else
    If cmbLista.ListIndex = 0 Then 'IPTU
        Sql = "SELECT   debitoparcela.codreduzido, debitoparcela.datadebase,debitoparcela.datavencimento,debitoparcela.numparcela,debitoparcela.codcomplemento, debitotributo.valortributo, parceladocumento.numdocumento, vwFULLIMOVEL.nomecidadao,vwFULLIMOVEL.CPF, vwFULLIMOVEL.CNPJ, vwFULLIMOVEL.LOGRADOURO, vwFULLIMOVEL.li_num, vwFULLIMOVEL.li_compl, vwFULLIMOVEL.descbairro,"
        Sql = Sql & "vwFULLIMOVEL.CodLogr FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
        Sql = Sql & "parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND "
        Sql = Sql & "debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.codcomplemento = parceladocumento.codcomplemento INNER JOIN vwFULLIMOVEL ON debitoparcela.codreduzido = vwFULLIMOVEL.codreduzido "
        'Sql = Sql & "WHERE (debitoparcela.codreduzido between 35001 and 40000)  AND (debitoparcela.anoexercicio = 2018) AND (debitoparcela.codlancamento = 1) AND (debitoparcela.seqlancamento = 0) AND (debitoparcela.codcomplemento = 0 or debitoparcela.codcomplemento = 91 or debitoparcela.codcomplemento = 92) AND (debitoparcela.statuslanc = 18) and (debitoparcela.usuario is null)  order by parceladocumento.numdocumento"
        Sql = Sql & "WHERE (debitoparcela.codreduzido between 25001 and 40000)  AND (debitoparcela.anoexercicio = 2018) AND (debitoparcela.codlancamento = 1) AND (debitoparcela.seqlancamento = 0) AND (debitoparcela.codcomplemento = 0 or debitoparcela.codcomplemento = 91 or debitoparcela.codcomplemento = 92) AND (debitoparcela.statuslanc = 3) and (debitoparcela.usuario is null)  order by parceladocumento.numdocumento"
       ' Sql = Sql & "where parceladocumento.numdocumento in (select numdoc from codtmp)"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nTotBar = .RowCount
            Do Until .EOF
                If .AbsolutePosition Mod 50 = 0 Then
                    CallPb CLng(.AbsolutePosition), nTotBar
                End If
                If Trim(SubNull(!CPF)) = "" And Trim(SubNull(!Cnpj)) = "" Then
                    GoTo PROXIMOIPTU
                End If
                
                If Not IsNull(!CPF) Then
                    If Trim(!CPF) <> "" Then
                        sCPFCNPJ = RetornaNumero(!CPF)
                        nTipoDoc = 1
                    Else
                        sCPFCNPJ = !Cnpj
                        nTipoDoc = 2
                    End If
                ElseIf Not IsNull(!Cnpj) Then
                   sCPFCNPJ = !Cnpj
                   nTipoDoc = 2
                End If
'                If !CODREDUZIDO = 973 Then MsgBox "tester"
                ReDim Preserve aBoletos(UBound(aBoletos) + 1)
                aBoletos(UBound(aBoletos)).sNossoNumero = "287353200" & CStr(!NumDocumento)
                aBoletos(UBound(aBoletos)).sNumDocumento = CStr(!NumDocumento)
                aBoletos(UBound(aBoletos)).sDataVencimento = Format(!DataVencimento, "dd") & Format(!DataVencimento, "mm") & Format(!DataVencimento, "yyyy")
                aBoletos(UBound(aBoletos)).sValorNominal = FillLeft(RetornaNumero(CStr(!ValorTributo * 100)), 15)
                If Val(aBoletos(UBound(aBoletos)).sValorNominal) <> RetornaNumero(FormatNumber(!ValorTributo, 2)) Then
                    Sql = "insert errocob values(" & !CODREDUZIDO & "," & !NumParcela & "," & !CODCOMPLEMENTO & ")"
                    cn.Execute Sql, rdExecDirect
                End If
                'aBoletos(UBound(aBoletos)).sDataBase = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
                aBoletos(UBound(aBoletos)).sDataBase = "01012018"
                aBoletos(UBound(aBoletos)).nTipoInscricao = CStr(nTipoDoc)
                aBoletos(UBound(aBoletos)).nNumeroInscricao = sCPFCNPJ
                aBoletos(UBound(aBoletos)).sNome = Left(!nomecidadao, 40)
                aBoletos(UBound(aBoletos)).sEndereco = Left(!Logradouro & ", " & CStr(!Li_Num) & " " & SubNull(!Li_Compl), 40)
                If InStr(1, aBoletos(UBound(aBoletos)).sEndereco, Chr(13)) > 0 Then
                    aBoletos(UBound(aBoletos)).sEndereco = Left(!Logradouro & ", " & CStr(!Li_Num), 40)
                End If
                aBoletos(UBound(aBoletos)).sBairro = Left(SubNull(!DescBairro), 15)
                sCep = RetornaCEP(!CodLogr, !Li_Num)
                If Len(Trim(sCep)) < 5 Then
                    sCep = 14870000
                End If
                aBoletos(UBound(aBoletos)).sCep = Left(sCep, 5)
                aBoletos(UBound(aBoletos)).sSufixoCep = Right(sCep, 3)
                aBoletos(UBound(aBoletos)).sCidade = "JABOTICABAL"
                aBoletos(UBound(aBoletos)).sUF = "SP"
                DoEvents
                
PROXIMOIPTU:
               .MoveNext
            Loop
           .Close
        End With
    ElseIf cmbLista.ListIndex = 1 Then 'ISS
        Sql = "SELECT debitoparcela.codreduzido, debitoparcela.datadebase, debitoparcela.datavencimento, parceladocumento.numdocumento, vwFULLEMPRESA.razaosocial, "
        Sql = Sql & "vwFULLEMPRESA.cnpj, vwFULLEMPRESA.cpf, vwFULLEMPRESA.LOGRADOURO, vwFULLEMPRESA.numero, vwFULLEMPRESA.complemento,vwFULLEMPRESA.CodLogradouro , vwFULLEMPRESA.Cep, vwFULLEMPRESA.DescBairro, vwFULLEMPRESA.desccidade, vwFULLEMPRESA.SiglaUF "
        Sql = Sql & "FROM debitoparcela INNER JOIN parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.codcomplemento = parceladocumento.codcomplemento INNER JOIN "
        Sql = Sql & "vwFULLEMPRESA ON debitoparcela.codreduzido = vwFULLEMPRESA.codigomob WHERE (debitoparcela.anoexercicio = 2018) AND (debitoparcela.codlancamento = 6 OR debitoparcela.codlancamento = 14) and (debitoparcela.seqlancamento = 0) and (debitoparcela.codcomplemento = 0) AND (debitoparcela.statuslanc = 18) "
        Sql = Sql & "GROUP BY debitoparcela.codreduzido, debitoparcela.datadebase, debitoparcela.datavencimento, parceladocumento.numdocumento, vwFULLEMPRESA.razaosocial,vwFULLEMPRESA.cnpj, vwFULLEMPRESA.cpf, vwFULLEMPRESA.LOGRADOURO, vwFULLEMPRESA.numero, vwFULLEMPRESA.complemento,"
        Sql = Sql & "vwFULLEMPRESA.codlogradouro, vwFULLEMPRESA.cep, vwFULLEMPRESA.descbairro, vwFULLEMPRESA.desccidade, vwFULLEMPRESA.descuf, vwFULLEMPRESA.SiglaUF "
        Sql = Sql & "HAVING (debitoparcela.codreduzido BETWEEN 100000 AND 300000) "
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                nNumDoc = !NumDocumento
                Sql = "SELECT SUM(debitotributo.valortributo) AS somatributo FROM parceladocumento INNER JOIN "
                Sql = Sql & "debitotributo ON parceladocumento.codreduzido = debitotributo.codreduzido AND parceladocumento.anoexercicio = debitotributo.anoexercicio AND "
                Sql = Sql & "parceladocumento.codlancamento = debitotributo.codlancamento AND parceladocumento.seqlancamento = debitotributo.seqlancamento AND "
                Sql = Sql & "parceladocumento.NumParcela = debitotributo.NumParcela And parceladocumento.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
                Sql = Sql & "Where parceladocumento.NumDocumento =" & nNumDoc
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                nValorGuia = Round(RdoAux2!somatributo, 2)
                RdoAux2.Close
                
                If Trim(SubNull(!CPF)) = "" And Trim(SubNull(!Cnpj)) = "" Then
                    GoTo PROXIMOISS
                End If
                
                If Not IsNull(!CPF) Then
                    If Trim(!CPF) <> "" Then
                        sCPFCNPJ = RetornaNumero(!CPF)
                        nTipoDoc = 1
                    Else
                        sCPFCNPJ = !Cnpj
                        nTipoDoc = 2
                    End If
                ElseIf Not IsNull(!Cnpj) Then
                   sCPFCNPJ = !Cnpj
                   nTipoDoc = 2
                End If
                
                ReDim Preserve aBoletos(UBound(aBoletos) + 1)
                aBoletos(UBound(aBoletos)).sNossoNumero = "287353200" & CStr(!NumDocumento)
                aBoletos(UBound(aBoletos)).sNumDocumento = CStr(!NumDocumento)
                aBoletos(UBound(aBoletos)).sDataVencimento = Format(!DataVencimento, "dd") & Format(!DataVencimento, "mm") & Format(!DataVencimento, "yyyy")
                aBoletos(UBound(aBoletos)).sValorNominal = FillLeft(RetornaNumero(CStr(nValorGuia * 100)), 15)
                If Val(aBoletos(UBound(aBoletos)).sValorNominal) <> RetornaNumero(FormatNumber(nValorGuia, 2)) Then
                    MsgBox "teste"
                End If
                aBoletos(UBound(aBoletos)).sDataBase = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
                aBoletos(UBound(aBoletos)).nTipoInscricao = CStr(nTipoDoc)
                aBoletos(UBound(aBoletos)).nNumeroInscricao = sCPFCNPJ
                aBoletos(UBound(aBoletos)).sNome = Left(!RazaoSocial, 40)
                aBoletos(UBound(aBoletos)).sEndereco = Left(!Logradouro & ", " & CStr(!Numero) & " " & SubNull(!Complemento), 40)
                aBoletos(UBound(aBoletos)).sBairro = Left(SubNull(!DescBairro), 15)
                sCep = Format(RetornaNumero(!Cep), "00000000")
                aBoletos(UBound(aBoletos)).sCep = Left(sCep, 5)
                aBoletos(UBound(aBoletos)).sSufixoCep = Right(sCep, 3)
                aBoletos(UBound(aBoletos)).sCidade = !descCidade
                aBoletos(UBound(aBoletos)).sUF = !SiglaUF
                
PROXIMOISS:
               .MoveNext
            Loop
           .Close
        End With
    ElseIf cmbLista.ListIndex = 2 Then 'VS
        Sql = "SELECT debitoparcela.codreduzido, debitoparcela.datadebase, debitoparcela.datavencimento, ROUND(SUM(distinct debitotributo.valortributo), 2) AS SomaTributo, "
        Sql = Sql & "parceladocumento.numdocumento, vwFULLEMPRESA.razaosocial, vwFULLEMPRESA.cnpj, vwFULLEMPRESA.cpf, vwFULLEMPRESA.LOGRADOURO,"
        Sql = Sql & "vwFULLEMPRESA.numero, vwFULLEMPRESA.complemento, vwFULLEMPRESA.codlogradouro,vwFULLEMPRESA.cep, vwFULLEMPRESA.descbairro, vwFULLEMPRESA.desccidade,vwFULLEMPRESA.SiglaUF "
        Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
        Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
        Sql = Sql & "parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND "
        Sql = Sql & "debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.codcomplemento = parceladocumento.codcomplemento INNER JOIN "
        Sql = Sql & "vwFULLEMPRESA ON debitoparcela.codreduzido = vwFULLEMPRESA.codigomob WHERE (debitoparcela.anoexercicio = 2018) AND (debitoparcela.codlancamento = 13) and (debitoparcela.seqlancamento = 0) and (debitoparcela.codcomplemento = 0) AND (debitoparcela.statuslanc = 18) "
        Sql = Sql & "GROUP BY debitoparcela.codreduzido, debitoparcela.datadebase, debitoparcela.datavencimento, parceladocumento.numdocumento, vwFULLEMPRESA.razaosocial,"
        Sql = Sql & "vwFULLEMPRESA.cnpj, vwFULLEMPRESA.cpf, vwFULLEMPRESA.LOGRADOURO, vwFULLEMPRESA.numero, vwFULLEMPRESA.complemento,"
        Sql = Sql & "vwFULLEMPRESA.CodLogradouro,vwFULLEMPRESA.cep , vwFULLEMPRESA.DescBairro, vwFULLEMPRESA.desccidade, vwFULLEMPRESA.descuf, vwFULLEMPRESA.SiglaUF "
        Sql = Sql & "HAVING (debitoparcela.codreduzido BETWEEN 100000 AND 300000) "
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                
                If Trim(SubNull(!CPF)) = "" And Trim(SubNull(!Cnpj)) = "" Then
                    GoTo PROXIMOVS
                End If
                
                If Not IsNull(!CPF) Then
                    If Trim(!CPF) <> "" Then
                        sCPFCNPJ = RetornaNumero(!CPF)
                        nTipoDoc = 1
                    Else
                        sCPFCNPJ = !Cnpj
                        nTipoDoc = 2
                    End If
                ElseIf Not IsNull(!Cnpj) Then
                   sCPFCNPJ = !Cnpj
                   nTipoDoc = 2
                End If
                
                ReDim Preserve aBoletos(UBound(aBoletos) + 1)
                aBoletos(UBound(aBoletos)).sNossoNumero = "287353200" & CStr(!NumDocumento)
   '             If aBoletos(UBound(aBoletos)).sNossoNumero = "28735320015572593" Then MsgBox "teste"
                aBoletos(UBound(aBoletos)).sNumDocumento = CStr(!NumDocumento)
                aBoletos(UBound(aBoletos)).sDataVencimento = Format(!DataVencimento, "dd") & Format(!DataVencimento, "mm") & Format(!DataVencimento, "yyyy")
                aBoletos(UBound(aBoletos)).sValorNominal = FillLeft(RetornaNumero(CStr(!somatributo * 100)), 15)
                If Val(aBoletos(UBound(aBoletos)).sValorNominal) <> RetornaNumero(FormatNumber(!somatributo, 2)) Then
                    MsgBox "teste"
                End If
                aBoletos(UBound(aBoletos)).sDataBase = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
                aBoletos(UBound(aBoletos)).nTipoInscricao = CStr(nTipoDoc)
                aBoletos(UBound(aBoletos)).nNumeroInscricao = sCPFCNPJ
                aBoletos(UBound(aBoletos)).sNome = Left(!RazaoSocial, 40)
                aBoletos(UBound(aBoletos)).sEndereco = Left(!Logradouro & ", " & CStr(!Numero) & " " & SubNull(!Complemento), 40)
                aBoletos(UBound(aBoletos)).sBairro = Left(SubNull(!DescBairro), 15)
                sCep = Format(RetornaNumero(!Cep), "00000000")
                aBoletos(UBound(aBoletos)).sCep = Left(sCep, 5)
                aBoletos(UBound(aBoletos)).sSufixoCep = Right(sCep, 3)
                aBoletos(UBound(aBoletos)).sCidade = !descCidade
                aBoletos(UBound(aBoletos)).sUF = !SiglaUF
                
PROXIMOVS:
               .MoveNext
            Loop
           .Close
        End With
    ElseIf cmbLista.ListIndex = 3 Then 'CIP
        Sql = "SELECT   debitoparcela.codreduzido, debitoparcela.datadebase,debitoparcela.datavencimento, debitotributo.valortributo, parceladocumento.numdocumento, vwFULLIMOVEL.nomecidadao,vwFULLIMOVEL.CPF, vwFULLIMOVEL.CNPJ, vwFULLIMOVEL.LOGRADOURO, vwFULLIMOVEL.li_num, vwFULLIMOVEL.li_compl, vwFULLIMOVEL.descbairro,"
        Sql = Sql & "vwFULLIMOVEL.CodLogr FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
        Sql = Sql & "parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND "
        Sql = Sql & "debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.codcomplemento = parceladocumento.codcomplemento INNER JOIN vwFULLIMOVEL ON debitoparcela.codreduzido = vwFULLIMOVEL.codreduzido "
        Sql = Sql & "WHERE (debitoparcela.codreduzido < 100000) AND (debitoparcela.anoexercicio = 2018) AND (debitoparcela.codlancamento = 79) AND (debitoparcela.statuslanc = 18)"
        'Sql = Sql & "WHERE (debitoparcela.codreduzido =4547) AND (debitoparcela.anoexercicio = 2017) AND (debitoparcela.codlancamento = 79) AND (debitoparcela.statuslanc = 18)"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                
                If Trim(SubNull(!CPF)) = "" And Trim(SubNull(!Cnpj)) = "" Then
                    DoEvents
                    GoTo PROXIMOCIP
                End If
                
                If Not IsNull(!CPF) Then
                    If Trim(!CPF) <> "" Then
                        sCPFCNPJ = RetornaNumero(!CPF)
                        nTipoDoc = 1
                    Else
                        sCPFCNPJ = !Cnpj
                        nTipoDoc = 2
                    End If
                ElseIf Not IsNull(!Cnpj) Then
                   sCPFCNPJ = !Cnpj
                   nTipoDoc = 2
                End If
                
                ReDim Preserve aBoletos(UBound(aBoletos) + 1)
                'aBoletos(UBound(aBoletos)).sNossoNumero = "2950230" & CStr(!NumDocumento) & CStr(Modulo11(!NumDocumento))
                aBoletos(UBound(aBoletos)).sNossoNumero = "295023000" & CStr(!NumDocumento)
                'aBoletos(UBound(aBoletos)).sNumDocumento = CStr(!NumDocumento) & CStr(RetornaDVNumDoc(!NumDocumento))
                aBoletos(UBound(aBoletos)).sNumDocumento = CStr(!NumDocumento)
                aBoletos(UBound(aBoletos)).sDataVencimento = Format(!DataVencimento, "dd") & Format(!DataVencimento, "mm") & Format(!DataVencimento, "yyyy")
                aBoletos(UBound(aBoletos)).sValorNominal = FillLeft(RetornaNumero(CStr(!ValorTributo * 100)), 15)
                aBoletos(UBound(aBoletos)).sDataBase = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
                aBoletos(UBound(aBoletos)).nTipoInscricao = CStr(nTipoDoc)
                aBoletos(UBound(aBoletos)).nNumeroInscricao = sCPFCNPJ
                aBoletos(UBound(aBoletos)).sNome = Left(!nomecidadao, 40)
                aBoletos(UBound(aBoletos)).sEndereco = Left(!Logradouro & ", " & CStr(!Li_Num) & " " & SubNull(!Li_Compl), 40)
                aBoletos(UBound(aBoletos)).sBairro = Left(SubNull(!DescBairro), 15)
                sCep = RetornaCEP(!CodLogr, !Li_Num)
                If Len(Trim(sCep)) < 5 Then
                    sCep = 14870000
                End If
                aBoletos(UBound(aBoletos)).sCep = Left(sCep, 5)
                aBoletos(UBound(aBoletos)).sSufixoCep = Right(sCep, 3)
                aBoletos(UBound(aBoletos)).sCidade = "JABOTICABAL"
                aBoletos(UBound(aBoletos)).sUF = "SP"
                
PROXIMOCIP:
               .MoveNext
            Loop
           .Close
        End With
    Else
    End If
End If


End Sub

Private Sub GeraArquivo(nTipo As Integer)

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
    If nTipo < 3 Then
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
    .nDataGeracao = "25042018"
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
    If nTipo < 3 Then
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
    .sDataGeracao = "25042018"
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
For nPosReg = 1 To UBound(aBoletos)
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
        If nTipo < 3 Or nTipo = 1 Or nTipo = 2 Then
            .nConta = FillLeft("74000", 12)
            .sDvConta = "4 "
        Else
            .nConta = FillLeft("34692", 12)
            .sDvConta = "6 "
        End If
        .sNossoNumero = FillSpace(aBoletos(nPosReg).sNossoNumero, 20)
        
        
        .nCodCarteira = "7"
        .nFormaCadastro = "1"
        .sTipoDocumento = "1"
        .nIdentificacaoEmissao = "2"
        .sIdentificacaoDistribuicao = "2"
        .sNumeroDocumento = FillSpace(aBoletos(nPosReg).sNumDocumento, 15)
        .nDataVencimento = aBoletos(nPosReg).sDataVencimento
        .nValorNominal = FillLeft(aBoletos(nPosReg).sValorNominal, 15)
        .nAgenciaCobranca = "00000"
        .sDvAgenciaCobranca = "0"
        .nEspecieTitulo = "01"
        .sAceite = "N"
        .nDataEmissao = aBoletos(nPosReg).sDataBase
        .nCodigoJuros = "0"
        .nDataJuros = FillLeft("0", 8)
        .nJurosMora = FillLeft("0", 15)
        .nCodigoDesconto1 = "0"
        .nDataDesconto1 = FillLeft("0", 8)
        .nValorConcedido = FillLeft("0", 15)
        .nValorIOF = FillLeft("0", 15)
        .nValorAbatimento = FillLeft("0", 15)
        .sIdentificaTitulo = FillSpace(aBoletos(nPosReg).sNumDocumento, 25)
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
        On Error Resume Next
        Sql = "insert registro_cobranca (numdocumento,dataregistro) values(" & CLng(Left(Trim(.sNumeroDocumento), 8)) & ",'" & Format(Now, "mm/dd/yyyy hh:mm") & "')"
        cn.Execute Sql, rdExecDirect
        
        
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
        .nTipoInscricao = aBoletos(nPosReg).nTipoInscricao
        .nNumeroInscricao = FillLeft(aBoletos(nPosReg).nNumeroInscricao, 15)
        .sNome = FillSpace(aBoletos(nPosReg).sNome, 40)
        .sEndereco = FillSpace(aBoletos(nPosReg).sEndereco, 40)
        .sBairro = FillSpace(aBoletos(nPosReg).sBairro, 15)
        If Len(aBoletos(nPosReg).sCep) < 5 Then
            .nCep = "00000"
            .nCepsufixo = "000"
        Else
            .nCep = aBoletos(nPosReg).sCep
            .nCepsufixo = aBoletos(nPosReg).sSufixoCep
        End If
        
        .sCidade = FillSpace(aBoletos(nPosReg).sCidade, 15)
        .sUF = aBoletos(nPosReg).sUF
        .nipoInscricaoSacado = "0"
        .nNumeroInscricaoSacado = FillLeft("0", 15)
        .sNomeSacado = FillSpace(" ", 40)
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
nQtdeRegistroLote = (UBound(aBoletos) * 2) + 2
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
PBar.value = 0
Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='COBRANCA'"
cn.Execute Sql, rdExecDirect

ret = Shell("C:\Program Files\PSPad editor\pspad.exe" & " " & sArquivo, vbNormalFocus)

AtualizaRemessa

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If
DoEvents
Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

