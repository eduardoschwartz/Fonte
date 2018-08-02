VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmRelAjuiza 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Ajuizamento"
   ClientHeight    =   3585
   ClientLeft      =   4995
   ClientTop       =   4020
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   3780
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1800
      Width           =   1035
   End
   Begin VB.TextBox txtAno2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1035
   End
   Begin VB.TextBox txtAno1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1035
   End
   Begin VB.TextBox txtCod2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   1035
   End
   Begin VB.TextBox txtCod1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   1035
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   750
      TabIndex        =   5
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   3120
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmRelAjuiza.frx":0000
      PICN            =   "frmRelAjuiza.frx":001C
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
      Height          =   345
      Left            =   1890
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmRelAjuiza.frx":0176
      PICN            =   "frmRelAjuiza.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar PbF 
      Height          =   240
      Left            =   645
      TabIndex        =   11
      Top             =   2760
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Ajuizamento:"
      Height          =   255
      Left            =   540
      TabIndex        =   14
      Top             =   1830
      Width           =   1395
   End
   Begin VB.Label lblPF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   60
      TabIndex        =   13
      Top             =   2775
      Width           =   390
   End
   Begin VB.Label lblTotF 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   3015
      TabIndex        =   12
      Top             =   2775
      Width           =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Final.............:"
      Height          =   255
      Left            =   540
      TabIndex        =   10
      Top             =   1470
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Inicial...........:"
      Height          =   255
      Left            =   540
      TabIndex        =   9
      Top             =   1110
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Final........:"
      Height          =   255
      Left            =   540
      TabIndex        =   8
      Top             =   750
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Inicial......:"
      Height          =   255
      Left            =   540
      TabIndex        =   7
      Top             =   390
      Width           =   1395
   End
End
Attribute VB_Name = "frmRelAjuiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xImovel As clsImovel

Private Sub cmdPrint_Click()
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, RdoG As rdoResultset, RdoD As rdoResultset
Dim nNumParc As Integer, nCompl As Integer, nValorTotal As Double, ax As Integer
Dim nValorLancado As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double
Dim nSomaLancado As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double, RdoSP As rdoResultset
Dim sDataVencto As String, nInscricao As Integer, dDataInscricao As Date, nLivro As Integer, nPagina As Integer, sValorTotal As String
Dim sNome As String, sEND1 As String, sCOMPL1 As String, sBAIRRO1 As String, sCIDADE1 As String
Dim sCEP1  As String, sUF1   As String, sInscricao As String, sEND2 As String, sCOMPL2 As String, sBAIRRO2 As String
Dim sQuadra As String, sLote As String, sCIDADE2 As String, sCEP2 As String, sUF2 As String, qd As New rdoQuery, nCodTributo As Integer, aCodTrib() As Integer
Dim sLANCAMENTO As String, sDOCUMENTO As String, aAno() As Integer, nNumRec As Long, RdoAux As rdoResultset, sTributo As String

If txtData.Text <> "" Then
    If Not IsDate(txtData.Text) Then
        MsgBox "Data Inválida.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If

If Val(txtCod1.Text) = 0 Or Val(txtCod2.Text) = 0 Then
    MsgBox "Digite código inicial e final.", vbExclamation, "Atenção"
    Exit Sub
End If

ReDim aAno(0)
Sql = "DELETE FROM RELATORIOAJUIZAMENTO"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM RELATORIOAJUIZAMENTODETALHE"
cn.Execute Sql, rdExecDirect

Set qd.ActiveConnection = cn

Sql = "SELECT DISTINCT codreduzido From debitoparcela WHERE  statuslanc=3 AND "
Sql = Sql & "(CODREDUZIDO BETWEEN " & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text) & ") "
'Sql = Sql & "(CODREDUZIDO BETWEEN " & Val(txtCod1.text) & " AND " & Val(txtCod2.text) & ")  AND DATAAJUIZA='12/11/2008' "
If Not IsDate(txtData.Text) Then
    Sql = Sql & "AND DATAAJUIZA IS NOT NULL"
Else
    Sql = Sql & "AND DATAAJUIZA='" & Format(txtData.Text, "mm/dd/yyyy") & "'"
End If
Sql = Sql & " order by codreduzido"
Set RdoG = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
lblTotF.Caption = RdoG.RowCount
nNumRec = lblTotF.Caption
Do Until RdoG.EOF

    'GAUGE
'    If xId Mod 100 = 0 Then
       CallPb RdoG.AbsolutePosition, nNumRec
 '   End If

    nCodReduz = RdoG!CODREDUZIDO
    sNome = ""
    sEND1 = ""
    sCOMPL1 = ""
    sBAIRRO1 = ""
    sCIDADE1 = ""
    sCEP1 = ""
    sUF1 = ""
    sInscricao = ""
    sEND2 = ""
    sCOMPL2 = ""
    sBAIRRO2 = ""
    sQuadra = ""
    sLote = ""
    sCIDADE2 = ""
    sCEP2 = ""
    sUF2 = ""
    nValorTotal = "2.000,54"
    'sDOCUMENTO = "254.356.784-85"
    
    If nCodReduz < 100000 Then 'IMOVEL
        With xImovel
           .CarregaImovel nCodReduz
            sInscricao = .Inscricao
            sQuadra = .Li_Quadras
            sLote = .Li_Lotes
            sNome = .NomePropPrincipal
            sEND1 = .EnderecoCompleto
            sCOMPL1 = .Li_Compl
            sBAIRRO1 = .DescBairro
            sCIDADE1 = "JABOTICABAL"
            sCEP1 = RetornaCEP(.CodLogr, .Li_Num)
            If Len(sCEP1) <> 9 Then
                sCEP1 = "00000-000"
            End If
            sUF1 = "SP"
            Sql = "SELECT ENDENTREGA.CODREDUZIDO,ENDENTREGA.EE_CODLOG, ENDENTREGA.EE_NOMELOG,ENDENTREGA.EE_NUMIMOVEL,"
            Sql = Sql & "ENDENTREGA.EE_COMPLEMENTO, ENDENTREGA.EE_UF,ENDENTREGA.EE_CIDADE, ENDENTREGA.EE_BAIRRO,"
            Sql = Sql & "ENDENTREGA.EE_CEP, ENDENTREGA.EE_LOTEAMENTO,ENDENTREGA.EE_DESCBAIRRO , Cidade.DESCCIDADE "
            Sql = Sql & "FROM ENDENTREGA INNER JOIN  CIDADE ON ENDENTREGA.EE_UF = CIDADE.SIGLAUF AND "
            Sql = Sql & "ENDENTREGA.Ee_Cidade = Cidade.CODCIDADE WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                If RdoAux2!Ee_NomeLog <> "" Then
                    sEND2 = RdoAux2!Ee_NomeLog & " Nº " & CStr(RdoAux2!Ee_NumImovel)
                    sCOMPL2 = RdoAux2!Ee_Complemento
                    sBAIRRO2 = SubNull(RdoAux2!Ee_DESCBairro)
                    sCIDADE2 = SubNull(RdoAux2!descCidade)
                    sCEP2 = Format(RdoAux2!Ee_Cep, "00000-000")
                    sUF2 = SubNull(RdoAux2!Ee_Uf)
                Else
                    sEND2 = sEND1 '205-250
                    sCOMPL2 = sCOMPL1 '251-270
                    sBAIRRO2 = sBAIRRO1
                    sCIDADE2 = sCIDADE1
                    sCEP2 = sCEP1
                    sUF2 = sUF1
                End If
            Else
                sEND2 = sEND1 '205-250
                sCOMPL2 = sCOMPL1 '251-270
                sBAIRRO2 = sBAIRRO1
                sCIDADE2 = sCIDADE1
                sCEP2 = sCEP1
                sUF2 = sUF1
            End If
            RdoAux2.Close
        End With
    ElseIf nCodReduz > 100000 And nCodReduz < 500000 Then 'EMPRESA
        Sql = "SELECT CODIGOMOB,DVMOB,RAZAOSOCIAL,NOMEFANTASIA,CNPJ,CPF,INSCESTADUAL,"
        Sql = Sql & "DATAABERTURA,NUMPROCESSO,DATAPROCESSO,ATIVEXTENSO,SIGLAUF,CODCIDADE,CODBAIRRO,"
        Sql = Sql & "DESCCIDADE,DESCBAIRRO,DESCUF,CODLOGRADOURO,NOMELOGR,"
        Sql = Sql & "NUMERO,COMPLEMENTO,CEP,DATAENCERRAMENTO,NUMPROCENCERRAMENTO,DATAPROCENCERRAMENTO,"
        Sql = Sql & "HORARIO,DESCHORARIO,HOMEPAGE,NOMECONTATO,FONECONTATO,FAXCONTATO,CARGOCONTATO,"
        Sql = Sql & "EMAILCONTATO,RESPCONTABIL,CAPITALSOCIAL,QTDEEMPREGADO,QTDEPROF,RG,ORGAO "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInscricao = Format(!codigomob, "0000000") & "-" & !DVMOB
            sQuadra = ""
            sLote = ""
            sNome = !razaosocial
            If Not IsNull(!CPF) Then
               sDOCUMENTO = Format(Trim(!CPF), "000\.000\.000-00")
            ElseIf Not IsNull(!Cnpj) Then
               sDOCUMENTO = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
            ElseIf Not IsNull(!rg) Then
               sDOCUMENTO = !rg
            Else
               sDOCUMENTO = ""
            End If
    
            Sql = "SELECT CODIGOMOB,INSCESTADUAL,CNPJ,CPF,RAZAOSOCIAL,CODLOGRADOURO,DESCCIDADE,SIGLAUF,CODCIDADE,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO,NOMELOGR "
            Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    sNumInsc = !INSCESTADUAL
                    sEND1 = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro) & " nº " & Val(SubNull(!Numero))
                    sCOMPL1 = SubNull(!Complemento)
                    Sql = "SELECT  DESCBAIRRO From BAIRRO WHERE  siglauf = '" & !siglaUF & "' And codcidade = " & !CodCidade & " And codbairro = " & !CodBairro
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        sBAIRRO1 = SubNull(!DescBairro)
                       .Close
                    End With
                    Sql = "SELECT  DESCCIDADE From CIDADE WHERE  siglauf = '" & !siglaUF & "' And codcidade = " & !CodCidade
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        sCIDADE1 = SubNull(!descCidade)
                       .Close
                    End With
                    sCEP1 = RetornaCEP(!CodLogradouro, !Numero)
                    If Len(sCEP1) <> 9 Then
                        sCEP1 = "00000-000"
                    End If
                    sUF1 = SubNull(!siglaUF)
                End If
            End With
            Sql = "SELECT * FROM MOBILIARIOENDENTREGA WHERE CODMOBILIARIO=" & nCodReduz
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                    If .RowCount > 0 Then
                        sEND2 = SubNull(!NomeLogradouro) & " nº " & !NUMIMOVEL
                        sCOMPL2 = SubNull(!Complemento)
                        sUF2 = SubNull(!UF)
                        If !CodCidade > 0 Then
                            Sql = "SELECT  DESCCIDADE From CIDADE WHERE  siglauf = '" & !UF & "' And codcidade = " & !CodCidade
                            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux3
                                sCIDADE2 = SubNull(!descCidade)
                               .Close
                            End With
                        Else
                            sCIDADE2 = SubNull(!descCidade)
                        End If
                        If !CodBairro > 0 Then
                            Sql = "SELECT  DESCBAIRRO From BAIRRO WHERE  siglauf = '" & !UF & "' And codcidade = " & !CodCidade & " And codbairro = " & !CodBairro
                            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux3
                                sBAIRRO2 = SubNull(!DescBairro)
                               .Close
                            End With
                        Else
                            sBAIRRO2 = SubNull(!DescBairro)
                        End If
                        sCEP2 = RetornaCEP(!CodLogradouro, !NUMIMOVEL)
                        If Len(sCEP2) <> 9 Then
                            sCEP2 = "00000-000"
                        End If
                        
                    
                    Else
                        sEND2 = sEND1 '205-250
                        sCOMPL2 = sCOMPL1 '251-270
                        sBAIRRO2 = sBAIRRO1
                        sCIDADE2 = sCIDADE1
                        sCEP2 = sCEP1
                        sUF2 = sUF1
                    End If
            End With
            
        End With
    Else 'OUTROS
        Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, "
        Sql = Sql & "vwLOGRADOURO.ABREVTITLOG, vwLOGRADOURO.NOMELOGRADOURO, cidadao.numimovel, cidadao.complemento, cidadao.codbairro,"
        Sql = Sql & "bairro.descbairro, cidadao.codcidade, cidade.desccidade, cidadao.siglauf, cidadao.cep, cidadao.nomelogradouro AS nomelogradouro2"
        Sql = Sql & " FROM cidadao RIGHT OUTER JOIN bairro ON cidadao.siglauf = bairro.siglauf AND "
        Sql = Sql & "cidadao.codcidade = bairro.codcidade AND cidadao.codbairro = bairro.codbairro INNER JOIN cidade ON bairro.siglauf = cidade.siglauf AND "
        Sql = Sql & "bairro.codcidade = cidade.codcidade LEFT OUTER JOIN  vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO "
        Sql = Sql & "WHERE Cidadao.codcidadao = " & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sInscricao = Format(!CodCidadao, "0000000")
                sNome = !nomecidadao
                sQuadra = ""
                sLote = ""
                If Not IsNull(!NomeLogradouro) Then
                   sEND1 = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & RdoAux!NUMIMOVEL
                Else
                   sEND1 = !NOMELOGRADOURO2 & " Nº " & RdoAux!NUMIMOVEL
                End If
                sCOMPL1 = SubNull(!Complemento)
                sBAIRRO1 = SubNull(!DescBairro)
                sCIDADE1 = SubNull(!descCidade)
                sCEP1 = !Cep
                If Len(sCEP1) <> 9 Then
                    sCEP1 = "00000-000"
                End If
                sUF1 = SubNull(!siglaUF)
            End If
            sEND2 = sEND1 '205-250
            sCOMPL2 = sCOMPL1 '251-270
            sBAIRRO2 = sBAIRRO1
            sCIDADE2 = sCIDADE1
            sCEP2 = sCEP1
            sUF2 = sUF1
        End With
    End If
    
    If Not IsDate(txtData.Text) Then
        Sql = "SELECT * From debitoparcela WHERE CODREDUZIDO=" & nCodReduz
    Else
        Sql = "SELECT * From debitoparcela WHERE CODREDUZIDO=" & nCodReduz & " AND (dataajuiza = '" & Format(txtData.Text, "mm/dd/yyyy") & "')"
    End If
    'Sql = "SELECT * From debitoparcela WHERE CODREDUZIDO=" & nCodReduz & " AND (dataajuiza = '12/11/2008')"
    Set RdoD = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    With RdoD
        Do Until .EOF
            sLANCAMENTO = ""
            nSomaLancado = 0: nSomaJuros = 0: nSomaMulta = 0: nSomaCorrecao = 0
            nAno = !AnoExercicio
            nLanc = !CodLancamento
            nSeq = !SeqLancamento
            nNumParc = !NumParcela
            nCompl = !CODCOMPLEMENTO
            'sDataVencto = .TextMatrix(x, 6)
            
            
            On Error Resume Next
            RdoSP.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = nCodReduz
            qd(2) = nAno
            qd(3) = nAno
            qd(4) = nLanc
            qd(5) = nLanc
            qd(6) = nSeq
            qd(7) = nSeq
            qd(8) = nNumParc
            qd(9) = nNumParc
            qd(10) = nCompl
            qd(11) = nCompl
            qd(12) = 1
            qd(13) = 99
            qd(14) = Format(Now, "mm/dd/yyyy")
            qd(15) = NomeDoUsuario
            Set RdoSP = qd.OpenResultset(rdOpenKeyset)
            With RdoSP
                ReDim aCodTrib(0)
                If .RowCount = 0 Then
                    RdoSP.Close
                    GoTo Proximo
                End If
                Do Until .EOF
                    ReDim Preserve aCodTrib(UBound(aCodTrib) + 1)
                    aCodTrib(UBound(aCodTrib)) = !CodTributo
                    
                    Sql = "INSERT RELATORIOAJUIZAMENTODETALHE (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,"
                    Sql = Sql & "CODCOMPLEMENTO,DATAVENCIMENTO,PRINCIPAL,CORRECAO,MULTA,JUROS,INSCRICAO,DATAINSCRICAO,"
                    Sql = Sql & "LIVRO,PAGINA,VALORTOTAL,LANCAMENTO,TRIBUTO) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & ","
                    Sql = Sql & !CODCOMPLEMENTO & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!ValorTributo)) & ","
                    Sql = Sql & Virg2Ponto(CStr(!VALORCORRECAO)) & "," & Virg2Ponto(CStr(!ValorMulta)) & "," & Virg2Ponto(CStr(!ValorJuros)) & ","
                    If Not IsNull(!Datainscricao) Then
                        Sql = Sql & Val(SubNull(!CERTIDAO)) & ",'" & Format(!Datainscricao, "mm/dd/yyyy") & "'," & Val(SubNull(!NUMLIVRO)) & "," & Val(SubNull(!PAGINA)) & ",'"
                    Else
                        Sql = Sql & Val(SubNull(!CERTIDAO)) & "," & "NULL" & " ," & Val(SubNull(!NUMLIVRO)) & "," & Val(SubNull(!PAGINA)) & ",'"
                    End If
                    Sql = Sql & Virg2Ponto(CStr(!ValorTotal)) & "','" & "" & "','" & !ABREVTRIBUTO & "')"
                    cn.Execute Sql, rdExecDirect
                    sLANCAMENTO = sLANCAMENTO & !ABREVTRIBUTO & ", "
                   .MoveNext
                Loop
               .Close
            End With
            sLANCAMENTO = Left$(sLANCAMENTO, Len(sLANCAMENTO) - 2)
            Sql = "UPDATE RELATORIOAJUIZAMENTODETALHE SET LANCAMENTO='" & sLANCAMENTO & "' WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND "
            Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
            cn.Execute Sql, rdExecDirect
        
            '*** BUSCA O ARTIGO ****
            nCodTributo = 0
            For z = 1 To UBound(aCodTrib)
                Sql = "SELECT * FROM TRIBUTOARTIGO WHERE CODTRIBUTO=" & aCodTrib(z)
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux2.RowCount > 0 Then
                    nCodTributo = aCodTrib(z)
                    RdoAux2.Close
                End If
            Next
            If nCodTributo > 0 Then
                Sql = "UPDATE RELATORIOAJUIZAMENTODETALHE SET CODTRIBUTO=" & nCodTributo & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
                Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND "
                Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                cn.Execute Sql, rdExecDirect
            End If
            '***********************
            .MoveNext
        Loop
    
    
        Sql = "SELECT DISTINCT(ANOEXERCICIO) FROM RELATORIOAJUIZAMENTODETALHE WHERE CODREDUZIDO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aAno(UBound(aAno) + 1)
                aAno(UBound(aAno)) = !AnoExercicio
               .MoveNext
            Loop
           .Close
        End With
        For ax = 1 To UBound(aAno)
            Sql = "SELECT SUM(PRINCIPAL) AS VALORLANC, SUM(CORRECAO) AS VALORCOR, SUM(MULTA) AS VALORMULTA,SUM(JUROS) AS VALORJUROS "
            Sql = Sql & "FROM RELATORIOAJUIZAMENTODETALHE WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & aAno(ax)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If Not IsNull(!VALORLANC) Then
                     nValorTotal = FormatNumber(!VALORLANC + !VALORCOR + !ValorMulta + !ValorJuros, 2)
                    
                    sValorTotal = "R$ " & nValorTotal & " (" & Extenso(nValorTotal) & ")"
                    Sql = "UPDATE RELATORIOAJUIZAMENTODETALHE SET VALORTOTAL='" & sValorTotal & "' WHERE ANOEXERCICIO=" & aAno(ax) & " AND CODREDUZIDO=" & nCodReduz
                    cn.Execute Sql, rdExecDirect
                End If
               .Close
            End With
        Next
    
        Sql = "SELECT SUM(PRINCIPAL) AS VALORLANC, SUM(CORRECAO) AS VALORCOR, SUM(MULTA) AS VALORMULTA,SUM(JUROS) AS VALORJUROS "
        Sql = Sql & "FROM RELATORIOAJUIZAMENTODETALHE WHERE CODREDUZIDO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nValorTotal = FormatNumber(!VALORLANC + !VALORCOR + !ValorMulta + !ValorJuros, 2)
           .Close
        End With
        Sql = "INSERT RELATORIOAJUIZAMENTO (NOME,END1,COMPL1,BAIRRO1,CIDADE1,CEP1,UF1,INSCRICAO,CODREDUZ,"
        Sql = Sql & "END2,COMPL2,BAIRRO2,QUADRA,LOTE,CIDADE2,CEP2,UF2,VALORTOTAL,LANCAMENTO,DOCUMENTO) VALUES('"
        Sql = Sql & Mask(sNome) & "','" & Mask(sEND1) & "','" & Mask(sCOMPL1) & "','" & Mask(sBAIRRO1) & "','" & Mask(sCIDADE1) & "','"
        Sql = Sql & sCEP1 & "','" & sUF1 & "','" & sInscricao & "'," & nCodReduz & ",'" & Mask(sEND2) & "','"
        Sql = Sql & Mask(sCOMPL2) & "','" & Mask(sBAIRRO2) & "','" & sQuadra & "','" & sLote & "','"
        Sql = Sql & Mask(sCIDADE2) & "','" & Left$(sCEP2, 9) & "','" & sUF2 & "'," & Virg2Ponto(CStr(nValorTotal)) & ",'" & Mask(sLANCAMENTO) & "','" & sDOCUMENTO & "')"
        cn.Execute Sql, rdExecDirect
    
    End With
Proximo:
    RdoG.MoveNext
Loop
    
'EXIBE RELATORIO

frmReport.ShowReport "AJUIZAMENTO", frmMdi.hwnd, Me.hwnd

Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(x).Description
Next
Resume Next
        
       
End Sub

Private Sub cmdPrintOld_Click()
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, RdoG As rdoResultset, RdoD As rdoResultset
Dim nNumParc As Integer, nCompl As Integer, nValorTotal As Double, ax As Integer
Dim nValorLancado As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double
Dim nSomaLancado As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double
Dim sDataVencto As String, nInscricao As Integer, dDataInscricao As Date, nLivro As Integer, nPagina As Integer, sValorTotal As String
Dim sNome As String, sEND1 As String, sCOMPL1 As String, sBAIRRO1 As String, sCIDADE1 As String
Dim sCEP1  As String, sUF1   As String, sInscricao As String, sEND2 As String, sCOMPL2 As String, sBAIRRO2 As String
Dim sQuadra As String, sLote As String, sCIDADE2 As String, sCEP2 As String, sUF2 As String
Dim sLANCAMENTO As String, sDOCUMENTO As String, aAno() As Integer, nNumRec As Long, RdoAux As rdoResultset, sTributo As String

ReDim aAno(0)
Sql = "DELETE FROM RELATORIOAJUIZAMENTO"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM RELATORIOAJUIZAMENTODETALHE"
cn.Execute Sql, rdExecDirect

Sql = "SELECT DISTINCT codreduzido From debitoparcela WHERE (dataajuiza = '12/11/2008')"


'Sql = "SELECT DISTINCT CODREDUZIDO From DEBITOPARCELA WHERE "
'Sql = Sql & "(CODREDUZIDO BETWEEN " & Val(txtCod1.text) & " AND " & Val(txtCod2.text) & ") AND "
'Sql = Sql & "(ANOEXERCICIO BETWEEN " & Val(txtAno1.text) & " AND " & Val(txtAno2.text) & ") AND "
'Sql = Sql & "(DATAAJUIZA IS NOT NULL) ORDER BY CODREDUZIDO"
'Sql = Sql & "(DATAAJUIZA='02/01/05') ORDER BY CODREDUZIDO"
Set RdoG = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
lblTotF.Caption = RdoG.RowCount
nNumRec = lblTotF.Caption
Do Until RdoG.EOF

    'GAUGE
'    If xId Mod 100 = 0 Then
       CallPb RdoG.AbsolutePosition, nNumRec
 '   End If

    nCodReduz = RdoG!CODREDUZIDO
    sNome = ""
    sEND1 = ""
    sCOMPL1 = ""
    sBAIRRO1 = ""
    sCIDADE1 = ""
    sCEP1 = ""
    sUF1 = ""
    sInscricao = ""
    sEND2 = ""
    sCOMPL2 = ""
    sBAIRRO2 = ""
    sQuadra = ""
    sLote = ""
    sCIDADE2 = ""
    sCEP2 = ""
    sUF2 = ""
    nValorTotal = "2.000,54"
    sDOCUMENTO = "254.356.784-85"
    
    If nCodReduz < 100000 Then 'IMOVEL
        With xImovel
           .CarregaImovel nCodReduz
            sInscricao = .Inscricao
            sQuadra = .Li_Quadras
            sLote = .Li_Lotes
            sNome = .NomePropPrincipal
            sEND1 = .EnderecoCompleto
            sCOMPL1 = .Li_Compl
            sBAIRRO1 = .DescBairro
            sCIDADE1 = "JABOTICABAL"
            sCEP1 = RetornaCEP(.CodLogr, .Li_Num)
            If Len(sCEP1) <> 9 Then
                sCEP1 = "00000-000"
            End If
            sUF1 = "SP"
            Sql = "SELECT ENDENTREGA.CODREDUZIDO,ENDENTREGA.EE_CODLOG, ENDENTREGA.EE_NOMELOG,ENDENTREGA.EE_NUMIMOVEL,"
            Sql = Sql & "ENDENTREGA.EE_COMPLEMENTO, ENDENTREGA.EE_UF,ENDENTREGA.EE_CIDADE, ENDENTREGA.EE_BAIRRO,"
            Sql = Sql & "ENDENTREGA.EE_CEP, ENDENTREGA.EE_LOTEAMENTO,ENDENTREGA.EE_DESCBAIRRO , Cidade.DESCCIDADE "
            Sql = Sql & "FROM ENDENTREGA INNER JOIN  CIDADE ON ENDENTREGA.EE_UF = CIDADE.SIGLAUF AND "
            Sql = Sql & "ENDENTREGA.Ee_Cidade = Cidade.CODCIDADE WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                If RdoAux2!Ee_NomeLog <> "" Then
                    sEND2 = RdoAux2!Ee_NomeLog & " Nº " & CStr(RdoAux2!Ee_NumImovel)
                    sCOMPL2 = RdoAux2!Ee_Complemento
                    sBAIRRO2 = SubNull(RdoAux2!Ee_DESCBairro)
                    sCIDADE2 = SubNull(RdoAux2!descCidade)
                    sCEP2 = Format(RdoAux2!Ee_Cep, "00000-000")
                    sUF2 = SubNull(RdoAux2!Ee_Uf)
                Else
                    sEND2 = sEND1 '205-250
                    sCOMPL2 = sCOMPL1 '251-270
                    sBAIRRO2 = sBAIRRO1
                    sCIDADE2 = sCIDADE1
                    sCEP2 = sCEP1
                    sUF2 = sUF1
                End If
            Else
                sEND2 = sEND1 '205-250
                sCOMPL2 = sCOMPL1 '251-270
                sBAIRRO2 = sBAIRRO1
                sCIDADE2 = sCIDADE1
                sCEP2 = sCEP1
                sUF2 = sUF1
            End If
            RdoAux2.Close
        End With
    ElseIf nCodReduz > 100000 And nCodReduz < 500000 Then 'EMPRESA
        Sql = "SELECT CODIGOMOB,DVMOB,RAZAOSOCIAL,NOMEFANTASIA,CNPJ,CPF,INSCESTADUAL,"
        Sql = Sql & "DATAABERTURA,NUMPROCESSO,DATAPROCESSO,ATIVEXTENSO,SIGLAUF,CODCIDADE,CODBAIRRO,"
        Sql = Sql & "DESCCIDADE,DESCBAIRRO,DESCUF,CODLOGRADOURO,NOMELOGR,"
        Sql = Sql & "NUMERO,COMPLEMENTO,CEP,DATAENCERRAMENTO,NUMPROCENCERRAMENTO,DATAPROCENCERRAMENTO,"
        Sql = Sql & "HORARIO,DESCHORARIO,HOMEPAGE,NOMECONTATO,FONECONTATO,FAXCONTATO,CARGOCONTATO,"
        Sql = Sql & "EMAILCONTATO,RESPCONTABIL,CAPITALSOCIAL,QTDEEMPREGADO,QTDEPROF,RG,ORGAO "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInscricao = Format(!codigomob, "0000000") & "-" & !DVMOB
            sQuadra = ""
            sLote = ""
            sNome = !razaosocial
            Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
            Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
            Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
            Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    sEND1 = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & RdoAux!Numero
                End If
               .Close
            End With
            sCOMPL1 = SubNull(!Complemento)
            sBAIRRO1 = SubNull(!DescBairro)
            sCIDADE1 = SubNull(!descCidade)
            sCEP1 = !Cep
            If Len(sCEP1) <> 9 Then
                sCEP1 = "00000-000"
            End If
            sUF1 = SubNull(!siglaUF)
        End With
    Else 'OUTROS
    
    End If
    
    '****** SEGUNDA PARTE
    
    Sql = "SELECT * From DEBITOPARCELA WHERE "
    Sql = Sql & "CODREDUZIDO=" & nCodReduz & " AND "
    Sql = Sql & "(ANOEXERCICIO BETWEEN " & Val(txtAno1.Text) & " AND " & Val(txtAno2.Text) & ") AND "
    Sql = Sql & "(DATAAJUIZA IS NOT NULL) AND NUMPARCELA>0 AND STATUSLANC=3"
    Set RdoD = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    
    Do Until RdoD.EOF
        sTributo = ""
        nSomaLancado = 0: nSomaJuros = 0: nSomaMulta = 0: nSomaCorrecao = 0
        nAno = RdoD!AnoExercicio
        nLanc = RdoD!CodLancamento
        nSeq = RdoD!SeqLancamento
        nNumParc = RdoD!NumParcela
        nCompl = RdoD!CODCOMPLEMENTO
        sDataVencto = RdoD!DataVencimento
    
        Sql = "SELECT DEBITOPARCELA.*, tributo.abrevtributo AS NOMETRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,LANCAMENTO.DESCFULL FROM DEBITOPARCELA LEFT OUTER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO AND "
        Sql = Sql & "DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO AND "
        Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND DEBITOPARCELA.NUMPARCELA = DEBITOTRIBUTO.NUMPARCELA AND "
        Sql = Sql & "DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO INNER JOIN LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
        Sql = Sql & "INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
        Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & nCodReduz & " AND DEBITOPARCELA.ANOEXERCICIO=" & nAno & " AND DEBITOPARCELA.CODLANCAMENTO=" & nLanc & " AND "
        Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO=" & nSeq & " AND DEBITOPARCELA.NUMPARCELA=" & nNumParc & " AND "
        Sql = Sql & "DEBITOPARCELA.CODCOMPLEMENTO=" & nCompl & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
        With RdoAux
            Do Until .EOF
                nValorLancado = RdoAux!ValorTributo
                nValorCorrecao = CalculaCorrecao(nValorLancado, RdoAux!DataVencimento)
                nValorJuros = CDbl(CalculaJuros(nValorLancado + nValorCorrecao, RdoAux!DataVencimento))
                nValorMulta = CDbl(CalculaMulta(nValorLancado + nValorCorrecao, RdoAux!DataVencimento))
                nSomaLancado = nSomaLancado + nValorLancado
                nSomaJuros = nSomaJuros + nValorJuros
                nSomaMulta = nSomaMulta + nValorMulta
                nSomaCorrecao = nSomaCorrecao + nValorCorrecao
                nInscricao = Val(SubNull(RdoAux!NUMCERTIDAO))
                If Not IsNull(RdoAux!Datainscricao) Then
                   dDataInscricao = RdoAux!Datainscricao
                Else
                   dDataInscricao = Format("01/01/2000", "dd/mm/yyyy")
                End If
                nPagina = Val(SubNull(RdoAux!PAGINALIVRO))
                nLivro = Val(SubNull(RdoAux!NUMEROLIVRO))
                sLANCAMENTO = RdoAux!DESCFULL
                sTributo = sTributo & !NOMETRIBUTO & ","
                RdoAux.MoveNext
            Loop
            sTributo = "(" & Left$(sTributo, Len(sTributo) - 1) & ")"
            sLANCAMENTO = sLANCAMENTO & " " & sTributo
            
            sValorTotal = ""
            Sql = "INSERT RELATORIOAJUIZAMENTODETALHE (CODREDUZIDO,ANOEXERCICIO,SEQLANCAMENTO,NUMPARCELA,"
            Sql = Sql & "CODCOMPLEMENTO,DATAVENCIMENTO,PRINCIPAL,CORRECAO,MULTA,JUROS,INSCRICAO,DATAINSCRICAO,"
            Sql = Sql & "LIVRO,PAGINA,VALORTOTAL) VALUES(" & nCodReduz & "," & nAno & "," & nSeq & "," & nNumParc & ","
            Sql = Sql & nCompl & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nSomaLancado)) & ","
            Sql = Sql & Virg2Ponto(CStr(nSomaCorrecao)) & "," & Virg2Ponto(CStr(nSomaMulta)) & "," & Virg2Ponto(CStr(nSomaJuros)) & ","
            Sql = Sql & nInscricao & ",'" & Format(dDataInscricao, "mm/dd/yyyy") & "'," & nLivro & "," & nPagina & ",'"
            Sql = Sql & sValorTotal & "')"
            cn.Execute Sql, rdExecDirect
            
           .Close
        End With
        
        RdoD.MoveNext
    Loop
    
       
    Sql = "SELECT DISTINCT(ANOEXERCICIO) FROM RELATORIOAJUIZAMENTODETALHE WHERE CODREDUZIDO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = !AnoExercicio
           .MoveNext
        Loop
       .Close
    End With
    For ax = 1 To UBound(aAno)
        Sql = "SELECT SUM(PRINCIPAL) AS VALORLANC, SUM(CORRECAO) AS VALORCOR, SUM(MULTA) AS VALORMULTA,SUM(JUROS) AS VALORJUROS "
        Sql = Sql & "FROM RELATORIOAJUIZAMENTODETALHE WHERE ANOEXERCICIO=" & aAno(ax) & " AND CODREDUZIDO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If Not IsNull(!VALORLANC) Then
                nValorTotal = FormatNumber(!VALORLANC + !VALORCOR + !ValorMulta + !ValorJuros, 2)
            End If
           .Close
           sValorTotal = "R$ " & nValorTotal & " (" & Extenso(nValorTotal) & ")"
           Sql = "UPDATE RELATORIOAJUIZAMENTODETALHE SET VALORTOTAL='" & sValorTotal & "' WHERE ANOEXERCICIO=" & aAno(ax) & " AND CODREDUZIDO=" & nCodReduz
           cn.Execute Sql, rdExecDirect
        End With
    Next

    nValorTotal = 0
    Sql = "SELECT SUM(PRINCIPAL) AS VALORLANC, SUM(CORRECAO) AS VALORCOR, SUM(MULTA) AS VALORMULTA,SUM(JUROS) AS VALORJUROS "
    Sql = Sql & "FROM RELATORIOAJUIZAMENTODETALHE WHERE CODREDUZIDO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If Not IsNull(!VALORLANC) Then
            nValorTotal = FormatNumber(!VALORLANC + !VALORCOR + !ValorMulta + !ValorJuros, 2)
        Else
            nValorTotal = 0
        End If
       .Close
    End With
    Sql = "INSERT RELATORIOAJUIZAMENTO (NOME,END1,COMPL1,BAIRRO1,CIDADE1,CEP1,UF1,INSCRICAO,CODREDUZ,"
    Sql = Sql & "END2,COMPL2,BAIRRO2,QUADRA,LOTE,CIDADE2,CEP2,UF2,VALORTOTAL,LANCAMENTO,DOCUMENTO) VALUES('"
    Sql = Sql & Left(Mask(sNome), 50) & "','" & sEND1 & "','" & sCOMPL1 & "','" & sBAIRRO1 & "','" & sCIDADE1 & "','"
    Sql = Sql & sCEP1 & "','" & sUF1 & "','" & sInscricao & "'," & nCodReduz & ",'" & Mask(sEND2) & "','"
    Sql = Sql & sCOMPL2 & "','" & sBAIRRO2 & "','" & Mask(sQuadra) & "','" & sLote & "','"
    Sql = Sql & sCIDADE2 & "','" & Left$(sCEP2, 9) & "','" & sUF2 & "'," & Virg2Ponto(CStr(nValorTotal)) & ",'" & sLANCAMENTO & "','" & sDOCUMENTO & "')"
    cn.Execute Sql, rdExecDirect
    
    
    '****** FIM *******
    
    RdoG.MoveNext
Loop

frmReport.ShowReport "AJUIZAMENTO", frmMdi.hwnd, Me.hwnd

Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(x).Description
Next
Resume Next
        
       
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Set xImovel = New clsImovel
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   PbF.Value = (nPosF * 100) / nTotal
Else
   PbF.Value = 100
End If
lblPF.Caption = nPosF

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

