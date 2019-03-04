VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDataBase 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Base"
   ClientHeight    =   3690
   ClientLeft      =   4965
   ClientTop       =   2835
   ClientWidth     =   3855
   Icon            =   "frmDataBase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   3855
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   3300
      Width           =   1035
      _ExtentX        =   1826
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDataBase.frx":08CA
      PICN            =   "frmDataBase.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.MonthView Mv 
      Height          =   3210
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   5662
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      ShowToday       =   0   'False
      StartOfWeek     =   159973377
      TitleBackColor  =   192
      TitleForeColor  =   16777215
      TrailingForeColor=   12632256
      CurrentDate     =   42736
      MaxDate         =   43830
      MinDate         =   42736
   End
   Begin VB.Label lblDB 
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
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   4080
      Width           =   2955
   End
End
Attribute VB_Name = "frmDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sData As String
Dim sOldData As String
Dim Sql As String

Private Sub cmdSair_Click()

If sOldData <> sData Then
    If MsgBox("Deseja alterar a Data Base para " & sData, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
       Sql = "UPDATE PARAMETROS SET VALPARAM='" & sData & "' WHERE NOMEPARAM='DATABASE'"
       cn.Execute Sql, rdExecDirect
       frmMdi.Sbar.Panels(6).Text = "Data Base: " & sData
    End If
    Ocupado
    Sql = "UPDATE PROCESSOGTI SET DATAARQUIVA='" & Format(Now, "mm/dd/yyyy") & "' "
    Sql = Sql & " Where (FISICO = 0) And (DATAARQUIVA Is Null) And (DATACANCEL Is Null) And (DATASUSPENSO Is Null)"
    cn.Execute Sql, rdExecDirect
    Liberado
End If
If frmMdi.frTeste.Visible = False Then
    AtualizaIntegrativa
End If

Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
sOldData = Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4)
sData = sOldData
End Sub

Private Sub Mv_DateClick(ByVal DateClicked As Date)

sData = CStr(Format(Mv.Day, "00") & "/" & Format(Mv.Month, "00") & "/" & Mv.Year)
lblDB.Caption = "Data Base: " & sData
End Sub

Public Sub AtualizaIntegrativa()


Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, Y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String, sObs As String
Dim clsImovel As New clsImovel, nCodReduz As Long, sSetor As String, sRG As String, dDataPrimeiraParc As String, nValorTotalHon As Double, RdoAux3 As rdoResultset
Dim nPagina As Integer, nLivro As Integer, xImovel As clsImovel, nQtdeParc As Integer, RdoAux4 As rdoResultset, bCancelado As Boolean, sStatus As String, nTotalRec As Long
Dim Data1 As Date, Data2 As Date, nNumCertidao As Long, dDataInscricao As Date, sNome As String, sCepImovel As String, sCidadeImovel As String, sUFImovel As String
Dim sQuadra As String, sLote As String, nId As Long, sFone As String, sEmail As String
Dim Sql As String, RdoAux As rdoResultset, RdoAcordo As rdoResultset, RdoAux5 As rdoResultset
Dim nPos As Integer, sNumProc As String, nNumproc As Long, nAnoproc As Integer

Data1 = Now - 6
Data2 = Now

If frmMdi.frTeste.Visible = True Then Exit Sub
Set xImovel = New clsImovel
Ocupado
ConectaIntegrativa
Sql = "select * from processoreparc wHERE (dataprocesso between '" & Format(Data1, "mm/dd/yyyy") & "' and '" & Format(Data2, "mm/dd/yyyy") & "') ORDER BY numprocesso"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotalRec = .RowCount
    Do Until .EOF
        nCodReduz = !CODIGORESP
        sNumProc = !numprocesso
      '  If sNumProc = "12673/2016" Then MsgBox "teste"
        bCancelado = !Cancelado
        nQtdeParc = !qtdeparcela
        dDataProc = !DATAPROCESSO
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        
        Sql = "SELECT origemreparc.numprocesso, debitoparcela.codreduzido, debitoparcela.dataajuiza "
        Sql = Sql & "FROM debitoparcela INNER JOIN origemreparc ON debitoparcela.codreduzido = origemreparc.codreduzido AND debitoparcela.anoexercicio = origemreparc.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = origemreparc.codlancamento AND debitoparcela.seqlancamento = origemreparc.numsequencia AND "
        Sql = Sql & "debitoparcela.NumParcela = origemreparc.NumParcela And debitoparcela.CODCOMPLEMENTO = origemreparc.CODCOMPLEMENTO "
        Sql = Sql & "WHERE origemreparc.numprocesso = '" & sNumProc & "' AND debitoparcela.codreduzido = " & nCodReduz
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux3!dataajuiza) Then
            RdoAux3.Close
            GoTo proximo
        Else
            RdoAux3.Close
        End If
        
        Sql = "select * from acordos where idacordo=" & nNumproc & " and anoacordo=" & nAnoproc
        Set RdoAux3 = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If RdoAux3.RowCount = 0 Then
                
            
                'ENDEREÇO DO CONTRIBUINTE
                Select Case nCodReduz
                    Case 1 To 99999
                        sTipoImposto = "REPARCEL."
                        sSetor = "IMOBILIÁRIO"
                        xImovel.CarregaImovel nCodReduz
                        sNumInsc = xImovel.Inscricao
                        sCodReduz = nCodReduz
                        sNomeResp = xImovel.NomePropPrincipal
                        sNome = sNomeResp
                        sQuadra = xImovel.Li_Quadras
                        sLote = xImovel.Li_Lotes
                        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
                        sEndImovel = xImovel.Endereco
                        nNumImovel = xImovel.Numero
                        sComplImovel = xImovel.Complemento
                        sBairroImovel = xImovel.Bairro
                        sCepImovel = xImovel.Cep
                        sCidadeImovel = xImovel.Cidade
                        sUFImovel = xImovel.UF
                        
                        sEndEntrega = xImovel.Ee_NomeLog
                        nNumEntrega = xImovel.Ee_NumImovel
                        sComplEntrega = xImovel.Ee_Complemento
                        sBairroEntrega = xImovel.Ee_Bairro
                        sCidadeEntrega = "JABOTICABAL"
                        sUFEntrega = "SP"
                        sCepEntrega = xImovel.Ee_Cep
                        Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg "
                        Sql = Sql & "FROM cidadao INNER JOIN proprietario ON cidadao.codcidadao = proprietario.codcidadao "
                        Sql = Sql & "WHERE(proprietario.codreduzido = " & nCodReduz & ") AND (proprietario.tipoprop = 'P') AND (proprietario.principal = 1)"
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                sCPF = SubNull(!CPF)
                                If Trim(sCPF) = "" Then
                                   sCPF = SubNull(!Cnpj)
                                End If
                             Else
                                sCPF = ""
                             End If
                             sRG = SubNull(!rg)
                            .Close
                        End With
                
                    Case 100000 To 500000
                        sSetor = "MOBILIÁRIO"
                        sTipoImposto = "REPARCEL."
                        sNomeResp = xImovel.NomePropPrincipal
                        sNumInsc = nCodReduz
                        sCodReduz = nCodReduz
                        sLote = ""
                        sQuadra = ""
                        
                        xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
                        sEndImovel = xImovel.Endereco
                        nNumImovel = xImovel.Numero
                        sComplImovel = xImovel.Complemento
                        sBairroImovel = xImovel.Bairro
                        sCepImovel = xImovel.Cep
                        sCidadeImovel = xImovel.Cidade
                        sUFImovel = xImovel.UF
                        
                        sEndEntrega = xImovel.Ee_NomeLog
                        nNumEntrega = xImovel.Ee_NumImovel
                        sComplEntrega = xImovel.Ee_Complemento
                        sBairroEntrega = xImovel.Bairro
                        sCidadeEntrega = xImovel.Cidade
                        sUFEntrega = xImovel.UF
                        sCepEntrega = xImovel.Ee_Cep
                        Sql = "SELECT codigomob, inscestadual,razaosocial, cnpj, cpf From mobiliario WHERE codigomob = " & nCodReduz
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                sNomeResp = !RazaoSocial
                                sCPF = SubNull(!CPF)
                                If Trim(sCPF) = "" Then
                                   sCPF = SubNull(!Cnpj)
                                End If
                             Else
                                sCPF = ""
                             End If
                             sRG = SubNull(!INSCESTADUAL)
                            .Close
                        End With
                        
                    Case 500000 To 800000
                        sSetor = "TAXAS DIVERSAS"
                        sTipoImposto = "REPARCEL."
                        sNomeResp = xImovel.NomePropPrincipal
                        sNome = sNomeResp
                        sNumInsc = nCodReduz
                        sCodReduz = nCodReduz
                        sLote = ""
                        sQuadra = ""
                        
                        xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                        sEndImovel = xImovel.Endereco
                        nNumImovel = Val(xImovel.Numero)
                        sComplImovel = xImovel.Complemento
                        sBairroImovel = xImovel.Bairro
                        sCepImovel = xImovel.Cep
                        sCidadeImovel = xImovel.Cidade
                        sUFImovel = xImovel.UF
                        
                        sEndEntrega = sEndImovel
                        nNumEntrega = nNumImovel
                        sComplEntrega = sComplImovel
                        sBairroEntrega = sBairroImovel
                        sCidadeEntrega = xImovel.Cidade
                        sUFEntrega = xImovel.UF
                        sCepEntrega = xImovel.Cep
                        
                        Sql = "SELECT codcidadao,nomecidadao,cpf,cnpj,rg from cidadao WHERE CODCIDADAO=" & nCodReduz
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                sCPF = SubNull(!CPF)
                                If Trim(sCPF) = "" Then
                                   sCPF = SubNull(!Cnpj)
                                End If
                             Else
                                sCPF = ""
                             End If
                             sRG = SubNull(!rg)
                            .Close
                        End With
                End Select
                
                
                Sql = "SELECT * FROM debitoparcela WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND "
                Sql = Sql & "(debitoparcela.numparcela = 1) AND (debitoparcela.numprocesso = '" & sNumProc & "')"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    dDataPrimeiraParc = !DataVencimento
                    nAno = !AnoExercicio
                    nLanc = !CodLancamento
                    nSeq = !SeqLancamento
                    nParc = !NumParcela
                    
                    nCompl = !CODCOMPLEMENTO
                    
                    Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO <> 3"
                    Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux4
                        nValorParc = FormatNumber(!soma, 2)
                       .Close
                    End With
                    
                    Sql = "SELECT valortributo FROM debitotributo WHERE codreduzido = " & nCodReduz & " and anoexercicio=" & !AnoExercicio & " and "
                    Sql = Sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and "
                    Sql = Sql & "codcomplemento=" & !CODCOMPLEMENTO & " and codtributo=90"
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        If .RowCount > 0 Then
                            nValorTotalHon = !ValorTributo * nQtdeParc
                        Else
                            nValorTotalHon = 0
                        End If
                       .Close
                    End With
                   .Close
                
                End With
   
                '*** VERIFICA SE O PARCELAMENTO JÁ EXISTE NA TABELA ACORDOS **
                Sql = "select * from acordos where idacordo=" & nNumproc & " and anoacordo=" & nAnoproc
                Set RdoAux4 = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux4.RowCount = 0 Then
                    
                    '***** fone e email ****
                    If sSetor = "IMOBILIÁRIO" Then
                        Sql = "SELECT proprietario.codcidadao, cidadao.telefone as fone, cidadao.email  FROM proprietario INNER JOIN "
                        Sql = Sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao WHERE (proprietario.tipoprop = 'P') AND (proprietario.principal = 1) AND proprietario.codreduzido = " & nCodReduz
                    ElseIf sSetor = "MOBILIÁRIO" Then
                        Sql = "SELECT codigomob, fonecontato as fone, emailcontato as email FROM mobiliario WHERE  codigomob = " & nCodReduz
                    Else
                        Sql = "SELECT  codcidadao, telefone as fone, email FROM cidadao WHERE codcidadao = " & nCodReduz
                    End If
                    Set RdoAux5 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux5.RowCount > 0 Then
                        sFone = SubNull(RdoAux5!fone)
                        sEmail = SubNull(RdoAux5!Email)
                    End If
                    RdoAux5.Close
                   '***********************
                    
                    
                    'GRAVA O ACORDO
                    Sql = "insert acordos(idacordo,anoacordo,dtparcelamento,setordevedor,iddevedor,nroprocessoadm,crcacordante,nomeacordante,cpfcnpj,rginscrestadual,"
                    Sql = Sql & "cep,endereco,numero,complemento,bairro, cidade,estado,vlrtotal,qtdparcelas,primeirovencimento,vlrtotalhonorarios,qtdparcelashonorarios,"
                    Sql = Sql & "vlrparcelahonorarios,dtvenctohonorarios,VlrTotalDespesas, QtdParcelasDespesas, VlrParcelaDespesas, DtVenctoDespesas, DtGeracao,telefone,email) values ("
                    Sql = Sql & nNumproc & "," & nAnoproc & ",'" & Format(dDataProc, "mm/dd/yyyy") & "','" & sSetor & "',"
                    Sql = Sql & nCodReduz & ",'" & nNumproc & RetornaDVProcesso(nNumproc) & "/" & nAnoproc & "'," & nCodReduz & ",'" & Mask(sNomeResp) & "','" & sCPF & "','" & Left(sRG, 20) & "','" & sCep & "','" & sEndImovel & "',"
                    Sql = Sql & nNumImovel & ",'" & Left(sComplImovel, 40) & "','" & sBairroImovel & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "'," & Virg2Ponto(Round((nValorParc * nQtdeParc), 2)) & "," & nQtdeParc & ",'"
                    Sql = Sql & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'," & Virg2Ponto(Round(nValorTotalHon, 2)) & "," & IIf(nValorTotalHon = 0, 0, nQtdeParc) & "," & Virg2Ponto(Round((nValorTotalHon / nQtdeParc), 2)) & ","
                    Sql = Sql & IIf(nValorTotalHon = 0, "Null", "'" & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'") & "," & "0,0,0," & "Null" & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sFone & "','" & sEmail & "')"
                    cnInt.Execute Sql, rdExecDirect
                    
                   'GRAVA NA TABELA ACORDOSTATUS
                    sStatus = IIf(bCancelado = False, "PARCELAMENTO EM DIA", "PARCEL.CANCELADO")
                    Sql = "insert acordostatus(idacordo,anoacordo,dtocorrencia,ocorrencia,dtgeracao) values("
                    Sql = Sql & nNumproc & "," & nAnoproc & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sStatus & "','" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute Sql, rdExecDirect
                    
                   'GRAVA OS DÉBITOS DO ACORDO
                    Sql = "SELECT DISTINCT origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
                    Sql = Sql & "origemreparc.numparcela, origemreparc.codcomplemento, SUM(debitotributo.valortributo) AS Total, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.datavencimento "
                    Sql = Sql & "FROM origemreparc INNER JOIN debitotributo ON origemreparc.codreduzido = debitotributo.codreduzido AND origemreparc.anoexercicio = debitotributo.anoexercicio AND "
                    Sql = Sql & "origemreparc.codlancamento = debitotributo.codlancamento AND origemreparc.numsequencia = debitotributo.seqlancamento AND "
                    Sql = Sql & "origemreparc.numparcela = debitotributo.numparcela AND origemreparc.codcomplemento = debitotributo.codcomplemento INNER JOIN debitoparcela ON origemreparc.codreduzido = debitoparcela.codreduzido AND "
                    Sql = Sql & "origemreparc.anoexercicio = debitoparcela.anoexercicio AND origemreparc.codlancamento = debitoparcela.codlancamento AND origemreparc.numsequencia = debitoparcela.seqlancamento AND "
                    Sql = Sql & "origemreparc.NumParcela = debitoparcela.NumParcela And origemreparc.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO GROUP BY origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
                    Sql = Sql & "origemreparc.NumParcela , origemreparc.CODCOMPLEMENTO, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.datavencimento "
                    Sql = Sql & "HAVING origemreparc.numprocesso = '" & sNumProc & "' AND origemreparc.codreduzido =" & nCodReduz
                    Sql = Sql & " ORDER BY origemreparc.anoexercicio, origemreparc.numparcela"
                    Set RdoAcordo = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAcordo
                        Do Until .EOF
                            
                            nAno = !AnoExercicio
                            nLanc = !CodLancamento
                            nSeq = !numsequencia
                            nParc = !NumParcela
                            nCompl = !CODCOMPLEMENTO
                            If IsNull(!numcertidao) Then GoTo ProximoCDA
                            nPagina = Val(SubNull(!paginalivro))
                            nLivro = Val(SubNull(!numerolivro))
                            nNumCertidao = Val(SubNull(!numcertidao))
                            dDataInscricao = !datainscricao
                            dDataVencto = !DataVencimento
                           'GRAVA NA TABELA ACORDODEBITO
                            Sql = "insert acordodebitos(idacordo,anoacordo,nrolivro,nrofolha,seq,lancamento,exercicio,vlroriginal,vlrcorrecao,vlrjuros,vlrmulta,vlrtotal,nroparcela,complparcela,ajuizado,dtgeracao) values("
                            Sql = Sql & nNumproc & "," & nAnoproc & "," & nLivro & "," & nPagina & ","
                            Sql = Sql & RdoAcordo!numsequencia & "," & RdoAcordo!CodLancamento & "," & RdoAcordo!AnoExercicio & "," & Virg2Ponto(Format(RdoAcordo!Total, "#0.##")) & ",0,0,0," & Virg2Ponto(Format(RdoAcordo!Total, "#0.##")) & ","
                            Sql = Sql & RdoAcordo!NumParcela & "," & RdoAcordo!CODCOMPLEMENTO & "," & IIf(IsNull(!dataajuiza), 0, 1) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                            cnInt.Execute Sql, rdExecDirect
                           
                           
                            '*** GRAVA AS CDA, CDADebito, PARTES E CADASTRO ***
                                
                            Sql = "INSERT CDAs(IdDevedor,SetorDevedor,DtInscricao,NroCertidao,NroLivro,NroFolha,DtGeracao) values("
                            Sql = Sql & nCodReduz & ",'" & sSetor & "','" & Format(dDataInscricao, "mm/dd/yyyy") & "'," & nNumCertidao & ","
                            Sql = Sql & nLivro & "," & nPagina & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                            cnInt.Execute Sql, rdExecDirect
                            
                            Sql = "select @@identity as LastKey"
                            Set RdoAux2 = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            nCDA = RdoAux2!lastkey
                            RdoAux2.Close
                                
                                
                            If nCodReduz < 500000 Then
                                Sql = "INSERT Cadastro(IdCDA,SetorDevedor,Crc,Nome,Inscricao,CPFCnpj,RgInscrEstadual,LocalCep,LocalEndereco,LocalNumero,LocalComplemento,"
                                Sql = Sql & "LocalBairro,LocalCidade,LocalEstado,Quadra,Lote,EntregaCep,EntregaEndereco,EntregaNumero,EntregaComplemento,EntregaBairro,"
                                Sql = Sql & "EntregaCidade,EntregaEstado,DtGeracao) values("
                                Sql = Sql & nCDA & ",'" & sSetor & "'," & nCodReduz & ",'" & SubNull(sNome) & "','" & sNumInsc & "','" & sCPF & "','" & sRG & "','"
                                Sql = Sql & sCepImovel & "','" & sEndImovel & "'," & nNumImovel & ",'" & sComplImovel & "','" & sBairroImovel & "','" & sCidadeImovel & "','" & sUFImovel & "','"
                                Sql = Sql & sQuadra & "','" & sLote & "','" & sCepEntrega & "','" & sEndEntrega & "'," & nNumEntrega & ",'" & sComplEntrega & "','"
                                Sql = Sql & sBairroEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Format(Now, "mm/dd/yyyy") & "')"
                                cnInt.Execute Sql, rdExecDirect
                            Else
                                Sql = "select * from vwFullCidadao where codcidadao=" & nCodReduz
                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    sCPF = SubNull(!Cnpj)
                                    If Trim(sCPF) = "" Then
                                        sCPF = SubNull(!CPF)
                                    End If
                                    Sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                                    Sql = Sql & nCDA & ",'Principal'," & !CodCidadao & ",'" & Mask(!NomeCidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                                    Sql = Sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                                    Sql = Sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                                    cnInt.Execute Sql, rdExecDirect
                                   .Close
                                End With
                            End If
                            
                            If nCodReduz < 100000 Then 'cadastra os proprietarios e compromissarios
                                Sql = "SELECT cadimob.codreduzido, proprietario.codcidadao, proprietario.tipoprop, vwFULLCIDADAO.nomecidadao, vwFULLCIDADAO.cpf,"
                                Sql = Sql & "vwFULLCIDADAO.cnpj, vwFULLCIDADAO.numimovel, vwFULLCIDADAO.complemento, vwFULLCIDADAO.siglauf, vwFULLCIDADAO.cep,"
                                Sql = Sql & "vwFULLCIDADAO.rg , vwFULLCIDADAO.orgao, vwFULLCIDADAO.DescBairro, vwFULLCIDADAO.desccidade, vwFULLCIDADAO.Endereco "
                                Sql = Sql & "FROM cadimob INNER JOIN proprietario ON cadimob.codreduzido = proprietario.codreduzido INNER JOIN vwFULLCIDADAO ON "
                                Sql = Sql & "proprietario.codcidadao = vwFULLCIDADAO.codcidadao where cadimob.codreduzido=" & nCodReduz
                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    Do Until .EOF
                                        sCPF = SubNull(!Cnpj)
                                        If Trim(sCPF) = "" Then
                                            sCPF = SubNull(!CPF)
                                        End If
                                        sTipoProp = IIf(!tipoprop = "P", "Principal", "Compromissário")
                                        Sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                                        Sql = Sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!NomeCidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                                        Sql = Sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                                        Sql = Sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                                        cnInt.Execute Sql, rdExecDirect
                                       .MoveNext
                                    Loop
                                   .Close
                                End With
                            ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then   'cadastra os socios
                                Sql = "SELECT * from vwmobiliarioproprietario where codmobiliario=" & nCodReduz
                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    Do Until .EOF
                                        sCPF = SubNull(!Cnpj)
                                        If Trim(sCPF) = "" Then
                                            sCPF = SubNull(!CPF)
                                        End If
                                        sTipoProp = "Sócio"
                                        Sql = "INSERT Partes(idCDA,Tipo,Crc,Nome,CpfCnpj,RgInscrEstadual,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,DtGeracao) Values("
                                        Sql = Sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!NomeCidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
                                        Sql = Sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
                                        Sql = Sql & Mask(SubNull(!descCidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
                                        cnInt.Execute Sql, rdExecDirect
                                       .MoveNext
                                    Loop
                                   .Close
                                End With
                            End If
                            
                            '***************************************************
                            'Carrega Tributos
                            Sql = "SELECT debitotributo.codtributo,valortributo,abrevtributo FROM debitotributo INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
                            Sql = Sql & "where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and "
                            Sql = Sql & "numparcela=" & nParc & " and codcomplemento=" & nCompl
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                Do Until .EOF
                                    
                                    Sql = "INSERT CDADebitos(idCDA,CodTributo,Tributo,Exercicio,Lancamento,Seq,NroParcela,ComplParcela,DtVencimento,VlrOriginal,DtGeracao) values("
                                    Sql = Sql & nCDA & "," & !CodTributo & ",'" & Mask(!ABREVTRIBUTO) & "'," & nAno & "," & nLanc & "," & nSeq & "," & nParc & ","
                                    Sql = Sql & nCompl & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!ValorTributo)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                                    cnInt.Execute Sql, rdExecDirect
                                    
                                   .MoveNext
                                Loop
                               .Close
                            End With
                            
                            '***************************************************
                                
                            '**************************************************
ProximoCDA:
                           .MoveNext
                        Loop
                       .Close
                    End With
                    
                End If
                
                RdoAux4.Close
                
            End If
        End With
proximo:
        DoEvents
       .MoveNext
    Loop
   .Close
End With


'*** PAGAMENTOS *****
Sql = "SELECT debitopago.codreduzido, debitopago.anoexercicio, debitopago.codlancamento, debitopago.seqlancamento, debitopago.numparcela, debitopago.codcomplemento, "
Sql = Sql & "debitopago.seqpag, debitopago.datapagamento, debitopago.datarecebimento, debitopago.valorpago, debitopago.codbanco, debitopago.codagencia,"
Sql = Sql & "debitopago.restituido, debitopago.numdocumento, debitopago.valorpagoreal, debitopago.intacto, debitopago.valortarifa, debitopago.arquivobanco,"
Sql = Sql & "debitopago.valordif , debitopago.datapagamentocalc, debitopago.dataintegracao, debitoparcela.numprocesso FROM debitopago INNER JOIN "
Sql = Sql & "debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND "
Sql = Sql & "debitopago.NumParcela = debitoparcela.NumParcela And debitopago.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO "
Sql = Sql & "WHERE debitopago.datarecebimento>='" & Format(Now - 30, "mm/dd/yyyy") & "' AND (debitopago.codlancamento = 20) AND (debitopago.restituido IS NULL) AND (debitopago.dataintegracao IS NULL)"
'Sql = Sql & "WHERE debitopago.datarecebimento>='08/01/2015' AND (debitopago.codlancamento = 20) AND (debitopago.restituido IS NULL) AND (debitopago.dataintegracao IS NULL)"

Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux3
    Do Until .EOF
        sNumProc = !numprocesso
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        'GRAVA NA TABELA ACORDOBAIXAS
        Sql = "insert acordobaixas(idAcordo, anoAcordo, DtBaixa, TipoBaixa, NroParcela, VlrOriginal, VlrCorrecao, VlrJuros, VlrMulta, VlrTotal, DtGeracao) values("
        Sql = Sql & nNumproc & "," & nAnoproc & ",'" & Format(!datarecebimento, "mm/dd/yyyy") & "','PAGAMENTO'," & !NumParcela & "," & Virg2Ponto(CStr(!valorpagoreal)) & ","
        Sql = Sql & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(!valorpagoreal)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cnInt.Execute Sql, rdExecDirect
        
        Sql = "update debitopago set dataintegracao='" & Format(Now, "mm/dd/yyyy") & "' where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and "
        Sql = Sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With


'Serasa
Sql = "SELECT DISTINCT IDDevedor, CONVERT(char(10), DATEADD(dd, DATEDIFF(DD, 0, DtGeracao), 0), 126) AS dtgeracao from negativacao where dtleitura is null"
Set RdoAux = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "insert serasa(codigo,dtentrada) values(" & !iddevedor & ",'" & Format(!DtGeracao, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
        Sql = "update negativacao set dtleitura='" & Format(Now, "mm/dd/yyyy") & "' where iddevedor=" & !iddevedor & " and datepart(dd,dtgeracao)=" & Day(!DtGeracao) & " and "
        Sql = Sql & "datepart(mm,dtgeracao)=" & Month(!DtGeracao) & " and datepart(yy,dtgeracao)=" & Year(!DtGeracao)
        cnInt.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With


'Parcelamentos Ajuizados
Sql = "SELECT AjuizamentoCdas.IdAjuizamentoCdas, AjuizamentoCdas.CodProcesso, AjuizamentoCdas.IdDevedor, AjuizamentoCdas.SetorDevedor, AjuizamentoCdas.NroCertidao, "
Sql = Sql & "AjuizamentoCdas.NroLivro, AjuizamentoCdas.NroFolha, AjuizamentoCdas.Seq, AjuizamentoCdas.Lancamento, AjuizamentoCdas.ComplParcela,"
Sql = Sql & "AjuizamentoCdas.Exercicio, AjuizamentoCdas.DtGeracao, AjuizamentoProcessos.DtLeitura, AjuizamentoCdas.NroParcela, AjuizamentoCdas.CodTributo,"
Sql = Sql & "AjuizamentoCdas.DtVencimento, AjuizamentoCdas.VlrOriginal, AjuizamentoProcessos.DtAjuizamento, AjuizamentoProcessos.Protocolo,"
Sql = Sql & "AjuizamentoProcessos.IdAjuizamentoProcesso , AjuizamentoProcessos.AnoProtocolo, AjuizamentoProcessos.ProcessoCNJ "
Sql = Sql & "FROM AjuizamentoCdas INNER JOIN AjuizamentoProcessos ON AjuizamentoCdas.CodProcesso = AjuizamentoProcessos.CodProcesso "
Sql = Sql & "WHERE AjuizamentoProcessos.DtLeitura IS NULL "
'Sql = Sql & "WHERE AjuizamentoCdas.DtLeitura IS NULL "
Sql = Sql & "ORDER BY AjuizamentoCdas.CodProcesso"

Set RdoAux = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        'If !NroCertidao > 0 Then
            Sql = "update debitoparcela set dataajuiza='" & Format(!DtAjuizamento, "mm/dd/yyyy") & "',processocnj='" & !ProcessoCNJ & "' where "
            Sql = Sql & "codreduzido=" & !iddevedor & " and anoexercicio=" & !exercicio & " and codlancamento=" & !lancamento & " and "
            'Sql = Sql & "seqlancamento=" & !Seq & " and numparcela=" & !nroparcela
            'Sql = Sql & " numparcela=" & !nroparcela & " and (statuslanc<5 or statuslanc=38 or statuslanc=39)"
            Sql = Sql & " numparcela=" & !nroparcela & " and (statuslanc=38 or statuslanc=39)"
            
            'If !NroCertidao > 0 Then
            '    Sql = Sql & " and  numcertidao=" & !NroCertidao
            'End If
            cn.Execute Sql, rdExecDirect
'            If cn.RowsAffected > 0 Then
'                MsgBox "teste"
 '           End If
            
            Sql = " update ajuizamentoprocessos set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where idajuizamentoprocesso=" & !IdAjuizamentoProcesso & " and dtleitura is null"
            cnInt.Execute Sql, rdExecDirect
            
            Sql = " update ajuizamentocdas set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' where idajuizamentocdas=" & !idajuizamentocdas & " and dtleitura is null"
            cnInt.Execute Sql, rdExecDirect
        'End If
DoEvents
       .MoveNext
    Loop
   .Close
End With

'A.R. DIGITAL
Sql = "SELECT idajuizamentodespesa,CodProcesso, DtDespesa, ValorDespesa, cnj From AjuizamentoDespesas WHERE (TipoDespesa = 'AR DIGITAL') AND (DtLeitura IS NULL)"
Set RdoAux = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nId = !idajuizamentodespesa
        Sql = " SELECT DISTINCT IdDevedor From AjuizamentoCdas Where CodProcesso=" & !CodProcesso
        Set RdoAux2 = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nCodReduz = RdoAux2!iddevedor
        RdoAux2.Close
        
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & 78 & " AND ANOEXERCICIO=" & Year(!DTDESPESA)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        RdoAux2.Close
        
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,dataajuiza,DATADEBASE,PROCESSOCNJ,USUARIO) VALUES(" & nCodReduz & "," & Year(!DTDESPESA) & ","
'        Sql = Sql & 78 & "," & nSeq & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "','"
'        Sql = Sql & !CNJ & "','INTEGRAÇÃO INTERLIS')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,dataajuiza,DATADEBASE,PROCESSOCNJ,USERID) VALUES(" & nCodReduz & "," & Year(!DTDESPESA) & ","
        Sql = Sql & 78 & "," & nSeq & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(!DTDESPESA, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & !CNJ & "',236)"
        cn.Execute Sql, rdExecDirect
        
        nValorTributo = Ponto2Virg(!valordespesa)
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
        Sql = Sql & nCodReduz & "," & Year(!DTDESPESA) & "," & 78 & "," & nSeq & "," & 1 & "," & 0 & "," & 667 & "," & Virg2Ponto(Format(nValorTributo, "#0.00")) & ")"
        cn.Execute Sql, rdExecDirect
        
        Sql = "update AjuizamentoDespesas set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where idajuizamentodespesa=" & nId
        cnInt.Execute Sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With

'RETORNO PROTESTO
'Sql = "SELECT Protesto_remessa.Codigo, Protesto_remessa.idDevedor, Protesto_Debitos.Exercicio, Protesto_Debitos.Lancamento, Protesto_Debitos.Seq,Protesto_Debitos.nroParcela, Protesto_Debitos.ComplParcela,"
'Sql = Sql & "Protesto_Debitos.dtaVencimento, Protesto_remessa.cod_protesto, Protesto_remessa.nroTitulo,Protesto_remessa.data_remessa FROM Protesto_remessa INNER JOIN Protesto_Debitos ON "
'Sql = Sql & "Protesto_remessa.cod_protesto = Protesto_Debitos.Cod_protesto Where (Protesto_remessa.dtLeitura Is Null) ORDER BY Protesto_remessa.idDevedor, Protesto_Debitos.Exercicio, Protesto_Debitos.Lancamento, "
'Sql = Sql & "Protesto_Debitos.Seq, Protesto_Debitos.nroParcela,Protesto_Debitos.ComplParcela"

Sql = "SELECT Protesto_remessa.Codigo, Protesto_remessa.idDevedor, Protesto_Debitos.Exercicio, Protesto_Debitos.Lancamento, Protesto_Debitos.Seq,Protesto_Debitos.nroParcela, Protesto_Debitos.ComplParcela, "
Sql = Sql & "Protesto_Debitos.dtaVencimento, Protesto_remessa.cod_protesto, Protesto_remessa.nroTitulo,Protesto_remessa.data_remessa , Protesto_Ocorrencia.Ocorrencia,Irregularidade FROM Protesto_remessa INNER JOIN "
Sql = Sql & "Protesto_Debitos ON Protesto_remessa.cod_protesto = Protesto_Debitos.Cod_protesto LEFT OUTER JOIN Protesto_Ocorrencia ON Protesto_remessa.cod_protesto = Protesto_Ocorrencia.cod_protesto "
Sql = Sql & "Where (Protesto_remessa.dtLeitura Is Null or Protesto_Ocorrencia.dtLeitura Is Null) ORDER BY Protesto_remessa.idDevedor, Protesto_Debitos.Exercicio, Protesto_Debitos.Lancamento, Protesto_Debitos.Seq, Protesto_Debitos.nroParcela,Protesto_Debitos.ComplParcela "
Set RdoAux = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If Trim(SubNull(!Ocorrencia)) = "" Then
            Sql = "update debitoparcela set STATUSLANC=39,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            Sql = Sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc=3"
        ElseIf Trim(SubNull(!Ocorrencia)) = "PROTESTADO" Then
            Sql = "update debitoparcela set STATUSLANC=38,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            Sql = Sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc in (3,39)"
        ElseIf Trim(SubNull(!Ocorrencia)) = "PAGO" Then
            Sql = "update debitoparcela set STATUSLANC=41,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            Sql = Sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc in (3,38,39)"
        Else
            Sql = "update debitoparcela set STATUSLANC=3,protesto_nro_titulo=" & Val(Trim(!nroTitulo)) & ",protesto_data_remessa='" & Format(!data_remessa, "mm/dd/yyyy") & "' where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            Sql = Sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc in (38,39)"
        End If
        cn.Execute Sql, rdExecDirect
        If Trim(SubNull(!irregularidade)) <> "" Then
            Sql = "update debitoparcela set STATUSLANC=3,protesto_nro_titulo=null,protesto_data_remessa=null where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
            Sql = Sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroparcela & " and codcomplemento=" & !complparcela & " and statuslanc in (38,39)"
            cn.Execute Sql, rdExecDirect
        End If
        
        Sql = "update protesto_remessa set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where codigo=" & !Codigo
        cnInt.Execute Sql, rdExecDirect
        
        Sql = "update protesto_ocorrencia set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where cod_protesto=" & !Cod_protesto
        cnInt.Execute Sql, rdExecDirect
        
        Sql = "update protesto_debitos set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where cod_protesto=" & !Cod_protesto & " and Exercicio=" & !exercicio & " and lancamento=" & !lancamento & " and "
        Sql = Sql & "seq=" & !Seq & " and nroparcela=" & !nroparcela & " and complparcela=" & !complparcela
        cnInt.Execute Sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Sql = "update Protesto_Ocorrencia set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where dtleitura is null"
cnInt.Execute Sql, rdExecDirect



Sql = "select * from CDADebitosNaoProtestados where DtLeitura is null"
Set RdoAux = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        'Sql = "update debitoparcela set STATUSLANC=3,protesto_nro_titulo=Null,protesto_data_remessa=Null where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
        Sql = "update debitoparcela set STATUSLANC=3 where codreduzido=" & Val(Trim(!iddevedor)) & " and anoexercicio="
        Sql = Sql & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nrparcela & " and codcomplemento=" & !complparcela & " and statuslanc=38"
        cn.Execute Sql, rdExecDirect
        
        Sql = "update CDADebitosNaoProtestados set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where idCDADebitosNaoProtestados=" & !idCDADebitosNaoProtestados
        cnInt.Execute Sql, rdExecDirect
        
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With

'Sql = "SELECT DISTINCT RTRIM(Protesto_remessa.idDevedor) AS IdDevedor, Protesto_Debitos.Exercicio, Protesto_Debitos.Lancamento, Protesto_Debitos.Seq, Protesto_Debitos.nroParcela, Protesto_Debitos.ComplParcela,  Protesto_Ocorrencia.cod_protesto "
'Sql = Sql & "FROM Protesto_Debitos INNER JOIN Protesto_Ocorrencia ON Protesto_Debitos.Cod_protesto = Protesto_Ocorrencia.cod_protesto INNER JOIN Protesto_remessa ON Protesto_Ocorrencia.cod_protesto = Protesto_remessa.cod_protesto "
'Sql = Sql & "WHERE (Protesto_Ocorrencia.dtLeitura IS NULL) AND (RTRIM(Protesto_Ocorrencia.Ocorrencia) <> 'PAGO') AND (RTRIM(Protesto_Ocorrencia.Ocorrencia) <> 'PROTESTADO')"
'Set RdoAux = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        Sql = "select * from debitoparcela where codreduzido=" & !iddevedor & " and anoexercicio=" & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroParcela & " and codcomplemento=" & !complparcela
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If RdoAux2.RowCount > 0 Then
'            If RdoAux2!statuslanc = 38 Then
'                Sql = "update debitoparcela set statuslanc=3,protesto_nro_titulo=null,protesto_data_remessa=null where codreduzido=" & !iddevedor & " and anoexercicio=" & !exercicio & " and codlancamento=" & !lancamento & " and seqlancamento=" & !Seq & " and numparcela=" & !nroParcela & " and codcomplemento=" & !complparcela
'                cn.Execute Sql, rdExecDirect
'            End If
'        End If
'        Sql = "update Protesto_Ocorrencia set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where codigo=" & !Cod_protesto
'        cnInt.Execute Sql, rdExecDirect

'       .MoveNext
'    Loop
'   .Close
'End With
'Sql = "update Protesto_Ocorrencia set dtleitura='" & Format(Now, "mm/dd/yyyy hh:mm") & "' where dtleitura is null"
'cnInt.Execute Sql, rdExecDirect

cnInt.Close
Liberado
End Sub
