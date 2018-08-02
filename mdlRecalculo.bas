Attribute VB_Name = "mdlRecalculo"
'Variáveis
Dim RdoAux As rdoResultset
Dim Sql As String

Dim nDistrito As Integer
Dim nSetor As Integer
Dim nQuadra As Integer
Dim nLote As Long
Dim nSeq As Integer
Dim nAreaTerreno As Double
Dim nComprimentoTestada As Double
Dim nComprimentoTestadaP As Double
Dim nCodReduzido As Long
Dim nCodSituacao As Integer
Dim nCodPedologia As Integer
Dim nCodTopografia As Integer
Dim nCodAgrupamento As Integer
Dim nValorAgrupamento As Double
Dim nFatorDistrito As Double
Dim nFatorGleba As Double
Dim nFatorProfundidade As Double
Dim nFatorSituacao As Double
Dim nFatorPedologia As Double
Dim nFatorTopografia As Double
Dim nValorVenalTerritorial As Double
Dim nValorVenalPredial As Double
Dim nValorVenalPredialSMAR As Double
Dim nValorVenalImovel As Double
Dim nValorVenalImovelSMAR As Double
Dim nValorFatores As Double
Dim nAliquotaTerritorial As Double
Dim nAliquotaPredial As Double
Dim nValorITU As Double
Dim nValorIPTU As Double
Dim nTaxaExpP As Double
Dim nTaxaExpU As Double
Dim nUfirAtual As Double
Dim nAnoCalculo As Integer
Dim nAnoCalculoOriginal As Integer
Dim nNumeroParcelas As Integer
Dim nDescontoUnica As Double
Dim sTemUnica As String
Dim bSucesso As Boolean

Public Function ExecutaCalculo19982002(nCodigoReduzidoImovel As Long, nAnoCalc As Integer) As Boolean
Dim bIPTU As Boolean
Dim nValorUFIR As Double, nValorIptuFinal As Double

bIPTU = False
'Conecta "MANE", "1"
'ROTINA PARA RECALCULO DOS IMOVEIS DE 1998 A 2002
'APLICANDO O REDUTOR DE 20%

'RETORNA VERDADEIRO SE O CALCULO FOI EFETUADO COM SUCESSO

bSucesso = True
nCodReduzido = nCodigoReduzidoImovel
If nAnoCalc > 1999 Then
     nAnoCalculo = 1999
     nAnoCalculoOriginal = nAnoCalc
Else
    nAnoCalculo = nAnoCalc
    nAnoCalculoOriginal = 0
End If

nValorVenalPredial = 0
nValorVenalPredialSMAR = 0

frmMostragemCalculo.Rtb.SelColor = VerdeEscuro
frmMostragemCalculo.Rtb.SelText = "=================================="
PTb "  AMOSTRA DE CÁLCULO ITU/IPTU", VerdeEscuro
PTb "==================================", VerdeEscuro
PTb "Código do Imóvel: "
PTb Format(nCodReduzido, "0000000"), vbBlue, False
PTb "      Ano de Cálculo: ", , False
PTb Format(IIf(nAnoCalculoOriginal = 0, nAnoCalculo, nAnoCalculoOriginal), "0000"), vbBlue, False

'CALCULA VALOR VENAL TERRITORIAL
nValorVenalTerritorial = CalculoValorVenalTerritorial()

'VERIFICA SE É PRECISO CALCULAR VALOR VENAL PREDIAL
Sql = "SELECT CODREDUZIDO, SEQAREA "
Sql = Sql & "FROM AREAS WHERE CODREDUZIDO =" & nCodReduzido
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       If .RowCount > 0 Then
            bIPTU = True
            nValorVenalPredial = CalculoValorVenalPredial()
            nValorVenalPredialSMAR = CalculoValorVenalPredialSMAR()
       End If
End With

PTb ""
PTb "=================================================", VerdeEscuro
PTb "CÁLCULO ITU/IPTU SEM O REDUTOR DE 20%", VerdeEscuro
PTb "=================================================", VerdeEscuro

nAliquotaPredial = 1.5 'por enquanto
nAliquotaTerritorial = 3 'por enquanto

If bIPTU Then
    PTb "Aliquota Predial = "
    PTb nAliquotaPredial & " %", vbBlue, False
Else
    PTb "Aliquota Territorial = "
    PTb nAliquotaTerritorial & " %", vbBlue, False
End If

'valor venal imovel
nValorVenalImovel = nValorVenalPredial + nValorVenalTerritorial
nValorVenalImovelSMAR = nValorVenalPredialSMAR + nValorVenalTerritorial

PTb ""
PTb "Valor Venal do Imóvel = Valor Venal Predial + Valor Venal Territorial"
PTb "Valor Venal do Imóvel = "
PTb "R$ " & Format(nValorVenalPredial, "#0.00") & " + R$ " & Format(nValorVenalTerritorial, "#0.00") & " ===> " & Format(nValorVenalImovel, "#0.00"), vbRed, False
PTb "Valor Venal do Imóvel (SMAR) = ", Roxo
PTb "R$ " & Format(nValorVenalPredialSMAR, "#0.00") & " + R$ " & Format(nValorVenalTerritorial, "#0.00") & " ===> " & Format(nValorVenalImovelSMAR, "#0.00"), vbRed, False

If bIPTU Then
     PTb ""
     PTb "Valor do IPTU = Valor Venal do Imóvel * Aliquota Predial"
     PTb "Valor do IPTU = "
     PTb Format(nValorVenalImovel, "#0.00") & " * " & Format(nAliquotaPredial, "#0.00") & " ===> R$ " & Format(nValorVenalImovel * (nAliquotaPredial / 100), "#0.00"), vbRed, False
     PTb "Valor do IPTU (SMAR) = ", Roxo
     PTb Format(nValorVenalImovelSMAR, "#0.00") & " * " & Format(nAliquotaPredial, "#0.00") & " ===> R$ " & Format(nValorVenalImovelSMAR * (nAliquotaPredial / 100), "#0.00"), vbRed, False
     nValorIPTU = nValorVenalImovelSMAR * (nAliquotaPredial / 100)
Else
     PTb ""
     PTb "Valor do ITU = Valor Venal do Imóvel * Aliquota Territoria"
     PTb "Valor do ITU = "
     PTb Format(nValorVenalImovel, "#0.00") & " * " & Format(nAliquotaTerritorial, "#0.00") & " ===> R$ " & Format(nValorVenalImovel * (nAliquotaTerritorial / 100), "#0.00"), vbRed, False
     PTb "Valor do ITU (SMAR) = ", Roxo
     PTb Format(nValorVenalImovelSMAR, "#0.00") & " * " & Format(nAliquotaTerritorial, "#0.00") & " ===> R$ " & Format(nValorVenalImovelSMAR * (nAliquotaTerritorial / 100), "#0.00"), vbRed, False
     nValorITU = nValorVenalImovelSMAR * (nAliquotaTerritorial / 100)
End If

If nAnoCalculoOriginal <> 0 Then
    PTb ""
    PTb "=================================================", VerdeEscuro
    PTb "CORREÇÃO MONETÁRIA DE 1999 À " & nAnoCalculoOriginal, VerdeEscuro
    PTb "=================================================", VerdeEscuro
    PTb ""
    
    Sql = "SELECT VALORUFIR FROM UFIR WHERE ANOUFIR=" & nAnoCalculoOriginal
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nUfirAtual = RdoAux!VALORUFIR
    RdoAux.Close
    
    If bIPTU Then
        PTb "Valor do IPTU = IPTU 1999 * UFIR " & nAnoCalculoOriginal & " / UFIR 1998"
        PTb "Valor do IPTU = "
        PTb Format(nValorIPTU, "#0.00") & " * " & Format(nUfirAtual, "#0.0000") & " / 0,9611 ==> R$ " & Format(nValorIPTU * nUfirAtual / 0.9611, "#0.00"), vbRed, False
        nValorIPTU = nValorIPTU * nUfirAtual / 0.9611
    Else
        PTb "Valor do ITU = ITU 1999 * UFIR " & nAnoCalculoOriginal & " / UFIR 1998"
        PTb "Valor do ITU = "
        PTb Format(nValorITU, "#0.00") & " * " & Format(nUfirAtual, "#0.0000") & " / 0,9611 ==> R$ " & Format(nValorITU * nUfirAtual / 0.9611, "#0.00"), vbRed, False
        nValorITU = nValorITU * nUfirAtual / 0.9611
    End If
End If
'EXECUTA CÁLCULO DE 1998
'ExecutaCalculo1998 nCodigoReduzidoImovel

If nAnoCalculo = 1998 Then
    PTb ""
    PTb "Valor  do ITU assumido será de :", Roxo
    PTb "R$ " & Format(nValorITUIPTU1998, "#0.00"), vbRed, False
    GoTo fim
End If

'APLICA O REDUTOR DE 20%
PTb ""
PTb "=================================================", VerdeEscuro
PTb "CÁLCULA O REDUTOR DE 20%", VerdeEscuro
PTb "=================================================", VerdeEscuro
If bIPTU Then
    PTb "Valor do IPTU em " & nAnoCalculo & " = R$ " & Format(nValorIPTU, "#0.00")
    PTb "Valor do IPTU em 1998 + 20% = R$ " & nValorITUIPTU1998 & " + " & nValorITUIPTU1998 * 0.2
    PTb "Valor do IPTU em 1998 + 20% = ", Roxo
    PTb "R$ " & Format(nValorITUIPTU1998 + nValorITUIPTU1998 * 0.2, "#0.00"), vbRed, False
Else
    PTb "Valor do ITU em " & nAnoCalculo & " = R$ " & Format(nValorITU, "#0.00")
    PTb "Valor do ITU em 1998 + 20% = R$ " & nValorITUIPTU1998 & " + " & nValorITUIPTU1998 * 0.2
    PTb "Valor do ITU em 1998 + 20% = ", Roxo
    PTb "R$ " & Format(nValorITUIPTU1998 + nValorITUIPTU1998 * 0.2, "#0.00"), vbRed, False
End If

'APLICA A CORRECAO
PTb ""
PTb "=================================================", VerdeEscuro
PTb "APLICA A CORREÇÃO MONETÁRIA", VerdeEscuro
PTb "=================================================", VerdeEscuro
PTb "Valor da UFIR em 1998 = "
PTb "0,9611", vbBlue, False
nValorUFIR = RetornaUFIR
PTb "Valor da UFIR em " & nAnoCalculo & " = "
PTb Format(nValorUFIR, "#0.0000"), vbBlue, False
PTb ""

If bIPTU Then
    PTb "Valor do IPTU em 1998 = (Valor do IPTU em 1998 (com 20%) / UFIR 1998) * UFIR " & IIf(nAnoCalculoOriginal = 0, nAnoCalculo, nAnoCalculoOriginal)
    PTb "Valor do IPTU em 1998 = (" & Format(nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2), "#0.00") & " / " & "0,9611) * " & Format(nValorUFIR, "#0.0000")
    PTb "Valor do IPTU em 1998 = ", Roxo
    PTb "R$ " & Format(((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR, "#0.00"), vbRed, False
Else
    PTb "Valor do ITU em 1998 = (Valor do ITU em 1998 (com 20%) / UFIR 1998) * UFIR " & IIf(nAnoCalculoOriginal = 0, nAnoCalculo, nAnoCalculoOriginal)
    PTb "Valor do ITU em 1998 = (" & Format(nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2), "#0.00") & " / " & "0,9611) * " & Format(nValorUFIR, "#0.0000")
    PTb "Valor do ITU em 1998 = ", Roxo
    PTb "R$ " & Format(((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR, "#0.00"), vbRed, False
End If

'VALIDAÇÃO
PTb ""
PTb "==============================================", VerdeEscuro
PTb "APLICA OS 20% + CORREÇÃO MONETÁRIA ", VerdeEscuro
PTb "==============================================", VerdeEscuro
PTb "Fórmula: ", Roxo
If bIPTU Then
    PTb "Se IPTU 1998 + 20% + Correção Monetária > IPTU Cálculado ====> IPTU Final = IPTU Calculado ", , False
    PTb "               Senão ===> IPTU Final = IPTU 1998 + 20% + Correção Monetária"
    PTb "IPTU Calculado= ", Roxo
    PTb "R$ " & Format(nValorIPTU, "#0.00"), vbRed, False
    PTb "      IPTU de 1998 + 20% + Correção Monetária= ", Roxo, False
    PTb "R$ " & Format(((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR, "#0.00"), vbRed, False
    PTb ""
    PTb "Valor  do IPTU assumido será de :", Roxo
    If ((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR > nValorIPTU Then
         PTb "R$ " & Format(nValorIPTU, "#0.00"), vbRed, False
         nValorIptuFinal = nValorIPTU
    Else
         PTb "R$ " & Format(((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR, "#0.00"), vbRed, False
         nValorIptuFinal = ((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR
    End If
Else
    PTb "Se ITU 1998 + 20% + Correção Monetária > ITU Cálculado ====> ITU Final = ITU Calculado ", , False
    PTb "               Senão ===> ITU Final = ITU 1998 + 20% + Correção Monetária"
    PTb "ITU Calculado= ", Roxo
    PTb "R$ " & Format(nValorITU, "#0.00"), vbRed, False
    PTb "      ITU de 1998 + 20% + Correção Monetária= ", Roxo, False
    PTb "R$ " & Format(((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR, "#0.00"), vbRed, False
    PTb ""
    PTb "Valor  do ITU assumido será de :", Roxo
    If ((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR > nValorITU Then
         PTb "R$ " & Format(nValorITU, "#0.00"), vbRed, False
         nValorIptuFinal = nValorITU
    Else
         PTb "R$ " & Format(((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR, "#0.00"), vbRed, False
         nValorIptuFinal = ((nValorITUIPTU1998 + (nValorITUIPTU1998 * 0.2)) / 0.9611) * nValorUFIR
    End If
End If

'PARCELAMENTO
PTb ""
PTb "==============================================", VerdeEscuro
PTb "PARCELAMENTO DO ITU/IPTU + TX.EXPEDIENTE", VerdeEscuro
PTb "==============================================", VerdeEscuro
PTb ""
Sql = "SELECT VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & IIf(nAnoCalculoOriginal = 0, nAnoCalculo, nAnoCalculoOriginal)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nTaxaExpP = RdoAux!VALORPARCELA
nTaxaExpU = RdoAux!VALORUNICA
RdoAux.Close

Sql = "SELECT QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA FROM PARAMPARCELA WHERE ANO=" & IIf(nAnoCalculoOriginal = 0, nAnoCalculo, nAnoCalculoOriginal)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nNumeroParcelas = RdoAux!QTDEPARCELA
nDescontoUnica = RdoAux!DESCONTOUNICA
sTemUnica = RdoAux!PARCELAUNICA
RdoAux.Close

PTb "Taxa de Expediente Parcelado = "
PTb "R$ " & Format(nTaxaExpP, "#0.00"), vbBlue, False
nValorUFIR = RetornaUFIR
If sTemUnica = "S" Then
     PTb "Taxa de Expediente Única = "
     PTb "R$ " & Format(nTaxaExpU, "#0.00"), vbBlue, False
Else
     PTb "Taxa de Expediente Única = "
     PTb "Não tem", vbBlue, False
End If

PTb ""
PTb nNumeroParcelas & " x Parcelas de R$ " & Format(nValorIptuFinal / 10, "#0.00") & " + R$ " & Format(nTaxaExpP / 10, "#0.00") & " = ", Roxo
PTb "R$ " & Format((nValorIptuFinal / 10) + (nTaxaExpP / 10), "#0.00"), vbRed, False
If sTemUnica = "S" Then
     PTb "Parcela Única de R$ " & Format(nValorIptuFinal - (nValorIptuFinal * (5 / 100)), "#0.00") & " + R$ " & Format(nTaxaExpU, "#0.00") & " com desconto de " & Format(nDescontoUnica, "#0.00") & " % = ", Roxo
     PTb "R$ " & Format((nValorIptuFinal - (nValorIptuFinal * nDescontoUnica / 100)) + nTaxaExpU, "#0.00"), vbRed, False
End If

fim:
ExecutaCalculo19982002 = bSucesso

End Function

Private Function RetornaUFIR() As Double

Sql = "SELECT VALORUFIR FROM UFIR WHERE ANOUFIR = " & IIf(nAnoCalculoOriginal = 0, nAnoCalculo, nAnoCalculoOriginal)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
RetornaUFIR = RdoAux!VALORUFIR
RdoAux.Close

End Function

Private Function CalculoValorVenalTerritorial() As Double
On Error GoTo Erro
Dim qd As New rdoQuery, bFracaoIdeal As Boolean, nNumTestadas As Integer, nSomaAreaTestada As Double

'CARREGA OS DADOS DO IMOVEL
Set qd.ActiveConnection = cn
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spDADOSDEUMIMOVEL(?) }"
qd(0) = nCodReduzido
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
       nDistrito = !Distrito
       nSetor = !Setor
       nQuadra = !Quadra
       nLote = !Lote
       nSeq = !Seq
       If Val(!Dt_FracaoIdeal) = 0 Then
            bFracaoIdeal = False
            nAreaTerreno = !Dt_AreaTerreno
       Else
           bFracaoIdeal = True
           nAreaTerreno = !Dt_FracaoIdeal
       End If
       nCodPedologia = !Dt_CodPedol
       nCodSituacao = !Dt_CodSituacao
       nCodTopografia = !Dt_CodTopog
      .Close
End With

Sql = "SELECT CODAGRUPA FROM vwFaceQuadra WHERE "
Sql = Sql & "CODDISTRITO=" & nDistrito & " AND "
Sql = Sql & "CODSETOR=" & nSetor & "AND "
Sql = Sql & "CODQUADRA=" & nQuadra & " AND "
Sql = Sql & "CODFACE=" & nSeq
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       nCodAgrupamento = !CODAGRUPA
      .Close
End With

'CARREGA O NUMERO DE TESTADAS
Sql = "SELECT AREATESTADA FROM TESTADA WHERE "
Sql = Sql & "CODREDUZIDO=" & nCodReduzido
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nNumTestadas = RdoAux.RowCount
RdoAux.Close

'CARREGA A METRAGEM LINEAR DA FACE PRINCIPAL
Sql = "SELECT AREATESTADA FROM TESTADA WHERE "
Sql = Sql & "CODREDUZIDO=" & nCodReduzido & " AND "
Sql = Sql & "NUMFACE=" & nSeq
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nComprimentoTestadaP = RdoAux!AREATESTADA
If nNumTestadas = 1 Then
    nComprimentoTestada = nComprimentoTestadaP 'Testada Principal
Else 'o comprimento da testada é a média de todas as testadas
    Sql = "SELECT AREATESTADA FROM TESTADA WHERE "
    Sql = Sql & "CODREDUZIDO=" & nCodReduzido
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
            nSomaAreaTestada = 0
           Do Until .EOF
                nSomaAreaTestada = nSomaAreaTestada + !AREATESTADA
               .MoveNext
           Loop
           nComprimentoTestada = nSomaAreaTestada / .RowCount
    End With
End If
RdoAux.Close


'SE HOUVER FRAÇÃO IDEAL O COMPRIMENTO DA TESTADA
'É CALCULADO POR ==> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
If bFracaoIdeal Then
    Sql = "SELECT AREACONSTR FROM AREAS WHERE CODREDUZIDO=" & nCodReduzido & " AND TIPOAREA='P'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nComprimentoTestada = nAreaTerreno * nComprimentoTestada / RdoAux2!AREACONSTR
End If

PTb "   Distrito: ", , False
PTb Format(nDistrito, "00"), vbBlue, False
PTb "  Setor: ", , False
PTb Format(nSetor, "00"), vbBlue, False
PTb "  Quadra: ", , False
PTb Format(nQuadra, "0000"), vbBlue, False
PTb "  Lote: ", , False
PTb Format(nLote, "00000"), vbBlue, False
PTb "  Face: ", , False
PTb Format(nSeq, "00"), vbBlue, False
PTb "Área do Terreno: "
PTb Format(nAreaTerreno, "#0.00") & " m²", vbBlue, False
PTb "  Testada Principal: ", , False
PTb Format(nComprimentoTestada, "#0.00") & " m", vbBlue, False
PTb "   Pedologia: ", , False
PTb Format(nCodPedologia, "00"), vbBlue, False
PTb "   Topografia: ", , False
PTb Format(nCodTopografia, "00"), vbBlue, False
PTb "   Situação: ", , False
PTb Format(nCodSituacao, "00"), vbBlue, False
PTb "   Zona: ", , False
PTb Format(nCodAgrupamento, "00"), vbBlue, False
PTb ""
PTb "======================================================", VerdeEscuro
If nAnoCalculoOriginal = 0 Then
     PTb "CÁLCULO VALOR VENAL TERRITORIAL", VerdeEscuro
Else
     PTb "CÁLCULO VALOR VENAL TERRITORIAL (CALCULADO EM 1999)", VerdeEscuro
End If
PTb "======================================================", VerdeEscuro

'RETORNA O VALOR DA PLANTA GENÉRICA DE VALORES
nValorAgrupamento = RetornaValorPlantaGenerica()
PTb "Valor na Planta Genérica de Valores: "
PTb "R$ " & Format(nValorAgrupamento, "#0.00"), vbBlue, False

'CALCULA OS FATORES
nValorFatores = CalculoDeFatores

'FORMULA FINAL DO VALOR VENAL TERRITORIAL
CalculoValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores

PTb ""
PTb "Valor Venal Territorial= Área do terreno * Valor do m² * Soma dos Fatores "
PTb "Valor Venal Territorial= "
PTb Format(nAreaTerreno, "#0.00") & " m²", vbBlue, False
PTb " * ", , False
PTb "R$ " & Format(nValorAgrupamento, "#0.00"), vbBlue, False
PTb " * ", , False
PTb Format(nValorFatores, "#0.00"), vbBlue, False
PTb " ===> R$ " & Format(CalculoValorVenalTerritorial, "#0.00"), vbRed, False

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function CalculoValorVenalPredial() As Double
On Error GoTo Erro

Dim RdoAux2 As rdoResultset
Dim nCodUso As Integer
Dim nCodTipo As Integer
Dim nCodCategoria As Integer
Dim nValorArea As Double
Dim nFatorCateg As Double
Dim nSomaValorVenalArea As Double

PTb ""
PTb "============================================================", VerdeEscuro
If nAnoCalculoOriginal = 0 Then
     PTb "CÁLCULO VALOR VENAL PREDIAL", VerdeEscuro
Else
     PTb "CÁLCULO VALOR VENAL PREDIAL (CALCULADO EM 1999)", VerdeEscuro
End If
PTb "============================================================", VerdeEscuro

'VALOR VENAL PREDIAL =  £(AREA CONSTRUIDA * PADRÃO CONSTRUÇÃO) DE TODAS AS AREAS DO IMÓVEL
PTb "Valor Venal Predial = £(Áreas Construidas * Padrão de Construção) de todas as Áreas do Imóvel * Fator Distrito"
PTb ""

'RETORNA FATOR DISTRITO
Sql = "SELECT FATORDISTRITO FROM FATORDISTRITO WHERE "
Sql = Sql & "ANODISTRITO=" & nAnoCalculo & " AND "
Sql = Sql & "CODDISTRITO=" & nDistrito
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       nFatorDistrito = !FATORDISTRITO
      .Close
End With

'RETORNA AS ÁREA DO IMOVEL
Sql = "SELECT CODREDUZIDO, SEQAREA, TIPOAREA, AREACONSTR, USOCONSTR , TIPOCONSTR, CATCONSTR "
Sql = Sql & "FROM AREAS WHERE CODREDUZIDO =" & nCodReduzido
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       nSomaValorVenalArea = 0
       PTb "Áreas do Imóvel:", Roxo
       PTb ""
       Do Until .EOF
            nFatorCateg = 0
           'CARREGA AS VARIAVEIS
            nValorArea = !AREACONSTR
            nCodUso = !USOCONSTR
            nCodTipo = !TIPOCONSTR
            nCodCategoria = !CATCONSTR
           'PARA CADA ÁREA RETORNA O FATOR CATEGORIA
            Sql = "SELECT FATORCATEG FROM FATORCATEG WHERE "
            Sql = Sql & "CODUSO =" & nCodUso & " AND CODTIPO =" & nCodTipo & " AND "
            Sql = Sql & "CODCATEG =" & nCodCategoria & " AND ANOCATEG =" & nAnoCalculo & " AND CODMOEDA = 1"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                 nFatorCateg = RdoAux2!FATORCATEG
                 nSomaValorVenalArea = nSomaValorVenalArea + (nValorArea * nFatorCateg)
            End If
            PTb "Tipo: "
            PTb !TIPOAREA, vbBlue, False
            PTb "   Área: ", , False
            PTb Format(!AREACONSTR, "#0.00") & " m²", vbBlue, False
            PTb "   Uso: ", , False
            PTb Format(!USOCONSTR, "00"), vbBlue, False
            PTb "   Tipo: ", , False
            PTb Format(!TIPOCONSTR, "00"), vbBlue, False
            PTb "   Categoria: ", , False
            PTb Format(!CATCONSTR, "00"), vbBlue, False
            PTb "   Fator Categoria:", , False
            PTb Format(nFatorCateg, "#0.00"), vbBlue, False
            PTb "   Subtotal ===> ", , False
            PTb "R$ " & Format(nFatorCateg * nValorArea, "#0.00"), vbRed, False
           .MoveNext
       Loop
       PTb ""
       PTb "Valor Venal Predial="
       PTb "R$ " & Format(nSomaValorVenalArea, "#0.00") & " * " & Format(nFatorDistrito, "#0.00") & " ===> " & Format(nSomaValorVenalArea * nFatorDistrito, "#0.00"), vbRed, False
End With

CalculoValorVenalPredial = nSomaValorVenalArea

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function CalculoValorVenalPredialSMAR() As Double
On Error GoTo Erro

Dim nCodUso As Integer
Dim nCodTipo As Integer
Dim nCodCategoria As Integer
Dim nFatorCateg As Double
Dim nSomaValorVenalArea As Double

PTb ""
PTb "=======================================================================", VerdeEscuro
If nAnoCalculoOriginal = 0 Then
    PTb "CÁLCULO VALOR VENAL PREDIAL (SMAR-APD)", VerdeEscuro
Else
    PTb "CÁLCULO VALOR VENAL PREDIAL (SMAR-APD) (CALCULADO EM 1999)", VerdeEscuro
End If
PTb "=======================================================================", VerdeEscuro

PTb "Valor Venal Predial = £(Áreas Construidas * Padrão de Construção da Área Principal) * Fator Distrito"
PTb ""

'NO SISTEMA  DA SMAR O CALCULO DO VALOR VENAL PREDIAL SE DAVA PELA SEGUINTE FORMULA
'VALOR VENAL PREDIAL = £(AREAS CONSTRUIDAS) * FATOR CATEGORIA DA AREA PRINCIPAL

'RETORNA A ÁREA PRINCIPAL DO IMOVEL
Sql = "SELECT CODREDUZIDO, SEQAREA, TIPOAREA, AREACONSTR, USOCONSTR , TIPOCONSTR, CATCONSTR "
Sql = Sql & "FROM AREAS WHERE CODREDUZIDO =" & nCodReduzido & " AND TIPOAREA='P'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        If .RowCount > 0 Then
            'CARREGA AS VARIAVEIS
             nCodUso = !USOCONSTR
             nCodTipo = !TIPOCONSTR
             nCodCategoria = !CATCONSTR
            'RETORNA O FATOR CATEGORIA
             Sql = "SELECT FATORCATEG FROM FATORCATEG WHERE "
             Sql = Sql & "CODUSO =" & nCodUso & " AND CODTIPO =" & nCodTipo & " AND "
             Sql = Sql & "CODCATEG =" & nCodCategoria & " AND ANOCATEG =" & nAnoCalculo & " AND CODMOEDA = 1"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             If RdoAux2.RowCount > 0 Then
                  nFatorCateg = RdoAux2!FATORCATEG
             Else
                 nFatorCateg = 0
             End If
             PTb "Área Principal: ", Roxo, , , , True
             PTb "   Uso: ", , False
             PTb Format(!USOCONSTR, "00"), vbBlue, False
             PTb "   Tipo: ", , False
             PTb Format(!TIPOCONSTR, "00"), vbBlue, False
             PTb "   Categoria: ", , False
             PTb Format(!CATCONSTR, "00"), vbBlue, False
             PTb "   Fator Categoria: ", , False
             PTb Format(nFatorCateg, "#0.00"), vbBlue, False
        Else
             nFatorCateg = 0
        End If
       .Close
End With

'SE NÃO HOUVER FATOR CATEGORIA RETORNA 0
If nFatorCateg = 0 Then
     CalculoValorVenalPredialSMAR = 0
     Exit Function
End If

'RETORNA FATOR DISTRITO
Sql = "SELECT FATORDISTRITO FROM FATORDISTRITO WHERE "
Sql = Sql & "ANODISTRITO=" & nAnoCalculo & " AND "
Sql = Sql & "CODDISTRITO=" & nDistrito
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       nFatorDistrito = !FATORDISTRITO
      .Close
End With

'SOMA AS AREAS CONSTRUIDAS DO IMOVEL
Sql = "SELECT SUM(AREACONSTR) AS SOMAAREAS "
Sql = Sql & "FROM AREAS WHERE CODREDUZIDO =" & nCodReduzido
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      PTb ""
      PTb "Soma das Áreas = "
      PTb Format(!SOMAAREAS, "#0.00") & " m²", vbBlue, False
       nSomaValorVenalArea = !SOMAAREAS * nFatorCateg
      .Close
End With

PTb "Valor Venal Predial= "
PTb Format(nSomaValorVenalArea, "#0.00") & " * " & Format(nFatorDistrito, "#0.00") & "===> ", vbRed, False
PTb "R$ " & Format(nSomaValorVenalArea * nFatorDistrito, "#0.00"), vbRed, False

CalculoValorVenalPredialSMAR = nSomaValorVenalArea

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function RetornaValorPlantaGenerica() As Double
On Error GoTo Erro
Dim qd As New rdoQuery

'CARREGA VALOR PLANTA GENERICA
Set qd.ActiveConnection = cn
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spVALORPLANTAGENERICA(?,?) }"
qd(0) = nCodAgrupamento
qd(1) = nAnoCalculo

Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
       If .RowCount = 0 Then
            RetornaValorPlantaGenerica = 0
       Else
            RetornaValorPlantaGenerica = Format(!VALORTERRENO, "#0.00")
       End If
End With
RdoAux.Close

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function


Private Function RetornaFatorProfundidade() As Double
On Error GoTo Erro
Dim nCodProfun As Integer
Dim nValorProfundidade As Double

nCodProfun = 0
RetornaFatorProfundidade = 0

'*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
nValorProfundidade = nAreaTerreno / nComprimentoTestadaP

'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
Sql = "SELECT CODDISTRITO, CODPROFUN, MINPROFUN, MAXPROFUN From PROFUNDIDADE "
Sql = Sql & "WHERE CODDISTRITO=" & nDistrito & " ORDER BY CODDISTRITO, MINPROFUN"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      Do Until .EOF
           If nValorProfundidade >= !MINPROFUN And nValorProfundidade <= !MAXPROFUN Then
                Exit Do
           ElseIf nValorProfundidade >= !MINPROFUN And !MAXPROFUN = 0 Then
                Exit Do
           End If
          .MoveNext
      Loop
      nCodProfun = !CODPROFUN
      PTb "  Código Profundidade: ", , False
      PTb Format(nCodProfun, "00"), vbBlue, False
      PTb "  no intervalo mín: ", , False
      PTb Format(!MINPROFUN, "#0.00") & " m", vbBlue, False
      PTb "  e máx: ", , False
      PTb Format(!MAXPROFUN, "#0.00") & " m", vbBlue, False
     .Close
End With

'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
Sql = "SELECT FATORPROFUN FROM FATORPROFUN WHERE "
Sql = Sql & "ANOPROFUN =" & nAnoCalculo & " AND "
Sql = Sql & "CODDISTRITO =" & nDistrito & " AND "
Sql = Sql & "CODPROFUN =" & nCodProfun
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       If .RowCount > 0 Then
            RetornaFatorProfundidade = !FATORPROFUN
       Else
            RetornaFatorProfundidade = 0
       End If
       PTb "  Fator: ", , False
       PTb Format(!FATORPROFUN, "#0.00"), vbBlue, False
      .Close
End With

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function RetornaFatorSituacao() As Double
On Error GoTo Erro

Sql = "SELECT FATORSITUACAO FROM FATORSITUACAO WHERE "
Sql = Sql & "ANOSITUACAO=" & nAnoCalculo & " AND "
Sql = Sql & "CODSITUACAO=" & nCodSituacao
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      If .RowCount > 0 Then
           RetornaFatorSituacao = !FATORSITUACAO
      Else
           RetornaFatorSituacao = 0
      End If
      PTb "  Código Situação: ", , False
      PTb Format(nCodSituacao, "00"), vbBlue, False
      PTb "  Fator Situação: ", , False
      PTb Format(!FATORSITUACAO, "#0.00"), vbBlue, False
     .Close
End With

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function RetornaFatorPedologia() As Double
On Error GoTo Erro

Sql = "SELECT FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE "
Sql = Sql & "ANOPEDOLOGIA=" & nAnoCalculo & " AND "
Sql = Sql & "CODPEDOLOGIA=" & nCodPedologia
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      If .RowCount > 0 Then
           RetornaFatorPedologia = !FATORPEDOLOGIA
      Else
           RetornaFatorPedologia = 0
      End If
      PTb "  Código Pedologia: ", , False
      PTb Format(nCodPedologia, "00"), vbBlue, False
      PTb "  Fator Pedologia: ", , False
      PTb Format(!FATORPEDOLOGIA, "#0.00"), vbBlue, False
     .Close
End With

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function RetornaFatorTopografia() As Double
On Error GoTo Erro

Sql = "SELECT FATORTOPOG FROM FATORTOPOGRAFIA WHERE "
Sql = Sql & "ANOTOPOG=" & nAnoCalculo & " AND "
Sql = Sql & "CODTOPOG=" & nCodTopografia
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      If .RowCount > 0 Then
           RetornaFatorTopografia = !FATORTOPOG
      Else
           RetornaFatorTopografia = 0
      End If
      PTb "  Código Topografia: ", , False
      PTb Format(nCodTopografia, "00"), vbBlue, False
      PTb "  Fator Topografia: ", , False
      PTb Format(!FATORTOPOG, "#0.00"), vbBlue, False
     .Close
End With

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function CalculoDeFatores() As Double
On Error GoTo Erro

PTb ""
PTb "*** CÁLCULO DOS FATORES ***", Roxo

'RETORNA O VALOR DO FATOR GLEBA
PTb ""
PTb "* Fator Gleba *", Roxo, , , , True
nFatorGleba = RetornaFatorGleba(nAreaTerreno, nAnoCalculo)

'RETORNA O VALOR DO FATOR PROFUNDIDADE
PTb ""
PTb "* Fator Profundidade *", Roxo, , , , True
nFatorProfundidade = RetornaFatorProfundidade()

'RETORNA O VALOR DO FATOR SITUAÇÃO
PTb ""
PTb "* Fator Situação *", Roxo, , , , True
nFatorSituacao = RetornaFatorSituacao()

'RETORNA O VALOR DO FATOR PEDOLOGIA
PTb ""
PTb "* Fator Pedologia *", Roxo, , , , True
nFatorPedologia = RetornaFatorPedologia()

'RETORNA O VALOR DO FATOR TOPOGRAFIA
PTb ""
PTb "* Fator Topografia *", Roxo, , , , True
nFatorTopografia = RetornaFatorTopografia()

CalculoDeFatores = nFatorGleba * nFatorProfundidade * nFatorSituacao * nFatorPedologia * nFatorTopografia
PTb ""
PTb "Cálculo dos Fatores = Fator Gleba * Fator Profundidade * Fator Situação * Fator Pedologia * Fator Topografia"
PTb "Cálculo dos Fatores = "
PTb Format(CalculoDeFatores, "#0.00"), vbBlue, False

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Sub PTb(sTexto As String, Optional nColor As OLE_COLOR = vbBlack, Optional NovaLinha As Boolean = True, Optional Negrito As Boolean = False, Optional Italico As Boolean = False, Optional Sublinhado As Boolean = False)
On Error GoTo Erro

With frmMostragemCalculo.Rtb
      .SelColor = nColor
      .SelBold = Negrito
      .SelItalic = Italico
      .SelUnderline = Sublinhado
       If Not NovaLinha Then
           .SelText = .SelText & sTexto
       Else
           .SelText = vbCrLf & .SelText & sTexto
       End If
End With

Exit Sub
Erro:
MsgBox Err.Description
bSucesso = False

End Sub
