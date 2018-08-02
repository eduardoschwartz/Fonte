Attribute VB_Name = "mdlRecalculo98"
'Variáveis
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset
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
Dim nValorTxConservacao1998 As Double
Dim nValorTxLimpeza1998 As Double
Dim nAliquotaTerritorial As Double
Dim nAliquotaPredial As Double
Dim nValorITU As Double
Dim nValorIPTU As Double
Public nValorITUIPTU1998 As Double
Dim nAnoCalculo As Integer
Dim bSucesso As Boolean

Public Function ExecutaCalculo1998(nCodigoReduzidoImovel As Long) As Boolean
Dim bIPTU As Boolean

bIPTU = False
Conecta "MANE", "1"
'ROTINA PARA RECALCULO DOS IMOVEIS DE 1998 A 2002
'APLICANDO O REDUTOR DE 20%

'RETORNA VERDADEIRO SE O CALCULO FOI EFETUADO COM SUCESSO

bSucesso = True
nCodReduzido = nCodigoReduzidoImovel
nValorTxConservacao1998 = 1.35
'nValorTxLimpeza1998 = 3.78
nAnoCalculo = 1998
nValorVenalPredial = 0
nValorVenalPredialSMAR = 0


PTb ""
PTb "==================================", VerdeEscuro
PTb "  EXECUÇÃO DO CÁLCULO EM 1998", VerdeEscuro
PTb "==================================", VerdeEscuro

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

nAliquotaPredial = 1.5 'por enquanto
nAliquotaTerritorial = 3 'por enquanto

PTb ""
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
PTb "Valor Venal do Imóvel = "
PTb "R$ " & FormatNumber(nValorVenalPredial, 2) & " + R$ " & FormatNumber(nValorVenalTerritorial, 2) & " ===> " & FormatNumber(nValorVenalImovel, 2), vbRed, False
PTb "Valor Venal do Imóvel (SMAR) = ", Roxo
PTb "R$ " & FormatNumber(nValorVenalPredialSMAR, 2) & " + R$ " & FormatNumber(nValorVenalTerritorial, 2) & " ===> " & FormatNumber(nValorVenalImovelSMAR, 2), vbRed, False

If bIPTU Then
     PTb ""
     PTb "Valor do IPTU = "
     PTb FormatNumber(nValorVenalImovel, 2) & " * " & FormatNumber(nAliquotaPredial, 2) & " ===> R$ " & FormatNumber(nValorVenalImovel * (nAliquotaPredial / 100), 2), vbRed, False
     PTb "Valor do IPTU (SMAR) = ", Roxo
     PTb Format(nValorVenalImovelSMAR, "#0.00") & " * " & Format(nAliquotaPredial, "#0.00") & " ===> R$ " & Format(nValorVenalImovelSMAR * (nAliquotaPredial / 100), "#0.00"), vbRed, False
     nValorIPTU = nValorVenalImovelSMAR * (nAliquotaPredial / 100)
Else
     PTb ""
     PTb "Valor do ITU = "
     PTb Format(nValorVenalImovel, "#0.00") & " * " & Format(nAliquotaTerritorial, "#0.00") & " ===> R$ " & Format(nValorVenalImovel * (nAliquotaTerritorial / 100), "#0.00"), vbRed, False
     PTb "Valor do ITU (SMAR) = ", Roxo
     PTb Format(nValorVenalImovelSMAR, "#0.00") & " * " & Format(nAliquotaTerritorial, "#0.00") & " ===> R$ " & Format(nValorVenalImovelSMAR * (nAliquotaTerritorial / 100), "#0.00"), vbRed, False
     nValorITU = nValorVenalImovelSMAR * (nAliquotaTerritorial / 100)
End If

PTb ""
PTb "Valor do ITU/IPTU em 1998 =  ((V.V.Ter. 1998 + V.V.Pred.1998) * Aliq.) + Tx.Conserv. 1998 + Tx.Limpeza 1998"
PTb "Valor do ITU/IPTU em 1998 =  IPTU 1998 + Tx.Conserv. 1998 + Tx.Limpeza 1998"
If bIPTU Then
     PTb "Valor do ITU/IPTU em 1998 = " & Format(nValorIPTU, "#0.00") & " + (" & Format(nComprimentoTestada, "#0.00") & " * " & Format(nValorTxConservacao1998, "#0.00") & ") + (" & Format(nComprimentoTestada, "#0.00") & " * " & Format(nValorTxLimpeza1998, "#0.00") & ")"
     PTb "Valor do ITU/IPTU em 1998 = ", Roxo
     PTb "R$ " & Format(nValorIPTU + ((nComprimentoTestada * nValorTxConservacao1998) + (nComprimentoTestada * nValorTxLimpeza1998)), "#0.00"), vbRed, False
     nValorITUIPTU1998 = Format(nValorIPTU + ((nComprimentoTestada * nValorTxConservacao1998) + (nComprimentoTestada * nValorTxLimpeza1998)), "#0.00")
Else
     PTb "Valor do ITU/IPTU em 1998 = " & Format(nValorITU, "#0.00") & " + (" & Format(nComprimentoTestada, "#0.00") & " * " & Format(nValorTxConservacao1998, "#0.00") & ") + (" & Format(nComprimentoTestada, "#0.00") & " * " & Format(nValorTxLimpeza1998, "#0.00") & ")"
     PTb "Valor do ITU/IPTU em 1998 = ", Roxo
     PTb "R$ " & Format(nValorITU + ((nComprimentoTestada * nValorTxConservacao1998) + (nComprimentoTestada * nValorTxLimpeza1998)), "#0.00"), vbRed, False
     nValorITUIPTU1998 = Format(nValorITU + ((nComprimentoTestada * nValorTxConservacao1998) + (nComprimentoTestada * nValorTxLimpeza1998)), "#0.00")
End If
'APLICA O REDUTOR DE 20%
fim:
ExecutaCalculo1998 = bSucesso

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
nComprimentoTestadaP = RdoAux!AREATESTADA 'Testada Principal
If nNumTestadas = 1 Then
    nComprimentoTestada = nComprimentoTestadaP
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

'CALCULA A TAXA DE LIMPEZA
Sql = "SELECT USOCONSTR FROM AREAS WHERE CODREDUZIDO=" & nCodReduzido & " AND TIPOAREA='P'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       If .RowCount > 0 Then
            Select Case !USOCONSTR
                    Case 1
                          nValorTxLimpeza1998 = 3.78
                    Case 2
                          nValorTxLimpeza1998 = 10.57
                    Case 3
                          nValorTxLimpeza1998 = 10.57
                    Case 4, 5
                          'nValorTxLimpeza1998 = 4.54
                          nValorTxLimpeza1998 = 10.57
            End Select
      Else
            nValorTxLimpeza1998 = 3.01
      End If
End With

'SE HOUVER FRAÇÃO IDEAL O COMPRIMENTO DA TESTADA
'É CALCULADO POR ==> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
If bFracaoIdeal Then
    Sql = "SELECT AREACONSTR FROM AREAS WHERE CODREDUZIDO=" & nCodReduzido & " AND TIPOAREA='P'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nComprimentoTestada = nAreaTerreno * nComprimentoTestada / RdoAux2!AREACONSTR
End If

''RETORNA O VALOR DA PLANTA GENÉRICA DE VALORES
nValorAgrupamento = RetornaValorPlantaGenerica()
PTb "Valor na Planta Genérica de Valores: "
PTb "R$ " & Format(nValorAgrupamento, "#0.00"), vbBlue, False

'CALCULA OS FATORES
nValorFatores = CalculoDeFatores

PTb "Valor do metro linear para Taxa de Conservação em 1998 = R$ "
PTb Format(nValorTxConservacao1998, "#0.00"), vbBlue, False
PTb "Valor do metro linear para Taxa de Limpeza  em 1998 = R$ "
PTb Format(nValorTxLimpeza1998, "#0.00"), vbBlue, False

'FORMULA FINAL DO VALOR VENAL TERRITORIAL
CalculoValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores

PTb ""
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
       PTb ""
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
Dim nValorArea As Double
Dim nFatorCateg As Double
Dim nSomaValorVenalArea As Double

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
       nSomaValorVenalArea = !SOMAAREAS * nFatorCateg
      .Close
End With

PTb "Valor Venal Predial (SMAR) = ", Roxo
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

Private Function RetornaFatorGleba() As Double
On Error GoTo Erro
Dim nCodGleba As Integer

nCodGleba = 0
RetornaFatorGleba = 0

'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
Sql = "SELECT CODGLEBA, MINGLEBA, MAXGLEBA From GLEBA ORDER BY MINGLEBA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      Do Until .EOF
           If nAreaTerreno >= !MINGLEBA And nAreaTerreno <= !MAXGLEBA Then
                Exit Do
           ElseIf nAreaTerreno >= !MINGLEBA And !MAXGLEBA = 0 Then
                Exit Do
           End If
          .MoveNext
      Loop
      nCodGleba = !CODGLEBA
     .Close
End With

'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
Sql = "SELECT FATORGLEBA FROM FATORGLEBA WHERE "
Sql = Sql & "ANOGLEBA =" & nAnoCalculo & " AND "
Sql = Sql & "CODGLEBA =" & nCodGleba
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       If .RowCount > 0 Then
            RetornaFatorGleba = !FATORGLEBA
       Else
            RetornaFatorGleba = 0
       End If
     .Close
End With

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
     .Close
End With

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Function CalculoDeFatores() As Double
On Error GoTo Erro

nFatorGleba = RetornaFatorGleba()

nFatorProfundidade = RetornaFatorProfundidade()

'RETORNA O VALOR DO FATOR SITUAÇÃO
nFatorSituacao = RetornaFatorSituacao()

'RETORNA O VALOR DO FATOR PEDOLOGIA
nFatorPedologia = RetornaFatorPedologia()

'RETORNA O VALOR DO FATOR TOPOGRAFIA
nFatorTopografia = RetornaFatorTopografia()

CalculoDeFatores = nFatorGleba * nFatorProfundidade * nFatorSituacao * nFatorPedologia * nFatorTopografia
PTb "Cálculo dos Fatores = "
PTb Format(CalculoDeFatores, "#0.00"), vbBlue, False

Exit Function
Erro:
MsgBox Err.Description
bSucesso = False

End Function

Private Sub PTb(sTexto As String, Optional nColor As OLE_COLOR = vbBlack, Optional NovaLinha As Boolean = True, Optional Negrito As Boolean = False, Optional Italico As Boolean = False, Optional Sublinhado As Boolean = False)
On Error GoTo Erro

With frmMostragemCalculo.RTb
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

