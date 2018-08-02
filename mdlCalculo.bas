Attribute VB_Name = "mdlCalculo"

Public Function RetornaFatorGleba(nAreaTerreno As Double, nAnoCalculo) As Double
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

Public Function RetornaFatorPedologia(nCodPedologia As Integer, nAnoCalculo As Integer) As Double
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

Public Function RetornaFatorProfundidade(nAreaTerreno As Double, nComprimentoTestadaP As Double, nDistrito As Integer, nAnoCalculo As Integer) As Double
On Error GoTo Erro
Dim nCodProfun As Integer
Dim nValorProfundidade As Double

nCodProfun = 0
RetornaFatorProfundidade = 0

'*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
nValorProfundidade = Format(nAreaTerreno / nComprimentoTestadaP, "#0.00")

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

Public Function RetornaFatorSituacao(nAnoCalculo As Integer, nCodSituacao As Integer) As Double
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

Public Function RetornaFatorTopografia(nAnoCalculo As Integer, nCodTopografia As Integer) As Double
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


Public Function RetornaValorPlantaGenerica(nAnoCalculo As Integer, nCodAgrupamento As Integer) As Double
Dim RdoAux As rdoResultset

'CARREGA VALOR PLANTA GENERICA
Sql = "SELECT VALORTERRENO From TERRENO Where CODAGRUPAMENTO=" & nCodAgrupamento & " AND "
Sql = Sql & "ANOFATOR=" & nAnoCalculo & " AND CODMOEDA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
         RetornaValorPlantaGenerica = 0
    Else
         RetornaValorPlantaGenerica = Format(!VALORTERRENO, "#0.00")
    End If
   .Close
End With

End Function

