Attribute VB_Name = "modRel"
Private Type tPlano
    nAno As Integer
    nPlano As Integer
    sPlano As String
    nPerc As Double
    sDam As String
    sParcelado As String
    sDI As String
End Type

Private Type tRefis
    nAno As Integer
    nPlano As Integer
    sPlano As String
    nCodReduz As Long
    nNumDocumento As Long
    nValorGuia As Double
    nValorPago As Double
    sDataVencto As String
    sDataPagto As String
End Type

Public Sub GeraRefis(nAno As Integer, bParcelado As Boolean, bDam As Boolean, bDI As Boolean)
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, aRefis() As tRefis, nTotal As Double, aPlano() As tPlano, x As Integer, nPlano As Integer, nPos As Integer

ReDim aRefis(0): ReDim aPlano(0)
nTotal = 0
Ocupado

Sql = "select * from plano where ano=" & nAno & " order by ano"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aPlano(UBound(aPlano) + 1)
        aPlano(UBound(aPlano)).nAno = !Ano
        aPlano(UBound(aPlano)).nPlano = !Codigo
        aPlano(UBound(aPlano)).sPlano = !Nome
        aPlano(UBound(aPlano)).nPerc = !desconto
        aPlano(UBound(aPlano)).sDam = IIf(!dam = True, "S", "N")
        aPlano(UBound(aPlano)).sParcelado = IIf(!dam = True, "N", "S")
        aPlano(UBound(aPlano)).sDI = IIf(!distrito_industrial = True, "S", "N")
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(aPlano) - 1
    With aPlano(x)
        If (bParcelado And .sParcelado = "N") Or (Not bParcelado And .sParcelado = "S") Then
            GoTo proximo
        End If
        If (bDam And .sDam = "N") Or (Not bDam And .sDam = "S") Then
            GoTo proximo
        End If
        If (bParcelado And .sDam = "S") Or (Not bParcelado And .sDam = "N") Then
            GoTo proximo
        End If
        nPlano = .nPlano
            
        Sql = "SELECT parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento,parceladocumento.numparcela, parceladocumento.codcomplemento, "
        Sql = Sql & "parceladocumento.numdocumento, parceladocumento.valorjuros, parceladocumento.codbanco,parceladocumento.valormulta, parceladocumento.valorcorrecao, parceladocumento.intacto, parceladocumento.plano, "
        Sql = Sql & "debitopago.datapagamento,debitopago.datarecebimento, debitopago.ValorPago, NumDocumento.valorguia FROM parceladocumento INNER JOIN numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
        Sql = Sql & "LEFT OUTER JOIN debitopago ON parceladocumento.codreduzido = debitopago.codreduzido AND parceladocumento.anoexercicio = debitopago.anoexercicio AND parceladocumento.codlancamento = debitopago.codlancamento AND "
        Sql = Sql & "parceladocumento.seqlancamento = debitopago.seqlancamento AND parceladocumento.NumParcela = debitopago.NumParcela And parceladocumento.CODCOMPLEMENTO = debitopago.CODCOMPLEMENTO Where parceladocumento.plano = " & nPlano
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aRefis(UBound(aRefis) + 1)
                nPos = UBound(aRefis)
                aRefis(nPos).nAno = aPlano(x).nAno
                aRefis(nPos).nPlano = nPlano
                aRefis(nPos).sPlano = aPlano(x).sPlano
                aRefis(nPos).nCodReduz = !CODREDUZIDO
                aRefis(nPos).nNumDocumento = !NumDocumento
                aRefis(nPos).nValorGuia = !VALORGUIA
                If Not IsNull(!ValorPago) Then
                    aRefis(nPos).nValorPago = !ValorPago
                Else
                    aRefis(nPos).nValorPago = 0
                End If
               .MoveNext
            Loop
           .Close
        End With
End With
proximo:
Next


Liberado
Exit Sub



Open sPathBin & "\RefisDAM.txt" For Output As #1
Print #1, "RELATÓRIO DO REFIS (DISTRITO INDUSTRIAL)"
Print #1, "REFIS - (" & nAno & ") - Impresso em " & Format(Now, "dd/mm/yyyy")
Print #1, "-------------------------------------------------"
Print #1, " "
Print #1, "Documento     Valor     Código   Dt.Pagto."
Print #1, "-------------------------------------------------"
Print #1, " "
Sql = "SELECT DISTINCT parceladocumento.numdocumento, debitoparcela.codreduzido, debitopago.datapagamento, SUM(debitopago.valorpagoreal) AS valorpago,numdocumento.valorpago AS valordoc "
Sql = Sql & "FROM numdocumento INNER JOIN parceladocumento ON numdocumento.numdocumento = parceladocumento.numdocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
Sql = Sql & "debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND debitoparcela.anoexercicio = debitopago.anoexercicio AND debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitopago.NumParcela And debitoparcela.CODCOMPLEMENTO = debitopago.CODCOMPLEMENTO WHERE parceladocumento.plano in (" & sPlano & ")"
Sql = Sql & "GROUP BY parceladocumento.numdocumento, debitoparcela.codreduzido, debitopago.datapagamento, numdocumento.valorpago "
Sql = Sql & "HAVING (numdocumento.valorpago > 0) and (debito) ORDER BY debitopago.datapagamento"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
'        If !DataPagamento < CDate("08/15/2016") Then GoTo proximo
        nTotal = nTotal + !ValorPago
        nNumDoc = !NumDocumento
'        ax = nNumDoc & " " & FillLeft(FormatNumber(!Valordoc, 2), 10) & "     " & FillLeft(!CODREDUZIDO, 6) & "  " & Format(!DataPagamento, "dd/mm/yyyy")
        Print #1, ax
'proximo:
       .MoveNext
    Loop
   .Close
End With
Print #1, ""
Print #1, "----------------------------------"
Print #1, "Total pago: R$ " & FormatNumber(nTotal, 2)
Close #1
Z = Shell("NOTEPAD" & " " & sPathBin & "\RefisDAM.txt", vbNormalFocus)
Liberado
End Sub


