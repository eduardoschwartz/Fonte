Attribute VB_Name = "gti_Eicon"

Public Sub AtualizaEmpresa()
Dim Sql As String, RdoAux As rdoResultset, RdoEmp As rdoResultset, RdoAux2 As rdoResultset, RdoProp As rdoResultset
Dim nCodigo As Long, sIE As String, sRazao As String, sFantasia As String, sNumProcesso As String, sTipoEmpresa As String, nArea As Double
Dim sDoc As String, sDataAbertura As String, sDataEncerramento As String, sTipoLog As String, sTitLog As String, sNomeLog As String, sRegime As String
Dim nNumImovel As Integer, sCompl As String, sBairro As String, sCep As String, sCidade As String, sUF As String, sFone As String, sFax As String, sEmail As String
Dim nCodCidadao As Long, sNome As String, nCodLogr As Long, nTipoEnd As String

ConectaEicon



Sql = "select codigo from eicon_empresa order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        Do Until .EOF
        If .AbsolutePosition > 150 Then Exit Do
        nCodigo = !Codigo
        
       '******* DADOS DA EMPRESA **************************
        
        Sql = "select * from vwfullempresa3 where codigomob=" & nCodigo
        Set RdoEmp = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoEmp
            sIE = RetornaNumero(SubNull(!INSCESTADUAL))
            sRazao = !RazaoSocial
            sFantasia = SubNull(!NOMEFANTASIA)
            sNumProcesso = SubNull(!NUMPROCESSO)
            If Len(SubNull(!Cnpj)) = 14 Then
                sTipoEmpresa = "J"
                sDoc = RetornaNumero(!Cnpj)
            Else
                sTipoEmpresa = "F"
                sDoc = RetornaNumero(SubNull(!CPF))
            End If
            If Len(sDoc) < 2 Then sTipoEmpresa = "J"
            sDataAbertura = Format(!DATAABERTURA, "dd/mm/yyyy")
            If IsNull(!dataencerramento) Then
                sDataEncerramento = ""
            Else
                sDataEncerramento = Format(!dataencerramento, "dd/mm/yyyy")
            End If
            sTipoLog = Trim(SubNull(!AbrevTipoLog))
            sTitLog = Trim(SubNull(!AbrevTitLog))
            sNomeLog = Trim(SubNull(!NomeLogradouro))
            nNumImovel = Val(SubNull(!Numero))
            sCompl = SubNull(!Complemento)
            sBairro = SubNull(!DescBairro)
            sCep = RetornaNumero(RetornaCEP(!CodLogradouro, !Numero))
            sCidade = SubNull(!desccidade)
            sUF = SubNull(!SiglaUF)
            sFone = Left(RetornaNumero(SubNull(!FONECONTATO)), 15)
            sFax = Left(RetornaNumero(SubNull(!faxcontato)), 15)
            sEmail = SubNull(!EMAILCONTATO)
            
            Sql = "select codtributo from mobiliarioatividadeiss where codmobiliario=" & nCodigo
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                If RdoAux2!CodTributo = 11 Then
                    sRegime = "F"
                ElseIf RdoAux2!CodTributo = 12 Then
                    sRegime = "E"
                ElseIf RdoAux2!CodTributo = 13 Then
                    sRegime = "V"
                Else
                    sRegime = "N"
                End If
            Else
                sRegime = "N"
            End If
            RdoAux2.Close
            If sRegime = "V" Then sRegime = "A"
            If sRegime = "E" Then segime = "T"
            nArea = IIf(IsNull(!AREATL), 0, !AREATL)
           .Close
           
        End With
        
        Sql = "insert tb_inter_empresas(cod_cliente,num_cadastro,timestamp,inscricao,inscricao_estadual,nome_empresa,nome_fantasia,"
        Sql = Sql & "num_processo,tipo_empresa,cpf_cnpj,data_abertura,data_encerramento,tipo_logradouro,titulo_logradouro,logradouro,"
        Sql = Sql & "num_imovel,complemento,bairro,cep,cidade,estado,telefone,fax,email,regime_empresa,status_empresa,classificacao,area_ocupada) "
        Sql = Sql & "values(2177," & nCodigo & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & nCodigo & "," & IIf(Val(sIE) > 0, Val(sIE), "Null") & ",'" & Mask(sRazao) & "',"
        Sql = Sql & IIf(sFantasia <> "", "'" & Mask(sFantasia) & "'", "Null") & "," & IIf(sNumProcesso <> "", "'" & sNumProcesso & "'", "Null") & ",'" & sTipoEmpresa & "'," & IIf(Val(sDoc) > 0, Val(sDoc), "Null") & ",'" & Format(sDataAbertura, "m/dd/yyyy") & "',"
        Sql = Sql & IIf(IsDate(sDataEncerramento), "'" & Format(sDataEncerramento, "mm/dd/yyyy") & "'", "Null") & "," & IIf(sTipoLog <> "", "'" & sTipoLog & "'", "Null") & ","
        Sql = Sql & IIf(sTitLog <> "", "'" & sTitLog & "'", "Null") & ",'" & Mask(sNomeLog) & "'," & IIf(nNumImovel > 0, "'" & CStr(nNumImovel) & "'", "Null") & "," & IIf(sCompl <> "", "'" & Left(sCompl, 40) & "'", "Null") & ",'"
        Sql = Sql & sBairro & "'," & IIf(Val(sCep) > 0, Val(sCep), "Null") & ",'" & sCidade & "','" & sUF & "'," & IIf(Val(sFone) > 0, Val(sFone), "Null") & "," & IIf(Val(sFax) > 0, Val(sFax), "Null") & "," & IIf(sEmail <> "", "'" & sEmail & "'", "Null") & ","
        Sql = Sql & IIf(sRegime <> "", "'" & sRegime & "'", "Null") & ",'" & IIf(IsDate(sDataEncerramento), "E", "A") & "','" & "N" & "'," & RetornaNumero(FormatNumber(nArea, 2)) & ")"
        cnEicon.Execute Sql, rdExecDirect
        
       '******* DADOS DOS SÓCIOS **************************
        Sql = "SELECT * FROM mobiliarioproprietario Where mobiliarioproprietario.codmobiliario = " & nCodigo
        Set RdoProp = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoProp
            Do Until .EOF
                nCodCidadao = !CodCidadao
                Sql = "select codigo from eicon_socio where codigo=" & nCodCidadao
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount = 0 Then
                    Sql = "insert eicon_socio(codigo) values(" & nCodCidadao & ")"
                    cn.Execute Sql, rdExecDirect
                    AtualizaSocio
                End If
                RdoAux2.Close
               .MoveNext
            Loop
            RdoProp.Close
        End With
                
        
        
        DoEvents
        RdoAux.MoveNext
    Loop
    RdoAux.Close
End With

Sql = "delete from eicon_empresa"
'cn.Execute Sql, rdExecDirect

cnEicon.Close
MsgBox "fim"
End Sub

Public Sub AtualizaSocio()

Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodigo As Long, sTipoLog As String, sTitLog As String, sNomeLog As String, nCodEmpresa As Long, sNome As String, nCodLogr As Long, nTipoEnd As String
Dim nNumImovel As Integer, sCompl As String, sBairro As String, sCep As String, sCidade As String, sUF As String, sFone As String, sFax As String, sEmail As String

ConectaEicon2

'****REMOVER NO FINAL DOS TESTES *****
Sql = "delete from tb_inter_socios"
'cnEicon2.Execute sql, rdExecDirect
'*************************************

Sql = "select codigo from eicon_socio order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        Do Until .EOF
        nCodigo = !Codigo
        
       '******* DADOS DOS SÓCIOS **************************
        Sql = "SELECT vwfullcidadao.*,codmobiliario FROM vwFULLCIDADAO inner join mobiliarioproprietario on vwfullcidadao.codcidadao=mobiliarioproprietario.codcidadao where mobiliarioproprietario.codcidadao = " & nCodigo
        Set RdoProp = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoProp
            Do Until .EOF
                nCodEmpresa = !codmobiliario
                sNome = !nomecidadao
                Sql = "select num_cadastro from tb_inter_socios where num_cadastro=" & nCodEmpresa & " and nome_socio='" & sNome & "' and controle is null"
                Set RdoAux2 = cnEicon2.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    RdoAux2.Close
                    GoTo PROXIMO
                End If
                RdoAux2.Close
                
                sDoc = RetornaNumero(SubNull(!CPF))
                
                If Not IsNull(etiqueta2) Then
                    sTipoEnd = "C"
                Else
                    sTipoEnd = "R"
                End If
                                
                If sTipoEnd = "R" Then
                    sTipoLog = SubNull(!AbrevTipoLog)
                    sTitLog = SubNull(!AbrevTitLog)
                    sNomeLog = SubNull(!NomeLogradouro)
                    If sNomeLog = "" Then
                        sNomeLog = SubNull(!NOMELOGRADOURO2)
                    End If
                    nNumImovel = Val(SubNull(!NUMIMOVEL))
                    sCompl = SubNull(!Complemento)
                    sBairro = SubNull(!DescBairro)
                    nCodLogr = Val(SubNull(!CodLogradouro))
                    If nCodLogr > 0 Then
                        sCep = RetornaNumero(RetornaCEP(nCodLogr, nNumImovel))
                    Else
                        sCep = RetornaNumero(SubNull(!Cep))
                    End If
                    sCidade = SubNull(!descidade)
                    sUF = SubNull(!SiglaUF)
                    sFone = SubNull(!telefone)
                    sEmail = SubNull(!email)
                Else
                    sTipoLog = SubNull(!AbrevTipoLogC)
                    sTitLog = SubNull(!AbrevTitLogC)
                    sNomeLog = SubNull(!NomeLogradouroC)
'                    If sNomeLog = "" Then
'                        sNomeLog = SubNull(!NOMELOGRADOURO2)
'                    End If
                    nNumImovel = Val(SubNull(!NUMIMOVEL2))
                    sCompl = SubNull(!Complemento2)
                    sBairro = SubNull(!DescBairroC)
                    nCodLogr = Val(SubNull(!CodLogradouro2))
                    If nCodLogr > 0 Then
                        sCep = RetornaNumero(RetornaCEP(nCodLogr, nNumImovel))
                    Else
                        sCep = RetornaNumero(SubNull(!Cep2))
                    End If
                    sCidade = SubNull(!desccidadeC)
                    sUF = SubNull(!SiglaUF2)
                    sFone = SubNull(!TELEFONE2)
                    sEmail = SubNull(!EMAIL2)
                End If
                
                Sql = "insert tb_inter_socios(cod_cliente,num_cadastro,inscricao,cod_socio,nome_socio,timestamp,cpf,tipo_logradouro,titulo_logradouro,logradouro,num_imovel,complemento,bairro,cep,"
                Sql = Sql & "cidade,estado,telefone,email) "
                Sql = Sql & "values(2177," & nCodEmpresa & "," & nCodEmpresa & "," & nCodigo & ",'" & Mask(sNome) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & IIf(Val(sDoc) > 0, Val(sDoc), "Null") & ","
                Sql = Sql & IIf(sTipoLog <> "", "'" & sTipoLog & "'", "Null") & "," & IIf(sTitLog <> "", "'" & sTitLog & "'", "Null") & ",'" & Mask(sNomeLog) & "'," & IIf(nNumImovel > 0, "'" & CStr(nNumImovel) & "'", "Null") & ","
                Sql = Sql & IIf(sCompl <> "", "'" & Left(sCompl, 40) & "'", "Null") & "," & IIf(sBairro <> "", "'" & sBairro & "'", "Null") & "," & IIf(Val(sCep) > 0, Val(sCep), "Null") & ",'" & sCidade & "','" & sUF & "',"
                Sql = Sql & IIf(Val(sFone) > 0, Val(sFone), "Null") & "," & IIf(sEmail <> "", "'" & sEmail & "'", "Null") & ")"
                cnEicon2.Execute Sql, rdExecDirect
                 
                Sql = "delete from eicon_socio where codigo=" & nCodigo
                cn.Execute Sql, rdExecDirect
                 
               .MoveNext
            Loop
            RdoProp.Close
        End With
        
PROXIMO:
        DoEvents
        RdoAux.MoveNext
    Loop
    RdoAux.Close
End With


cnEicon2.Close


End Sub
