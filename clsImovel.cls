VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim sCodLogradouro As String
Dim sEndereco As String
Dim sNumero As String
Dim sComplemento As String
Dim sCodBairro As String
Dim sBairro As String
Dim sCodCidade As String
Dim sCidade As String
Dim sUF As String
Dim sCep As String

Dim Sql As String
Dim nCodImovel As Long
Dim nDV As Integer
Dim nDistrito As Integer
Dim nSetor As Integer
Dim nQuadra As Integer
Dim nLote As Integer
Dim nSeq As Integer
Dim nUnidade As Integer
Dim nSubUnidade As Integer
Dim sInscricao As String
Dim bInativo As Boolean
Dim nLi_Num As Integer
Dim sLi_Compl As String
Dim nLi_CodBairro As Integer
Dim sLi_Quadras As String
Dim sLi_Lotes As String
Dim nDt_AreaTerreno As Double
Dim nDt_CodUsoTerreno As Integer
Dim nDt_CodBenf As Integer
Dim nDt_CodTopog As Integer
Dim nDt_CodCategProp As Integer
Dim nDt_CodSituacao As Integer
Dim nDt_CodPedol As Integer
Dim sDt_NumAgua As String
Dim nDt_FracaoIdeal As Double
Dim nEe_TipoEnd As Integer
Dim nDc_QtdeEdif As Integer
Dim nDc_QtdePav As Integer
Dim sAbrevTitLog As String
Dim sAbrevTipoLog As String
Dim sNomeLogradouro As String
Dim nCodLogr As Long
Dim nCodCondominio As Integer
Dim sLi_Cep As String
Dim sDescBenfeitoria As String
Dim sDescSituacao As String
Dim sDescUsoTerreno As String
Dim sDescCategProp As String
Dim sDescPedologia As String
Dim sDescTopografia As String
Dim nCodPropPrincipal As Long
Dim sNomePropPrincipal As String
Dim sDescBairro As String
Dim sDescDistrito As String
Dim nEe_CodLog As Long
Dim sEe_NomeLog As String
Dim nEe_NumImovel As Integer
Dim sEe_Complemento As String
Dim sEe_Cep As String
Dim sEe_UF As String
Dim nEe_Cidade As Integer
Dim nEe_Bairro As Integer
Dim sEnderecoCompleto As String
Dim sEnderecoCompleto2 As String
Dim nCodProp As Long
Dim sNomeProp As String
Dim nFracao As Double
Dim aTestada(1 To 10, 1 To 2) As Double
Dim vTestada As Variant
Dim nQtdeTestada As Integer
Dim aTestadaCond(1 To 10, 1 To 2) As Double
Dim vTestadaCond As Variant
Dim nQtdeTestadaCond As Integer
Dim aProp(1 To 30, 1 To 6) As Variant
Dim vProp As Variant
Dim nQtdeProp As Integer
Dim nNumUnidades As Integer
Dim nAreaConstruida As Double
Dim sNomeCondominio As String
Dim sTipoMat As String
Dim nNumMat As Long
Dim bImune  As Boolean
Dim bConjugado As Boolean
Dim bReside As Boolean
Dim bCip As Boolean

Public Sub CarregaImovel(nValor As Long)
Dim RdoC As rdoResultset
Dim qd As New rdoQuery

Set qd.ActiveConnection = cn
On Error Resume Next
RdoC.Close
On Error GoTo 0
qd.Sql = "{ Call spDADOSDEUMIMOVEL(?) }"
qd(0) = nValor
Set RdoC = qd.OpenResultset(rdOpenKeyset)
With RdoC
     If .RowCount > 0 Then
         CodigoImovel = !CODREDUZIDO
         DV = !DV
         Distrito = !Distrito
         Setor = !Setor
         Quadra = !Quadra
         Lote = !Lote
         Seq = !Seq
         Unidade = !Unidade
         SubUnidade = !SubUnidade
         Inscricao = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
         bInativo = !Inativo
         Li_Num = !Li_Num
         Li_Compl = SubNull(!Li_Compl)
         Li_CodBairro = IIf(IsNull(!Li_CodBairro), 999, !Li_CodBairro)
         Li_Quadras = SubNull(!Li_Quadras)
         Li_Lotes = SubNull(!Li_Lotes)
         Dt_AreaTerreno = !Dt_AreaTerreno
         Dt_CodUsoTerreno = !Dt_CodUsoTerreno
         Dt_CodBenf = !Dt_CodBenf
         Dt_CodTopog = !Dt_CodTopog
         Dt_CodCategProp = !Dt_CodCategProp
         Dt_CodSituacao = !Dt_CodSituacao
         Dt_CodPedol = !Dt_CodPedol
         Dt_NumAgua = SubNull(!Dt_NumAgua)
         If Not IsNull(!Dt_FracaoIdeal) Then
            Dt_FracaoIdeal = !Dt_FracaoIdeal
         End If
         TipoMat = SubNull(!TipoMat)
         NumMat = Val(SubNull(!NumMat))
         Imune = IIf(IsNull(!Imune), False, !Imune)
         Conjugado = IIf(IsNull(!Conjugado), False, True)
         ResideImovel = IIf(IsNull(!ResideImovel), 1, !ResideImovel)
         IsentoCIP = IIf(IsNull(!Cip), 0, !Cip)
         Ee_TipoEnd = IIf(IsNull(!Ee_TipoEnd), 0, !Ee_TipoEnd)
         Dc_QtdeEdif = Val(SubNull(!Dc_QtdeEdif))
         Dc_QtdePav = Val(SubNull(!Dc_QtdePav))
         AbrevTitLog = SubNull(!AbrevTitLog)
         AbrevTipoLog = SubNull(!AbrevTipoLog)
         NomeLogradouro = SubNull(!NomeLogradouro)
         CodLogr = !CodLogr
         CodCondominio = !CodCondominio
         Li_Cep = SubNull(!Li_Cep)
         DescBenfeitoria = !DescBenfeitoria
         DescSituacao = !DescSituacao
         DescUsoTerreno = !DescUsoTerreno
         DescCategProp = !DescCategProp
         DescPedologia = !DescPedologia
         DescTopografia = !DescTopografia
         CodPropPrincipal = !CodCidadao
         NomePropPrincipal = !nomecidadao
         DescBairro = SubNull(!DescBairro)
         DescDistrito = !DescDistrito
         Ee_CodLog = Val(SubNull(!Ee_CodLog))
         Ee_NomeLog = SubNull(!Ee_NomeLog)
         Ee_NumImovel = Val(SubNull(!Ee_NumImovel))
         Ee_Complemento = SubNull(!Ee_Complemento)
         Ee_Cep = SubNull(!Ee_Cep)
         Ee_Uf = SubNull(!Ee_Uf)
         Ee_Cidade = Val(SubNull(!Ee_Cidade))
         Ee_Bairro = Val(SubNull(!Ee_Bairro))
         EnderecoCompleto = Format(CodLogr, "0000") & " - " & Trim$(AbrevTipoLog) & IIf(IsNull(AbrevTitLog), "", " " & Trim$(AbrevTitLog)) & " " & NomeLogradouro & ", " & Li_Num
         EnderecoCompleto2 = Trim$(AbrevTipoLog) & IIf(IsNull(AbrevTitLog), "", " " & Trim$(AbrevTitLog)) & " " & NomeLogradouro & ", " & Li_Num
     Else
        CodigoImovel = 0
     End If
    .Close
     Set qd = Nothing
End With

End Sub

Public Sub CarregaTestada()
Dim RdoC As rdoResultset
Dim x As Integer
Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA "
Sql = Sql & "WHERE CODREDUZIDO=" & CodigoImovel
Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoC
    x = 1
    Do Until .EOF
       aTestada(x, 1) = !NUMFACE
       aTestada(x, 2) = !AREATESTADA
       x = x + 1
      .MoveNext
    Loop
   .Close
End With
QtdeTestada = x - 1
Testada = aTestada

End Sub

Public Sub CarregaTestadaCond(nValor As Long)
Dim RdoC As rdoResultset
Dim x As Integer
Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADACONDOMINIO "
Sql = Sql & "WHERE CODCOND=" & nValor
Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoC
    x = 1
    Do Until .EOF
       aTestadaCond(x, 1) = !NUMFACE
       aTestadaCond(x, 2) = !AREATESTADA
       x = x + 1
      .MoveNext
    Loop
   .Close
End With
QtdeTestadaCond = x - 1
TestadaCond = aTestadaCond

End Sub

Public Sub CarregaProprietario()
Dim RdoC As rdoResultset, RdoE As rdoResultset
Dim x As Integer

Sql = "SELECT PROPRIETARIO.CODREDUZIDO,PROPRIETARIO.CODCIDADAO,CIDADAO.NOMECIDADAO,"
Sql = Sql & "PROPRIETARIO.TIPOPROP,PROPRIETARIO.PRINCIPAL "
Sql = Sql & "FROM PROPRIETARIO INNER JOIN CIDADAO ON "
Sql = Sql & "PROPRIETARIO.CODCIDADAO = CIDADAO.CODCIDADAO "
Sql = Sql & "WHERE CODREDUZIDO=" & CodigoImovel & " ORDER BY proprietario.tipoprop, proprietario.principal DESC"
Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoC
    x = 1
    Do Until .EOF
       aProp(x, 1) = !CodCidadao
       aProp(x, 2) = !nomecidadao
       aProp(x, 3) = !tipoprop
       aProp(x, 4) = !principal
       Sql = "SELECT  espolio.codigo, tipousuario.nome FROM  espolio INNER JOIN  tipousuario ON espolio.tipo = tipousuario.codigo where espolio.codigo=" & !CodCidadao
       Set RdoE = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       If RdoE.RowCount > 0 Then
          aProp(x, 5) = 1
          aProp(x, 6) = "(" & RdoE!Nome & ")"
       Else
          aProp(x, 5) = 0
          aProp(x, 6) = ""
       End If
       RdoE.Close
       x = x + 1
      .MoveNext
    Loop
   .Close
End With
QtdeProp = x - 1
prop = aProp
End Sub

Public Sub CarregaProprietarioCondominio(nCodigo As Integer)
Dim RdoC As rdoResultset
Dim x As Integer

Sql = "SELECT CONDOMINIO.CD_PROP, CIDADAO.CODCIDADAO, CIDADAO.NOMECIDADAO "
Sql = Sql & "FROM CONDOMINIO INNER JOIN CIDADAO ON CONDOMINIO.CD_PROP = CIDADAO.CODCIDADAO "
Sql = Sql & "Where CONDOMINIO.CD_CODIGO = " & nCodigo
Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoC
    x = 1
    Do Until .EOF
       aProp(x, 1) = !CodCidadao
       aProp(x, 2) = !nomecidadao
       aProp(x, 3) = "P"
       aProp(x, 4) = 1
       x = x + 1
      .MoveNext
    Loop
   .Close
End With
QtdeProp = x - 1
prop = aProp
End Sub

Public Sub CarregaNomeCondominio(nDistrito As Integer, nSetor As Integer, nQuadra As Integer, nLote As Integer, nSeq As Integer)
Dim RdoC As rdoResultset

Sql = "SELECT CD_CODIGO, CD_NOMECOND "
Sql = Sql & "FROM CONDOMINIO "
Sql = Sql & "WHERE CD_DISTRITO=" & nDistrito & " AND "
Sql = Sql & "CD_SETOR=" & nSetor & " AND "
Sql = Sql & "CD_QUADRA=" & nQuadra & " AND "
Sql = Sql & "CD_LOTE=" & nLote & " AND "
Sql = Sql & "CD_SEQ=" & nSeq
Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoC
    If .RowCount = 0 Then
       CodCondominio = 0
       NomeCondominio = "Não Selecionado"
    Else
       CodCondominio = !CD_CODIGO
       NomeCondominio = !cd_nomecond
    End If
   .Close
End With

End Sub

Public Sub CarregaCondominio(nValor As Long)
Dim RdoC As rdoResultset, RdoAux As rdoResultset
Dim RdoAux2 As rdoResultset

Sql = "SELECT CONDOMINIO.CD_CODIGO, CONDOMINIO.CD_NOMECOND, CONDOMINIO.CD_DISTRITO, CONDOMINIO.CD_SETOR,"
Sql = Sql & "CONDOMINIO.CD_QUADRA, CONDOMINIO.CD_LOTE, CONDOMINIO.CD_SEQ, CONDOMINIO.CD_NUM, CONDOMINIO.CD_COMPL,"
Sql = Sql & "CONDOMINIO.CD_UF,CONDOMINIO.CD_CODCIDADE, CONDOMINIO.CD_CODBAIRRO, CONDOMINIO.CD_CEP,"
Sql = Sql & "CONDOMINIO.CD_QUADRAS, CONDOMINIO.CD_LOTES, CONDOMINIO.CD_AREATERRENO, CONDOMINIO.CD_CODUSOTERRENO,"
Sql = Sql & "CONDOMINIO.CD_CODBENF, CONDOMINIO.CD_CODTOPOG, CONDOMINIO.CD_CODCATEGPROP, CONDOMINIO.CD_CODSITUACAO,"
Sql = Sql & "CONDOMINIO.CD_CODPEDOL, CONDOMINIO.CD_AREATOTCONSTR,  CONDOMINIO.CD_NUMUNID, "
Sql = Sql & "CD_PROP,CD_FRACAO FROM CONDOMINIO WHERE CD_CODIGO=" & nValor
Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoC
    If .RowCount = 0 Then
       CodCondominio = 999
    Else
       CodCondominio = !CD_CODIGO
       NomeCondominio = !cd_nomecond
       Distrito = !CD_DISTRITO
       Setor = !CD_SETOR
       Quadra = !CD_QUADRA
       Lote = !CD_LOTE
       Seq = !CD_SEQ
       Li_Num = !CD_NUM
       Li_Compl = SubNull(!CD_COMPL)
       Li_CodBairro = !CD_CODBAIRRO
       Li_Cep = SubNull(!CD_CEP)
       Li_Quadras = SubNull(!CD_QUADRAS)
       Li_Lotes = SubNull(!CD_LOTES)
       Dt_AreaTerreno = !CD_AREATERRENO
       Dt_CodUsoTerreno = !CD_CODUSOTERRENO
       Dt_CodBenf = !CD_CODBENF
       Dt_CodTopog = !CD_CODTOPOG
       Dt_CodCategProp = !CD_CODCATEGPROP
       Dt_CodSituacao = !CD_CODSITUACAO
       Dt_CodPedol = !CD_CODPEDOL
       NumUnidades = !CD_NUMUNID
       AreaConstruida = !CD_AREATOTCONSTR
       Sql = "SELECT CODLOGR FROM FACEQUADRA WHERE CODDISTRITO=" & Distrito & " AND CODSETOR=" & Setor
       Sql = Sql & " AND CODQUADRA=" & Quadra & " AND CODFACE=" & Seq
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
          If .RowCount > 0 Then
              CodLogr = !CodLogr
          End If
         .Close
       End With
       FracaoIdeal = FormatNumber(SubNull(!CD_FRACAO), 2)
       If IsNull(!CD_PROP) Then
          CodProp = 0
       Else
          CodProp = !CD_PROP
       End If
       Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & CodProp
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       If RdoAux.RowCount = 0 Then
            Proprietario = ""
       Else
            Proprietario = RdoAux!nomecidadao
       End If
       RdoAux.Close
       
    End If
   .Close
End With

End Sub

Property Let CodLogradouro(sValor As String)
sCodLogradouro = sValor
End Property

Property Get CodLogradouro() As String
CodLogradouro = sCodLogradouro
End Property

Property Let Endereco(sValor As String)
sEndereco = sValor
End Property

Property Get Endereco() As String
Endereco = sEndereco
End Property

Property Let Numero(sValor As String)
sNumero = sValor
End Property

Property Get Numero() As String
Numero = sNumero
End Property

Property Let Complemento(sValor As String)
sComplemento = sValor
End Property

Property Get Complemento() As String
Complemento = sComplemento
End Property


Property Let CodBairro(sValor As String)
sCodBairro = sValor
End Property

Property Get CodBairro() As String
CodBairro = sCodBairro
End Property

Property Let Bairro(sValor As String)
sBairro = sValor
End Property

Property Get Bairro() As String
Bairro = sBairro
End Property

Property Let CodCidade(sValor As String)
sCodCidade = sValor
End Property

Property Get CodCidade() As String
CodCidade = sCodCidade
End Property

Property Let Cidade(sValor As String)
sCidade = sValor
End Property

Property Get Cidade() As String
Cidade = sCidade
End Property

Property Let UF(sValor As String)
sUF = sValor
End Property

Property Get UF() As String
UF = sUF
End Property

Property Let Cep(sValor As String)
sCep = sValor
End Property

Property Get Cep() As String
Cep = sCep
End Property





Property Let CodigoImovel(nValor As Long)
nCodImovel = nValor
End Property

Property Get CodigoImovel() As Long
CodigoImovel = nCodImovel
End Property

Property Let DV(nValor As Integer)
nDV = nValor
End Property

Property Get DV() As Integer
DV = nDV
End Property

Property Let Distrito(nValor As Integer)
nDistrito = nValor
End Property

Property Get Distrito() As Integer
Distrito = nDistrito
End Property

Property Let Setor(nValor As Integer)
nSetor = nValor
End Property

Property Get Setor() As Integer
Setor = nSetor
End Property

Property Let Quadra(nValor As Integer)
nQuadra = nValor
End Property

Property Get Quadra() As Integer
Quadra = nQuadra
End Property

Property Let Lote(nValor As Integer)
nLote = nValor
End Property

Property Get Lote() As Integer
Lote = nLote
End Property

Property Let Seq(nValor As Integer)
nSeq = nValor
End Property

Property Get Seq() As Integer
Seq = nSeq
End Property

Property Let Unidade(nValor As Integer)
nUnidade = nValor
End Property

Property Get Unidade() As Integer
Unidade = nUnidade
End Property

Property Let SubUnidade(nValor As Integer)
nSubUnidade = nValor
End Property

Property Get SubUnidade() As Integer
SubUnidade = nSubUnidade
End Property

Property Let Inscricao(sValor As String)
sInscricao = sValor
End Property

Property Get Inscricao() As String
Inscricao = sInscricao
End Property

Property Let Inativo(bValor As Boolean)
bInativo = bValor
End Property

Property Get Inativo() As Boolean
Inativo = bInativo
End Property

Property Let Li_Num(nValor As Integer)
nLi_Num = nValor
End Property

Property Get Li_Num() As Integer
Li_Num = nLi_Num
End Property

Property Let Li_Compl(sValor As String)
sLi_Compl = sValor
End Property

Property Get Li_Compl() As String
Li_Compl = sLi_Compl
End Property

Property Let Li_CodBairro(nValor As Integer)
nLi_CodBairro = nValor
End Property

Property Get Li_CodBairro() As Integer
Li_CodBairro = nLi_CodBairro
End Property

Property Let Li_Quadras(sValor As String)
sLi_Quadras = sValor
End Property

Property Get Li_Quadras() As String
Li_Quadras = sLi_Quadras
End Property

Property Let Li_Lotes(sValor As String)
sLi_Lotes = sValor
End Property

Property Get Li_Lotes() As String
Li_Lotes = sLi_Lotes
End Property

Property Let Dt_AreaTerreno(nValor As Double)
nDt_AreaTerreno = nValor
End Property

Property Get Dt_AreaTerreno() As Double
Dt_AreaTerreno = nDt_AreaTerreno
End Property

Property Let Dt_CodUsoTerreno(nValor As Integer)
nDt_CodUsoTerreno = nValor
End Property

Property Get Dt_CodUsoTerreno() As Integer
Dt_CodUsoTerreno = nDt_CodUsoTerreno
End Property

Property Let Dt_CodBenf(nValor As Integer)
nDt_CodBenf = nValor
End Property

Property Get Dt_CodBenf() As Integer
Dt_CodBenf = nDt_CodBenf
End Property

Property Let Dt_CodTopog(nValor As Integer)
nDt_CodTopog = nValor
End Property

Property Get Dt_CodTopog() As Integer
Dt_CodTopog = nDt_CodTopog
End Property

Property Let Dt_CodCategProp(nValor As Integer)
nDt_CodCategProp = nValor
End Property

Property Get Dt_CodCategProp() As Integer
Dt_CodCategProp = nDt_CodCategProp
End Property

Property Let Dt_CodSituacao(nValor As Integer)
nDt_CodSituacao = nValor
End Property

Property Get Dt_CodSituacao() As Integer
Dt_CodSituacao = nDt_CodSituacao
End Property

Property Let Dt_CodPedol(nValor As Integer)
nDt_CodPedol = nValor
End Property

Property Get Dt_CodPedol() As Integer
Dt_CodPedol = nDt_CodPedol
End Property

Property Let Dt_NumAgua(sValor As String)
sDt_NumAgua = sValor
End Property

Property Get Dt_NumAgua() As String
Dt_NumAgua = sDt_NumAgua
End Property

Property Let Dt_FracaoIdeal(nValor As Double)
nDt_FracaoIdeal = nValor
End Property

Property Get Dt_FracaoIdeal() As Double
Dt_FracaoIdeal = nDt_FracaoIdeal
End Property

Property Let Ee_TipoEnd(nValor As Integer)
nEe_TipoEnd = nValor
End Property

Property Get Ee_TipoEnd() As Integer
Ee_TipoEnd = nEe_TipoEnd
End Property

Property Let Dc_QtdeEdif(nValor As Integer)
nDc_QtdeEdif = nValor
End Property

Property Get Dc_QtdeEdif() As Integer
Dc_QtdeEdif = nDc_QtdeEdif
End Property

Property Let Dc_QtdePav(nValor As Integer)
nDc_QtdePav = nValor
End Property

Property Get Dc_QtdePav() As Integer
Dc_QtdePav = nDc_QtdePav
End Property

Property Let AbrevTitLog(sValor As String)
sAbrevTitLog = sValor
End Property

Property Get AbrevTitLog() As String
AbrevTitLog = sAbrevTitLog
End Property

Property Let AbrevTipoLog(sValor As String)
sAbrevTipoLog = sValor
End Property

Property Get AbrevTipoLog() As String
AbrevTipoLog = sAbrevTipoLog
End Property

Property Let NomeLogradouro(sValor As String)
sNomeLogradouro = sValor
End Property

Property Get NomeLogradouro() As String
NomeLogradouro = sNomeLogradouro
End Property

Property Let CodLogr(nValor As Long)
nCodLogr = nValor
End Property

Property Get CodLogr() As Long
CodLogr = nCodLogr
End Property

Property Let FracaoIdeal(nValor As Double)
nFracao = nValor
End Property

Property Get FracaoIdeal() As Double
FracaoIdeal = nFracao
End Property

Property Let CodProp(nValor As Long)
nCodProp = nValor
End Property

Property Get CodProp() As Long
CodProp = nCodProp
End Property

Property Let Proprietario(sValor As String)
sNomeProp = sValor
End Property

Property Get Proprietario() As String
Proprietario = sNomeProp
End Property

Property Let TipoMat(sValor As String)
sTipoMat = sValor
End Property

Property Get TipoMat() As String
TipoMat = sTipoMat
End Property

Property Let NumMat(nValor As Long)
nNumMat = nValor
End Property

Property Get NumMat() As Long
NumMat = nNumMat
End Property

Property Let Imune(nValor As Boolean)
bImune = nValor
End Property

Property Get Imune() As Boolean
Imune = bImune
End Property

Property Let Conjugado(nValor As Boolean)
bConjugado = nValor
End Property

Property Get Conjugado() As Boolean
Conjugado = bConjugado
End Property

Property Let ResideImovel(nValor As Boolean)
bReside = nValor
End Property

Property Get ResideImovel() As Boolean
ResideImovel = bReside
End Property

Property Let IsentoCIP(nValor As Boolean)
bCip = nValor
End Property

Property Get IsentoCIP() As Boolean
IsentoCIP = bCip
End Property


Property Let CodCondominio(nValor As Integer)
nCodCondominio = nValor
End Property

Property Get CodCondominio() As Integer
CodCondominio = nCodCondominio
End Property

Property Let Li_Cep(sValor As String)
sLi_Cep = sValor
End Property

Property Get Li_Cep() As String
Li_Cep = sLi_Cep
End Property

Property Let DescBenfeitoria(sValor As String)
sDescBenfeitoria = sValor
End Property

Property Get DescBenfeitoria() As String
DescBenfeitoria = sDescBenfeitoria
End Property

Property Let DescSituacao(sValor As String)
sDescSituacao = sValor
End Property

Property Get DescSituacao() As String
DescSituacao = sDescSituacao
End Property

Property Let DescUsoTerreno(sValor As String)
sDescUsoTerreno = sValor
End Property

Property Get DescUsoTerreno() As String
DescUsoTerreno = sDescUsoTerreno
End Property

Property Let DescCategProp(sValor As String)
sDescCategProp = sValor
End Property

Property Get DescCategProp() As String
DescCategProp = sDescCategProp
End Property

Property Let DescPedologia(sValor As String)
sDescPedologia = sValor
End Property

Property Get DescPedologia() As String
DescPedologia = sDescPedologia
End Property

Property Let DescTopografia(sValor As String)
sDescTopografia = sValor
End Property

Property Get DescTopografia() As String
DescTopografia = sDescTopografia
End Property

Property Let CodPropPrincipal(nValor As Long)
nCodPropPrincipal = nValor
End Property

Property Get CodPropPrincipal() As Long
CodPropPrincipal = nCodPropPrincipal
End Property

Property Let NomePropPrincipal(sValor As String)
sNomePropPrincipal = sValor
End Property

Property Get NomePropPrincipal() As String
NomePropPrincipal = sNomePropPrincipal
End Property

Property Let DescBairro(sValor As String)
sDescBairro = sValor
End Property

Property Get DescBairro() As String
DescBairro = sDescBairro
End Property

Property Let DescDistrito(sValor As String)
sDescDistrito = sValor
End Property

Property Get DescDistrito() As String
DescDistrito = sDescDistrito
End Property

Property Let Ee_CodLog(nValor As Long)
nEe_CodLog = nValor
End Property

Property Get Ee_CodLog() As Long
Ee_CodLog = nEe_CodLog
End Property

Property Let Ee_NomeLog(sValor As String)
sEe_NomeLog = sValor
End Property

Property Get Ee_NomeLog() As String
Ee_NomeLog = sEe_NomeLog
End Property

Property Let Ee_NumImovel(nValor As Integer)
nEe_NumImovel = nValor
End Property

Property Get Ee_NumImovel() As Integer
Ee_NumImovel = nEe_NumImovel
End Property

Property Let Ee_Complemento(sValor As String)
sEe_Complemento = sValor
End Property

Property Get Ee_Complemento() As String
Ee_Complemento = sEe_Complemento
End Property

Property Let Ee_Cep(sValor As String)
sEe_Cep = sValor
End Property

Property Get Ee_Cep() As String
Ee_Cep = sEe_Cep
End Property

Property Let Ee_Uf(sValor As String)
sEe_UF = sValor
End Property

Property Get Ee_Uf() As String
Ee_Uf = sEe_UF
End Property

Property Let Ee_Cidade(nValor As Integer)
nEe_Cidade = nValor
End Property

Property Get Ee_Cidade() As Integer
Ee_Cidade = nEe_Cidade
End Property

Property Let Ee_Bairro(nValor As Integer)
nEe_Bairro = nValor
End Property

Property Get Ee_Bairro() As Integer
Ee_Bairro = nEe_Bairro
End Property

Property Let EnderecoCompleto(sValor As String)
sEnderecoCompleto = sValor
End Property

Property Get EnderecoCompleto() As String
EnderecoCompleto = sEnderecoCompleto
End Property

Property Let EnderecoCompleto2(sValor As String)
sEnderecoCompleto2 = sValor
End Property

Property Get EnderecoCompleto2() As String
EnderecoCompleto2 = sEnderecoCompleto2
End Property

Property Let Testada(vValor As Variant)
vTestada = vValor
End Property

Property Get Testada() As Variant
Testada = vTestada
End Property

Property Let QtdeTestada(nValor As Integer)
nQtdeTestada = nValor
End Property

Property Get QtdeTestada() As Integer
QtdeTestada = nQtdeTestada
End Property

Property Let prop(vValor As Variant)
vProp = vValor
End Property

Property Get prop() As Variant
prop = vProp
End Property

Property Let QtdeProp(nValor As Integer)
nQtdeProp = nValor
End Property

Property Get QtdeProp() As Integer
QtdeProp = nQtdeProp
End Property

Property Let NomeCondominio(sValor As String)
sNomeCondominio = sValor
End Property

Property Get NomeCondominio() As String
NomeCondominio = sNomeCondominio
End Property

Property Let TestadaCond(vValor As Variant)
vTestadaCond = vValor
End Property

Property Get TestadaCond() As Variant
TestadaCond = vTestadaCond
End Property

Property Let QtdeTestadaCond(nValor As Integer)
nQtdeTestadaCond = nValor
End Property

Property Get QtdeTestadaCond() As Integer
QtdeTestadaCond = nQtdeTestadaCond
End Property

Property Let NumUnidades(nValor As Integer)
nNumUnidades = nValor
End Property

Property Get NumUnidades() As Integer
NumUnidades = nNumUnidades
End Property

Property Let AreaConstruida(nValor As Double)
nAreaConstruida = nValor
End Property

Property Get AreaConstruida() As Double
AreaConstruida = nAreaConstruida
End Property

Public Sub RetornaEndereco(nCodReduz As Long, TySeqEndereco As SeqEndereco, TyTipoEndereco As TipoEndereco)
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim sCodLogradouro, sEndereco As String, sNumero As String, sComplemento As String, sCodBairro As String
Dim sBairro As String, sCodCidade As String, sCidade As String, sCep As String, sEstado As String, sValorTotal As String
DoEvents
If TySeqEndereco = Imobiliario Then
    Sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount = 0 Then Exit Sub
        If TyTipoEndereco = Localizacao Then 'imovel
            sCodLogradouro = SubNull(!CodLogr)
            Sql = "select logradouro,endereco_resumido from vwlogradouro where codlogradouro=" & !CodLogr
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            'sEndereco = SubNull(RdoAux3!Logradouro)
            sEndereco = SubNull(RdoAux3!endereco_resumido)
            If sEndereco = "" Then sEndereco = SubNull(!Logradouro)
            RdoAux3.Close
            sNumero = SubNull(!Li_Num)
            sComplemento = SubNull(!Li_Compl)
            sCodBairro = SubNull(!Li_CodBairro)
            If Val(sCodBairro) <> 999 Then
                sBairro = SubNull(!DescBairro)
            Else
                sCodBairro = ""
                Bairro = ""
            End If
            sCodCidade = SubNull(!LI_CODCIDADE)
            sCidade = SubNull(!descCidade)
            If Val(sCodLogradouro) > 0 Then
                sCep = RetornaNumero(RetornaCEP(Val(sCodLogradouro), Val(sNumero)))
            Else
                sCep = "14870000"
            End If
            sEstado = SubNull(!li_uf)
        ElseIf TyTipoEndereco = Cadastrocidadao Then 'proprietario
            Sql = "select * from vwfullcidadao where codcidadao=" & !CodCidadao
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If (SubNull(!etiqueta) = "S" And SubNull(!etiqueta2) = "S") Then
                    'sEndereco = SubNull(!Endereco)
                    sEndereco = SubNull(!endereco_resumido)
                    If sEndereco = "" Then sEndereco = SubNull(!Endereco)
                    sNumero = SubNull(!NUMIMOVEL)
                    sComplemento = SubNull(!Complemento)
                    sCodBairro = SubNull(!CodBairro)
                    If Val(sCodBairro) <> 999 Then
                        sBairro = SubNull(!DescBairro)
                    Else
                        sCodBairro = ""
                        Bairro = ""
                    End If
                    sCodCidade = SubNull(!CodCidade)
                    sCidade = SubNull(!descCidade)
                    'If sCidade = "" Then
                    '    sCidade = SubNull(!desccidade2)
                    'End If
                    If Val(sCodLogradouro) > 0 Then
                        sCep = RetornaNumero(RetornaCEP(Val(sCodLogradouro), Val(sNumero)))
                    Else
                        sCep = RetornaNumero(SubNull(!Cep))
                    End If
                    sEstado = SubNull(!SiglaUF)
                    GoTo fim
                End If
            
                If (IsNull(!etiqueta2) Or SubNull(!etiqueta2) = "N") Then
                    'sEndereco = SubNull(!Endereco)
                    sEndereco = SubNull(!endereco_resumido)
                    If sEndereco = "" Then sEndereco = SubNull(!Endereco)
                    sNumero = SubNull(!NUMIMOVEL)
                    sComplemento = SubNull(!Complemento)
                    sCodBairro = SubNull(!CodBairro)
                    If Val(sCodBairro) <> 999 Then
                        sBairro = SubNull(!DescBairro)
                    Else
                        sCodBairro = ""
                        Bairro = ""
                    End If
                    sCodCidade = SubNull(!CodCidade)
                    sCidade = SubNull(!descCidade)
                    'If sCidade = "" Then
                    '    sCidade = SubNull(!desccidade2)
                    'End If
                    If Val(sCodLogradouro) > 0 Then
                        sCep = RetornaNumero(RetornaCEP(Val(sCodLogradouro), Val(sNumero)))
                    Else
                        sCep = RetornaNumero(SubNull(!Cep))
                    End If
                    sEstado = SubNull(!SiglaUF)
                Else
                    sEndereco = SubNull(!NomeLogradouroC)
                    If sEndereco = "" Then
                        'sEndereco = SubNull(!Enderecoc)
                        sEndereco = SubNull(!Endereco_resumidoC)
                    End If
                    sNumero = SubNull(!NUMIMOVEL2)
                    sComplemento = SubNull(!Complemento2)
                    sCodBairro = SubNull(!CodBairro2)
                    If Val(sCodBairro) <> 999 Then
                        sBairro = SubNull(!DescBairroC)
                    Else
                        sCodBairro = ""
                        Bairro = ""
                    End If
                    sCodCidade = SubNull(!codcidade2)
                    sCidade = SubNull(!desccidadeC)
                    If Val(SubNull(!codcidade2)) = 413 Then
                        sCep = RetornaNumero(RetornaCEP(Val(!codcidade2), Val(!NUMIMOVEL2)))
                    Else
                        sCep = RetornaNumero(SubNull(!Cep2))
                    End If
                    sEstado = SubNull(!SiglaUF2)
                End If
            End With
        ElseIf TyTipoEndereco = Entrega Then 'entrega
            sCodLogradouro = SubNull(!Ee_CodLog)
            If Val(sCodLogradouro) = 0 Then
                If SubNull(!Ee_NomeLog) <> "" Then
                    Sql = "select codlogradouro from vwlogradouro where logradouro='" & Mask(RemovePonto(Trim(!Ee_NomeLog))) & "'"
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux3.RowCount > 0 Then
                        sCodLogradouro = SubNull(RdoAux3!CodLogradouro)
                        sEndereco = RemovePonto(Trim(!Ee_NomeLog))
                    Else
                        sEndereco = !Ee_NomeLog
                    End If
                    RdoAux3.Close
                End If
            Else
                Sql = "select logradouro,endereco_resumido from vwlogradouro where codlogradouro=" & !Ee_CodLog
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                sEndereco = SubNull(RdoAux3!Logradouro)
                sEndereco = SubNull(RdoAux3!endereco_resumido)
                RdoAux3.Close
            End If
            sNumero = SubNull(!Ee_NumImovel)
            sComplemento = SubNull(!Ee_Complemento)
            sCodBairro = SubNull(!Ee_Bairro)
            If Val(sCodBairro) <> 999 Then
                sBairro = SubNull(!BairroEE)
            Else
                sCodBairro = ""
                Bairro = ""
            End If
            sCodCidade = SubNull(!Ee_Cidade)
            If sCodCidade = "" Then sCodCidade = "413"
            If IsNull(!Ee_Uf) Then
                sEstado = "SP"
            Else
                sEstado = !Ee_Uf
            End If
            Sql = "select desccidade from cidade where siglauf='" & sEstado & "' and codcidade=" & Val(sCodCidade)
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux3.RowCount > 0 Then
                sCidade = SubNull(RdoAux3!descCidade)
            Else
                sCidade = ""
            End If
            RdoAux3.Close
            
            If Val(sCodLogradouro) > 0 Then
                sCep = RetornaNumero(RetornaCEP(Val(sCodLogradouro), Val(sNumero)))
            Else
                sCep = RetornaNumero(SubNull(!Ee_Cep))
                If Val(sCep) = 0 And sCidade = "JABOTICABAL" Then
                    sCep = "14870000"
                ElseIf Trim(sCep) = "" Then
                    sCep = "00000000"
                End If
            End If
            sEstado = SubNull(!Ee_Uf)
            If Trim(sEstado) = "" Then
                sEstado = "SP"
            End If
        End If
       .Close
    End With
ElseIf TySeqEndereco = Mobiliario Then
    Sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        sEndereco = SubNull(!eenomelogr)
        If IsNull(!eenomelogr) Then
            TyTipoEndereco = Localizacao
        End If
        If TyTipoEndereco = Localizacao Then
            'local da empresa
            sCodLogradouro = SubNull(!CodLogradouro)
            If Val(sCodLogradouro) > 0 Then
                sEndereco = Trim(SubNull(!Endereco))
            Else
                sEndereco = Trim(SubNull(!Logradouro))
            End If
            sNumero = SubNull(!Numero)
            sComplemento = SubNull(!Complemento)
            sCodBairro = SubNull(!CodBairro)
            sBairro = SubNull(!DescBairro)
            sCodCidade = SubNull(!CodCidade)
            sCidade = SubNull(!descCidade)
            If Val(sCodLogradouro) > 0 Then
                sCep = RetornaNumero(RetornaCEP(Val(sCodLogradouro), Val(sNumero)))
            Else
                sCep = RetornaNumero(SubNull(!Cep))
                If Val(sCep) = 0 And sCidade = "JABOTICABAL" Then
                    sCep = "14870000"
                End If
            End If
            sEstado = SubNull(!SiglaUF)
            If sEstado = "" And sCidade = "JABOTICABAL" Then
                sEstado = "SP"
            End If
        Else
            'endereco entrega
            sCodLogradouro = SubNull(!eecodlogr)
            
            sNumero = SubNull(!eenumero)
            sComplemento = SubNull(!eecompl)
            sCodBairro = SubNull(!eecodbairro)
            Sql = "select descbairro from bairro where siglauf='" & SubNull(!EeUf) & "' and codcidade=" & Val(SubNull(!eecodcidade)) & " and codbairro=" & Val(SubNull(!eecodbairro))
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux3.RowCount > 0 Then
                sBairro = SubNull(RdoAux3!DescBairro)
            Else
                sBairro = ""
            End If
            sCodCidade = SubNull(!eecodcidade)
            Sql = "select desccidade from cidade where siglauf='" & SubNull(!EeUf) & "' and codcidade=" & Val(SubNull(!eecodcidade))
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux3.RowCount > 0 Then
                sCidade = SubNull(RdoAux3!descCidade)
            Else
                sCidade = SubNull(!eeDesccidade)
            End If
            RdoAux3.Close
            If Val(sCodLogradouro) > 0 Then
                sCep = RetornaNumero(RetornaCEP(Val(sCodLogradouro), Val(sNumero)))
            Else
                sCep = RetornaNumero(SubNull(!EeCep))
                If Val(sCep) = 0 And sCidade = "JABOTICABAL" Then
                    sCep = "14870000"
                End If
            End If
            sEstado = SubNull(!EeUf)
            If sEstado = "" And sCidade = "JABOTICABAL" Then
                sEstado = "SP"
            End If
        End If
        If Val(sCodBairro) = 999 Then
            sCodBairro = ""
            sBairro = ""
        End If
       
       .Close
    End With
ElseIf TySeqEndereco = cidadao Then
    Sql = "select * from vwfullcidadao where codcidadao=" & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        sEndereco = SubNull(!Endereco)
        sNumero = SubNull(!NUMIMOVEL)
        sCodBairro = SubNull(!CodBairro)
        sComplemento = SubNull(!Complemento)
        sBairro = SubNull(!DescBairro)
        If Val(sCodBairro) = 999 Then
            sCodBairro = ""
            sBairro = ""
        End If
        sCodCidade = SubNull(!CodCidade)
        sCidade = SubNull(!descCidade)
        sCep = RetornaNumero(SubNull(!Cep))
        sEstado = SubNull(!SiglaUF)
       .Close
    End With
End If

fim:
CodLogradouro = sCodLogradouro
Endereco = sEndereco
Numero = sNumero
Complemento = sComplemento
CodBairro = sCodBairro
Bairro = sBairro
CodCidade = sCodCidade
Cidade = sCidade
UF = sEstado
Cep = sCep

End Sub

