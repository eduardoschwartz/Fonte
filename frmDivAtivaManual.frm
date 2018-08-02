VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmDivAtivaManual 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inscrição na Divida Ativa"
   ClientHeight    =   3630
   ClientLeft      =   3105
   ClientTop       =   2925
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5550
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   345
      Left            =   3270
      TabIndex        =   8
      ToolTipText     =   "Gravar os Dados"
      Top             =   3150
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmDivAtivaManual.frx":0000
      PICN            =   "frmDivAtivaManual.frx":001C
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
      Left            =   4395
      TabIndex        =   0
      ToolTipText     =   "Sair da Tela"
      Top             =   3150
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
      MICON           =   "frmDivAtivaManual.frx":03C1
      PICN            =   "frmDivAtivaManual.frx":03DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   2385
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   4207
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Ano      |^Lanc |^Seq|^Par  |^Com|^Sit   |^Vencimento  |^D |^A |>Principal     "
   End
   Begin esMaskEdit.esMaskedEdit mskDataInsc 
      Height          =   285
      Left            =   1410
      TabIndex        =   7
      Top             =   3165
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmDivAtivaManual.frx":044B
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
   Begin VB.Label lblPag 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1425
      TabIndex        =   6
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label lblLivro 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1410
      TabIndex        =   5
      Top             =   2580
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inscrição..:"
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   3195
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da Página....:"
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Livro.......:"
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   2580
      Width           =   1200
   End
End
Attribute VB_Name = "frmDivAtivaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim sAno As String, sLanc As String, sSeq As String, sParc As String
Dim sComp As String, nCodReduz As Long, nNumCert As Integer
Dim nLivro As Integer, nPagina As Integer, sTypeBook As String, nPos As Integer


Private Sub cmdGravar_Click()
Dim sTipoDivida As String, nCDA As Long, sRG As String, sCPF As String
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String, sComplementoEntrega As String
Dim sBairro As String, sCidade As String, sUF As String, sCep As String, sEnderecoEntrega As String, nNumEntrega As Integer, sBairroEntrega As String
Dim sCidadeEntrega As String, sUFEntrega As String, sCepEntrega As String, xImovel As clsImovel, sQuadras As String, sLotes As String, nTipoEnd As Integer
Dim sTipoProp As String

Set xImovel = New clsImovel

If Not IsDate(mskDataInsc.Text) Then
    MsgBox "Data inválida.", vbCritical, "atenção"
    Exit Sub
End If

If MsgBox("Confirme a inscrição destes lançamentos na Divida Ativa !.", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

ConectaIntegrativa

Sql = "SELECT max(numcertidao) as maximo from debitoparcela where numerolivro=" & Val(lblLivro.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!MAXIMO) Then
        nNumCert = 1
    Else
    nNumCert = !MAXIMO + 1
    End If
   .Close
End With


If nCodReduz < 100000 Then
    sTipoDivida = "Imobiliário"
    Sql = "select * from vwfullimovel where codreduzido=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        sNome = !nomecidadao
        sInscricao = !Inscricao
        sRG = SubNull(!rg)
        sCPF = SubNull(!CPF)
        If Trim(sCPF) = "" Then
            sCPF = SubNull(!Cnpj)
        End If
        sEndereco = !Logradouro
        nNumero = !Li_Num
        sCep = RetornaNumero(RetornaCEP(!CodLogr, !Li_Num))
        sComplemento = SubNull(!Li_Compl)
        sBairro = SubNull(!DescBairro)
        sCidade = "JABOTICABAL"
        sUF = "SP"
        sQuadras = Left(SubNull(!Li_Quadras), 5)
        sLotes = Left(SubNull(!Li_Lotes), 20)
        nTipoEnd = RdoAux!Ee_TipoEnd
       .Close
    End With
    If nTipoEnd = 0 Then
        sEnderecoEntrega = sEndereco
        nNumEntrega = nNumero
        sCepEntrega = sCep
        sComplementoEntrega = sComplemento
        sBairroEntrega = sBairro
        sCidadeEntrega = sCidade
        sUFEntrega = sUF
    Else
        xImovel.RetornaEndereco nCodReduz, Imobiliario, IIf(nTipoEnd = 1, cadastrocidadao, Entrega)
        sEnderecoEntrega = xImovel.Endereco
        nNumEntrega = xImovel.Numero
        sCepEntrega = xImovel.Cep
        sComplementoEntrega = xImovel.Complemento
        sBairroEntrega = xImovel.Bairro
        sCidadeEntrega = xImovel.Cidade
        sUFEntrega = xImovel.UF
    End If
ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
    sTipoDivida = "Mobiliário"
    Sql = "select * from vwfullempresa3 where codigomob=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        sNome = !RazaoSocial
        sInscricao = nCodReduz
        sRG = SubNull(!INSCESTADUAL)
        If Trim(sRG) = "" Then
            sRG = SubNull(!rg)
        End If
        sCPF = SubNull(!Cnpj)
        If Trim(sCPF) = "" Then
            sCPF = SubNull(!CPF)
        End If
        sEndereco = !Logradouro
        nNumero = !Numero
        sCep = SubNull(!Cep)
        sComplemento = SubNull(!Complemento)
        sBairro = SubNull(!DescBairro)
        sCidade = SubNull(!desccidade)
        sUF = SubNull(!SiglaUF)
        sQuadras = ""
        sLotes = ""
       .Close
    End With
    Sql = "SELECT MOBILIARIOENDENTREGA.CODMOBILIARIO, MOBILIARIOENDENTREGA.CODLOGRADOURO, MOBILIARIOENDENTREGA.NOMELOGRADOURO, "
    Sql = Sql & "MOBILIARIOENDENTREGA.NUMIMOVEL,MOBILIARIOENDENTREGA.COMPLEMENTO, MOBILIARIOENDENTREGA.UF,MOBILIARIOENDENTREGA.CODCIDADE,"
    Sql = Sql & "MOBILIARIOENDENTREGA.CODBAIRRO, MOBILIARIOENDENTREGA.CEP,MOBILIARIOENDENTREGA.DESCBAIRRO, MOBILIARIOENDENTREGA.DESCCIDADE, BAIRRO.DESCBAIRRO AS DESCBAIRRO2,"
    Sql = Sql & "CIDADE.DESCCIDADE AS DESCCIDADE2 FROM dbo.bairro INNER JOIN dbo.cidade ON dbo.bairro.siglauf = dbo.cidade.siglauf AND dbo.bairro.codcidade = dbo.cidade.codcidade RIGHT OUTER JOIN "
    Sql = Sql & "dbo.mobiliarioendentrega ON dbo.bairro.siglauf = dbo.mobiliarioendentrega.uf AND dbo.bairro.codcidade = dbo.mobiliarioendentrega.codcidade AND dbo.Bairro.CodBairro = dbo.mobiliarioendentrega.CodBairro "
    Sql = Sql & "Where CODMOBILIARIO = " & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If !CodLogradouro = 0 Then
                sEnderecoEntrega = !NomeLogradouro
                nNumEntrega = !NUMIMOVEL
                If !CodBairro = 0 Then
                    sBairroEntrega = SubNull(!DescBairro)
                ElseIf !CodBairro = 999 Then
                    sBairroEntrega = ""
                Else
                    sBairroEntrega = SubNull(!DescBairro2)
                End If
                sCidadeEntrega = SubNull(!desccidade)
                If Trim(sCidadeEntrega) = "" Then
                    sCidadeEntrega = SubNull(!desccidade2)
                End If
                If !CodCidade = 413 Then
                    sCepEntrega = Format(!Cep, "00000000")
                Else
                    sCepEntrega = RetornaNumero(RetornaCEP(!CodLogradouro, !NUMIMOVEL))
                End If
                sComplEntrega = SubNull(!Complemento)
                sUFEntrega = !UF
            Else
                Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sEnderecoEntrega = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                        nNumEntrega = SubNull(RdoAux!NUMIMOVEL)
                        sCepEntrega = RetornaNumero(RetornaCEP(RdoAux!CodLogradouro, RdoAux!NUMIMOVEL))
                    End If
                    sBairroEntrega = SubNull(RdoAux!DescBairro2)
                    sCidadeEntrega = SubNull(RdoAux!desccidade2)
                    sComplementoEntrega = SubNull(RdoAux!Complemento)
                    sUFEntrega = RdoAux!UF
                   .Close
                End With
            End If
        Else
            sEnderecoEntrega = sEndereco
            nNumEntrega = nNumero
            sCepEntrega = sCep
            sComplementoEntrega = sComplemento
            sBairroEntrega = sBairro
            sCidadeEntrega = sCidade
            sUFEntrega = sUF
        End If
       .Close
    End With

Else
    sTipoDivida = "Taxas Diversas"
End If

Sql = "INSERT CDAs(IdDevedor,SetorDevedor,DtInscricao,NroCertidao,NroLivro,NroFolha,DtGeracao) values("
Sql = Sql & nCodReduz & ",'" & sTipoDivida & "','" & Format(mskDataInsc.Text, "mm/dd/yyyy") & "'," & nNumCert & ","
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
    Sql = Sql & nCDA & ",'" & sTipoDivida & "'," & nCodReduz & ",'" & SubNull(sNome) & "','" & sInscricao & "','" & sCPF & "','" & sRG & "','"
    Sql = Sql & sCep & "','" & sEndereco & "'," & nNumero & ",'" & sComplemento & "','" & sBairro & "','" & sCidade & "','" & sUF & "','"
    Sql = Sql & sQuadras & "','" & sLotes & "','" & sCepEntrega & "','" & sEnderecoEntrega & "'," & nNumEntrega & ",'" & sComplementoEntrega & "','"
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
        Sql = Sql & nCDA & ",'Principal'," & !CodCidadao & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
        Sql = Sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
        Sql = Sql & Mask(SubNull(!desccidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
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
            Sql = Sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
            Sql = Sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
            Sql = Sql & Mask(SubNull(!desccidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
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
            Sql = Sql & nCDA & ",'" & sTipoProp & "'," & nCodReduz & ",'" & Mask(!nomecidadao) & "','" & sCPF & "','" & SubNull(!rg) & "','" & SubNull(!Cep) & "','"
            Sql = Sql & Mask(SubNull(!Endereco)) & "'," & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "','" & Mask(SubNull(!DescBairro)) & "','"
            Sql = Sql & Mask(SubNull(!desccidade)) & "','" & SubNull(!SiglaUF) & "','" & Format(Now, "mm/dd/yyyy") & "')"
            cnInt.Execute Sql, rdExecDirect
           .MoveNext
        Loop
       .Close
    End With
End If

For x = 1 To grdTemp.Rows - 1
    
    sAno = grdTemp.TextMatrix(x, 0)
    sLanc = Left$(grdTemp.TextMatrix(x, 1), 3)
    sSeq = grdTemp.TextMatrix(x, 2)
    sParc = IIf(grdTemp.TextMatrix(x, 3) = "Unica", "00", grdTemp.TextMatrix(x, 3))
    sComp = grdTemp.TextMatrix(x, 4)
    nPos = nPos + 1
    If nPos >= 31 Then
        nPos = 1
        nPagina = nPagina + 1
    End If
    
    Sql = "UPDATE DEBITOPARCELA SET NUMEROLIVRO=" & nLivro & " ,PAGINALIVRO=" & nPagina & " ,NUMCERTIDAO=" & nNumCert
    Sql = Sql & " ,DATAINSCRICAO='" & Format(mskDataInsc.Text, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & nCodReduz
    Sql = Sql & " AND ANOEXERCICIO=" & Val(sAno) & " AND CODLANCAMENTO=" & Val(sLanc)
    Sql = Sql & " AND SEQLANCAMENTO=" & Val(sSeq) & " AND NUMPARCELA=" & Val(sParc) & " AND CODCOMPLEMENTO=" & Val(sComp)
    cn.Execute Sql, rdExecDirect
    
    Sql = "UPDATE PARAMETROS SET VALPARAM='" & nNumCert & "' WHERE NOMEPARAM='NUMEROCERTIDAO'"
    cn.Execute Sql, rdExecDirect
    
    Sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
    Sql = Sql & "debitoparcela.CODCOMPLEMENTO , debitoparcela.statuslanc, debitoparcela.DataVencimento, debitotributo.CodTributo, debitotributo.ValorTributo, TRIBUTO.abrevtributo "
    Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
    Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
    Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
    Sql = Sql & "tributo ON debitotributo.codtributo = tributo.codtributo where debitotributo.codtributo<>3 and debitoparcela.codreduzido=" & nCodReduz & " and "
    Sql = Sql & "debitoparcela.anoexercicio=" & Val(sAno) & " and debitoparcela.codlancamento=" & Val(sLanc) & " and debitoparcela.seqlancamento=" & Val(sSeq) & " and "
    Sql = Sql & "debitoparcela.numparcela=" & Val(sParc) & " and debitoparcela.codcomplemento=" & Val(sComp)
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
       Do Until .EOF
            Sql = "INSERT CDADebitos(idCDA,CodTributo,Tributo,Exercicio,Lancamento,Seq,NroParcela,ComplParcela,DtVencimento,VlrOriginal,DtGeracao) values("
            Sql = Sql & nCDA & "," & !CodTributo & ",'" & Mask(!ABREVTRIBUTO) & "'," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & ","
            Sql = Sql & !CODCOMPLEMENTO & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!ValorTributo)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
            cnInt.Execute Sql, rdExecDirect
           .MoveNext
        Loop
       .Close
    End With
    


Next

Unload Me
frmDebitoImob.txtCod.SetFocus
frmDebitoImob.grdExtrato.SetFocus

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Centraliza Me
Me.Top = Me.Top + 1000
CarregaLista
mskDataInsc.Text = Format(Now, "dd/mm/yyyy")

Liberado
End Sub

Private Sub CarregaLista()
Dim x As Integer
Dim sSit As String, sVencto As String, sDA As String
Dim sAj As String, nValorPrincipal As Double, nTipo As Integer


With frmDebitoImob.grdExtrato
    nCodReduz = Val(frmDebitoImob.txtCod.Text)
    For x = 1 To .Rows
        If .CellText(x, 12) = "S" Then
           sAno = .CellText(x, 1)
           sLanc = Left$(.CellText(x, 2), 3)
           sSeq = .CellText(x, 3)
           sParc = IIf(.CellText(x, 4) = "Unica", "00", .CellText(x, 4))
           sComp = .CellText(x, 5)
           sSit = Left$(.CellText(x, 6), 2)
           sVencto = .CellText(x, 7)
           sDA = .CellText(x, 8)
           sAj = .CellText(x, 9)
           nValorPrincipal = .CellText(x, 10)
           
           grdTemp.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
             sComp & Chr(9) & sSit & Chr(9) & sVencto & Chr(9) & sDA & Chr(9) & sAj & Chr(9) & _
             FormatNumber(nValorPrincipal, 2)
           
        End If
    Next
End With

If nCodReduz < 100000 Then
   nTipo = 1
ElseIf nCodReduz > 100000 And nCodReduz < 500000 Then
   nTipo = 2
ElseIf nCodReduz > 500000 Then
   nTipo = 3
End If

If nTipo = 1 Then 'IPTU
    sTypeBook = "CODREDUZIDO < 100000"
ElseIf nTipo = 2 Then 'ISS
    sTypeBook = "CODREDUZIDO > 100000 AND CODREDUZIDO < 500000"
ElseIf nTipo = 3 Then 'TAXAS
    sTypeBook = "CODREDUZIDO > 500000"
End If


'Sql = "SELECT NUMERO From LIVRO WHERE CODTIPO = " & nTipo & " AND ANO =  " & Val(grdTemp.TextMatrix(1, 0))
'Sql = "SELECT NUMERO From LIVRO WHERE CODTIPO = " & nTipo & " AND ANO =  " & Year(Now)

Sql = "SELECT livro.codtipo, livro.numero, livro.ano, lancamento.codlancamento FROM lancamento INNER JOIN "
'Sql = Sql & "livro ON lancamento.tipolivro = livro.codtipo where ano=" & Val(grdTemp.TextMatrix(1, 0)) & " and codlancamento=" & Val(grdTemp.TextMatrix(1, 1))
Sql = Sql & "livro ON lancamento.tipolivro = livro.codtipo where ano=" & Year(Now) & " and codlancamento=" & Val(grdTemp.TextMatrix(1, 1))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     nTipo = !CodTipo
     nLivro = !Numero
    .Close
End With

cn.QueryTimeout = 0
Sql = "SELECT MAX(PAGINALIVRO) AS MAXIMO FROM DEBITOPARCELA WHERE numerolivro=" & nLivro
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!MAXIMO) Then
       nPagina = 1
    Else
        nPagina = !MAXIMO + 1
    End If
   .Close
End With

Sql = "SELECT DESCTIPO FROM TIPOLIVRO WHERE CODTIPO=" & nTipo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblLivro.Caption = nLivro & " - " & !DESCTIPO
   .Close
End With

lblPag.Caption = nPagina

End Sub

Private Sub CloseBook(nTipo As Integer)
Dim sTributosDA As String, sLancamentoDA As String, sTypeBook As String
Dim nPos As Integer, nPagina As Integer
Pb.value = 0

If nTipo = 1 Then 'IPTU
    sLancamentoDA = "CODLANCAMENTO=1 OR CODLANCAMENTO=29"
    sTypeBook = "CODREDUZIDO < 100000"
ElseIf nTipo = 2 Then 'ISS
    sLancamentoDA = "CODLANCAMENTO=2 OR CODLANCAMENTO=3 OR CODLANCAMENTO=5  OR CODLANCAMENTO=6 OR CODLANCAMENTO=13"
    sTypeBook = "CODREDUZIDO > 100000 AND CODREDUZIDO < 500000"
ElseIf nTipo = 3 Then 'TAXAS
    sLancamentoDA = "CODLANCAMENTO > 0"
    sTypeBook = "CODREDUZIDO > 500000"
End If

Sql = "SELECT CODTRIBUTO FROM TRIBUTO WHERE DA=1 ORDER BY CODTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sTributosDA = sTributosDA & !CodTributo & ","
       .MoveNext
    Loop
End With
sTributosDA = Chomp(sTributosDA, chomp_righT, 1)

Sql = "SELECT MAX(PAGINALIVRO) AS MAXIMO FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Val(cmbAno.Text) & " AND " & sTypeBook
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!MAXIMO) Then
       lblPag.Caption = 1
    Else
        If !MAXIMO > 0 Then
            Sql = "SELECT COUNT(CODREDUZIDO) AS CONTADOR FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Val(cmbAno.Text) & " AND " & sTypeBook & " AND PAGINALIVRO=" & !MAXIMO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nPos = !contador
                If nPos < 31 Then
                   nPagina = RdoAux!MAXIMO
                Else
                   nPagina = RdoAux!MAXIMO + 1
                End If
               .Close
            End With
        Else
            nPagina = 1
        End If
    End If
   .Close
End With


Sql = "SELECT DISTINCT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO FROM vwDIVIDAATIVA WHERE " & sTypeBook & " AND ANOEXERCICIO=2003 AND (" & sLancamentoDA & ") AND NUMPARCELA>0 AND STATUSLANC=3"
Sql = Sql & " AND NUMEROLIVRO=0 and CODTRIBUTO IN (SELECT CODTRIBUTO FROM TRIBUTO WHERE DA=1) ORDER BY CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotal = .RowCount
    Do Until .EOF
        Sql = "UPDATE DEBITOPARCELA SET NUMEROLIVRO=" & Val(lblNumero.Caption) & " ,PAGINALIVRO=" & nPagina
        Sql = Sql & " ,DATAINSCRICAO='" & Format(mskDataFim.Text, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & !CODREDUZIDO
        Sql = Sql & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND STATUSLANC=3 "
        Sql = Sql & " AND NUMEROLIVRO=0"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        If nPos > 31 Then
            nPos = 1
            nPagina = nPagina + 1
        End If
        nTotal = nTotal - 1
       .MoveNext
    Loop
   .Close
End With


End Sub

