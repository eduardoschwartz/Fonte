VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmAjuizamento 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuizamento de Débitos"
   ClientHeight    =   3900
   ClientLeft      =   8430
   ClientTop       =   3825
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5550
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   2190
      TabIndex        =   7
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   3480
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
      MICON           =   "frmAjuizamento.frx":0000
      PICN            =   "frmAjuizamento.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   345
      Left            =   2190
      TabIndex        =   5
      ToolTipText     =   "Gravar os Dados"
      Top             =   3480
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAjuizamento.frx":0176
      PICN            =   "frmAjuizamento.frx":0192
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
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   3480
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
      MICON           =   "frmAjuizamento.frx":0537
      PICN            =   "frmAjuizamento.frx":0553
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
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   5490
      _ExtentX        =   9684
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
   Begin esMaskEdit.esMaskedEdit mskDataAj 
      Height          =   285
      Left            =   1740
      TabIndex        =   0
      Top             =   2940
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmAjuizamento.frx":05C1
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
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   345
      Left            =   60
      TabIndex        =   8
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   3480
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Cancelar Ajuizamento"
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
      MICON           =   "frmAjuizamento.frx":05DD
      PICN            =   "frmAjuizamento.frx":05F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1740
      TabIndex        =   6
      Top             =   2625
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Ajuizamento..:"
      Height          =   225
      Left            =   165
      TabIndex        =   4
      Top             =   2970
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da Certidão.......:"
      Height          =   225
      Left            =   165
      TabIndex        =   3
      Top             =   2610
      Width           =   1470
   End
End
Attribute VB_Name = "frmAjuizamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, sql As String, RdoAux2 As rdoResultset
Dim sAno As String, sLanc As String, sSeq As String, sParc As String
Dim sComp As String, nCodReduz As Long
Dim nLivro As Integer, nPagina As Integer, sTypeBook As String, nPos As Integer
Dim xImovel As clsImovel

Private Sub cmdCancel_Click()

If MsgBox("Cancelar o ajuizamento destes lançamentos !.", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
nCodReduz = Val(frmDebitoImob.txtCod.Text)

For x = 1 To grdTemp.Rows - 1
    
    sAno = grdTemp.TextMatrix(x, 0)
    sLanc = Left$(grdTemp.TextMatrix(x, 1), 3)
    sSeq = grdTemp.TextMatrix(x, 2)
    sParc = IIf(grdTemp.TextMatrix(x, 3) = "Unica", "00", grdTemp.TextMatrix(x, 3))
    sComp = grdTemp.TextMatrix(x, 4)
    
    sql = "UPDATE DEBITOPARCELA SET DATAAJUIZA=NULL WHERE CODREDUZIDO=" & nCodReduz
    sql = sql & " AND ANOEXERCICIO=" & Val(sAno) & " AND CODLANCAMENTO=" & Val(sLanc)
    sql = sql & " AND SEQLANCAMENTO=" & Val(sSeq) & " AND NUMPARCELA=" & Val(sParc) & " AND CODCOMPLEMENTO=" & Val(sComp)
    cn.Execute sql, rdExecDirect
Next

Unload Me
frmDebitoImob.txtCod.SetFocus
frmDebitoImob.grdExtrato.SetFocus

End Sub

Private Sub cmdGravar_Click()

If Not IsDate(mskDataAj.Text) Then
    MsgBox "Data inválida.", vbCritical, "atenção"
    Exit Sub
End If

If MsgBox("Confirme o ajuizamento destes lançamentos !.", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub


For x = 1 To grdTemp.Rows - 1
    
    sAno = grdTemp.TextMatrix(x, 0)
    sLanc = Left$(grdTemp.TextMatrix(x, 1), 3)
    sSeq = grdTemp.TextMatrix(x, 2)
    sParc = IIf(grdTemp.TextMatrix(x, 3) = "Unica", "00", grdTemp.TextMatrix(x, 3))
    sComp = grdTemp.TextMatrix(x, 4)
    
    'sql = "UPDATE DEBITOPARCELA SET NUMCERTIDAO=" & Val(lblNumero.Caption) & " ,DATAAJUIZA='" & Format(mskDataAj.text, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & nCodReduz
    sql = "UPDATE DEBITOPARCELA SET DATAAJUIZA='" & Format(mskDataAj.Text, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & nCodReduz
    sql = sql & " AND ANOEXERCICIO=" & Val(sAno) & " AND CODLANCAMENTO=" & Val(sLanc)
    sql = sql & " AND SEQLANCAMENTO=" & Val(sSeq) & " AND NUMPARCELA=" & Val(sParc) & " AND CODCOMPLEMENTO=" & Val(sComp)
    cn.Execute sql, rdExecDirect
Next

Unload Me
frmDebitoImob.txtCod.SetFocus
frmDebitoImob.grdExtrato.SetFocus

End Sub

Private Sub cmdPrint_Click()
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, RdoAux3 As rdoResultset
Dim nNumParc As Integer, nCompl As Integer, nValorTotal As Double, ax As Integer
Dim nValorLancado As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double
Dim nSomaLancado As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double
Dim sDataVencto As String, nInscricao As Integer, dDataInscricao As Date, nLivro As Integer, nPagina As Integer, sValorTotal As String
Dim sNome As String, sEND1 As String, sCOMPL1 As String, sBAIRRO1 As String, sCIDADE1 As String
Dim sCEP1  As String, sUF1   As String, sInscricao As String, sEND2 As String, sCOMPL2 As String, sBAIRRO2 As String
Dim sQuadra As String, sLote As String, sCIDADE2 As String, sCEP2 As String, sUF2 As String, qd As New rdoQuery, nCodTributo As Integer, aCodTrib() As Integer
Dim sLANCAMENTO As String, sDOCUMENTO As String, aAno() As Integer, sTributo As String, bJuros As Boolean, bMulta As Boolean

ReDim aAno(0)
sql = "DELETE FROM RELATORIOAJUIZAMENTO"
cn.Execute sql, rdExecDirect
sql = "DELETE FROM RELATORIOAJUIZAMENTODETALHE"
cn.Execute sql, rdExecDirect

nCodReduz = Val(frmDebitoImob.txtCod.Text)
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
        sql = "SELECT ENDENTREGA.CODREDUZIDO,ENDENTREGA.EE_CODLOG, ENDENTREGA.EE_NOMELOG,ENDENTREGA.EE_NUMIMOVEL,"
        sql = sql & "ENDENTREGA.EE_COMPLEMENTO, ENDENTREGA.EE_UF,ENDENTREGA.EE_CIDADE, ENDENTREGA.EE_BAIRRO,"
        sql = sql & "ENDENTREGA.EE_CEP, ENDENTREGA.EE_LOTEAMENTO,ENDENTREGA.EE_DESCBAIRRO , Cidade.DESCCIDADE "
        sql = sql & "FROM ENDENTREGA INNER JOIN  CIDADE ON ENDENTREGA.EE_UF = CIDADE.SIGLAUF AND "
        sql = sql & "ENDENTREGA.Ee_Cidade = Cidade.CODCIDADE WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!Ee_NomeLog <> "" Then
                sEND2 = RdoAux2!Ee_NomeLog & " Nº " & CStr(RdoAux2!Ee_NumImovel)
                sCOMPL2 = RdoAux2!Ee_Complemento
                sBAIRRO2 = SubNull(RdoAux2!Ee_DESCBairro)
                sCIDADE2 = SubNull(RdoAux2!desccidade)
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
    sql = "SELECT CODIGOMOB,DVMOB,RAZAOSOCIAL,NOMEFANTASIA,CNPJ,CPF,INSCESTADUAL,"
    sql = sql & "DATAABERTURA,NUMPROCESSO,DATAPROCESSO,ATIVEXTENSO,SIGLAUF,CODCIDADE,CODBAIRRO,"
    sql = sql & "DESCCIDADE,DESCBAIRRO,DESCUF,CODLOGRADOURO,NOMELOGR,"
    sql = sql & "NUMERO,COMPLEMENTO,CEP,DATAENCERRAMENTO,NUMPROCENCERRAMENTO,DATAPROCENCERRAMENTO,"
    sql = sql & "HORARIO,DESCHORARIO,HOMEPAGE,NOMECONTATO,FONECONTATO,FAXCONTATO,CARGOCONTATO,"
    sql = sql & "EMAILCONTATO,RESPCONTABIL,CAPITALSOCIAL,QTDEEMPREGADO,QTDEPROF,RG,ORGAO "
    sql = sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        sInscricao = Format(!codigomob, "0000000") & "-" & !DVMOB
        sQuadra = ""
        sLote = ""
        sNome = !RazaoSocial
        If Not IsNull(!CPF) Then
           sDOCUMENTO = Format(Trim(!CPF), "000\.000\.000-00")
        ElseIf Not IsNull(!Cnpj) Then
           sDOCUMENTO = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
        ElseIf Not IsNull(!rg) Then
           sDOCUMENTO = !rg
        Else
           sDOCUMENTO = ""
        End If

        sql = "SELECT CODIGOMOB,INSCESTADUAL,CNPJ,CPF,RAZAOSOCIAL,CODLOGRADOURO,DESCCIDADE,SIGLAUF,CODCIDADE,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO,NOMELOGR "
        sql = sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sNumInsc = !INSCESTADUAL
                sEND1 = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro) & " nº " & Val(SubNull(!Numero))
                sCOMPL1 = SubNull(!Complemento)
                sql = "SELECT  DESCBAIRRO From BAIRRO WHERE  siglauf = '" & !SiglaUF & "' And codcidade = " & !CodCidade & " And codbairro = " & !CodBairro
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    sBAIRRO1 = SubNull(!DescBairro)
                   .Close
                End With
                sql = "SELECT  DESCCIDADE From CIDADE WHERE  siglauf = '" & !SiglaUF & "' And codcidade = " & !CodCidade
                Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    sCIDADE1 = SubNull(!desccidade)
                   .Close
                End With
                sCEP1 = RetornaCEP(!CodLogradouro, !Numero)
                If Len(sCEP1) <> 9 Then
                    sCEP1 = "00000-000"
                End If
                sUF1 = SubNull(!SiglaUF)
            End If
        End With
        sql = "SELECT * FROM MOBILIARIOENDENTREGA WHERE CODMOBILIARIO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
                If .RowCount > 0 Then
                    sEND2 = SubNull(!NomeLogradouro) & " nº " & !numimovel
                    sCOMPL2 = SubNull(!Complemento)
                    sUF2 = SubNull(!UF)
                    If !CodCidade > 0 Then
                        sql = "SELECT  DESCCIDADE From CIDADE WHERE  siglauf = '" & !UF & "' And codcidade = " & !CodCidade
                        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            sCIDADE2 = SubNull(!desccidade)
                           .Close
                        End With
                    Else
                        sCIDADE2 = SubNull(!desccidade)
                    End If
                    If !CodBairro > 0 Then
                        sql = "SELECT  DESCBAIRRO From BAIRRO WHERE  siglauf = '" & !UF & "' And codcidade = " & !CodCidade & " And codbairro = " & !CodBairro
                        Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            sBAIRRO2 = SubNull(!DescBairro)
                           .Close
                        End With
                    Else
                        sBAIRRO2 = SubNull(!DescBairro)
                    End If
                    sCEP2 = RetornaCEP(!CodLogradouro, !numimovel)
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
    sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, "
    sql = sql & "vwLOGRADOURO.ABREVTITLOG, vwLOGRADOURO.NOMELOGRADOURO, cidadao.numimovel, cidadao.complemento, cidadao.codbairro,"
    sql = sql & "bairro.descbairro, cidadao.codcidade, cidade.desccidade, cidadao.siglauf, cidadao.cep, cidadao.nomelogradouro AS nomelogradouro2"
    sql = sql & " FROM cidadao RIGHT OUTER JOIN bairro ON cidadao.siglauf = bairro.siglauf AND "
    sql = sql & "cidadao.codcidade = bairro.codcidade AND cidadao.codbairro = bairro.codbairro INNER JOIN cidade ON bairro.siglauf = cidade.siglauf AND "
    sql = sql & "bairro.codcidade = cidade.codcidade LEFT OUTER JOIN  vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO "
    sql = sql & "WHERE Cidadao.codcidadao = " & nCodReduz
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            sInscricao = Format(!CodCidadao, "0000000")
            sNome = !nomecidadao
            sQuadra = ""
            sLote = ""
            If Not IsNull(!NomeLogradouro) Then
               sEND1 = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & RdoAux!numimovel
            Else
               sEND1 = !NOMELOGRADOURO2 & " Nº " & RdoAux!numimovel
            End If
            sCOMPL1 = SubNull(!Complemento)
            sBAIRRO1 = SubNull(!DescBairro)
            sCIDADE1 = SubNull(!desccidade)
            sCEP1 = !Cep
            If Len(sCEP1) <> 9 Then
                sCEP1 = "00000-000"
            End If
            sUF1 = SubNull(!SiglaUF)
        End If
        sEND2 = sEND1 '205-250
        sCOMPL2 = sCOMPL1 '251-270
        sBAIRRO2 = sBAIRRO1
        sCIDADE2 = sCIDADE1
        sCEP2 = sCEP1
        sUF2 = sUF1
    End With
End If

With grdTemp
    For x = 1 To .Rows - 1
        sLANCAMENTO = ""
        nSomaLancado = 0: nSomaJuros = 0: nSomaMulta = 0: nSomaCorrecao = 0
        nAno = .TextMatrix(x, 0)
        nLanc = .TextMatrix(x, 1)
        nSeq = .TextMatrix(x, 2)
        nNumParc = .TextMatrix(x, 3)
        nCompl = .TextMatrix(x, 4)
        sDataVencto = .TextMatrix(x, 6)
        
        Set qd.ActiveConnection = cn
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
        qd.sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
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
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        With RdoAux
            ReDim aCodTrib(0)
            Do Until .EOF
                ReDim Preserve aCodTrib(UBound(aCodTrib) + 1)
                aCodTrib(UBound(aCodTrib)) = !CodTributo
                
                sql = "INSERT RELATORIOAJUIZAMENTODETALHE (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,"
                sql = sql & "CODCOMPLEMENTO,DATAVENCIMENTO,PRINCIPAL,CORRECAO,MULTA,JUROS,INSCRICAO,DATAINSCRICAO,"
                sql = sql & "LIVRO,PAGINA,VALORTOTAL,LANCAMENTO,TRIBUTO) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & ","
                sql = sql & !CODCOMPLEMENTO & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(!ValorTributo)) & ","
                sql = sql & Virg2Ponto(CStr(!ValorCorrecao)) & "," & Virg2Ponto(CStr(!ValorMulta)) & "," & Virg2Ponto(CStr(!ValorJuros)) & ","
                sql = sql & SubNull(!CERTIDAO) & ",'" & Format(!datainscricao, "mm/dd/yyyy") & "'," & !NUMLIVRO & "," & !PAGINA & ",'"
                sql = sql & CStr(!ValorTotal) & "','" & "" & "','" & !abrevtributo & "')"
                cn.Execute sql, rdExecDirect
                sLANCAMENTO = sLANCAMENTO & !abrevtributo & ", "
               .MoveNext
            Loop
           .Close
        End With
        sLANCAMENTO = Left$(sLANCAMENTO, Len(sLANCAMENTO) - 2)
        sql = "UPDATE RELATORIOAJUIZAMENTODETALHE SET LANCAMENTO='" & sLANCAMENTO & "' WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND "
        sql = sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
        cn.Execute sql, rdExecDirect
    
        '*** BUSCA O ARTIGO ****
        nCodTributo = 0
        For z = 1 To UBound(aCodTrib)
            sql = "SELECT * FROM TRIBUTOARTIGO WHERE CODTRIBUTO=" & aCodTrib(z)
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
            If RdoAux2.RowCount > 0 Then
                nCodTributo = aCodTrib(z)
                RdoAux2.Close
            End If
        Next
        If nCodTributo > 0 Then
            sql = "UPDATE RELATORIOAJUIZAMENTODETALHE SET CODTRIBUTO=" & nCodTributo & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
            sql = sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND "
            sql = sql & "CODCOMPLEMENTO=" & nCompl
            cn.Execute sql, rdExecDirect
        End If
        '***********************
    
    Next


    sql = "SELECT DISTINCT(ANOEXERCICIO) FROM RELATORIOAJUIZAMENTODETALHE"
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = !AnoExercicio
           .MoveNext
        Loop
       .Close
    End With
    For ax = 1 To UBound(aAno)
        sql = "SELECT SUM(PRINCIPAL) AS VALORLANC, SUM(CORRECAO) AS VALORCOR, SUM(MULTA) AS VALORMULTA,SUM(JUROS) AS VALORJUROS "
        sql = sql & "FROM RELATORIOAJUIZAMENTODETALHE WHERE ANOEXERCICIO=" & aAno(ax)
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nValorTotal = FormatNumber(!VALORLANC + !VALORCOR + !ValorMulta + !ValorJuros, 2)
           .Close
           sValorTotal = "R$ " & nValorTotal & " (" & Extenso(nValorTotal) & ")"
           sql = "UPDATE RELATORIOAJUIZAMENTODETALHE SET VALORTOTAL='" & sValorTotal & "' WHERE ANOEXERCICIO=" & aAno(ax)
           cn.Execute sql, rdExecDirect
        End With
    Next

    sql = "SELECT SUM(PRINCIPAL) AS VALORLANC, SUM(CORRECAO) AS VALORCOR, SUM(MULTA) AS VALORMULTA,SUM(JUROS) AS VALORJUROS "
    sql = sql & "FROM RELATORIOAJUIZAMENTODETALHE "
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nValorTotal = FormatNumber(!VALORLANC + !VALORCOR + !ValorMulta + !ValorJuros, 2)
       .Close
    End With
    sql = "INSERT RELATORIOAJUIZAMENTO (NOME,END1,COMPL1,BAIRRO1,CIDADE1,CEP1,UF1,INSCRICAO,CODREDUZ,"
    sql = sql & "END2,COMPL2,BAIRRO2,QUADRA,LOTE,CIDADE2,CEP2,UF2,VALORTOTAL,LANCAMENTO,DOCUMENTO) VALUES('"
    sql = sql & Mask(sNome) & "','" & Mask(sEND1) & "','" & Mask(sCOMPL1) & "','" & Mask(sBAIRRO1) & "','" & Mask(sCIDADE1) & "','"
    sql = sql & sCEP1 & "','" & sUF1 & "','" & sInscricao & "'," & nCodReduz & ",'" & Mask(sEND2) & "','"
    sql = sql & Mask(sCOMPL2) & "','" & Mask(sBAIRRO2) & "','" & sQuadra & "','" & sLote & "','"
    sql = sql & Mask(sCIDADE2) & "','" & Left$(sCEP2, 9) & "','" & sUF2 & "'," & Virg2Ponto(CStr(nValorTotal)) & ",'" & Mask(sLANCAMENTO) & "','" & sDOCUMENTO & "')"
    cn.Execute sql, rdExecDirect

End With

'EXIBE RELATORIO

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

Private Sub Form_Activate()

If frmDebitoImob.lblAjuiza.Caption = "S" Then
    cmdPrint.Visible = True
    cmdCancel.Visible = True
    cmdGravar.Visible = False
    mskDataAj.Enabled = False
Else
    cmdCancel.Visible = False
    cmdPrint.Visible = False
    cmdGravar.Visible = True
    mskDataAj.Enabled = True
End If

End Sub

Private Sub Form_Load()
Set xImovel = New clsImovel
Centraliza Me
Me.Top = Me.Top + 1000
CarregaLista
If frmDebitoImob.lblAjuiza.Caption = "N" Then
   mskDataAj.Text = Format(Now, "dd/mm/yyyy")
End If

Liberado
End Sub

Private Sub CarregaLista()
Dim x As Integer
Dim sSit As String, sVencto As String, sDA As String
Dim sAj As String, nValorPrincipal As Double


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


If frmDebitoImob.lblAjuiza.Caption = "N" Then
    cn.QueryTimeout = 0
    sql = "SELECT MAX(NUMCERTIDAO) AS MAXIMO From DEBITOPARCELA WHERE ANOEXERCICIO = " & Year(Now)
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         If IsNull(!MAXIMO) Then
            lblNumero.Caption = 1
         Else
            lblNumero.Caption = !MAXIMO + 1
         End If
        .Close
    End With
Else
    lblNumero.Caption = ""
End If
End Sub

Private Sub CloseBook(nTipo As Integer)
Dim sTributosDA As String, sLancamentoDA As String, sTypeBook As String
Dim nPos As Integer, nPagina As Integer
Pb.Value = 0

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

sql = "SELECT CODTRIBUTO FROM TRIBUTO WHERE DA=1 ORDER BY CODTRIBUTO"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sTributosDA = sTributosDA & !CodTributo & ","
       .MoveNext
    Loop
End With
sTributosDA = Chomp(sTributosDA, chomp_righT, 1)

sql = "SELECT MAX(PAGINALIVRO) AS MAXIMO FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Val(cmbAno.Text) & " AND " & sTypeBook
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!MAXIMO) Then
       lblPag.Caption = 1
    Else
        If !MAXIMO > 0 Then
            sql = "SELECT COUNT(CODREDUZIDO) AS CONTADOR FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Val(cmbAno.Text) & " AND " & sTypeBook & " AND PAGINALIVRO=" & !MAXIMO
            Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nPos = !CONTADOR
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


sql = "SELECT DISTINCT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO FROM vwDIVIDAATIVA WHERE " & sTypeBook & " AND ANOEXERCICIO=2003 AND (" & sLancamentoDA & ") AND NUMPARCELA>0 AND STATUSLANC=3"
sql = sql & " AND NUMEROLIVRO=0 and CODTRIBUTO IN (SELECT CODTRIBUTO FROM TRIBUTO WHERE DA=1) ORDER BY CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotal = .RowCount
    Do Until .EOF
        sql = "UPDATE DEBITOPARCELA SET NUMEROLIVRO=" & Val(lblNumero.Caption) & " ,PAGINALIVRO=" & nPagina
        sql = sql & " ,DATAINSCRICAO='" & Format(mskDataFim.Text, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & !CODREDUZIDO
        sql = sql & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND STATUSLANC=3 "
        sql = sql & " AND NUMEROLIVRO=0"
        cn.Execute sql, rdExecDirect
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

Private Sub Form_Unload(Cancel As Integer)
Set xImovel = Nothing
End Sub
