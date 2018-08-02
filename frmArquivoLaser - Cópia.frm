VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmArquivoLaser 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Arquivo Laser"
   ClientHeight    =   2310
   ClientLeft      =   4395
   ClientTop       =   3435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   6000
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Arquivo de IPTU"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   300
      Width           =   1605
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Arquivo de ISS Est/Variavel"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   585
      Width           =   2895
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Etiqueta"
      Height          =   195
      Index           =   2
      Left            =   2085
      TabIndex        =   7
      Top             =   270
      Width           =   1095
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Arquivo de ISS Fixo/TLL"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   885
      Value           =   -1  'True
      Width           =   2265
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Arquivo de Vig.Sanitária"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1170
      Width           =   2265
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "Sair da Tela"
      Top             =   495
      Width           =   1125
      _ExtentX        =   1984
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmArquivoLaser.frx":0000
      PICN            =   "frmArquivoLaser.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Height          =   315
      Left            =   4770
      TabIndex        =   1
      ToolTipText     =   "Cancelar Edição"
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gerar"
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
      MICON           =   "frmArquivoLaser.frx":008A
      PICN            =   "frmArquivoLaser.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   240
      Left            =   870
      TabIndex        =   2
      Top             =   1845
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdRel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   180
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   2580
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Rel"
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
      MICON           =   "frmArquivoLaser.frx":0145
      PICN            =   "frmArquivoLaser.frx":0161
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   285
      TabIndex        =   4
      Top             =   1860
      Width           =   390
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   3240
      TabIndex        =   3
      Top             =   1860
      Width           =   720
   End
End
Attribute VB_Name = "frmArquivoLaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type aReg
    sTexto As String * 2334
End Type
Private Type UNICA
    nCodTributo As Integer
    nValorTributo As String
End Type

Private Type Endereco
    nCodigo As Integer
    sNome As String
    sLogradouro As String
    nNumero As Integer
    sBairro As String
    sCEP As String
    sCidade As String
    sUF As String
    bRecebe As Boolean
End Type

Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, aEnd() As Endereco
Dim xImovel As clsImovel

Private Sub cmdGerar_Click()

If MsgBox("Deseja confirmar a geração do Arquivo Laser.", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Ocupado
If Opt(0).Value = True Then
   GeraIPTU
ElseIf Opt(1).Value = True Then
   GeraISSVarEst
ElseIf Opt(2).Value = True Then
   GeraEtiqueta
ElseIf Opt(3).Value = True Then
   GeraISSFixo
ElseIf Opt(4).Value = True Then
   GeraVS
End If
Liberado

End Sub

Private Sub cmdRel_Click()
Dim dData As String, sCNPJ As String, dDataS As String, sDivida As String
Open sPathBin & "\EMPRESAS.TXT" For Output As #1
Sql = "SELECT CODIGOMOB,RAZAOSOCIAL,CNPJ,DATAENCERRAMENTO FROM MOBILIARIO ORDER BY CODIGOMOB"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If Not IsNull(!DATAENCERRAMENTO) Then
           dData = Format(!DATAENCERRAMENTO, "dd/mm/yyyy")
        Else
           dData = "00/00/0000"
        End If
        If IsNull(!Cnpj) Or !Cnpj = 0 Or !Cnpj = "" Then
            sCNPJ = "0"
        Else
            sCNPJ = Left(!Cnpj, 2) & "." & Mid(!Cnpj, 3, 3) & "." & Mid(!Cnpj, 6, 3) & "/" & Mid(!Cnpj, 9, 4) & "-" & Right(!Cnpj, 2)
        End If
        
        Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & !codigomob
        Sql = Sql & " ORDER BY DATAEVENTO DESC"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If !CODTIPOEVENTO = 2 Then
                    dDataS = Format(!DATAPROCEVENTO, "dd/mm/yyyy")
                Else
                    dDataS = "00/00/0000"
                End If
            Else
                dDataS = "00/00/0000"
            End If
           .Close
        End With

        Sql = "SELECT DISTINCT codreduzido From debitoparcela WHERE codreduzido=" & !codigomob & "  and (codlancamento = 2 OR codlancamento = 3 OR codlancamento = 5 OR  codlancamento = 13) AND (statuslanc = 3) AND (datavencimento <= CONVERT(DATETIME, '2006-12-31 00:00:00', 102))"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If RdoAux2.RowCount > 0 Then
                sDivida = "S"
            Else
                sDivida = "N"
            End If
            RdoAux2.Close
        End With


        ax = !codigomob & "#" & !RazaoSocial & "#" & sCNPJ & "#" & dData & "#" & dDataS & "#" & sDivida
        Print #1, ax
       .MoveNext
    Loop
   .Close
End With
Close #1
MsgBox "fim"

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Set xImovel = New clsImovel
CarregaEnderecoContabil
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub GeraISSVarEst()
Dim xId As Long, nNumRec As Long, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset, x As Integer
Dim nCodLogr As Long, nNum As Integer, nQtdeParcF As Integer, nQtdeParcE As Integer, nQtdeParcS As Integer
Dim nExpParc As Double, nExpUnica As Double, t As Integer, nCodEsc As Integer

'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sFantasia As String, sEnd As String, sCompl As String, sBairro As String, sCEP As String
Dim sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim sTipoImposto As String, sInscricao As String, sQtdeParc As String, sCodAtiv As String, sDescAtiv As String, sCodInscricao As String
Dim bISSFixo As Boolean, bISSEstimado As Boolean, bISSVariavel As Boolean, bVigSanit As Boolean, bTxLic As Boolean
Dim aDescParc(0 To 12) As String, sDescParc As String
Dim aCodTrib(0 To 10) As String, sCodTrib As String
Dim aDescTrib(0 To 10) As String, sDescTrib As String
Dim aVencParc(0 To 12) As String, sVencParc As String
Dim aValorTributoUnica(0 To 12) As String, sValorTributoUnica As String
Dim aValorTributo(0 To 12) As String, sValorTributo As String
Dim aValorParc(0 To 12) As String, sValorParc As String
Dim aMesAno(0 To 12) As String, sMesAno As String, sMes As String
Dim aNumDoc(0 To 12) As String, sNumDoc As String
Dim nTotalTrib As Double, sTotalTrib As String
Dim nTotalTribUnica As Double, sTotalTribUnica As String
Dim aCodBarra(0 To 12) As String, sCodBarra As String
Dim dDataBase As Date, nUfir As Double
Dim tDado As String, tEnd As String, tNum As Integer, tTipo As Integer
Dim tCidade As String, tBairro As String

'GoTo ORDENA
'********************************
' PARAMETROS DAS PARCELAS
'********************************
'PARCELAS PARA ISS FIXO E TLL
Sql = "SELECT QTDEPARCELA FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & 2014 & " AND CODTIPO=2"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nQtdeParcF = !qtdeparcela
   .Close
End With

nUfir = RetornaUFIR(2014)
sAgencia = "02345000024"

'PARCELAS PARA ISS ESTIMADO E VARIÁVEL
Sql = "SELECT QTDEPARCELA FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & 2014 & " AND CODTIPO=3"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nQtdeParcE = !qtdeparcela
   .Close
End With

'PARCELAS PARA VIGILÂNCIA SANITÁRIA
Sql = "SELECT QTDEPARCELA FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & 2014 & " AND CODTIPO=5"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nQtdeParcS = !qtdeparcela
   .Close
End With

Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & 2014
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nExpParc = FormatNumber(RdoAux!VALORPARCELA, 2)
nExpUnica = FormatNumber(RdoAux!ValorUnica, 2)

'## ************************************************ ##
'## ***************** G U I A ********************** ##
'## ************************************************ ##

Sql = "TRUNCATE TABLE LASERTMP"
cn.Execute Sql, rdExecDirect

Open sPathBin & "\LASERISSEST.TXT" For Output As #1
'Open sPathBin & "\LASERISSVAR.TXT" For Output As #2
'Open sPathBin & "\LASERISSFIXOTL.TXT" For Output As #3
'Open sPathBin & "\LASERVIGSANIT.TXT" For Output As #4

Sql = "SELECT DISTINCT CODREDUZIDO From DEBITOPARCELA "
Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 3 OR CODLANCAMENTO = 5 ) "
'Sql = Sql & "AND CODREDUZIDO=111801"
'Sql = Sql & "AND (DATAVENCIMENTO > '02/20/2014') ORDER BY CODREDUZIDO"
'Sql = Sql & "CODLANCAMENTO = 6 OR CODLANCAMENTO = 13) AND (DATAVENCIMENTO > '02/20/2014') ORDER BY CODREDUZIDO"
'Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 2 OR CODLANCAMENTO=6 ) AND (DATAVENCIMENTO > '01/10/2014') ORDER BY CODREDUZIDO"
'Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 13) AND (DATAVENCIMENTO > '02/20/2014') ORDER BY CODREDUZIDO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        If !CODREDUZIDO = 115554 Then MsgBox "TESTE"
        If xId Mod 50 = 0 Then
           CallPb xId, nNumRec
        End If
        Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
        Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,MOBILIARIO.RESPCONTABIL,"
        Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
        Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
        Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
        Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                GoTo Proximo
            End If
            nCodEsc = Val(SubNull(!RESPCONTABIL))
            nCodLogr = !CodLogradouro
            sExercicio = "2014" '1-4
            sCodInscricao = Format(!codigomob, "00000000000000")
            sContribuinte = FillSpace(!RazaoSocial, 40) '5-44
            sFantasia = FillSpace(SubNull(!NOMEFANTASIA), 40) '45-84
            sCodAtiv = Format(!codatividade, "00000000000000")
            sDescAtiv = FillSpace(Left$(!ATIVEXTENSO, 50), 50)
            Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    sEnd = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero), 46) '212-257
                    nNum = RdoAux2!Numero
                Else
                    nNum = 0
                End If
               .Close
            End With
            sCEP = RetornaCEP(nCodLogr, nNum)
            sCompl = FillSpace(SubNull(Left(!Complemento, 20)), 20) '258-277
            sBairro = FillSpace(SubNull(!DescBairro), 30) '278-307
            Sql = "SELECT MOBILIARIOENDENTREGA.CODMOBILIARIO, MOBILIARIOENDENTREGA.CODLOGRADOURO, MOBILIARIOENDENTREGA.NOMELOGRADOURO, "
            Sql = Sql & "MOBILIARIOENDENTREGA.NUMIMOVEL,MOBILIARIOENDENTREGA.COMPLEMENTO, MOBILIARIOENDENTREGA.UF,MOBILIARIOENDENTREGA.CODCIDADE,"
            Sql = Sql & "MOBILIARIOENDENTREGA.CODBAIRRO, MOBILIARIOENDENTREGA.CEP,MOBILIARIOENDENTREGA.DESCBAIRRO, MOBILIARIOENDENTREGA.DESCCIDADE, BAIRRO.DESCBAIRRO AS DESCBAIRRO2,"
            Sql = Sql & "CIDADE.DESCCIDADE AS DESCCIDADE2 FROM BAIRRO INNER JOIN MOBILIARIOENDENTREGA ON BAIRRO.SIGLAUF = MOBILIARIOENDENTREGA.UF AND BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO INNER JOIN "
            Sql = Sql & "CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE Where CODMOBILIARIO = " & !codigomob
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    If !CodLogradouro = 0 Then
                        sEndEntrega = FillSpace(!NomeLogradouro & " Nº " & CStr(!numimovel), 46) '85-130
                        sNumEntrega = SubNull(!numimovel)
                        If !CodBairro = 0 Then
                            sBairroEntrega = FillSpace(SubNull(!DescBairro), 30) '151-180
                        ElseIf !CodBairro = 999 Then
                            sBairroEntrega = FillSpace(" ", 30) '151-180
                        Else
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30)
                        End If
                        If !CodCidade = 0 Then
                            sCidEntrega = FillSpace(SubNull(!desccidade), 20) '181-200
                        Else
                            sCidEntrega = FillSpace(SubNull(!desccidade2), 20)
                        End If
                        If !CodCidade = 413 Then
                            sCepEntrega = Format(!Cep, "00000-000") '201-209
                        Else
                            sCepEntrega = RetornaCEP(!CodLogradouro, Val(sNumEntrega))
                        End If
                        sComplEntrega = FillSpace(SubNull(!Complemento), 20)
                        sUFEntrega = !UF '210-211
                    Else
                        Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
                        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux4
                            If .RowCount > 0 Then
                                sEndEntrega = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(SubNull(RdoAux3!numimovel)), 46) '85-130
                                sNumEntrega = SubNull(RdoAux3!numimovel)
                                sCepEntrega = RetornaCEP(RdoAux3!CodLogradouro, Val(sNumEntrega))
                            End If
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30) '151-180
                            sCidEntrega = FillSpace(SubNull(RdoAux3!desccidade2), 20) '181-200
                            sComplEntrega = FillSpace(SubNull(RdoAux2!Complemento), 20)
                            sUFEntrega = RdoAux3!UF '210-211
                           .Close
                        End With
                    End If
                Else
                    nCodLogr = 0
                    sEndEntrega = sEnd '85-130
                    sBairroEntrega = FillSpace(sBairro, 30) '151-180
                    sCidEntrega = FillSpace("JABOTICABAL", 20) '181-200
                    sCepEntrega = sCEP '201-209
                    sComplEntrega = FillSpace(" ", 20)
                    sUFEntrega = "SP" '210-211
                End If
               .Close
            End With
            
           .Close
        End With
        
        If nCodEsc > 0 Then
            '***ENDERECO CONTADOR***
            For t = 1 To UBound(aEnd)
                With aEnd(t)
                    If aEnd(t).nCodigo = nCodEsc Then
                        If aEnd(t).bRecebe Then
                            sEndEntrega = FillSpace(.sLogradouro & " Nº " & CStr(.nNumero), 46) '85-130
                            sNumEntrega = .nNumero
                            sBairroEntrega = FillSpace(.sBairro, 30)
                            sCidEntrega = FillSpace(.sCidade, 20)
                            sUFEntrega = .sUF
                            sComplEntrega = FillSpace(" ", 20)
                            sCepEntrega = .sCEP
                        End If
                        Exit For
                    End If
                End With
            Next
            '***********************
        End If
        
        
        tEnd = sEndEntrega
        tNum = Val(sNumEntrega)
        tBairro = sBairroEntrega
        tCidade = sCidEntrega
        If Left(sCepEntrega, 1) = "_" Then sCepEntrega = "         "
        If Trim(sCepEntrega) = "" Then sCepEntrega = "14870-000"
        
        'INSCRICAO
        sInscricao = Format(Val(sCodInscricao), "000000000000000")
'        If nCodLogr > 0 Then
'            Sql = "SELECT DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,"
'            Sql = Sql & "CODLOGR,LI_NUM FROM vwCnsImovel WHERE CODLOGR=" & nCodLogr
'            Sql = Sql & " AND LI_NUM=" & nNum
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                If .RowCount > 0 Then
'                   sInscricao = Format(!Distrito, "00") & Format(!Setor, "00") & Format(!Quadra, "0000") & Format(!Lote, "00000") & Format(!Seq, "00")
'                End If
'               .Close
'            End With
'        Else
'            sInscricao = "000000000000000"
'        End If
        'TIPOS DE CARNE A GERAR
        Sql = "SELECT DISTINCT CODLANCAMENTO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=2014"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            bISSEstimado = False
            bISSFixo = False
            bISSVariavel = False
            bVigSanit = False
            bTxLic = False
            Do Until .EOF
                Select Case !CodLancamento
                Case 2
'                    bISSFixo = True
                Case 3
                    bISSEstimado = True
                Case 5
                    bISSVariavel = True
                Case 6
'                    bTxLic = True
                Case 13
 '                   bVigSanit = True
                End Select
              .MoveNext
            Loop
           .Close
        End With
        If bTxLic Then 'TAXA DE LICENÇA
'            GoTo Proximo
            sTipoImposto = FillSpace("TAXA DE LICENÇA", 20)  '308-327
            sQtdeParc = Format(nQtdeParcF, "0000000000") '427-428
            Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
            Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
            Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
            Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
            Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
            Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 6  "
            'Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 6  AND (DATAVENCIMENTO > '01/10/2014')"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSomaTrib = 0
                For x = 0 To 12
                    aMesAno(x) = ""
                Next
                For x = 1 To 10
                    aValorTributo(x) = ""
                    aValorTributoUnica(x) = ""
                Next
                If .RowCount < 4 Then
                    GoTo FIMTL
                End If
                Do Until .EOF
                   'Select Case Month(!DATAVENCIMENTO)
                   Select Case !NumParcela
                        Case 1
                            sMes = "Jan"
                        Case 2
                            sMes = "Fev"
                        Case 3
                            sMes = "Mar"
                        Case 4
                            sMes = "Abr"
                        Case 5
                            sMes = "Mai"
                        Case 6
                            sMes = "Jun"
                        Case 7
                            sMes = "Jul"
                        Case 8
                            sMes = "Ago"
                        Case 9
                            sMes = "Set"
                        Case 10
                            sMes = "Out"
                        Case 11
                            sMes = "Nov"
                        Case 12
                            sMes = "Dez"
                   End Select
                   If !NumParcela = 0 Then
                       aMesAno(!NumParcela) = "Jan/14"
                   Else
                       aMesAno(!NumParcela) = sMes & "/14"
                   End If
'                   aMesAno(!NumParcela) = sMes & "/" & Right$(Year(Now), 2)
                   aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9)
                   If !NumParcela = 0 Then
                      aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
'                      aMesAno(0) = Month(!DATAVENCIMENTO) & "/" & Right$(Year(!DATAVENCIMENTO), 2)
                   Else
                      aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                   End If
                   If !NumParcela = 1 Then
                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                        Sql = Sql & " ORDER BY DESCTRIBUTO"
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            x = 1
                            Do Until .EOF
                                aDescTrib(x) = FillSpace(!ABREVTRIBUTO, 15)
                                aValorTributo(x) = FillLeft(!ValorTributo, 17)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                        For s = x To 10
                            aDescTrib(s) = FillSpace(" ", 15)
                        Next
                        sDescTrib = ""
                        For x = 0 To 10
                            sDescTrib = sDescTrib & aDescTrib(x) '449-598
                        Next
                   ElseIf !NumParcela = 0 Then
                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                        Sql = Sql & " ORDER BY DESCTRIBUTO"
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            x = 1
                            Do Until .EOF
                                aValorTributoUnica(x) = FillLeft(!ValorTributo, 17)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                   End If
                  .MoveNext
                Loop
               .Close
            End With
            aDescParc(0) = "UNICA"
            
            If sTipoImposto = "ISS VARIAVEL" Or sTipoImposto = "ISS ESTIMADO" Then
                For x = 1 To 12
                    aDescParc(x) = Format(x, "00") & "/12"
                Next
                sDescParc = ""
                For x = 0 To 12
                    sDescParc = sDescParc & aDescParc(x)
                Next
            Else
                For x = 1 To Val(sQtdeParc)
                    aDescParc(x) = Format(x, "00") & "/03"
                Next
                For x = Val(sQtdeParc) + 1 To 12
                    aDescParc(x) = "00/00"
                Next
                sDescParc = ""
                For x = 0 To 12
                    sDescParc = sDescParc & aDescParc(x)
                Next
            End If
            
            sMesAno = ""
            If aMesAno(0) = "" Then aMesAno(0) = "      "
            For x = Val(sQtdeParc) + 1 To 12
                If aMesAno(x) = "" Then
                   aMesAno(x) = FillLeft(" ", 13)
                End If
            Next
            For x = 0 To 12
                sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
            Next
            
            For x = Val(sQtdeParc) + 1 To 12
                aVencParc(x) = "00/00/0000"
            Next
            sVencParc = ""
            For x = 0 To 12
                sVencParc = sVencParc & aVencParc(x) '662-791
            Next
                    
            For x = Val(sQtdeParc) + 1 To 12
                aNumDoc(x) = "000000000"
            Next
            sNumDoc = ""
'            aNumDoc(0) = "0"
            For x = 0 To 12
                sNumDoc = sNumDoc & Format(aNumDoc(x), "000000000") '662-791
            Next
                    
            sValorTributoUnica = ""
            sValorTributo = ""
            For x = 1 To 10
                If aValorTributo(x) = "" Then
                   aValorTributo(x) = FillLeft("0,00", 17)
                End If
            Next
            
            For x = 1 To 10
                If aValorTributoUnica(x) = "" Then
                   aValorTributoUnica(x) = FillLeft("0,00", 17)
                End If
            Next
            
            For x = 1 To 10
                sValorTributo = sValorTributo & FillLeft(aValorTributo(x), 17)
                sValorTributoUnica = sValorTributoUnica & FillLeft(aValorTributoUnica(x), 17)
            Next
            
            nTotalTrib = 0
            nTotalTribUnica = 0
            For x = 0 To 10
                If x > 0 Then
                    If aValorTributo(x) = "" Then Exit For
                    nTotalTrib = nTotalTrib + CDbl(aValorTributo(x))
                End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
            Next
            sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
            For x = 0 To 10
                If x > 0 Then
                    If aValorTributoUnica(x) = "" Then Exit For
                    nTotalTribUnica = nTotalTribUnica + CDbl(aValorTributoUnica(x))
                End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
            Next
            sTotalTribUnica = FillLeft(CStr(FormatNumber(nTotalTribUnica, 2)), 17)


FIMTL:
        End If 'fim da Taxa Licença
        
        If bISSFixo Then 'ISS FIXO TEM APENAS UM CARNE
'            GoTo Proximo
            sTipoImposto = FillSpace("ISS FIXO/TLL", 20)  '308-327
            sQtdeParc = Format(nQtdeParcF, "0000000000") '427-428
            Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
            Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
            Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
            Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
            Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
            Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 2  "
            'Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 2  AND (DATAVENCIMENTO > '01/10/2014')"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSomaTrib = 0
                For x = 0 To 12
                    aMesAno(x) = ""
                Next
                For x = 1 To 10
                    aValorTributo(x) = ""
                    aValorTributoUnica(x) = ""
                Next
                If .RowCount < 4 Then
                    GoTo FIMFIXO
                End If
                Do Until .EOF
                   'Select Case Month(!DATAVENCIMENTO)
                   Select Case !NumParcela
                        Case 1
                            sMes = "Jan"
                        Case 2
                            sMes = "Fev"
                        Case 3
                            sMes = "Mar"
                        Case 4
                            sMes = "Abr"
                        Case 5
                            sMes = "Mai"
                        Case 6
                            sMes = "Jun"
                        Case 7
                            sMes = "Jul"
                        Case 8
                            sMes = "Ago"
                        Case 9
                            sMes = "Set"
                        Case 10
                            sMes = "Out"
                        Case 11
                            sMes = "Nov"
                        Case 12
                            sMes = "Dez"
                   End Select
                   If !NumParcela = 0 Then
                       aMesAno(!NumParcela) = "Jan/14"
                   Else
                       aMesAno(!NumParcela) = sMes & "/14"
                   End If
'                  aMesAno(!NumParcela) = sMes & "/" & Right$(Year(Now), 2)
                   aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9)
                   If !NumParcela = 0 Then
                      aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
'                      aMesAno(0) = Month(!DATAVENCIMENTO) & "/" & Right$(Year(!DATAVENCIMENTO), 2)
                   Else
                      aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                   End If
                   If !NumParcela = 1 Then
                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                        Sql = Sql & " ORDER BY DESCTRIBUTO"
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            x = 1
                            Do Until .EOF
                                aDescTrib(x) = FillSpace(!ABREVTRIBUTO, 15)
                                aValorTributo(x) = FillLeft(!ValorTributo, 17)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                        For s = x To 10
                            aDescTrib(s) = FillSpace(" ", 15)
                        Next
                        sDescTrib = ""
                        For x = 0 To 10
                            sDescTrib = sDescTrib & aDescTrib(x) '449-598
                        Next
                   ElseIf !NumParcela = 0 Then
                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                        Sql = Sql & " ORDER BY DESCTRIBUTO"
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            x = 1
                            Do Until .EOF
                                aValorTributoUnica(x) = FillLeft(!ValorTributo, 17)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                   End If
                  .MoveNext
                Loop
               .Close
            End With
            aDescParc(0) = "UNICA"
            
            If sTipoImposto = "ISS VARIAVEL" Or sTipoImposto = "ISS ESTIMADO" Then
                For x = 1 To 12
                    aDescParc(x) = Format(x, "00") & "/12"
                Next
                sDescParc = ""
                For x = 0 To 12
                    sDescParc = sDescParc & aDescParc(x)
                Next
            Else
                For x = 1 To Val(sQtdeParc)
                    aDescParc(x) = Format(x, "00") & "/03"
                Next
                For x = Val(sQtdeParc) + 1 To 12
                    aDescParc(x) = "00/00"
                Next
                sDescParc = ""
                For x = 0 To 12
                    sDescParc = sDescParc & aDescParc(x)
                Next
            End If
            
            sMesAno = ""
            If aMesAno(0) = "" Then aMesAno(0) = "      "
            For x = Val(sQtdeParc) + 1 To 12
                If aMesAno(x) = "" Then
                   aMesAno(x) = FillLeft(" ", 13)
                End If
            Next
            For x = 0 To 12
                sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
            Next
            
            For x = Val(sQtdeParc) + 1 To 12
                aVencParc(x) = "00/00/0000"
            Next
            sVencParc = ""
            For x = 0 To 12
                sVencParc = sVencParc & aVencParc(x) '662-791
            Next
                    
            For x = Val(sQtdeParc) + 1 To 12
                aNumDoc(x) = "000000000"
            Next
            sNumDoc = ""
'            aNumDoc(0) = "0"
            For x = 0 To 12
                sNumDoc = sNumDoc & Format(aNumDoc(x), "000000000") '662-791
            Next
                    
            sValorTributoUnica = ""
            sValorTributo = ""
            For x = 1 To 10
                If aValorTributo(x) = "" Then
                   aValorTributo(x) = FillLeft("0,00", 17)
                End If
            Next
            
            For x = 1 To 10
                If aValorTributoUnica(x) = "" Then
                   aValorTributoUnica(x) = FillLeft("0,00", 17)
                End If
            Next
            
            For x = 1 To 10
                sValorTributo = sValorTributo & FillLeft(aValorTributo(x), 17)
                sValorTributoUnica = sValorTributoUnica & FillLeft(aValorTributoUnica(x), 17)
            Next
            
            nTotalTribUnica = 0
            nTotalTrib = 0
            For x = 0 To 10
                If x > 0 Then
                    If aValorTributo(x) = "" Then Exit For
                    nTotalTrib = nTotalTrib + CDbl(aValorTributo(x))
                End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
            Next
            sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
            For x = 0 To 10
                If x > 0 Then
                    If aValorTributoUnica(x) = "" Then Exit For
                    nTotalTribUnica = nTotalTribUnica + CDbl(aValorTributoUnica(x))
                End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
            Next
            sTotalTribUnica = FillLeft(CStr(FormatNumber(nTotalTribUnica, 2)), 17)


FIMFIXO:
        End If 'fim do fixo
        If bISSEstimado Then
            DoEvents
            sTipoImposto = FillSpace("ISS ESTIMADO", 20)  '308-327
            sQtdeParc = Format(nQtdeParcE, "00") '427-428
            Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
            Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
            Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
            Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
            Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
            Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 3 "
            'Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 3 AND (DATAVENCIMENTO > '01/10/2014')"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSomaTrib = 0
                Do Until .EOF
                   'Select Case Month(!DATAVENCIMENTO)
                   Select Case !NumParcela
                        Case 1
                            sMes = "Jan"
                        Case 2
                            sMes = "Fev"
                        Case 3
                            sMes = "Mar"
                        Case 4
                            sMes = "Abr"
                        Case 5
                            sMes = "Mai"
                        Case 6
                            sMes = "Jun"
                        Case 7
                            sMes = "Jul"
                        Case 8
                            sMes = "Ago"
                        Case 9
                            sMes = "Set"
                        Case 10
                            sMes = "Out"
                        Case 11
                            sMes = "Nov"
                        Case 12
                            sMes = "Dez"
                   End Select
                   If !NumParcela = 0 Then
                       aMesAno(!NumParcela) = FillSpace(" ", 6)
                   Else
                       aMesAno(!NumParcela) = sMes & "/14"
                   End If
                   aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9) 'este
                   If !NumParcela = 0 Then
                      aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
                      aMesAno(0) = FillSpace(" ", 6)
                   Else
'                          sDataDoc = Format(!DATADOCUMENTO, "dd/mm/yyyy") '561-570
                      aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                   End If
                   If !NumParcela = 1 Then
'                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
'                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
'                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
'                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
'                        Sql = Sql & "ORDER BY DESCTRIBUTO"
'                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                        With RdoAux3
'                            x = 1
'                        Sql = "SELECT MOBILIARIOATIVIDADEISS.CODATIVIDADE,VALORISS,DESCATIVIDADE FROM MOBILIARIOATIVIDADEISS INNER JOIN "
'                        Sql = Sql & "ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE WHERE MOBILIARIOATIVIDADEISS.CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
                        Sql = "SELECT DISTINCT MOBILIARIOATIVIDADEISS.CODATIVIDADE, MOBILIARIOATIVIDADEISS.VALORISS, ATIVIDADEISS.DESCATIVIDADE FROM MOBILIARIOATIVIDADEISS INNER JOIN "
                        Sql = Sql & "ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE INNER JOIN TABELAISS ON ATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV  "
                        Sql = Sql & "Where (MOBILIARIOATIVIDADEISS.CODMOBILIARIO = " & Val(RdoAux!CODREDUZIDO) & ") And (TABELAISS.TipoISS = 12)"
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            For x = 1 To 10
                                aCodTrib(x) = FillLeft("", 15)
                                aDescTrib(x) = FillLeft("", 50)
                                aValorTributo(x) = FillLeft("0,00", 17)
                            Next
                            x = 1
                            Do Until .EOF
                                If x > 10 Then Exit Do
                                
                                aValorTributo(x) = FillLeft(FormatNumber(!valoriss * RetornaAliquotaISS(!codatividade, Now) * nUfir, 2), 17)
                                aCodTrib(x) = FillSpace(Left$(!codatividade, 15), 15)
                                aDescTrib(x) = FillSpace(Left$(!descatividade, 50), 50)
'                                aValorTributo(x) = FillLeft(!VALORTRIBUTO, 17)
'                                aDescTrib(x) = FillSpace(!ABREVTRIBUTO, 15)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                        For s = x To 10
                            aCodTrib(s) = FillSpace(" ", 15)
                            aDescTrib(s) = FillSpace(" ", 50)
                        Next
                        sCodTrib = ""
                        sDescTrib = ""
                        For x = 0 To 10
                            sCodTrib = sCodTrib & aCodTrib(x)
                            sDescTrib = sDescTrib & aDescTrib(x)
                        Next
                   Else
                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                        Sql = Sql & " ORDER BY DESCTRIBUTO"
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            x = 1
                            Do Until .EOF
                                aValorTributoUnica(x) = FillLeft(!ValorTributo, 17)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                   End If

                  
                  .MoveNext
                Loop
               .Close
            End With
            aDescParc(0) = "     "
            For x = 1 To 12
            'For x = 1 To Val(sQtdeParc)
                aDescParc(x) = Format(x, "00") & "/12"
            Next
            sDescParc = ""
            'For x = 0 To Val(sQtdeParc)
            For x = 0 To 12
                sDescParc = sDescParc & aDescParc(x)
            Next
'            For x = Val(sQtdeParc) + 1 To 10
'                aDescParc(x) = "00/00"
'            Next
            
            sMesAno = ""
            'If aMesAno(0) = "" Then aMesAno(0) = "      "
            aMesAno(0) = FillLeft(" ", 6)
            For x = Val(sQtdeParc) To 12
                If aMesAno(x) = "" Then
                   aMesAno(x) = "      "
                End If
            Next
            For x = 0 To Val(sQtdeParc)
                sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
            Next
            
            For x = Val(sQtdeParc) + 1 To 12
                aVencParc(x) = "00/00/0000"
            Next
            sVencParc = ""
            For x = 0 To 12
                sVencParc = sVencParc & aVencParc(x) '662-791
            Next
                    
            For x = Val(sQtdeParc) + 1 To 12
                aNumDoc(x) = "000000000"
            Next
            sNumDoc = ""
            aNumDoc(0) = "0"
            For x = 0 To 12
                sNumDoc = sNumDoc & Format(aNumDoc(x) & Modulo11(aNumDoc(x)), "000000000") '662-791
            Next
                    
            sValorTributoUnica = ""
            sValorTributo = ""
            For x = 1 To 10
                If aValorTributo(x) = "" Then
                   aValorTributo(x) = FillLeft("0,00", 17)
                End If
            Next
            
            For x = 1 To 10
                sValorTributo = sValorTributo & FillLeft(FormatNumber(aValorTributo(x) * 12, 2), 17)
            Next
            
            nTotalTrib = 0
            For x = 0 To 10
                If x > 0 Then
                    If aValorTributo(x) = "" Then Exit For
                    nTotalTrib = nTotalTrib + CDbl(aValorTributo(x))
                End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
            Next
            sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
        
        End If 'fim do estimado
        
        If bISSVariavel Then
        GoTo Proximo
            sTipoImposto = FillSpace("ISS VARIAVEL", 20)  '308-327
            sQtdeParc = Format(nQtdeParcE, "00") '427-428
            Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
            Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
            Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
            Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
            Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
            Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO = " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 5  "
            'Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO = " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 5  AND (DATAVENCIMENTO > '01/10/2014')"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSomaTrib = 0
                Do Until .EOF
                   'Select Case Month(!DATAVENCIMENTO)
                   Select Case !NumParcela
                        Case 1
                            sMes = "Jan"
                        Case 2
                            sMes = "Fev"
                        Case 3
                            sMes = "Mar"
                        Case 4
                            sMes = "Abr"
                        Case 5
                            sMes = "Mai"
                        Case 6
                            sMes = "Jun"
                        Case 7
                            sMes = "Jul"
                        Case 8
                            sMes = "Ago"
                        Case 9
                            sMes = "Set"
                        Case 10
                            sMes = "Out"
                        Case 11
                            sMes = "Nov"
                        Case 12
                            sMes = "Dez"
                   End Select
                   If !NumParcela = 0 Then
                      aMesAno(!NumParcela) = FillSpace(" ", 6)
                   Else
                      aMesAno(!NumParcela) = sMes & "/10"
                   End If
                   aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9)
                   If !NumParcela = 0 Then
                      aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
                      aMesAno(0) = FillSpace(" ", 6)
                   Else
'                          sDataDoc = Format(!DATADOCUMENTO, "dd/mm/yyyy") '561-570
                      aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                   End If
                   If !NumParcela = 1 Then
'                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
'                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
'                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
'                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
'                        Sql = Sql & "ORDER BY DESCTRIBUTO"
'                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                        With RdoAux3
'                            x = 1
                        Sql = "SELECT MOBILIARIOATIVIDADEISS.CODATIVIDADE,VALORISS,DESCATIVIDADE FROM MOBILIARIOATIVIDADEISS INNER JOIN "
                        Sql = Sql & "ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE WHERE MOBILIARIOATIVIDADEISS.CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            For x = 1 To 10
                                aCodTrib(x) = FillLeft("", 15)
                                aDescTrib(x) = FillLeft("", 50)
                                aValorTributo(x) = FillLeft("0,00", 17)
                            Next
                            x = 1
                            Do Until .EOF
                                If x > 10 Then Exit Do
                                aValorTributo(x) = FillLeft(!valoriss, 17)
                                aCodTrib(x) = FillSpace(Left$(!codatividade, 15), 15)
                                aDescTrib(x) = FillSpace(Left$(!descatividade, 50), 50)
'                                aValorTributo(x) = FillLeft(!VALORTRIBUTO, 17)
'                                aDescTrib(x) = FillSpace(!ABREVTRIBUTO, 15)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                        For s = x To 10
                            aCodTrib(s) = FillSpace(" ", 15)
                            aDescTrib(s) = FillSpace(" ", 50)
                        Next
                        sCodTrib = ""
                        sDescTrib = ""
                        For x = 0 To 10
                            sCodTrib = sCodTrib & aCodTrib(x)
                            sDescTrib = sDescTrib & aDescTrib(x)
                        Next
                   Else
                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                        Sql = Sql & " ORDER BY DESCTRIBUTO"
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            x = 1
                            Do Until .EOF
                                aValorTributoUnica(x) = FillLeft(!ValorTributo, 17)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                   End If

                  
                  .MoveNext
                Loop
               .Close
            End With
            aDescParc(0) = "     "
            For x = 1 To 12
            'For x = 1 To Val(sQtdeParc)
                aDescParc(x) = Format(x, "00") & "/12"
            Next
            sDescParc = ""
            'For x = 0 To Val(sQtdeParc)
            For x = 0 To 12
                sDescParc = sDescParc & aDescParc(x)
            Next
'            For x = Val(sQtdeParc) + 1 To 10
'                aDescParc(x) = "00/00"
'            Next
            
            sMesAno = ""
            'If aMesAno(0) = "" Then aMesAno(0) = "      "
            aMesAno(0) = FillLeft(" ", 6)
'            For x = Val(sQtdeParc) To 12
'                If aMesAno(x) = "" Then
'                   aMesAno(x) = "      "
'                End If
'            Next
            For x = 0 To Val(sQtdeParc)
                sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
            Next
            
            For x = Val(sQtdeParc) + 1 To 12
                aVencParc(x) = "00/00/0000"
            Next
            sVencParc = ""
            For x = 0 To 12
                sVencParc = sVencParc & aVencParc(x) '662-791
            Next
                    
            For x = Val(sQtdeParc) + 1 To 12
                aNumDoc(x) = "000000000"
            Next
            sNumDoc = ""
            aNumDoc(0) = "0"
            For x = 0 To 12
                sNumDoc = sNumDoc & Format(aNumDoc(x), "000000000") '662-791
            Next
                    
            sValorTributoUnica = ""
            sValorTributo = ""
            For x = 1 To 10
                If aValorTributo(x) = "" Then
                   aValorTributo(x) = FillLeft("0,00", 17)
                End If
            Next
            
            For x = 1 To 10
                sValorTributo = sValorTributo & FillLeft(aValorTributo(x), 17)
            Next
            
            If Trim(sTipoImposto) <> "ISS VARIAVEL" And Trim(sTipoImposto) <> "ISS ESTIMADO" Then
                nTotalTrib = 0
                For x = 0 To 10
                    If x > 0 Then
                        If aValorTributo(x) = "" Then Exit For
                        nTotalTrib = nTotalTrib + CDbl(aValorTributo(x))
                    End If
    '                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
    '                sValorTributo = sValorTributo & aValorTributo(x) '449-598
                Next
            Else
                nTotalTrib = 0
            End If
            sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
        
        End If 'fim do variavel

        If bVigSanit Then
            sTipoImposto = FillSpace("VIGIL.SANITÁRIA", 20)  '308-327
            sQtdeParc = Format(nQtdeParcS, "00") '427-428
            Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
            Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
            Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
            Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
            Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
            Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO = " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 13"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSomaTrib = 0
                For x = 0 To 12
                    aMesAno(x) = ""
                Next
                For x = 1 To 10
                    aValorTributo(x) = ""
                    aValorTributoUnica(x) = ""
                Next
                If .RowCount < 4 Then
                    GoTo FIMVS
                End If
                Do Until .EOF
                   'Select Case Month(!DATAVENCIMENTO)
                   Select Case !NumParcela
                        Case 1
                            sMes = "Jan"
                        Case 2
                            sMes = "Fev"
                        Case 3
                            sMes = "Mar"
                        Case 4
                            sMes = "Abr"
                        Case 5
                            sMes = "Mai"
                        Case 6
                            sMes = "Jun"
                        Case 7
                            sMes = "Jul"
                        Case 8
                            sMes = "Ago"
                        Case 9
                            sMes = "Set"
                        Case 10
                            sMes = "Out"
                        Case 11
                            sMes = "Nov"
                        Case 12
                            sMes = "Dez"
                   End Select
                   If !NumParcela = 0 Then
                       aMesAno(!NumParcela) = "Jan/11"
                   Else
                       aMesAno(!NumParcela) = sMes & "/11"
                   End If
                   aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9)
                   If !NumParcela = 0 Then
                      aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
'                      aMesAno(0) = FillSpace(" ", 6)
                   Else
'                          sDataDoc = Format(!DATADOCUMENTO, "dd/mm/yyyy") '561-570
                      aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                   End If
                   If !NumParcela = 1 Then
'                        Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
'                        Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
'                        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
'                        Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
'                        Sql = Sql & "ORDER BY DESCTRIBUTO"
'                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                        With RdoAux3
'                            x = 1
                        Sql = "SELECT MOBILIARIOATIVIDADEVS.*,VIGSANITARIA.VALORALIQ,DESCVIGSANITARIA FROM MOBILIARIOATIVIDADEVS INNER JOIN "
                        Sql = Sql & " VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT AND "
                        Sql = Sql & " MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT WHERE MOBILIARIOATIVIDADEVS.CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            For x = 1 To 10
                                aCodTrib(x) = FillLeft("", 15)
                                aDescTrib(x) = FillLeft("", 50)
                                aValorTributo(x) = FillLeft("0,00", 17)
                            Next
                            x = 1
                            Do Until .EOF
                                If x > 10 Then Exit Do
                                aValorTributo(x) = FillLeft((FormatNumber(!VALORALIQ, 2)), 17)
                                aCodTrib(x) = FillSpace(Left$(!CODVIGSANIT & "-" & !SUBCODVIGSANIT, 15), 15)
                                aDescTrib(x) = FillSpace(Left$(!DESCVIGSANITARIA, 50), 50)
'                                aValorTributo(x) = FillLeft(!VALORTRIBUTO, 17)
'                                aDescTrib(x) = FillSpace(!ABREVTRIBUTO, 15)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                        For s = x To 10
                            aCodTrib(s) = FillSpace(" ", 15)
                            aDescTrib(s) = FillSpace(" ", 50)
                        Next
                        sCodTrib = ""
                        sDescTrib = ""
                        For x = 0 To 10
                            sCodTrib = sCodTrib & aCodTrib(x)
                            sDescTrib = sDescTrib & aDescTrib(x)
                        Next
                   Else
                        Sql = "SELECT MOBILIARIOATIVIDADEVS.*,VIGSANITARIA.VALORALIQ,DESCVIGSANITARIA FROM MOBILIARIOATIVIDADEVS INNER JOIN "
                        Sql = Sql & " VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT AND "
                        Sql = Sql & " MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT WHERE MOBILIARIOATIVIDADEVS.CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            x = 1
                            Do Until .EOF
                                aValorTributoUnica(x) = FillLeft(FormatNumber(!VALORALIQ, 2), 17)
                                x = x + 1
                               .MoveNext
                            Loop
                           .Close
                        End With
                   End If

                  
                  .MoveNext
                Loop
               .Close
            End With
            aDescParc(0) = "UNICA"
            'For x = 1 To 12
            For x = 1 To Val(sQtdeParc)
                aDescParc(x) = Format(x, "00") & "/04"
            Next
            For x = Val(sQtdeParc) + 1 To 12
                aDescParc(x) = "00/00"
            Next
            sDescParc = ""
            'For x = 0 To Val(sQtdeParc)
            For x = 0 To 12
                sDescParc = sDescParc & aDescParc(x)
            Next
            
            sMesAno = ""
            'If aMesAno(0) = "" Then aMesAno(0) = "      "
            aMesAno(0) = FillLeft(" ", 6)
            For x = Val(sQtdeParc) + 1 To 12
                If aMesAno(x) = "" Then
                   aMesAno(x) = "      "
                End If
            Next
            For x = 0 To 12
                sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
            Next
            
            For x = Val(sQtdeParc) + 1 To 12
                aVencParc(x) = "00/00/0000"
            Next
            sVencParc = ""
            For x = 0 To 12
                sVencParc = sVencParc & aVencParc(x) '662-791
            Next
                    
            For x = Val(sQtdeParc) + 1 To 12
                aNumDoc(x) = "000000000"
            Next
            sNumDoc = ""
            aNumDoc(0) = "0"
            For x = 0 To 12
                sNumDoc = sNumDoc & Format(aNumDoc(x), "000000000") '662-791
            Next
                    
            sValorTributoUnica = ""
            sValorTributo = ""
            For x = 1 To 10
                If aValorTributo(x) = "" Then
                   aValorTributo(x) = FillLeft("0,00", 17)
                End If
            Next
            For x = 1 To 10
                If aValorTributoUnica(x) = "" Then
                   aValorTributoUnica(x) = FillLeft("0,00", 17)
                End If
            Next
            
            For x = 1 To 10
                sValorTributo = sValorTributo & FillLeft(FormatNumber(aValorTributo(x), 2), 17)
                sValorTributoUnica = sValorTributoUnica & FillLeft(aValorTributoUnica(x), 17)
            Next
            
            nTotalTrib = 0
            nTotalTribUnica = 0
            For x = 0 To 10
                If x > 0 Then
                    If aValorTributo(x) = "" Then Exit For
                    nTotalTrib = nTotalTrib + CDbl(aValorTributo(x))
                End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
            Next
            sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
            For x = 0 To 10
                If x > 0 Then
                    If aValorTributoUnica(x) = "" Then Exit For
                    nTotalTribUnica = nTotalTribUnica + CDbl(aValorTributoUnica(x))
                End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
            Next
            sTotalTribUnica = FillLeft(CStr(FormatNumber(nTotalTribUnica, 2)), 17)
FIMVS:
        End If 'fim da vig.sanitaria


'If Trim(sTipoImposto) <> "ISS ESTIMADO" And Trim(sTipoImposto) <> "ISS VARIAVEL" Then GoTo PROXIMO
        If Trim(sTipoImposto) = "ISS ESTIMADO" Then
            nCodLanc = 3
        ElseIf Trim(sTipoImposto) = "ISS VARIAVEL" Then
            nCodLanc = 5
        ElseIf Trim(sTipoImposto) = "ISS FIXO/TLL" Then
            nCodLanc = 2
        ElseIf Trim(sTipoImposto) = "TAXA DE LICENÇA" Then
            nCodLanc = 6
        ElseIf Trim(sTipoImposto) = "VIGIL.SANITÁRIA" Then
            nCodLanc = 13
        End If
        If Trim(sTipoImposto) <> "ISS ESTIMADO" And Trim(sTipoImposto) <> "ISS VARIAVEL" Then
           'aValorParc(0) = ((nTotalTribUnica * nUfir) - (nTotalTribUnica * nUfir * 0.05)) + nExpUnica
           
           'aValorParc(0) = FormatNumber(((nTotalTribUnica * nUfir) - (nTotalTribUnica * nUfir * 0.05)) + nExpUnica, 2)
           aValorParc(0) = FormatNumber(nTotalTribUnica + nExpUnica, 2)
        Else
           aValorParc(0) = "0,00"
        End If
        If Trim(sTipoImposto) = "VIGIL.SANITÁRIA" Then
            For x = 1 To Val(sQtdeParc)
                aValorParc(x) = FormatNumber(((nTotalTrib * nUfir) / 4) + nExpParc, 2)
            Next
        Else
            For x = 1 To Val(sQtdeParc)
                aValorParc(x) = FormatNumber(nTotalTrib + nExpParc, 2)
            Next
        End If
        For x = Val(sQtdeParc) + 1 To 12
            aValorParc(x) = "0,00"
        Next
        sValorParc = ""
        For x = 0 To 12
            sValorParc = sValorParc & FillLeft(aValorParc(x), 18) '662-791
        Next
        
        sDataProc = Format(Now, "dd/mm/yyyy") '439-448
        'sDataDoc = FillLeft(Format(Now, "dd/mm/yyyy"), 15) '561-570
        
        If Trim(sTipoImposto) <> "ISS ESTIMADO" And Trim(sTipoImposto) <> "ISS VARIAVEL" Then
            dDataBase = "07/10/1997"
            aCodBarra(0) = Format("0", "00000000000000000000000000000000000000000000")
            For x = 0 To Val(sQtdeParc)
                If Not IsDate(aVencParc(x)) Then Exit For
                nFatorVencto = CDate(aVencParc(x)) - dDataBase
                'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "02345000024" & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "0000000") & "0003300"
                aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "023" & Format(aNumDoc(x), "00000000") & Format(Val(Chomp(aNumDoc(x), chomp_righT, 1)), "0000000") & "0003300"
                'aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
            Next
            For x = 4 To 12
                aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
            Next
        Else
            dDataBase = "07/10/1997"
            aCodBarra(0) = Format("0", "00000000000000000000000000000000000000000000")
            For x = 1 To Val(sQtdeParc)
                If Not IsDate(aVencParc(x)) Then Exit For
                nFatorVencto = CDate(aVencParc(x)) - dDataBase
                If Trim(sTipoImposto) = "ISS ESTIMADO" Then
'                   aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & sAgencia & Format(Val(aNumDoc(x)), "0000000") & "0003300"
                   aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "9" & "1235028" & Format(aNumDoc(x) & Modulo11(aNumDoc(x)), "0000000000000") & "0102"
                Else
                    'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(0, 2))), "0000000000") & sAgencia & Format(Val(aNumDoc(x)), "0000000") & "0003300"
                   aCodBarra(x) = "XXXXXXXXXXXXXXXXXX"
                   'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(0, 2))), "0000000000") & "023" & Format(aNumDoc(x), "00000000") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "0000000") & "0003300"
                End If
                'aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(aNumDoc(x)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
            Next
            For x = Val(sQtdeParc) + 1 To 12
                aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
            Next
        End If
        sCodBarra = ""
        For x = 0 To 12
            sCodBarra = sCodBarra & aCodBarra(x) '792-1012
        Next
        sValorEXP = FillLeft(FormatNumber(nExpParc, 2), 17)
        If Trim(sTipoImposto) = "ISS ESTIMADO" Then
           ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
           'ax = sExercicio & sContribuinte & sFantasia & sEnd & sCompl & sBairro & sCidEntrega & sCep
           
           'ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCep & sTipoImposto & sInscricao & sCodInscricao & sDescAtiv
           ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sDescAtiv
           'ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataDoc & sDataProc & sNumDoc & sCodTrib & sValorTributoUnica & sValorTributo
           ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataProc & sDataProc & sNumDoc & sCodTrib & sValorTributoUnica & sValorTributo
           ax = ax & sTotalTrib & "          " & sVencParc & sValorParc & sCodBarra & sDescTrib & "0" & sValorEXP & FillLeft(FormatNumber("20,28", 2), 17)
           tTipo = "1"
           Print #1, ax
        ElseIf Trim(sTipoImposto) = "ISS VARIAVEL" Then
           sTotalTrib = FillLeft("0,00", 17)
           ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
           ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sCodInscricao & sDescAtiv
           ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataDoc & sDataProc & sNumDoc & sCodTrib & sValorTributoUnica & sValorTributo
           ax = ax & sTotalTrib & "          " & sVencParc & sValorParc & sCodBarra & sDescTrib & "0" & sValorEXP & FillLeft(FormatNumber("20,28", 2), 17)
           tTipo = "2"
           Print #2, ax
        ElseIf Trim(sTipoImposto) = "VIGIL.SANITÁRIA" Then
           sTotalTrib = FillLeft("0,00", 17)
           'ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
           ax = sExercicio & sContribuinte & sFantasia & sEnd & sCompl & sBairro & sCidEntrega & sCepEntrega
           ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sCodInscricao & sDescAtiv
           ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataDoc & sDataProc & sNumDoc & sCodTrib & sValorTributoUnica & sValorTributo
           ax = ax & sTotalTrib & "          " & sVencParc & sValorParc & sCodBarra & sDescTrib & "0"
           tTipo = "4"
           'Print #4, ax
        ElseIf Trim(sTipoImposto) = "ISS FIXO/TLL" Then
           ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
           ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sCodInscricao & sDescAtiv
           ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataDoc & sDataProc & sNumDoc & sDescTrib & sValorTributoUnica & sValorTributo
           ax = ax & sTotalTrib & "          " & sVencParc & sValorParc & sCodBarra
           tTipo = "3"
           'Print #3, ax
        ElseIf Trim(sTipoImposto) = "TAXA DE LICENÇA" Then
           ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
           ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sCodInscricao & sDescAtiv
           ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataDoc & sDataProc & sNumDoc & sDescTrib & sValorTributoUnica & sValorTributo
           ax = ax & sTotalTrib & "          " & sVencParc & sValorParc & sCodBarra
           tTipo = "3"
          ' Print #3, ax
        End If

        tDado = ax
        ax = tDado & "," & tEnd & "," & tNum & "," & tTipo
        
        Sql = "INSERT LASERTMP (DADO,CIDADE,BAIRRO,ENDERECO,NUMERO,TIPO) VALUES('" & Mask(tDado) & "','"
        Sql = Sql & Trim(Mask(tBairro)) & "','" & Trim(Mask(tCidade)) & "','" & Trim(Mask(tEnd)) & "','" & tNum & "','" & tTipo & "')"
        cn.Execute Sql, rdExecDirect
        
Proximo:
        xId = xId + 1
       .MoveNext
 '      Exit Do
    Loop
   .Close
End With

'Close #4
'Close #3
'Close #2
Close #1
'Exit Sub
ORDENA:

Open sPathBin & "\LASERISSEST.TXT" For Output As #1
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=1 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #1, Trim(!dado) & Format(xId, "00000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #1

Open sPathBin & "\LASERISSVAR.TXT" For Output As #1
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=2 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #1, Trim(!dado) & Format(xId, "00000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #1

'Open sPathBin & "\LASERISSFIXOTL.TXT" For Output As #1
'Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=3 ORDER BY CIDADE,ENDERECO,NUMERO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    xId = 1
'    Do Until .EOF
'        Print #1, Trim(!dado) & Format(xId, "000000")
 '       xId = xId + 1
 '      .MoveNext
 '   Loop
'End With
'Close #1

'Open sPathBin & "\LASERVIGSANIT.TXT" For Output As #1
'Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=4 ORDER BY CIDADE,ENDERECO,NUMERO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    xId = 1
'    Do Until .EOF
'        Print #1, Trim(!dado) & Format(xId, "000000")
'        xId = xId + 1
'       .MoveNext
'    Loop
'End With
'Close #1
MsgBox "fim"
End Sub
Private Sub GeraVS()
Dim xId As Long, nNumRec As Long, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset, x As Integer
Dim nCodLogr As Long, nNum As Integer, nQtdeParcS As Integer, t As Integer, nCodEsc As Integer
Dim nExpParc As Double, nExpUnica As Double

'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sFantasia As String, sEnd As String, sCompl As String, sBairro As String, sCEP As String
Dim sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim sTipoImposto As String, sInscricao As String, sQtdeParc As String, sCodAtiv As String, sDescAtiv As String, sCodInscricao As String
Dim aDescParc(0 To 12) As String, sDescParc As String
Dim aCodTrib(0 To 10) As String, sCodTrib As String
Dim aDescTrib(0 To 10) As String, sDescTrib As String
Dim aVencParc(0 To 12) As String, sVencParc As String
Dim aValorTributoUnica(0 To 12) As String, sValorTributoUnica As String
Dim aValorTributo(0 To 12) As String, sValorTributo As String
Dim aValorParc(0 To 12) As String, sValorParc As String
Dim aValorParcSEXP(0 To 12) As String, sValorParcSEXP As String
Dim aMesAno(0 To 12) As String, sMesAno As String, sMes As String
Dim aNumDoc(0 To 12) As String, sNumDoc As String
Dim nTotalTrib As Double, sTotalTrib As String
Dim nTotalTribUnica As Double, sTotalTribUnica As String
Dim aCodBarra(0 To 12) As String, sCodBarra As String
Dim dDataBase As Date, nUfir As Double
Dim tDado As String, tEnd As String, tNum As Integer, tTipo As Integer
Dim tCidade As String, tBairro As String, bAchou As Boolean
Dim sValorEXP As String, strLinha1 As String, strLinha2 As String, aCodDup() As Long, sCod As String, l As Integer, k As Integer

'GoTo ORDENA
'********************************
' PARAMETROS DAS PARCELAS
'********************************
k = 0
ReDim aCodDup(0)
'Open sPathBin & "\LASERVIGSANIT3.TXT" For Input As #1
'   Do While Not EOF(1)
'        Line Input #1, strLinha1
'        sCod = Mid(strLinha1, 346, 6)
'        ReDim Preserve aCodDup(UBound(aCodDup) + 1)
'        aCodDup(UBound(aCodDup)) = CLng(sCod)
'   Loop
'Close #1

nUfir = RetornaUFIR(2014)
sAgencia = "02345000024"

'PARCELAS PARA VIGILÂNCIA SANITÁRIA
Sql = "SELECT QTDEPARCELA FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & 2014 & " AND CODTIPO=5"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nQtdeParcS = !qtdeparcela
   .Close
End With

Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & 2014
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nExpParc = FormatNumber(RdoAux!VALORPARCELA, 2)
nExpUnica = FormatNumber(RdoAux!ValorUnica, 2)

'## ************************************************ ##
'## ***************** G U I A ********************** ##
'## ************************************************ ##

Sql = "TRUNCATE TABLE LASERTMP"
cn.Execute Sql, rdExecDirect

Open sPathBin & "\LASERVIGSANIT.TXT" For Output As #1

Sql = "SELECT DISTINCT CODREDUZIDO From DEBITOPARCELA "
'Sql = Sql & "WHERE  CODREDUZIDO=106842 AND (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 13)   ORDER BY CODREDUZIDO"
Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 13)   ORDER BY CODREDUZIDO"
 
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        
'        If !CODREDUZIDO = 100102 Then MsgBox "TESTE"
'        bAchou = False
'        For l = 1 To UBound(aCodDup)
'            If aCodDup(l) = !CODREDUZIDO Then
'                bAchou = True
'                Exit For
'            End If
'        Next
'        If bAchou Then
'            GoTo proximo
'        End If

        If xId Mod 50 = 0 Then
           CallPb xId, nNumRec
        End If
        Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
        Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,MOBILIARIO.RESPCONTABIL,"
        Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
        Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
        Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
        Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                GoTo Proximo
            End If
            nCodEsc = Val(SubNull(!RESPCONTABIL))
            nCodLogr = !CodLogradouro
            sExercicio = "2014" '1-4
            sCodInscricao = Format(!codigomob, "00000000000000")
            sContribuinte = FillSpace(!RazaoSocial, 40) '5-44
            sFantasia = FillSpace(SubNull(!NOMEFANTASIA), 40) '45-84
            sCodAtiv = Format(!codatividade, "00000000000000")
            sDescAtiv = FillSpace(Left$(!ATIVEXTENSO, 50), 50)
            Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    sEnd = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero), 46) '212-257
                    nNum = RdoAux2!Numero
                Else
                    nNum = 0
                End If
               .Close
            End With
            sCEP = RetornaCEP(nCodLogr, nNum)
            sCompl = FillSpace(SubNull(Left(!Complemento, 20)), 20) '258-277
            sBairro = FillSpace(SubNull(!DescBairro), 30) '278-307
            Sql = "SELECT MOBILIARIOENDENTREGA.CODMOBILIARIO, MOBILIARIOENDENTREGA.CODLOGRADOURO, MOBILIARIOENDENTREGA.NOMELOGRADOURO, "
            Sql = Sql & "MOBILIARIOENDENTREGA.NUMIMOVEL,MOBILIARIOENDENTREGA.COMPLEMENTO, MOBILIARIOENDENTREGA.UF,MOBILIARIOENDENTREGA.CODCIDADE,"
            Sql = Sql & "MOBILIARIOENDENTREGA.CODBAIRRO, MOBILIARIOENDENTREGA.CEP,MOBILIARIOENDENTREGA.DESCBAIRRO, MOBILIARIOENDENTREGA.DESCCIDADE, BAIRRO.DESCBAIRRO AS DESCBAIRRO2,"
            Sql = Sql & "CIDADE.DESCCIDADE AS DESCCIDADE2 FROM BAIRRO INNER JOIN MOBILIARIOENDENTREGA ON BAIRRO.SIGLAUF = MOBILIARIOENDENTREGA.UF AND BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO INNER JOIN "
            Sql = Sql & "CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE Where CODMOBILIARIO = " & !codigomob
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    If !CodLogradouro > 0 Then
'                        GoTo proximo
                        sEndEntrega = FillSpace(!NomeLogradouro & " Nº " & CStr(!numimovel), 46) '85-130
                        sNumEntrega = SubNull(!numimovel)
                        If !CodBairro = 0 Then
                            sBairroEntrega = FillSpace(SubNull(!DescBairro), 30) '151-180
                        ElseIf !CodBairro = 999 Then
                            sBairroEntrega = FillSpace(" ", 30) '151-180
                        Else
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30)
                        End If
                        If !CodCidade = 0 Then
                            sCidEntrega = FillSpace(SubNull(!desccidade), 20) '181-200
                        Else
                            sCidEntrega = FillSpace(SubNull(!desccidade2), 20)
                        End If
                        If !CodCidade = 413 Then
                            sCepEntrega = Format(!Cep, "00000-000") '201-209
                        Else
                            sCepEntrega = RetornaCEP(!CodLogradouro, Val(sNumEntrega))
                        End If
                        sComplEntrega = FillSpace(SubNull(!Complemento), 20)
                        sUFEntrega = !UF '210-211
                    Else
                        Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
                        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux4
                            If .RowCount > 0 Then
                                sNumEntrega = SubNull(RdoAux3!numimovel)
                                sEndEntrega = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & sNumEntrega, 46) '85-130
                                sCepEntrega = RetornaCEP(RdoAux3!CodLogradouro, Val(sNumEntrega))
                            End If
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30) '151-180
                            sCidEntrega = FillSpace(SubNull(RdoAux3!desccidade2), 20) '181-200
                            sComplEntrega = FillSpace(SubNull(RdoAux2!Complemento), 20)
                            sUFEntrega = RdoAux3!UF '210-211
                           .Close
                        End With
                    End If
                Else
'                    GoTo proximo
                    nCodLogr = 0
                    sEndEntrega = sEnd '85-130
                    sBairroEntrega = FillSpace(sBairro, 30) '151-180
                    sCidEntrega = FillSpace("JABOTICABAL", 20) '181-200
                    sCepEntrega = "14870-000" '201-209
                    sComplEntrega = FillSpace(" ", 20)
                    sUFEntrega = "SP" '210-211
                End If
               .Close
            End With
           .Close
        End With
        
        
        If nCodEsc > 0 Then
            '***ENDERECO CONTADOR***
            For t = 1 To UBound(aEnd)
                With aEnd(t)
                    If aEnd(t).nCodigo = nCodEsc Then
                        If aEnd(t).bRecebe Then
                            sEndEntrega = FillSpace(.sLogradouro & " Nº " & CStr(.nNumero), 46) '85-130
                            sNumEntrega = .nNumero
                            sBairroEntrega = FillSpace(.sBairro, 30)
                            sCidEntrega = FillSpace(.sCidade, 20)
                            sUFEntrega = .sUF
                            sComplEntrega = FillSpace(" ", 20)
                            sCepEntrega = .sCEP
                        End If
                        Exit For
                    End If
                End With
            Next
            '***********************
        End If
        
        tEnd = sEndEntrega
        tNum = Val(sNumEntrega)
        tBairro = sBairroEntrega
        tCidade = sCidEntrega
        If Left(sCepEntrega, 1) = "_" Then sCepEntrega = "         "
        If Trim(sCepEntrega) = "" Then sCepEntrega = "14870-000"
        
        'INSCRICAO
        sInscricao = Format(Val(sCodInscricao), "000000000000000")
'        If nCodLogr > 0 Then
'            Sql = "SELECT DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,"
'            Sql = Sql & "CODLOGR,LI_NUM FROM vwCnsImovel WHERE CODLOGR=" & nCodLogr
'            Sql = Sql & " AND LI_NUM=" & nNum
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                If .RowCount > 0 Then
'                   sInscricao = Format(!Distrito, "00") & Format(!Setor, "00") & Format(!Quadra, "0000") & Format(!Lote, "00000") & Format(!Seq, "00")
'                End If
'               .Close
'            End With
'        Else
'            sInscricao = "000000000000000"
'        End If


        sTipoImposto = FillSpace("VIGIL.SANITÁRIA", 20)  '308-327
        sQtdeParc = Format(nQtdeParcS, "00") '427-428
        Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
        Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
        Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
        Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
        Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
        Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO = " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = 13"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nSomaTrib = 0
            For x = 0 To 12
                aMesAno(x) = ""
            Next
            For x = 1 To 10
                aValorTributo(x) = ""
                aValorTributoUnica(x) = ""
            Next
            If .RowCount < 4 Then
                GoTo Proximo
            End If
            Do Until .EOF
               Select Case !NumParcela
                    Case 1
                        sMes = " 01"
                    Case 2
                        sMes = " 02"
                    Case 3
                        sMes = " 03"
                    Case 4
                        sMes = " 04"
                    Case 5
                        sMes = " 05"
                    Case 6
                        sMes = " 06"
                    Case 7
                        sMes = " 07"
                    Case 8
                        sMes = " 08"
                    Case 9
                        sMes = " 09"
                    Case 10
                        sMes = " 10"
                    Case 11
                        sMes = " 11"
                    Case 12
                        sMes = " 12"
               End Select
               If !NumParcela = 0 Then
                   aMesAno(!NumParcela) = "UNICA"
               Else
                   aMesAno(!NumParcela) = sMes & "/04"
               End If
               aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9)
               If !NumParcela = 0 Then
                  aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
               Else
                  aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
               End If
               
               
               If !NumParcela = 1 Then
                    Sql = "SELECT * FROM vwMOBILIARIOATIVIDADEVS2 WHERE CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
'                    Sql = "SELECT MOBILIARIOATIVIDADEVS.*,VIGSANITARIA.VALORALIQ,DESCVIGSANITARIA FROM MOBILIARIOATIVIDADEVS INNER JOIN "
'                    Sql = Sql & " VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT AND "
'                    Sql = Sql & " MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT WHERE MOBILIARIOATIVIDADEVS.CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        If .RowCount = 0 Then
                            k = k + 1
                        End If
                        For x = 1 To 10
                            aCodTrib(x) = FillLeft("", 15)
                            aDescTrib(x) = FillLeft("", 50)
                            aValorTributo(x) = FillLeft("0,00", 17)
                        Next
                        x = 1
                        Do Until .EOF
                            If x > 10 Then Exit Do
                            'aValorTributo(x) = FillLeft((FormatNumber(!VALORALIQ * nUfir, 2)), 17)
                            'aValorTributo(x) = FillLeft((FormatNumber(!valor * nUfir, 2)), 17)
                            aValorTributo(x) = FillLeft((FormatNumber(!valor * !QTDE, 2)), 17)
                            'aCodTrib(x) = FillSpace(Left$(!CODVIGSANIT & "-" & !SUBCODVIGSANIT, 15), 15)
                            aCodTrib(x) = FillSpace(Left$(!CNAE, 15), 15)
                            If !DESC2 = "não especificado" Then
                                aDescTrib(x) = FillSpace(Left$(!DESCRICAO, 50), 50)
                            Else
                                aDescTrib(x) = FillSpace(Left$(!DESCRICAO & " - " & !DESC2, 50), 50)
                            End If
'                            If !SUBCODVIGSANIT = 0 Then
'                                aDescTrib(x) = FillSpace(Left$(!DESCVIGSANITARIA, 50), 50)
'                            Else
'                                Sql = "SELECT DESCVIGSANITARIA FROM VIGSANITARIA WHERE CODVIGSANIT=" & !CODVIGSANIT
'                                Sql = Sql & " AND SUBCODVIGSANIT=0"
'                                Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                                aDescTrib(x) = FillSpace(Left$(RdoAux4!DESCVIGSANITARIA & " - " & !DESCVIGSANITARIA, 50), 50)
'                            End If
                            x = x + 1
                           .MoveNext
                        Loop
                       .Close
                    End With
                    For s = x To 10
                        aCodTrib(s) = FillSpace(" ", 15)
                        aDescTrib(s) = FillSpace(" ", 50)
                    Next
                    sCodTrib = ""
                    sDescTrib = ""
                    For x = 0 To 10
                        sCodTrib = sCodTrib & aCodTrib(x)
                        sDescTrib = sDescTrib & aDescTrib(x)
                    Next
               Else
                    Sql = "SELECT * FROM MOBILIARIOATIVIDADEVS2 WHERE CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
'                    Sql = "SELECT MOBILIARIOATIVIDADEVS.*,VIGSANITARIA.VALORALIQ,DESCVIGSANITARIA FROM MOBILIARIOATIVIDADEVS INNER JOIN "
'                    Sql = Sql & " VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT AND "
'                    Sql = Sql & " MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT WHERE MOBILIARIOATIVIDADEVS.CODMOBILIARIO=" & Val(RdoAux!CODREDUZIDO)
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        x = 1
                        Do Until .EOF
                            'aValorTributoUnica(x) = FillLeft(FormatNumber(!VALORALIQ * nUfir, 2), 17)
                            'aValorTributoUnica(x) = FillLeft(FormatNumber(!valor * nUfir, 2), 17)
                            aValorTributoUnica(x) = FillLeft(FormatNumber(!valor * !QTDE, 2), 17)
                            x = x + 1
                           .MoveNext
                        Loop
                       .Close
                    End With
               End If

              
              .MoveNext
            Loop
           .Close
        End With
        aDescParc(0) = "UNICA"
        For x = 1 To Val(sQtdeParc)
            aDescParc(x) = Format(x, "00") & "/04"
        Next
        For x = nQtdeParcS + 1 To 12
            aDescParc(x) = "00/00"
        Next
        sDescParc = ""
        For x = 0 To 12
            sDescParc = sDescParc & aDescParc(x)
        Next
        
        sMesAno = ""
'        aMesAno(0) = FillLeft(" ", 6)
        If aMesAno(0) = "" Then GoTo Proximo
        For x = nQtdeParcS + 1 To 12
            If aMesAno(x) = "" Then
               aMesAno(x) = "      "
            End If
        Next
        For x = 0 To 12
            sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
        Next
        
        For x = nQtdeParcS + 1 To 12
            aVencParc(x) = "00/00/0000"
        Next
        sVencParc = ""
        For x = 0 To 12
            sVencParc = sVencParc & aVencParc(x) '662-791
        Next
                
        For x = Val(sQtdeParc) + 1 To 12
            aNumDoc(x) = "000000000"
        Next
        sNumDoc = ""
'        aNumDoc(0) = "0"
        For x = 0 To 12
            sNumDoc = sNumDoc & Format(aNumDoc(x) & Modulo11(aNumDoc(x)), "000000000")  '662-791
        Next
                
        sValorTributoUnica = ""
        sValorTributo = ""
        For x = 1 To 10
            If aValorTributo(x) = "" Then
               aValorTributo(x) = FillLeft("0,00", 17)
            End If
        Next
        For x = 1 To 10
            If aValorTributoUnica(x) = "" Then
               aValorTributoUnica(x) = FillLeft("0,00", 17)
            End If
        Next
        
        For x = 1 To 10
            sValorTributo = sValorTributo & FillLeft(FormatNumber(aValorTributo(x), 2), 17)
            sValorTributoUnica = sValorTributoUnica & FillLeft(aValorTributoUnica(x), 17)
        Next
        
        nTotalTrib = 0
        nTotalTribUnica = 0
        For x = 0 To 10
            If x > 0 Then
                If aValorTributo(x) = "" Then Exit For
                nTotalTrib = nTotalTrib + CDbl(aValorTributo(x))
            End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
        Next
        sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
        For x = 0 To 10
            If x > 0 Then
                If aValorTributoUnica(x) = "" Then Exit For
                nTotalTribUnica = nTotalTribUnica + CDbl(aValorTributoUnica(x))
            End If
'                sValorTributoUnica = sValorTributoUnica & aValorTributoUnica(x) '449-598
'                sValorTributo = sValorTributo & aValorTributo(x) '449-598
        Next
        sTotalTribUnica = FillLeft(CStr(FormatNumber(nTotalTribUnica, 2)), 17)
FIMVS:


        nCodLanc = 13
        aValorParc(0) = FormatNumber(nTotalTribUnica - (nTotalTribUnica * 0.05) + nExpUnica, 2)
        'aValorParc(0) = FormatNumber(nTotalTribUnica - (nTotalTribUnica * 0.05) + nExpUnica, 2)
        aValorParcSEXP(0) = FormatNumber(nTotalTribUnica - (nTotalTribUnica * 0.05), 2)
        For x = 1 To nQtdeParcS
            aValorParc(x) = FormatNumber(((nTotalTrib) / 4) + nExpParc, 2)
            aValorParcSEXP(x) = FormatNumber(((nTotalTrib) / 4), 2)
        Next
        For x = nQtdeParcS + 1 To 12
            aValorParc(x) = "0,00"
            aValorParcSEXP(x) = "0,00"
        Next
        sValorParc = ""
        sValorParcSEXP = ""
        For x = 0 To 12
            sValorParc = sValorParc & FillLeft(aValorParc(x), 17) '662-791
            sValorParcSEXP = sValorParcSEXP & FillLeft(aValorParcSEXP(x), 17) '662-791
        Next
        
        sValorEXP = FillLeft(FormatNumber(nExpParc, 2), 17)
        
        sDataProc = Format(Now, "dd/mm/yyyy") '439-448
        sDataDoc = FillLeft(Format(Now, "dd/mm/yyyy"), 15) '561-570
        

'        dDataBase = "07/10/1997"
'        For x = 0 To nQtdeParcS
'            If Not IsDate(aVencParc(x)) Then Exit For
'            If x = 0 Then
'                'aCodBarra(x) = "8160" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'                aCodBarra(x) = "8160" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(aNumDoc(x)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'            Else
'                'aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'                aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(aNumDoc(x)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'            End If
'        Next
'        For x = 5 To 12
'            aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
'        Next
        
'        sCodBarra = ""
'        For x = 0 To 12
'            sCodBarra = sCodBarra & aCodBarra(x) '792-1012
'        Next

        dDataBase = "07/10/1997"
        aCodBarra(0) = Format("0", "00000000000000000000000000000000000000000000")
        For x = 0 To nQtdeParcS
            If Not IsDate(aVencParc(x)) Then Exit For
            nFatorVencto = CDate(aVencParc(x)) - dDataBase
            'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "023" & Format(aNumDoc(x), "00000000") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "0000000") & "0003300"
            'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & sAgencia & Format(Val(aNumDoc(x)), "00000000") & "00"
            'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & sAgencia & Format(Val(aNumDoc(x)), "0000000") & "0003300"
            aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "9" & "1235028" & Format(aNumDoc(x) & Modulo11(aNumDoc(x)), "0000000000000") & "0102"
        Next
        For x = 5 To 12
            aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
        Next
'        sCodTrib = FillSpace(" ", 150) '?????????????????
        sCodBarra = ""
        For x = 0 To 12
             sCodBarra = sCodBarra & aCodBarra(x) '792-1012
        Next


        sTotalTrib = FillLeft("0,00", 17) '??????????????
        ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
        'ax = sExercicio & sContribuinte & sFantasia & sEnd & sCompl & sBairro & sCidEntrega & sCepEntrega
        ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sDescAtiv
        ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataProc & sDataProc & sNumDoc & sCodTrib & sValorTributoUnica & sValorParcSEXP
        ax = ax & sVencParc & sValorParc & sCodBarra & sDescTrib & sValorEXP & FillLeft(FormatNumber("6,76", 2), 17)
        tTipo = "4"
        Print #1, ax

        tDado = ax
        ax = tDado & "," & tEnd & "," & tNum & "," & tTipo
        
        Sql = "INSERT LASERTMP (DADO,CIDADE,BAIRRO,ENDERECO,NUMERO,TIPO) VALUES('" & Mask(tDado) & "','"
        Sql = Sql & Trim(Mask(tBairro)) & "','" & Trim(Mask(tCidade)) & "','" & Trim(Mask(tEnd)) & "','" & tNum & "','" & tTipo & "')"
        cn.Execute Sql, rdExecDirect
        
Proximo:
        xId = xId + 1
       .MoveNext
     Loop
   .Close
End With

Close #1
'Exit Sub
ORDENA:

Open sPathBin & "\LASERVIGSANIT.TXT" For Output As #1
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=4 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #1, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #1

MsgBox "fim"

End Sub

Private Sub GeraISSFixo()
Dim xId As Long, nNumRec As Long, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset, x As Integer
Dim nCodLogr As Long, nNum As Integer, nQtdeParcF As Integer, t As Integer, nCodEsc As Integer
Dim nExpParc As Double, nExpUnica As Double, nValorUnicaInteiraISS As Double, nValorUnicaInteiraTLL As Double

'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sFantasia As String, sEnd As String, sCompl As String, sBairro As String, sCEP As String
Dim sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim sTipoImposto As String, sInscricao As String, sQtdeParc As String, sCodAtiv As String, sDescAtiv As String, sCodInscricao As String
Dim bISSFixo As Boolean, bTxLic As Boolean
Dim aDescParc(0 To 12) As String, sDescParc As String
Dim sCodTrib As String
Dim aDescTrib(0 To 10) As String, sDescTrib As String
Dim aVencParc(0 To 12) As String, sVencParc As String
Dim aValorTributoUnica(0 To 12) As UNICA, sValorTributoUnica As String
Dim aValorTributo(0 To 12) As String, sValorTributo As String
Dim aValorParc(0 To 12) As String, sValorParc As String
Dim aValorParcSEXP(0 To 12) As String, sValorParcSEXP As String
Dim aMesAno(0 To 12) As String, sMesAno As String, sMes As String
Dim aNumDoc(0 To 12) As String, sNumDoc As String
Dim nTotalTrib As Double, sTotalTrib As String
Dim nTotalTribUnica As Double, sTotalTribUnica As String
Dim aCodBarra(0 To 12) As String, sCodBarra As String
Dim dDataBase As Date, nUfir As Double
Dim tDado As String, tEnd As String, tNum As Integer, tTipo As Integer
Dim tCidade As String, tBairro As String, nDesc5Perc As Double
Dim sAgencia As String, nValorTribUnica As Double
Dim sValorEXP As String

'GoTo ORDENA
'********************************
' PARAMETROS DAS PARCELAS
'********************************
'PARCELAS PARA ISS FIXO E TLL
Sql = "SELECT QTDEPARCELA FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & 2014 & " AND CODTIPO=2"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nQtdeParcF = !qtdeparcela
   .Close
End With

nUfir = RetornaUFIR(2014)
sAgencia = "02345000024"

Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & 2014
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nExpParc = FormatNumber(RdoAux!VALORPARCELA, 2)
nExpUnica = FormatNumber(RdoAux!ValorUnica, 2)

'## ************************************************ ##
'## ***************** G U I A ********************** ##
'## ************************************************ ##

Sql = "TRUNCATE TABLE LASERTMP"
cn.Execute Sql, rdExecDirect

Open sPathBin & "\LASERISSFIXOTL.TXT" For Output As #1

Sql = "SELECT DISTINCT CODREDUZIDO From DEBITOPARCELA "
'Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 14 OR CODLANCAMENTO=6 ) AND CODREDUZIDO in (SELECT CODREDUZIDO FROM ISSFIXOTMP) "
Sql = Sql & "WHERE (ANOEXERCICIO = 2013) AND (CODLANCAMENTO = 14 OR CODLANCAMENTO=6 )"
'Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 2 OR CODLANCAMENTO=6 ) AND (DATAVENCIMENTO > '01/10/2014') "
'Sql = Sql & "AND (CODREDUZIDO=103313) "
'Sql = Sql & " and codreduzido =100022"
Sql = Sql & " ORDER BY CODREDUZIDO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
'        If !CODREDUZIDO = 103313 Then MsgBox "TESTE"
        If xId Mod 50 = 0 Then
            DoEvents
           CallPb xId, nNumRec
        End If
        
        nValorUnicaInteira = 0
        Sql = "SELECT mobiliario.codigomob, mobiliario.dvmob, mobiliario.razaosocial, mobiliario.nomefantasia, mobiliario.numero, mobiliario.codlogradouro,"
        Sql = Sql & "mobiliario.complemento, bairro.descbairro, cidade.desccidade, mobiliario.codatividade, mobiliario.ativextenso, mobiliario.nomelogradouro,MOBILIARIO.RESPCONTABIL,"
        Sql = Sql & "mobiliario.cep , Cidade.SiglaUF FROM mobiliario LEFT OUTER JOIN cidade ON mobiliario.siglauf = cidade.siglauf AND mobiliario.codcidade = cidade.codcidade LEFT OUTER JOIN "
        Sql = Sql & "bairro ON mobiliario.siglauf = bairro.siglauf AND mobiliario.codcidade = bairro.codcidade AND mobiliario.codbairro = bairro.codbairro  "
        Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then GoTo Proximo
            nCodEsc = Val(SubNull(!RESPCONTABIL))
            nCodLogr = !CodLogradouro
            sExercicio = "2014" '1-4
            sCodInscricao = Format(!codigomob, "00000000000000")
            sContribuinte = FillSpace(!RazaoSocial, 40) '5-44
            sFantasia = FillSpace(SubNull(!NOMEFANTASIA), 40) '45-84
            sCodAtiv = Format(!codatividade, "00000000000000")
            sDescAtiv = FillSpace(Left$(!ATIVEXTENSO, 50), 50)
            Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                
                If .RowCount > 0 Then
                    sEnd = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero), 46) '212-257
                    nNum = RdoAux2!Numero
                    sBairro = FillSpace(SubNull(RdoAux2!DescBairro), 30)
                Else
                    If RdoAux2!desccidade <> "JABOTICABAL" Then
                        If Not IsNull(RdoAux2!NomeLogradouro) Then
                            sEndEntrega = FillSpace(RdoAux2!NomeLogradouro & " Nº " & CStr(RdoAux2!Numero), 46) '212-257
                            sEnd = sEndEntrega
                            sNumEntrega = RdoAux2!Numero
                            sCidEntrega = FillSpace(RdoAux2!desccidade, 20)
                            sBairro = FillSpace(SubNull(RdoAux2!DescBairro), 30)
                            sUFEntrega = RdoAux2!SiglaUF
                            If Trim(RdoAux2!Cep) <> "" Then
                                sCepEntrega = Format(RdoAux2!Cep, "00000-000")
                            Else
                                sCepEntrega = "00000-000"
                            End If
                            GoTo fimend
                        Else
                            sEndEntrega = ""
                            sNumEntrega = ""
                            sBairro = ""
                            nNum = 0
                            sCidEntrega = ""
                            sCepEntrega = ""
                        End If
                    Else
                        sEndEntrega = ""
                        sBairro = ""
                        sNumEntrega = ""
                        nNum = 0
                        sCidEntrega = ""
                        sCepEntrega = ""
                    End If
                    
                End If
               .Close
            End With
            sCEP = RetornaCEP(nCodLogr, nNum)
            sCompl = FillSpace(SubNull(Left(!Complemento, 20)), 20) '258-277
            sBairro = FillSpace(SubNull(!DescBairro), 30) '278-307
'            GoTo fimend
            Sql = "SELECT MOBILIARIOENDENTREGA.CODMOBILIARIO, MOBILIARIOENDENTREGA.CODLOGRADOURO, MOBILIARIOENDENTREGA.NOMELOGRADOURO, "
            Sql = Sql & "MOBILIARIOENDENTREGA.NUMIMOVEL,MOBILIARIOENDENTREGA.COMPLEMENTO, MOBILIARIOENDENTREGA.UF,MOBILIARIOENDENTREGA.CODCIDADE,"
            Sql = Sql & "MOBILIARIOENDENTREGA.CODBAIRRO, MOBILIARIOENDENTREGA.CEP,MOBILIARIOENDENTREGA.DESCBAIRRO, MOBILIARIOENDENTREGA.DESCCIDADE, BAIRRO.DESCBAIRRO AS DESCBAIRRO2,"
            Sql = Sql & "CIDADE.DESCCIDADE AS DESCCIDADE2 FROM BAIRRO INNER JOIN MOBILIARIOENDENTREGA ON BAIRRO.SIGLAUF = MOBILIARIOENDENTREGA.UF AND BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO INNER JOIN "
            Sql = Sql & "CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE Where CODMOBILIARIO = " & !codigomob
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    If !CodLogradouro = 0 Then
                        sEndEntrega = FillSpace(!NomeLogradouro & " Nº " & CStr(!numimovel), 46) '85-130
                        sNumEntrega = SubNull(!numimovel)
                        If !CodBairro = 0 Then
                            sBairroEntrega = FillSpace(SubNull(!DescBairro), 30) '151-180
                        ElseIf !CodBairro = 999 Then
                            sBairroEntrega = FillSpace(" ", 30) '151-180
                        Else
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30)
                        End If
                        If !CodCidade = 0 Then
                            sCidEntrega = FillSpace(SubNull(!desccidade), 20) '181-200
                        Else
                            sCidEntrega = FillSpace(SubNull(!desccidade2), 20)
                        End If
                        If !CodCidade = 413 Then
                            sCepEntrega = Format(!Cep, "00000-000") '201-209
                        Else
                            sCepEntrega = RetornaCEP(!CodLogradouro, Val(sNumEntrega))
                        End If
                        sComplEntrega = FillSpace(SubNull(!Complemento), 20)
                        sUFEntrega = !UF '210-211
                    Else
                        Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
                        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux4
                            If .RowCount > 0 Then
                                sEndEntrega = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux3!numimovel), 46) '85-130
                                'sEndEntrega = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero), 46) '85-130
                                sNumEntrega = SubNull(RdoAux3!numimovel)
                                sCepEntrega = RetornaCEP(RdoAux3!CodLogradouro, Val(sNumEntrega))
                            End If
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30) '151-180
                            sCidEntrega = FillSpace(SubNull(RdoAux3!desccidade2), 20) '181-200
                            sComplEntrega = FillSpace(SubNull(RdoAux2!Complemento), 20)
                            sUFEntrega = RdoAux3!UF '210-211
                           .Close
                        End With
                    End If
                Else
                    nCodLogr = 0
                    sEndEntrega = sEnd '85-130
                    sBairroEntrega = FillSpace(sBairro, 30) '151-180
                    sCidEntrega = FillSpace("JABOTICABAL", 20) '181-200
                    sCepEntrega = "14870-000" '201-209
                    sComplEntrega = FillSpace(" ", 20)
                    sUFEntrega = "SP" '210-211
                End If
               .Close
            End With
           .Close
        End With
        
        If nCodEsc > 0 Then
            '***ENDERECO CONTADOR***
            For t = 1 To UBound(aEnd)
                With aEnd(t)
                    If aEnd(t).nCodigo = nCodEsc Then
                        If aEnd(t).bRecebe Then
                            sEndEntrega = FillSpace(.sLogradouro & " Nº " & CStr(.nNumero), 46) '85-130
                            sNumEntrega = .nNumero
                            sBairroEntrega = FillSpace(.sBairro, 30)
                            sCidEntrega = FillSpace(.sCidade, 20)
                            sUFEntrega = .sUF
                            sComplEntrega = FillSpace(" ", 20)
                            sCepEntrega = .sCEP
                        End If
                        Exit For
                    End If
                End With
            Next
            '***********************
        End If
        
fimend:
        tEnd = sEnd
        tNum = nNum
        tBairro = sBairro
        tCidade = sCidEntrega
        tEnd = sEndEntrega
        If Left(sCEP, 1) = "_" Then sCEP = "         "
        If Trim(sCEP) = "" Then sCEP = "00000-000"

       
        'INSCRICAO
        sInscricao = Format(Val(sCodInscricao), "000000000000000")
        
        'TIPOS DE CARNE A GERAR
        Sql = "SELECT DISTINCT CODLANCAMENTO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=2014"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            bISSFixo = False
            bTxLic = False
            Do Until .EOF
                Select Case !CodLancamento
                Case 14
                    bISSFixo = True
                Case 6
                    bTxLic = True
                End Select
              .MoveNext
            Loop
           .Close
        End With
        If bTxLic Then 'TAXA DE LICENÇA
            nCodLanc = 6
            Sql = "SELECT COUNT(CODREDUZIDO) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=2014 "
            Sql = Sql & " AND CODLANCAMENTO=14"
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If !CONTADOR = 0 Then
                    sTipoImposto = FillSpace("TAXA DE LICENÇA", 20)
                Else
                    sTipoImposto = FillSpace("ISS FIXO/TLL", 20)
                End If
               .Close
            End With
        Else
            nCodLanc = 14
            sTipoImposto = FillSpace("ISS FIXO/TLL", 20)
        End If
        
        sQtdeParc = Format(nQtdeParcF, "0000000000") '427-428
        Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
        Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
        Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
        Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
        Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
        Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = " & nCodLanc
        'Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = " & nCodLanc & "  AND (DATAVENCIMENTO > '01/10/2014')"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nSomaTrib = 0
            For x = 0 To 12
                aMesAno(x) = ""
            Next
            For x = 1 To 10
                aValorTributo(x) = ""
                aValorTributoUnica(x).nValorTributo = ""
            Next
            If .RowCount < 4 Then
                GoTo Proximo
            End If
            Do Until .EOF
               Select Case !NumParcela
                    Case 1
                        sMes = " 01"
                    Case 2
                        sMes = " 02"
                    Case 3
                        sMes = " 03"
                    Case 4
                        sMes = " 04"
                    Case 5
                        sMes = " 05"
                    Case 6
                        sMes = " 06"
                    Case 7
                        sMes = " 07"
                    Case 8
                        sMes = " 08"
                    Case 9
                        sMes = " 09"
                    Case 10
                        sMes = " 10"
                    Case 11
                        sMes = " 11"
                    Case 12
                        sMes = " 12"
               End Select
               
               If !NumParcela = 0 Then
                   aMesAno(!NumParcela) = "UNICA"
               Else
                   aMesAno(!NumParcela) = sMes & "/03"
               End If
               aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9)
               If !NumParcela = 0 Then
                  aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
               Else
                  aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
               End If
               If !NumParcela = 1 Then
                    Sql = "SELECT CODLANCAMENTO,DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                    Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                    Sql = Sql & " ORDER BY DESCTRIBUTO"
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        x = 1
                        Do Until .EOF
'                            If !CodLancamento <> 14 Then
                              aDescTrib(x) = FillSpace(!ABREVTRIBUTO, 15)
                              aValorTributo(x) = FillLeft(FormatNumber(!ValorTributo, 2), 17)
                              If !CodTributo = 14 Or !CodTributo = 11 Then
                    '              aValorTributoUnica(x) = FillLeft(FormatNumber(!VALORTRIBUTO * 3, 2), 17)
                    '              aValorTributoUnica(x) = FormatNumber(aValorTributoUnica(x) - (aValorTributoUnica(x) * 5 / 100), 2)
                                  'nDesc5Perc = FormatNumber(aValorTributoUnica(x) * 0.05, 2)
                              End If
                              x = x + 1
 '                           End If
                           .MoveNext
                        Loop
                       .Close
                    End With
                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                    Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=14 AND "
                    Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                    Sql = Sql & " ORDER BY DESCTRIBUTO"
                    Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    With RdoAux4

                       If .RowCount > 0 And bTxLic Then
                           aDescTrib(x) = FillSpace("ISS FIXO", 15)
                           aValorTributo(x) = FillLeft(FormatNumber(!ValorTributo, 2), 17)
                           x = x + 1
                       End If
                      .Close
                   End With
                    
                    For s = x To 10
                        aDescTrib(s) = FillSpace(" ", 15)
                    Next
                    sDescTrib = ""
                    For x = 0 To 10
                        sDescTrib = sDescTrib & aDescTrib(x) '449-598
                    Next
               ElseIf !NumParcela = 0 Then
                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                    Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                    Sql = Sql & " ORDER BY DESCTRIBUTO"
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        For x = 1 To 10
                            aValorTributoUnica(x).nCodTributo = 0
                            aValorTributoUnica(x).nValorTributo = 0
                        Next
                         x = 0
                        Do Until .EOF
                            If !CodTributo <> 11 Then
                                x = x + 1
 '                               aValorTributoUnica(x) = FillLeft(FormatNumber(!VALORTRIBUTO + (!VALORTRIBUTO * 0.05), 2), 17)
 '                           Else
                                If !CodTributo = 14 Then
                                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                                    Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                                    Sql = Sql & "ANOEXERCICIO=" & RdoAux2!AnoExercicio & " AND CODLANCAMENTO=" & RdoAux2!CodLancamento & " AND "
                                    Sql = Sql & "SEQLANCAMENTO=" & RdoAux2!SeqLancamento & " AND NUMPARCELA=" & 0 & " AND CODCOMPLEMENTO=" & RdoAux2!CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO=14"
                                    Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                    nValorTribUnica = RdoAux4!ValorTributo * 3
                                    nValorUnicaInteiraTLL = FormatNumber(RdoAux4!ValorTributo, 2)
                                    'nValorUnicaInteiraTLL = FormatNumber(nValorTribUnica - (nValorTribUnica * 5 / 100), 2)
                                    'nValorUnicaInteiraTLL = FormatNumber(RdoAux4!valortributo * 3, 2)
                                End If
                                aValorTributoUnica(x).nCodTributo = !CodTributo
                                aValorTributoUnica(x).nValorTributo = FillLeft(FormatNumber(!ValorTributo, 2), 17)
                            End If
                            
                           .MoveNext
                        Loop
                       .Close
                    End With
                    Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=14   AND NUMPARCELA=1 AND CODTRIBUTO<>3"
                    Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    With RdoAux4
                        If .RowCount > 0 Then
                            x = x + 1
                            aValorTributoUnica(x).nCodTributo = !CodTributo
                            aValorTributoUnica(x).nValorTributo = FillLeft(FormatNumber(!ValorTributo * 3, 2), 17)
                            nValorTribUnica = !ValorTributo * 3
'If !CodTributo = 24 Then MsgBox "teste"
                            If !CodTributo = 14 Then
'                                nValorUnicaInteiraTLL = FormatNumber(!VALORTRIBUTO * 3, 2)
                                nValorUnicaInteiraTLL = nValorTribUnica - (nValorTribUnica * 5 / 100)
                            ElseIf !CodTributo = 11 Then
                                'nValorUnicaInteiraISS = FormatNumber(!VALORTRIBUTO * 3, 2)
                                nValorUnicaInteiraISS = nValorTribUnica - (nValorTribUnica * 5 / 100)
                            Else
                                aValorTributoUnica(x).nValorTributo = FormatNumber(aValorTributoUnica(x).nValorTributo - (aValorTributoUnica(x).nValorTributo * 5 / 100), 2)
                            End If
                        End If
                       .Close
                    End With
               End If
              .MoveNext
            Loop
           .Close
        End With
        aDescParc(0) = "UNICA"
        For x = 1 To Val(sQtdeParc)
            aDescParc(x) = Format(x, "00") & "/03"
        Next
        For x = nQtdeParcF + 1 To 12
            aDescParc(x) = "00/00"
        Next
        sDescParc = ""
        For x = 0 To 12
            sDescParc = sDescParc & aDescParc(x)
        Next
        
        sMesAno = ""
'        aMesAno(0) = FillLeft(" ", 6)
        If aMesAno(0) = "" Then GoTo Proximo
        For x = nQtdeParcF + 1 To 12
            If aMesAno(x) = "" Then
               aMesAno(x) = "      "
            End If
        Next
        For x = 0 To 12
            sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
        Next
        
        For x = nQtdeParcF + 1 To 12
            aVencParc(x) = "00/00/0000"
        Next
        sVencParc = ""
        For x = 0 To 12
            sVencParc = sVencParc & aVencParc(x) '662-791
        Next
                
        For x = Val(sQtdeParc) + 1 To 12
            aNumDoc(x) = "000000000"
        Next
        sNumDoc = ""
        'aNumDoc(0) = "0"
        For x = 0 To 12
            sNumDoc = sNumDoc & Format(aNumDoc(x) & Modulo11(aNumDoc(x)), "000000000") '662-791
        Next
                
        sValorTributoUnica = ""
        sValorTributo = ""
        For x = 1 To 10
            If aValorTributo(x) = "" Then
               aValorTributo(x) = FillLeft("0,00", 17)
            End If
        Next
        For x = 1 To 10
            If aValorTributoUnica(x).nValorTributo = "" Then
               aValorTributoUnica(x).nValorTributo = FillLeft("0,00", 17)
            End If
        Next
        
        For x = 1 To 10
            sValorTributo = sValorTributo & FillLeft(FormatNumber(aValorTributo(x), 2), 17)
            'If aValorTributoUnica(x).nCodTributo = 11 Or aValorTributoUnica(x).nCodTributo = 14 Then
            If aValorTributoUnica(x).nCodTributo = 11 Then
                'sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(nValorUnicaInteiraISS, 2), 17)
                sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(aValorTributo(x) * 3, 2), 17)
            ElseIf aValorTributoUnica(x).nCodTributo = 14 Then
                'sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(nValorUnicaInteiraTLL, 2), 17)
                sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(aValorTributo(x) * 3, 2), 17)
            Else
                'sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(aValorTributoUnica(x).nValorTributo, 2), 17)
                If aValorTributoUnica(x).nCodTributo > 0 Then
                    If aValorTributoUnica(x).nCodTributo = 24 Then
                        sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(25, 2), 17)
                    Else
                        sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(aValorTributo(x) * 3, 2), 17)
                    End If
                Else
                    sValorTributoUnica = sValorTributoUnica & FillLeft(FormatNumber(0, 2), 17)
                End If
            End If
        Next
        
        nTotalTrib = 0
        nTotalTribUnica = 0
        For x = 0 To 10
            If x > 0 Then
'                If aValorTributo(x) = "" Then Exit For
                nTotalTrib = nTotalTrib + CDbl(aValorTributo(x)) '//TROCAR DEPOIS
'                If CDbl(aValorTributo(x)) > 1.52 Then
'                    nTotalTrib = CDbl(aValorTributo(x))
'                End If
            End If
        Next
        sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
        For x = 0 To 10
            If x > 0 Then
                If aValorTributoUnica(x).nValorTributo = "" Then Exit For
                If aValorTributoUnica(x).nCodTributo = 11 Then
                    nTotalTribUnica = nTotalTribUnica + CDbl(FormatNumber(nValorUnicaInteiraISS, 2))
                ElseIf aValorTributoUnica(x).nCodTributo = 14 Then
                    nTotalTribUnica = nTotalTribUnica + CDbl(FormatNumber(nValorUnicaInteiraTLL, 2))
                Else
                    nTotalTribUnica = nTotalTribUnica + CDbl(FormatNumber(aValorTributoUnica(x).nValorTributo, 2))
                End If
                'nTotalTribUnica = nTotalTribUnica + CDbl(FormatNumber(aValorTributoUnica(x) - (aValorTributoUnica(x) * 5 / 100), 2))
            End If
        Next
        'nTotalTribUnica = nTotalTribUnica - nDesc5Perc
        nTotalTribUnica = nTotalTribUnica
        sTotalTribUnica = FillLeft(CStr(FormatNumber(nTotalTribUnica, 2)), 17)
FIMVS:


        'nCodLanc = 13
        aValorParc(0) = FormatNumber(nTotalTribUnica + nExpParc, 2)
'        aValorParc(0) = FormatNumber(nTotalTrib - (nTotalTribUnica * 0.05) + nExpUnica, 2)
        aValorParcSEXP(0) = FormatNumber(nTotalTribUnica, 2)
        For x = 1 To nQtdeParcF
            aValorParc(x) = FormatNumber(((nTotalTrib + nExpParc)), 2) '//VER
            aValorParcSEXP(x) = FormatNumber(((nTotalTrib)), 2)
        Next
        For x = nQtdeParcF + 1 To 12
            aValorParc(x) = "0,00"
            aValorParcSEXP(x) = "0,00"
        Next
        sValorParc = ""
        sValorParcSEXP = ""
        For x = 0 To 12
            sValorParc = sValorParc & FillLeft(aValorParc(x), 17) '662-791
            sValorParcSEXP = sValorParcSEXP & FillLeft(aValorParcSEXP(x), 17) '662-791
        Next
        
        sValorEXP = FillLeft(FormatNumber(nExpParc, 2), 17)
        
        sDataProc = Format(Now, "dd/mm/yyyy") '439-448
        sDataDoc = FillLeft(Format(Now, "dd/mm/yyyy"), 15) '561-570
'        sDataProc = "01/01/2014" '439-448
'        sDataDoc = FillLeft("01/01/2014", 15) '561-570
        

'        For x = 0 To nQtdeParcF
'            If Not IsDate(aVencParc(x)) Then Exit For
'            If x = 0 Then
'            '    aCodBarra(x) = "8160" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
 '               aCodBarra(x) = "8160" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(aNumDoc(x)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
 '           Else
'             '   aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'                aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(aNumDoc(x)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'            End If
'        Next
        
        dDataBase = "07/10/1997"
        aCodBarra(0) = Format("0", "00000000000000000000000000000000000000000000")
        For x = 0 To nQtdeParcF
            If Not IsDate(aVencParc(x)) Then Exit For
            nFatorVencto = CDate(aVencParc(x)) - dDataBase
'            aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & sAgencia & Format(Val(aNumDoc(x)), "0000000") & "0003300"
            aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "9" & "1235028" & Format(aNumDoc(x) & Modulo11(aNumDoc(x)), "0000000000000") & "0102"
        Next
        For x = 4 To 12
            aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
        Next
        sCodTrib = FillSpace(" ", 150) '?????
        sCodBarra = ""
        For x = 0 To 12
             sCodBarra = sCodBarra & aCodBarra(x) '792-1012
        Next
 
        sTotalTrib = FillLeft("0,00", 17) '????
        ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
'        ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCep & sTipoImposto & sInscricao & sDescAtiv
        'ax = sExercicio & sContribuinte & sFantasia & sEnd & sCompl & sBairro & sCidEntrega & sCep
        ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sDescAtiv
        
        ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataProc & sDataProc & sNumDoc & sCodTrib & sValorTributoUnica & sValorParcSEXP
        'ax = ax & sTotalTrib & "          " & sVencParc & sValorParc & sCodBarra & sDescTrib & sValorEXP & FillLeft(FormatNumber("3,66", 2), 17)
        ax = ax & sVencParc & sValorParc & sCodBarra & sDescTrib & sValorEXP & FillLeft(FormatNumber("5,07", 2), 17)
        tTipo = "3"
        Print #1, ax


        tDado = ax
        ax = tDado & "," & tEnd & "," & tNum & "," & tTipo
        
        Sql = "INSERT LASERTMP (DADO,CIDADE,BAIRRO,ENDERECO,NUMERO,TIPO) VALUES('" & Mask(tDado) & "','"
        Sql = Sql & Trim(Mask(tBairro)) & "','" & Trim(Mask(tCidade)) & "','" & Trim(Mask(tEnd)) & "','" & tNum & "','" & tTipo & "')"
        cn.Execute Sql, rdExecDirect
        
Proximo:
        xId = xId + 1
        DoEvents
       .MoveNext
 '     Exit Do
    Loop
   .Close
End With

Close #1
'Exit Sub
ORDENA:

Open sPathBin & "\LASERISSFIXOTL.TXT" For Output As #1
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=3 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #1, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #1

MsgBox "FIM"
End Sub

Private Sub GeraIPTU()
Dim xId As Long, nNumRec As Long, nTipoEnd As Integer, nFatorVencto As Long, nCodCid As Long
Dim nSomaTrib As Double, nSomaUnica As Double, nValorParc As Double
Dim aDescParc(0 To 12) As String, sDescParc As String
Dim aVencParc(0 To 12) As String, sVencParc As String
Dim aValorParc(0 To 12) As String, sValorParc As String
Dim aNumDoc(0 To 12) As String, sNumDoc As String
Dim aNosNum(0 To 12) As String, sNosNum As String
Dim aCodProc(0 To 12) As String, sCodProc As String
Dim aCodBarra(0 To 12) As String, sCodBarra As String
Dim dDataBase As Date, aDataVencto(0 To 12) As Date, bPredial As Boolean
Dim tDado As String, tEnd As String, tNum As Integer, tTipo As Integer
Dim tCidade As String, tBairro As String

'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sSacado As String, sEnd As String, sCompl As String, sBairro As String
Dim sQuadra As String, sCEP As String, sLote As String, sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String
Dim sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sCodContribuinte As String, sInscricao As String
Dim sCodMunicipio As String, sNatureza As String, sAreaTerreno As String, sAreaConstrucao As String, sTestadaPrincipal As String
Dim sVVTerreno As String, sVVConstrucao As String, sVVImovel As String, sValorIPU As String, sValorITU As String
Dim sTxExpUnica As String, sTxExpParc As String, sValorTotal As String, sValorUnica As String, sDataDoc As String
Dim sDataProc As String, sQtdeParcela As String, sAgencia As String
Dim RdoAux3 As rdoResultset

'GoTo ORDENA
Sql = "TRUNCATE TABLE LASERTMP"
cn.Execute Sql, rdExecDirect
cmdGerar.Enabled = False
Open sPathBin & "\LASERIPTU.TXT" For Output As #1

Sql = "SELECT CODREDUZIDO, VVT, VVC, VVI, IMPOSTOPREDIAL,IMPOSTOTERRITORIAL, NATUREZA, AREACONSTRUCAO,"
Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA From LASERIPTU WHERE ANO=2013 "
'Sql = Sql & "and CODREDUZIDO =34787"
Sql = Sql & " ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        If xId Mod 50 = 0 Then
           CallPb xId, nNumRec
        End If
        
        'If xId = 50 Then GoTo ORDENA
        If !vvc = 0 Then
            bPredial = False
        Else
            bPredial = True
        End If
        With xImovel
           .CarregaImovel RdoAux!CODREDUZIDO
            If .Inativo Then
                GoTo Proximo
            End If
            sExercicio = "2014" '1-4
            sContribuinte = FillSpace(.NomePropPrincipal, 40)  '5-44
            sSacado = FillSpace(.NomePropPrincipal, 40)  '45-84
            sEnd = FillSpace(Chomp(.EnderecoCompleto, chomp_left, 7), 46)   '85-130
            sCompl = FillSpace(.Li_Compl, 20) '131-150
            sBairro = FillSpace(.DescBairro, 30) '151-180
            sQuadra = FillSpace(Left$(.Li_Quadras, 6), 6) '181-186
            sCEP = RetornaCEP(.CodLogr, .Li_Num)
            If Len(sCEP) <> 9 Then
                sCEP = "00000-000"
            End If
            
            '####mudar lote para 11 posições #####
            sLote = FillSpace(Left$(.Li_Lotes, 12), 12) '193-204
           
            sEndEntrega = ""
            sComplEntrega = ""
            sBairroEntrega = ""
            sCidEntrega = ""
            sCepEntrega = ""
            sUFEntrega = ""
            nTipoEnd = .Ee_TipoEnd
           
            If nTipoEnd = 0 Then
                .RetornaEndereco RdoAux!CODREDUZIDO, Imobiliario, Localizacao
            ElseIf nTipoEnd = 1 Then
                .RetornaEndereco RdoAux!CODREDUZIDO, Imobiliario, cadastrocidadao
            Else
                .RetornaEndereco RdoAux!CODREDUZIDO, Imobiliario, Entrega
            End If
            
            sEndEntrega = FillSpace(.Endereco & " Nº " & CStr(.Numero), 46) '205-250
            tNum = .Numero
            tEnd = .Endereco
            sComplEntrega = FillSpace(.Complemento, 20) '251-270
            sBairroEntrega = FillSpace(Trim(.Bairro), 30) '271-300
            sCidEntrega = FillSpace(.Cidade, 20) '301-320
            sCepEntrega = Format(.Cep, "00000-000") '321-329
            sUFEntrega = SubNull(.UF) '330-331
            
            If UCase$(Trim(sCidEntrega)) <> "JABOTICABAL" Then
                tTipo = 2
            Else
                If sCepEntrega = "00000-000" Or sCepEntrega = "" Then
                    sCepEntrega = "14870-000"
                End If
                If sCEP = "00000-000" Or sCEP = "" Then
                    sCEP = "14870-000"
                End If
                If nTipoEnd = 0 And Not bPredial Then
                    tTipo = 1
                Else
                    tTipo = 0
                End If
            End If
                    
            
            
            sCodContribuinte = Format(.CodigoImovel, "000000")   '332-342
            'sCodContribuinte = Format(.CodigoImovel, "000000000") & "-X"  '332-342
            'sInscricao = Format(RemovePonto(Chomp(.Inscricao, chomp_righT, 7)), "00\.00\.0000\.00000\.00") '343-361
            sInscricao = .Inscricao
            sCodMunicipio = "391-8" '362-366
            sNatureza = FillSpace(RdoAux!Natureza, 11) '367-377
            If .Dt_FracaoIdeal = 0 Then
                sAreaTerreno = FillLeft(FormatNumber(.Dt_AreaTerreno, 2), 10)
            Else
                sAreaTerreno = FillLeft(FormatNumber(.Dt_FracaoIdeal, 2), 10)
            End If
            sAreaConstrucao = IIf(Val(RdoAux!areaconstrucao) = 0, Space(6) & "0,00", FillLeft(FormatNumber(RdoAux!areaconstrucao, 2), 10))
            sTestadaPrincipal = IIf(Val(RdoAux!TESTADAPRINC) = 0, Space(6) & "0,00", FillLeft(FormatNumber(RdoAux!TESTADAPRINC, 2), 10))
            sVVTerreno = IIf(Val(RdoAux!VVT) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!VVT, 2), 17))
            sVVConstrucao = IIf(Val(RdoAux!vvc) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!vvc, 2), 17))
            sVVImovel = IIf(Val(RdoAux!VVI) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!VVI, 2), 17))
            sValorIPU = IIf(Val(RdoAux!IMPOSTOPREDIAL) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!IMPOSTOPREDIAL, 2), 17))
            sValorITU = IIf(Val(RdoAux!IMPOSTOTERRITORIAL) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!IMPOSTOTERRITORIAL, 2), 17))
            sQtdeParcela = Format(RdoAux!qtdeparc, "00") '581-582

            If Val(sQtdeParcela) < 12 Then MsgBox "teste"
            sTxExpUnica = FillLeft(FormatNumber(1.69, 2), 17)
            sTxExpParc = FillLeft(FormatNumber(1.69, 2), 17)
        
            Sql = "SELECT DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,"
            Sql = Sql & "DEBITOPARCELA.DATAVENCIMENTO,DEBITOTRIBUTO.CODTRIBUTO,DEBITOPARCELA.DATADEBASE,"
            Sql = Sql & "DEBITOTRIBUTO.VALORTRIBUTO,PARCELADOCUMENTO.NUMDOCUMENTO,NumDocumento.DATADOCUMENTO FROM DEBITOTRIBUTO INNER JOIN "
            Sql = Sql & "PARCELADOCUMENTO ON DEBITOTRIBUTO.CODREDUZIDO = PARCELADOCUMENTO.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = PARCELADOCUMENTO.ANOEXERCICIO AND "
            Sql = Sql & "DEBITOTRIBUTO.CODLANCAMENTO = PARCELADOCUMENTO.CODLANCAMENTO AND DEBITOTRIBUTO.SEQLANCAMENTO = PARCELADOCUMENTO.SEQLANCAMENTO AND "
            Sql = Sql & "DEBITOTRIBUTO.CODCOMPLEMENTO = PARCELADOCUMENTO.CODCOMPLEMENTO AND DEBITOTRIBUTO.NUMPARCELA = PARCELADOCUMENTO.NUMPARCELA Inner Join "
            Sql = Sql & "NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join DEBITOPARCELA ON DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND "
            Sql = Sql & "DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
            Sql = Sql & "DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
            Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOTRIBUTO.ANOEXERCICIO = " & Val(sExercicio) & " AND DEBITOTRIBUTO.CODLANCAMENTO = 1 AND DEBITOTRIBUTO.CODTRIBUTO <> 3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSomaTrib = 0: nSomaUnica = 0: aNumDoc(0) = 999
                Do Until .EOF
                   aCodProc(!NumParcela) = Format(!CODREDUZIDO, "00000000") & ".00." & !AnoExercicio & "." & Format(!CodLancamento, "00") & "." & Format(!SeqLancamento, "00") & "." & Format(!NumParcela, "00") & "." & !CODCOMPLEMENTO & ".0"
                   If !NumParcela = 0 Then
                      nSomaUnica = !ValorTributo + CDbl(sTxExpParc)
                      aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
                      aNumDoc(0) = Format(!NumDocumento & RetornaDVNumDoc(!NumDocumento), "000000000000000")
                      aNosNum(0) = FillLeft(!NumDocumento, 13)
                      aDataVencto(0) = Format(!DataVencimento, "dd/mm/yyyy")
                   Else
                      nValorParc = !ValorTributo + CDbl(sTxExpParc)
                      sDataDoc = Format(!DATADOCUMENTO, "dd/mm/yyyy") '561-570
                      nSomaTrib = nSomaTrib + !ValorTributo + CDbl(sTxExpParc)
                      aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                      aNumDoc(!NumParcela) = Format(!NumDocumento & RetornaDVNumDoc(!NumDocumento), "000000000000000")
                      aNosNum(!NumParcela) = FillLeft(!NumDocumento, 13)
                      aDataVencto(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                   End If
                  .MoveNext
                Loop
               .Close
            End With
            
            'tirar
            '*******************
            If nSomaUnica = 0 Then
                nSomaUnica = nSomaTrib - ((nSomaTrib * 5) / 100)
            End If
            '*******************
            
            aDescParc(0) = "ÚNICA"
            For x = 1 To Val(sQtdeParcela)
                aDescParc(x) = Format(x, "00") & "/" & sQtdeParcela
            Next
            For x = Val(sQtdeParcela) + 1 To 12
                aDescParc(x) = "00/00"
            Next
            sDescParc = ""
            For x = 0 To 12
                sDescParc = sDescParc & aDescParc(x) '597-661
            Next
            
            For x = Val(sQtdeParcela) + 1 To 12
                aVencParc(x) = "00/00/0000"
            Next
            sVencParc = ""
            For x = 0 To 12
                sVencParc = sVencParc & aVencParc(x) '662-791
            Next
            
            sNosNum = ""
            On Error Resume Next
            For x = 0 To 12
                If x <= Val(sQtdeParcela) Then
                    sNosNum = sNosNum & FillLeft(Val(aNosNum(x)) & Modulo11(aNosNum(x)), 13)
                Else
                    sNosNum = sNosNum & "    000000000"
                End If
            Next

            For x = Val(sQtdeParcela) + 1 To 12
                aNumDoc(x) = "000000000"
            Next
            sNumDoc = ""
            For x = 0 To 12
                sNumDoc = sNumDoc & aNumDoc(x) '1013-1168
            Next
            For x = Val(sQtdeParcela) + 1 To 12
                aCodProc(x) = "00000000000000000000000000000"
            Next
'            sCodProc = ""
'            For x = 0 To 12
'                sCodProc = sCodProc & aCodProc(x) '1013-1168
'            Next

            sValorTotal = FillLeft(FormatNumber(nSomaTrib, 2), 17)  '527-543
            sValorUnica = FillLeft(FormatNumber(nSomaUnica, 2), 17)  '544-560
            sDataProc = Format(Now, "dd/mm/yyyy") '571-580
            sAgencia = "023 45 00002 4"
            
            aValorParc(0) = sValorUnica
            For x = 1 To Val(sQtdeParcela)
                aValorParc(x) = nValorParc
            Next
            For x = Val(sQtdeParcela) + 1 To 12
                aValorParc(x) = "0"
            Next
            sValorParc = ""
            For x = 0 To 12
                sValorParc = sValorParc & FillLeft(FormatNumber(aValorParc(x), 2), 17)
            Next
                        
'            dDataBase = "07/10/1997"
'            For x = 0 To Val(sQtdeParcela)
'                nFatorVencto = aDataVencto(x) - dDataBase
'                aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "9" & "1235028" & Format(aNosNum(x) & Modulo11(aNosNum(x)), "0000000000000") & "0102"
'            Next
'            For x = Val(sQtdeParcela) + 1 To 12
'                aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
'            Next
'            sCodBarra = ""
'            For x = 0 To 12
'                sCodBarra = sCodBarra & aCodBarra(x) '792-1012
'            Next
            sCompl = FillSpace(Replace(sCompl, vbNewLine, "", , , vbTextCompare), 20)
            sComplEntrega = FillSpace(Replace(sComplEntrega, vbNewLine, "", , , vbTextCompare), 20)
            
            ax = sExercicio & sContribuinte & sEnd & sCompl & sBairro & sQuadra & sCEP & sLote & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega & sUFEntrega
            'ax = sExercicio & sContribuinte & sSacado & sEnd & sCompl & sBairro & sQuadra & sCEP & sLote & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega & sUFEntrega
            ax = ax & sCodContribuinte & sInscricao & sCodMunicipio & sNatureza & sAreaTerreno & sAreaConstrucao & sTestadaPrincipal & sVVTerreno & sVVConstrucao & sVVImovel & sValorIPU
            ax = ax & sValorITU & sTxExpUnica & FillLeft(FormatNumber(sTxExpParc * Val(sQtdeParcela), 2), 17) & sValorTotal & sValorUnica & sDataProc & sQtdeParcela & sDescParc & sVencParc & sValorParc
            'ax = ax & sValorITU & sTxExpUnica & FillLeft(FormatNumber(sTxExpParc * Val(sQtdeParcela), 2), 17) & sValorTotal & sValorUnica & sDataProc & sDataProc & sQtdeParcela & sAgencia & sDescParc & sVencParc & sValorParc & sNumDoc
            ax = ax & sNosNum
            'ax = ax & sCodProc & sNosNum & sCodBarra

            tDado = ax
            
            ax = tDado & "," & tEnd & "," & tNum & "," & tTipo
            
            Print #1, ax
            
            Sql = "INSERT LASERTMP (DADO,BAIRRO,CIDADE,ENDERECO,NUMERO,TIPO) VALUES('" & Mask(tDado) & "','"
            Sql = Sql & Trim(Mask(tBairro)) & "','" & Trim(Mask(tCidade)) & "','" & Trim(Mask(tEnd)) & "','" & tNum & "','" & tTipo & "')"
            cn.Execute Sql, rdExecDirect
            
        End With
Proximo:
        xId = xId + 1
        DoEvents
       .MoveNext
    Loop
Sair:
   .Close
End With

Close #1
'Exit Sub
ORDENA:

Open sPathBin & "\IPTU_JABOTICABAL.TXT" For Output As #1
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=0 ORDER BY ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #1, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #1

Open sPathBin & "\IPTU_TERRENO.TXT" For Output As #2
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=1 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #2, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #2

Open sPathBin & "\IPTU_FORA.TXT" For Output As #2
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=2 ORDER BY BAIRRO,CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #2, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #2
cmdGerar.Enabled = True
MsgBox "FIM"
End Sub

Private Sub GeraIPTUParcial()
Dim xId As Long, nNumRec As Long, nTipoEnd As Integer, nFatorVencto As Long
Dim nSomaTrib As Double, nSomaUnica As Double, nValorParc As Double
Dim aDescParc(0 To 12) As String, sDescParc As String
Dim aVencParc(0 To 12) As String, sVencParc As String
Dim aValorParc(0 To 12) As String, sValorParc As String
Dim aNumDoc(0 To 12) As String, sNumDoc As String
Dim aNosNum(0 To 12) As String, sNosNum As String
Dim aCodProc(0 To 12) As String, sCodProc As String
Dim aCodBarra(0 To 12) As String, sCodBarra As String
Dim dDataBase As Date, aDataVencto(0 To 12) As Date, bPredial As Boolean
Dim tDado As String, tEnd As String, tNum As Integer, tTipo As Integer
Dim tCidade As String, tBairro As String

'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sSacado As String, sEnd As String, sCompl As String, sBairro As String
Dim sQuadra As String, sCEP As String, sLote As String, sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String
Dim sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sCodContribuinte As String, sInscricao As String
Dim sCodMunicipio As String, sNatureza As String, sAreaTerreno As String, sAreaConstrucao As String, sTestadaPrincipal As String
Dim sVVTerreno As String, sVVConstrucao As String, sVVImovel As String, sValorIPU As String, sValorITU As String
Dim sTxExpUnica As String, sTxExpParc As String, sValorTotal As String, sValorUnica As String, sDataDoc As String
Dim sDataProc As String, sQtdeParcela As String, sAgencia As String
Dim RdoAux3 As rdoResultset, RdoAux4 As rdoResultset, nValorParcIPTU As Double

'GoTo ORDENA
Sql = "TRUNCATE TABLE LASERTMP"
cn.Execute Sql, rdExecDirect

Open sPathBin & "\LASERIPTU.TXT" For Output As #1

Sql = "SELECT CODREDUZIDO, VVT, VVC, VVI, IMPOSTOPREDIAL,IMPOSTOTERRITORIAL, NATUREZA, AREACONSTRUCAO,"
Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA From LASERIPTU WHERE ANO=2014 "
Sql = Sql & "and CODREDUZIDO BETWEEN 30297 AND 30326"
Sql = Sql & " ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        If xId Mod 50 = 0 Then
           CallPb xId, nNumRec
        End If
        
        Sql = "SELECT VALOR FROM CARNE WHERE CODIGO=" & !CODREDUZIDO
        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nValorParcIPTU = RdoAux4!valor / 9
        RdoAux4.Close
        
        'If xId = 50 Then GoTo ORDENA
        If !vvc = 0 Then
            bPredial = False
        Else
            bPredial = True
        End If
        With xImovel
           .CarregaImovel RdoAux!CODREDUZIDO
            If .Inativo Then
                GoTo Proximo
            End If
            'If RdoAux!CODREDUZIDO = 9283 Then MsgBox "TESTE"
            sExercicio = "2014" '1-4
            sContribuinte = FillSpace(.NomePropPrincipal, 40)  '5-44
            sSacado = FillSpace(.NomePropPrincipal, 40)  '45-84
            sEnd = FillSpace(Chomp(.EnderecoCompleto, chomp_left, 7), 46)   '85-130
            sCompl = FillSpace(.Li_Compl, 20) '131-150
            sBairro = FillSpace(.DescBairro, 30) '151-180
'            sQuadra = FillSpace(.Quadra, 6) '181-186
            sQuadra = FillSpace(Left$(.Li_Quadras, 6), 6) '181-186
            sCEP = RetornaCEP(.CodLogr, .Li_Num)
            If Len(sCEP) <> 9 Then
                sCEP = "00000-000"
            End If
            'sCep = Format(RetornaNumero(.Li_Cep), "00000-000") '187-192 ?????
            'sLote = FillSpace(.Lote, 12) '193-204
            
            '####mudar lote para 11 posições #####
            sLote = FillSpace(Left$(.Li_Lotes, 12), 12) '193-204
           
            sEndEntrega = ""
            sComplEntrega = ""
            sBairroEntrega = ""
            sCidEntrega = ""
            sCepEntrega = ""
            sUFEntrega = ""
            nTipoEnd = .Ee_TipoEnd
           
            Sql = "SELECT ENDENTREGA.CODREDUZIDO,ENDENTREGA.EE_CODLOG, ENDENTREGA.EE_NOMELOG,ENDENTREGA.EE_NUMIMOVEL,"
            Sql = Sql & "ENDENTREGA.EE_COMPLEMENTO, ENDENTREGA.EE_UF,ENDENTREGA.EE_CIDADE, ENDENTREGA.EE_BAIRRO,"
            Sql = Sql & "ENDENTREGA.EE_CEP, ENDENTREGA.EE_LOTEAMENTO,ENDENTREGA.EE_DESCBAIRRO , Cidade.DESCCIDADE "
            Sql = Sql & "FROM ENDENTREGA INNER JOIN  CIDADE ON ENDENTREGA.EE_UF = CIDADE.SIGLAUF AND "
            Sql = Sql & "ENDENTREGA.Ee_Cidade = Cidade.CODCIDADE WHERE CODREDUZIDO = " & RdoAux!CODREDUZIDO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                If RdoAux2!Ee_NomeLog <> "" And nTipoEnd = 2 Then
                    sEndEntrega = FillSpace(RdoAux2!Ee_NomeLog & " Nº " & CStr(RdoAux2!Ee_NumImovel), 46) '205-250
                    tNum = RdoAux2!Ee_NumImovel
                    tEnd = RdoAux2!Ee_NomeLog
                    sComplEntrega = FillSpace(RdoAux2!Ee_Complemento, 20) '251-270
                    sBairroEntrega = FillSpace(Trim(SubNull(RdoAux2!Ee_DESCBairro)), 30) '271-300
                    sCidEntrega = FillSpace(SubNull(RdoAux2!desccidade), 20) '301-320
                    If Left$(RdoAux2!Ee_Cep, 1) <> " " Then
                        sCepEntrega = Format(Replace(RdoAux2!Ee_Cep, "-", ""), "00000-000") '321-329
                    Else
                        sCepEntrega = "00000-000"
                    End If
                    If Left$(sCepEntrega, 1) = "_" Then sCepEntrega = "00000-000"
                    sUFEntrega = SubNull(RdoAux2!Ee_Uf) '330-331
                    If UCase$(Trim(sCidEntrega)) <> "JABOTICABAL" Then
                        tTipo = 2
                    Else
                        tTipo = 0
                    End If
                Else
                    If RdoAux2!Ee_CodLog > 0 Then
                        Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & RdoAux2!Ee_CodLog
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If RdoAux3.RowCount = 0 Then
                           sEndEntrega = FillSpace(" ", 46)
                        Else
                           sEndEntrega = FillSpace(Trim$(SubNull(RdoAux3!AbrevTipoLog)) & " " & Trim$(SubNull(RdoAux3!AbrevTitLog)) & " " & RdoAux3!NomeLogradouro & " Nº " & CStr(RdoAux2!Ee_NumImovel), 46)
                        End If
                        RdoAux3.Close
                        tNum = RdoAux2!Ee_NumImovel
                        tEnd = RdoAux2!Ee_NomeLog
                        sComplEntrega = FillSpace(RdoAux2!Ee_Complemento, 20) '251-270
                        sBairroEntrega = FillSpace(SubNull(RdoAux2!Ee_DESCBairro), 30) '271-300
                        sCidEntrega = FillSpace(SubNull(RdoAux2!desccidade), 20) '301-320
                        sCepEntrega = Format(Replace(RdoAux2!Ee_Cep, "-", ""), "00000-000") '321-329
                        If Left$(sCepEntrega, 1) = "_" Then sCepEntrega = "00000-000"
                        sUFEntrega = SubNull(RdoAux2!Ee_Uf) '330-331
                        If UCase$(Trim(sCidEntrega)) <> "JABOTICABAL" Then
                            tTipo = 2
                        Else
                            tTipo = 0
                        End If
                    Else
                        If Not bPredial Then
                            tTipo = 1
                        Else
                            tTipo = 0
                        End If
                        sEndEntrega = sEnd '205-250
                        If RdoAux2.RowCount = 0 Then
                           tEnd = .NomeLogradouro
                        Else
                           tEnd = Trim$(SubNull(.AbrevTipoLog)) & " " & Trim$(SubNull(.AbrevTitLog)) & " " & .NomeLogradouro
                        End If
                        tNum = .Li_Num
                        sComplEntrega = FillSpace(sCompl, 20) '251-270
                        sBairroEntrega = FillSpace(sBairro, 30) '271-300
                        sCidEntrega = FillSpace("JABOTICABAL", 20) '301-320
                        sCepEntrega = Format("14870-000", "00000-000") '321-329
                        sUFEntrega = "SP" '330-331
                    End If
                End If
            Else
                sEndEntrega = FillSpace(" ", 46) '205-250
                sComplEntrega = FillSpace(" ", 20) '251-270
                sBairroEntrega = FillSpace(" ", 30) '271-300
                sCidEntrega = FillSpace(" ", 20) '301-320
                sCepEntrega = "00000-000"  '321-329
                sUFEntrega = "  " '330-331
                tBairro = ""
                tCidade = ""
                If Not bPredial Then
                    tTipo = 1
                Else
                    tTipo = 0
                    sEndEntrega = sEnd '205-250
                    tEnd = Trim$(SubNull(.AbrevTipoLog)) & " " & Trim$(SubNull(.AbrevTitLog)) & " " & .NomeLogradouro
                    tNum = .Li_Num
                    sComplEntrega = FillSpace(sCompl, 20) '251-270
                    sBairroEntrega = FillSpace(sBairro, 30) '271-300
                    sCidEntrega = FillSpace("JABOTICABAL", 20) '301-320
                    sCepEntrega = Format("14870-000", "00000-000") '321-329
                    sUFEntrega = "SP" '330-331
                    tBairro = sBairroEntrega
                    tCidade = sCidEntrega
                    If UCase$(Trim(sCidEntrega)) <> "JABOTICABAL" Then
                        tTipo = 2
                    End If
                End If
            End If
            RdoAux2.Close
           
            
            sCodContribuinte = Format(.CodigoImovel, "000000000") & "-" & RetornaDVCodReduzido(.CodigoImovel) '332-342
            sInscricao = Format(RemovePonto(Chomp(.Inscricao, chomp_righT, 7)), "00\.00\.0000\.00000\.00") '343-361
            sCodMunicipio = "391-8" '362-366
            sNatureza = FillSpace(RdoAux!Natureza, 11) '367-377
            If .Dt_FracaoIdeal = 0 Then
                sAreaTerreno = FillLeft(FormatNumber(.Dt_AreaTerreno, 2), 10)
            Else
                sAreaTerreno = FillLeft(FormatNumber(.Dt_FracaoIdeal, 2), 10)
            End If
            sAreaConstrucao = IIf(Val(RdoAux!areaconstrucao) = 0, Space(6) & "0,00", FillLeft(FormatNumber(RdoAux!areaconstrucao, 2), 10))
            sTestadaPrincipal = IIf(Val(RdoAux!TESTADAPRINC) = 0, Space(6) & "0,00", FillLeft(FormatNumber(RdoAux!TESTADAPRINC, 2), 10))
            sVVTerreno = IIf(Val(RdoAux!VVT) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!VVT, 2), 17))
            sVVConstrucao = IIf(Val(RdoAux!vvc) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!vvc, 2), 17))
            sVVImovel = IIf(Val(RdoAux!VVI) = 0, Space(13) & "0,00", FillLeft(FormatNumber(RdoAux!VVI, 2), 17))
            sValorIPU = FillLeft(FormatNumber(nValorParcIPTU * 9, 2), 17)
            sValorITU = FillLeft(FormatNumber(0, 2), 17)
            sQtdeParcela = Format(9, "00") '581-582
'            sTxExpUnica = FillLeft(FormatNumber(Rdoaux!TXEXPUNICA, 2), 17)
'            sTxExpParc = FillLeft(FormatNumber(Rdoaux!txexpparc / Val(sQtdeParcela), 2), 17)

            
            sTxExpUnica = FillLeft(FormatNumber(1.36, 2), 17)
            sTxExpParc = FillLeft(FormatNumber(1.36, 2), 17)
            aVencParc(0) = "00/00/0000"
            aNumDoc(0) = Format(0, "000000000000000")
            nSomaUnica = 0
            aNosNum(0) = FillLeft(0, 13)
            aDataVencto(0) = "01/01/1900"
            
            Sql = "SELECT DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,"
            Sql = Sql & "DEBITOPARCELA.DATAVENCIMENTO,DEBITOTRIBUTO.CODTRIBUTO,DEBITOPARCELA.DATADEBASE,"
            Sql = Sql & "DEBITOTRIBUTO.VALORTRIBUTO,PARCELADOCUMENTO.NUMDOCUMENTO,NumDocumento.DATADOCUMENTO FROM DEBITOTRIBUTO INNER JOIN "
            Sql = Sql & "PARCELADOCUMENTO ON DEBITOTRIBUTO.CODREDUZIDO = PARCELADOCUMENTO.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = PARCELADOCUMENTO.ANOEXERCICIO AND "
            Sql = Sql & "DEBITOTRIBUTO.CODLANCAMENTO = PARCELADOCUMENTO.CODLANCAMENTO AND DEBITOTRIBUTO.SEQLANCAMENTO = PARCELADOCUMENTO.SEQLANCAMENTO AND "
            Sql = Sql & "DEBITOTRIBUTO.CODCOMPLEMENTO = PARCELADOCUMENTO.CODCOMPLEMENTO AND DEBITOTRIBUTO.NUMPARCELA = PARCELADOCUMENTO.NUMPARCELA Inner Join "
            Sql = Sql & "NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join DEBITOPARCELA ON DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND "
            Sql = Sql & "DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
            Sql = Sql & "DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
            Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOTRIBUTO.ANOEXERCICIO = " & Val(sExercicio) & " AND DEBITOTRIBUTO.CODLANCAMENTO = 1 AND DEBITOTRIBUTO.CODCOMPLEMENTO=1  AND DEBITOTRIBUTO.CODTRIBUTO <> 3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nSomaTrib = 0: nSomaUnica = 0: aNumDoc(0) = 0
                Do Until .EOF
                   aCodProc(!NumParcela) = Format(!CODREDUZIDO, "00000000") & ".00." & !AnoExercicio & "." & Format(!CodLancamento, "00") & "." & Format(!SeqLancamento, "00") & "." & Format(!NumParcela, "00") & "." & !CODCOMPLEMENTO & ".0"
                   If !NumParcela = 0 Then
                      nSomaUnica = !ValorTributo + CDbl(sTxExpParc)
                      aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
                      aNumDoc(0) = Format(!NumDocumento & RetornaDVNumDoc(!NumDocumento), "000000000000000")
                      'aNosNum(0) = "023 " & Format(!NumDocumento, "0000000") & " " & "0"
                      aNosNum(0) = FillLeft(!NumDocumento, 13)
'                      aNosNum(0) = FillLeft(!NumDocumento & Modulo11(!NumDocumento), 13)
                      'sNosNum = sNosNum & FillLeft(aNosNum(x) & Modulo11(aNosNum(x)), 13)
                      aDataVencto(0) = Format(!DataVencimento, "dd/mm/yyyy")
                   Else
                      nValorParc = !ValorTributo + CDbl(sTxExpParc)
                      sDataDoc = Format(!DATADOCUMENTO, "dd/mm/yyyy") '561-570
                      nSomaTrib = nSomaTrib + !ValorTributo + CDbl(sTxExpParc)
                      aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                      aNumDoc(!NumParcela) = Format(!NumDocumento & RetornaDVNumDoc(!NumDocumento), "000000000000000")
                      'aNosNum(!NumParcela) = "023 " & Format(!NumDocumento, "0000000") & " " & "0"
                      aNosNum(!NumParcela) = FillLeft(!NumDocumento, 13)
                      'aNosNum(!NumParcela) = FillLeft(!NumDocumento & Modulo11(!NumDocumento), 13)
                      aDataVencto(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
                      'tirar
                      '**************
'                        If !NumParcela = 1 Then
'                            aNumDoc(0) = Format(!NumDocumento - 1 & RetornaDVNumDoc(!NumDocumento - 1), "000000000000000")
'                            aNosNum(0) = FillLeft(!NumDocumento - 1, 13)
'                        End If
                      '**************
                   End If
                  .MoveNext
                Loop
               .Close
            End With
            
            'tirar
            '*******************
            If nSomaUnica = 0 Then
                'nSomaUnica = nSomaTrib - ((nSomaTrib * 5) / 100)
            End If
            '*******************
            
            aDescParc(0) = "ÚNICA"
            For x = 1 To Val(sQtdeParcela)
                aDescParc(x) = Format(x, "00") & "/" & sQtdeParcela
            Next
            For x = Val(sQtdeParcela) + 1 To 12
                aDescParc(x) = "00/00"
            Next
            sDescParc = ""
            For x = 0 To 12
                sDescParc = sDescParc & aDescParc(x) '597-661
            Next
            
            For x = Val(sQtdeParcela) + 1 To 12
                aVencParc(x) = "00/00/0000"
            Next
            sVencParc = ""
            For x = 0 To 12
                sVencParc = sVencParc & aVencParc(x) '662-791
            Next
            
            'If Val(sQtdeParcela) < 12 Then MsgBox "teste"
            'aNosNum(0) = "0000000000000"
            sNosNum = "0000000000000"
            On Error Resume Next
'            If sQtdeParcela < 12 Then MsgBox "teste"
            For x = 1 To 12
                If x <= Val(sQtdeParcela) Then
                    If aNosNum(x) = "" Then
                        sNosNum = sNosNum & "0000000000000"
                    Else
                        sNosNum = sNosNum & FillLeft(Val(aNosNum(x)) & Modulo11(aNosNum(x)), 13)
                    End If
                Else
                    sNosNum = sNosNum & "0000000000000"
                End If
            Next

            For x = 0 To 12
                aNumDoc(x) = "000000000000000"
'                aNosNum(x) = "0000000000000"
            Next
            sNumDoc = ""
'            sNosNum = ""
            For x = 0 To 12
                sNumDoc = sNumDoc & aNumDoc(x) '1013-1168
                'sNosNum = sNosNum & aNosNum(x) '1546-1714
'                sNosNum = sNosNum & aNosNum(x) & Modulo11(aNosNum(x)) '1546-1714
            Next
            For x = Val(sQtdeParcela) + 1 To 12
                aCodProc(x) = "00000000000000000000000000000"
            Next
'            aCodProc(0) = "00000000000000000000000000000"
            sCodProc = ""
            For x = 0 To 12
                If aCodProc(x) = "" Then
                    aCodProc(x) = "00000000000000000000000000000"
                End If
                sCodProc = sCodProc & aCodProc(x) '1013-1168
            Next
            
'            sValorTotal = Format(nSomaTrib, "00\.000\.000\.000.00") '527-543
'            sValorUnica = Format(nSomaUnica, "00\.000\.000\.000.00") '544-560
            sValorTotal = FillLeft(FormatNumber(nSomaTrib, 2), 17)  '527-543
            sValorUnica = FillLeft(FormatNumber(nSomaUnica, 2), 17)  '544-560
            sDataProc = Format(Now, "dd/mm/yyyy") '571-580
            sAgencia = "023 45 00002 4"
            
            aValorParc(0) = sValorUnica
            For x = 1 To Val(sQtdeParcela)
                'aValorParc(x) = Format(nValorParc, "00\.000\.000\.000.00")
                aValorParc(x) = nValorParc
            Next
            For x = Val(sQtdeParcela) + 1 To 12
                'aValorParc(x) = Format("0", "00\.000\.000\.000.00")
                aValorParc(x) = "0"
            Next
            sValorParc = ""
            For x = 0 To 12
                sValorParc = sValorParc & FillLeft(FormatNumber(aValorParc(x), 2), 17)
            Next
                        
            dDataBase = "07/10/1997"
            'aDataVencto(0) = aDataVencto(1)
            aCodBarra(0) = Format("0", "00000000000000000000000000000000000000000000")
            For x = 1 To Val(sQtdeParcela)
                nFatorVencto = aDataVencto(x) - dDataBase
                'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & RetornaNumero(sAgencia) & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "0000000") & "0003300"
                aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "9" & "1235028" & Format(aNosNum(x) & Modulo11(aNosNum(x)), "0000000000000") & "0102"
            Next
            For x = Val(sQtdeParcela) + 1 To 12
                aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
            Next
            sCodBarra = ""
            For x = 0 To 12
                sCodBarra = sCodBarra & aCodBarra(x) '792-1012
            Next
            
            ax = sExercicio & sContribuinte & sSacado & sEnd & sCompl & sBairro & sQuadra & sCEP & sLote & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega & sUFEntrega
            ax = ax & sCodContribuinte & sInscricao & sCodMunicipio & sNatureza & sAreaTerreno & sAreaConstrucao & sTestadaPrincipal & sVVTerreno & sVVConstrucao & sVVImovel & sValorIPU
            ax = ax & sValorITU & sTxExpUnica & FillLeft(FormatNumber(sTxExpParc * Val(sQtdeParcela), 2), 17) & sValorTotal & sValorUnica & sDataProc & sDataProc & sQtdeParcela & sAgencia & sDescParc & sVencParc & sValorParc & sNumDoc
            'ax = ax & sCodProc & sNosNum & sCodBarra & Format(xId + 1, "000000")
            ax = ax & sCodProc & sNosNum & sCodBarra

            tDado = ax
            
            ax = tDado & "," & tEnd & "," & tNum & "," & tTipo
            
            Print #1, ax
            
            Sql = "INSERT LASERTMP (DADO,CIDADE,BAIRRO,ENDERECO,NUMERO,TIPO) VALUES('" & Mask(tDado) & "','"
            Sql = Sql & Trim(Mask(tBairro)) & "','" & Trim(Mask(tCidade)) & "','" & Trim(Mask(tEnd)) & "','" & tNum & "','" & tTipo & "')"
            cn.Execute Sql, rdExecDirect
            
        End With
Proximo:
        xId = xId + 1
       .MoveNext
    Loop
Sair:
   .Close
End With

Close #1
'Exit Sub
ORDENA:

Open sPathBin & "\LASERCORREIO.TXT" For Output As #1
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=0 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #1, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #1

Open sPathBin & "\LASERBALCAO.TXT" For Output As #2
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=1 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #2, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #2

Open sPathBin & "\LASERFORA.TXT" For Output As #2
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=2 ORDER BY BAIRRO,CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #2, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #2
MsgBox "FIM"
End Sub

Private Sub GeraEtiqueta()

Dim xId As Long, nNumRec As Long, nTipoEnd As Integer
Dim tEnd As String, tNum As Integer, tTipo As Integer
Dim tCidade As String, tBairro As String, Posicao As Long, Registro As aReg


'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sSacado As String, sEnd As String, sCompl As String, sBairro As String
Dim sCEP As String, sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String
Dim sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sCodContribuinte As String, sInscricao As String

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Open sPathBin & "\LASERFORA.TXT" For Binary Access Read Write As #1
    Posicao = 1
    nCount = 0
    xId = 1
    nNumRec = LOF(1) / 806
    Do While Not EOF(1)
         If xId Mod 100 = 0 Then
            CallPb xId, nNumRec
         End If
         
         Get #1, Posicao, Registro
         If EOF(1) Then GoTo fim
         sCodContribuinte = Mid$(Registro.sTexto, 335, 9)
         sInscricao = Mid$(Registro.sTexto, 346, 19)
         sContribuinte = Trim$(Mid$(Registro.sTexto, 5, 40))
         sEndEntrega = Trim$(Mid$(Registro.sTexto, 208, 45))
         sComplEntrega = Trim$(Mid$(Registro.sTexto, 254, 20))
         sBairroEntrega = Trim$(Mid$(Registro.sTexto, 274, 30))
         sCidEntrega = Trim$(Mid$(Registro.sTexto, 304, 20))
         sUFEntrega = Trim$(Mid$(Registro.sTexto, 333, 2))
         sCepEntrega = Trim$(Mid$(Registro.sTexto, 324, 9))
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodContribuinte & " - " & sInscricao & "','" & Mask(sContribuinte) & "','"
        Sql = Sql & Left$(Trim(Mask(sEndEntrega)) & "  " & Mask(sComplEntrega), 60) & "','" & FillSpace(Mask(sBairroEntrega), 30) & "  " & Mask(Trim$(sCidEntrega)) & "','"
        Sql = Sql & FillSpace(sUFEntrega, 50) & " " & sCepEntrega & "')"
        cn.Execute Sql, rdExecDirect
         Posicao = Posicao + Len(Registro) + 2
         xId = xId + 1
    Loop
 Close #1

fim:
Pb.Value = 100

frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
On Error Resume Next
Close #1
'Sql = "SELECT CODREDUZIDO, VVT, VVC, VVI, IMPOSTOPREDIAL,IMPOSTOTERRITORIAL, NATUREZA, AREACONSTRUCAO,"
'Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA From LASERIPTU "
''Sql = Sql & "WHERE CODREDUZIDO=115 OR CODREDUZIDO=116 "
'Sql = Sql & " ORDER BY CODREDUZIDO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    nNumRec = .RowCount
'    Do Until .EOF
'        If xId Mod 100 = 0 Then
'           CallPb xId, nNumRec
'        End If
'
'        With xImovel
'           .CarregaImovel RdoAux!CODREDUZIDO
'
'            sExercicio = "2014" '1-4
'            sContribuinte = FillSpace(.NomePropPrincipal, 40)  '5-44
'            sSacado = FillSpace(.NomePropPrincipal, 40)  '45-84
'            sEnd = FillSpace(Chomp(.EnderecoCompleto, chomp_left, 7), 46)   '85-130
'            sCompl = FillSpace(.Li_Compl, 20) '131-150
'            sBairro = FillSpace(.DescBairro, 30) '151-180
'            sCep = RetornaCEP(.CodLogr, .Li_Num)
'            If Len(sCep) <> 9 Then
'                sCep = "00000-000"
'            End If
'
'            sEndEntrega = ""
'            sComplEntrega = ""
'            sBairroEntrega = ""
'            sCidEntrega = ""
'            sCepEntrega = ""
'            sUFEntrega = ""
'            nTipoEnd = .Ee_TipoEnd
'            tTipo = 0
'            Sql = "SELECT ENDENTREGA.CODREDUZIDO,ENDENTREGA.EE_CODLOG, ENDENTREGA.EE_NOMELOG,ENDENTREGA.EE_NUMIMOVEL,"
'            Sql = Sql & "ENDENTREGA.EE_COMPLEMENTO, ENDENTREGA.EE_UF,ENDENTREGA.EE_CIDADE, ENDENTREGA.EE_BAIRRO,"
'            Sql = Sql & "ENDENTREGA.EE_CEP, ENDENTREGA.EE_LOTEAMENTO,ENDENTREGA.EE_DESCBAIRRO , Cidade.DESCCIDADE "
'            Sql = Sql & "FROM ENDENTREGA INNER JOIN  CIDADE ON ENDENTREGA.EE_UF = CIDADE.SIGLAUF AND "
'            Sql = Sql & "ENDENTREGA.Ee_Cidade = Cidade.CODCIDADE WHERE CODREDUZIDO = " & RdoAux!CODREDUZIDO
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount > 0 Then
'                If RdoAux2!Ee_NomeLog <> "" And nTipoEnd = 2 Then
'                    sEndEntrega = FillSpace(RdoAux2!Ee_NomeLog & " Nº " & CStr(RdoAux2!Ee_NumImovel), 46) '205-250
'                    tNum = RdoAux2!Ee_NumImovel
'                    tEnd = RdoAux2!Ee_NomeLog
'                    sComplEntrega = FillSpace(RdoAux2!Ee_Complemento, 20) '251-270
'                    sBairroEntrega = FillSpace(SubNull(RdoAux2!Ee_DESCBairro), 30) '271-300
'                    sCidEntrega = FillSpace(SubNull(RdoAux2!desccidade), 20) '301-320
'                    sCepEntrega = Format(RdoAux2!Ee_Cep, "00000-000") '321-329
'                    sUFEntrega = SubNull(RdoAux2!Ee_Uf) '330-331
'                    tTipo = 0
'                Else
'                    If Not bPredial Then
'                        tTipo = 1
'                    Else
'                        tTipo = 0
'                    End If
'                    sEndEntrega = sEnd '205-250
'                    If RdoAux2.RowCount = 0 Then
'                       tEnd = .NomeLogradouro
'                    Else
'                       tEnd = Trim$(SubNull(.AbrevTipoLog)) & " " & Trim$(SubNull(.AbrevTitLog)) & " " & .NomeLogradouro
'                    End If
'                    tNum = .Li_Num
'                    sComplEntrega = FillSpace(sCompl, 20) '251-270
'                    sBairroEntrega = FillSpace(sBairro, 30) '271-300
'                    sCidEntrega = FillSpace("JABOTICABAL", 20) '301-320
'                    sCepEntrega = Format("14870-000", "00000-000") '321-329
'                    sUFEntrega = "SP" '330-331
'                End If
'            Else
'                GoTo PROXIMO
'                sEndEntrega = sEnd '205-250
'                tEnd = Trim$(SubNull(.AbrevTipoLog)) & " " & Trim$(SubNull(.AbrevTitLog)) & " " & .NomeLogradouro
'                tNum = .Li_Num
'                sComplEntrega = FillSpace(sCompl, 20) '251-270
'                sBairroEntrega = FillSpace(sBairro, 30) '271-300
'                sCidEntrega = FillSpace("JABOTICABAL", 20) '301-320
'                sCepEntrega = Format("14870-000", "00000-000") '321-329
'                sUFEntrega = "SP" '330-331
'            End If
'            If UCase$(Trim(sCidEntrega)) <> "JABOTICABAL" Then
'                tTipo = 2
'            End If
'            tBairro = sBairroEntrega
'            tCidade = sCidEntrega
'            RdoAux2.Close
'
'            sCodContribuinte = Format(.CodigoImovel, "000000\.00") & "-" & RetornaDVCodReduzido(.CodigoImovel) '332-342
'            sInscricao = Format(RemovePonto(Chomp(.Inscricao, chomp_right, 7)), "00\.00\.0000\.00000\.00") '343-361
'
'            If tTipo = 2 Then
'                Sql = "INSERT ETIQUETAIPTU (CODREDUZIDO,INSCRICAO,PROPRIETARIO,LOGRADOURO,NUMERO,COMPLEMENTO,BAIRRO,CIDADE,UF,CEP) VALUES("
'                Sql = Sql & RdoAux!CODREDUZIDO & ",'" & sInscricao & "','" & sContribuinte & "','" & Mask(sEndEntrega) & "'," & "0" & ",'" & sComplEntrega & "','"
'                Sql = Sql & Mask(sBairroEntrega) & "','" & sCidEntrega & "','" & sUFEntrega & "','" & sCepEntrega & "')"
'                cn.Execute Sql, rdExecDirect
'            End If
'        End With
'PROXIMO:
'        xId = xId + 1
'       .MoveNext
'    Loop
'   .Close
'End With

MsgBox "FIM"


End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function


Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.Value = (nPosF * 100) / nTotal
Else
   Pb.Value = 100
End If
lblPF.Caption = FormatNumber(Pb.Value, 2)

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub GeraISSFixoOld()
Dim xId As Long, nNumRec As Long, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset, x As Integer
Dim nCodLogr As Long, nNum As Integer, nQtdeParcF As Integer
Dim nExpParc As Double, nExpUnica As Double

'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sFantasia As String, sEnd As String, sCompl As String, sBairro As String, sCEP As String
Dim sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim sTipoImposto As String, sInscricao As String, sQtdeParc As String, sCodAtiv As String, sDescAtiv As String, sCodInscricao As String
Dim bISSFixo As Boolean, bTxLic As Boolean
Dim aDescParc(0 To 12) As String, sDescParc As String
Dim sCodTrib As String
Dim aDescTrib(0 To 10) As String, sDescTrib As String
Dim aVencParc(0 To 12) As String, sVencParc As String
Dim aValorTributoUnica(0 To 12) As String, sValorTributoUnica As String
Dim aValorTributo(0 To 12) As String, sValorTributo As String
Dim aValorParc(0 To 12) As String, sValorParc As String
Dim aValorParcSEXP(0 To 12) As String, sValorParcSEXP As String
Dim aMesAno(0 To 12) As String, sMesAno As String, sMes As String
Dim aNumDoc(0 To 12) As String, sNumDoc As String
Dim nTotalTrib As Double, sTotalTrib As String
Dim nTotalTribUnica As Double, sTotalTribUnica As String
Dim aCodBarra(0 To 12) As String, sCodBarra As String
Dim dDataBase As Date, nUfir As Double
Dim tDado As String, tEnd As String, tNum As Integer, tTipo As Integer
Dim tCidade As String, tBairro As String, nDesc5Perc As Double
Dim sAgencia As String
Dim sValorEXP As String

'GoTo ORDENA
'********************************
' PARAMETROS DAS PARCELAS
'********************************
'PARCELAS PARA ISS FIXO E TLL
Sql = "SELECT QTDEPARCELA FROM PARAMPARCELA "
Sql = Sql & "WHERE ANO=" & 2014 & " AND CODTIPO=2"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nQtdeParcF = !qtdeparcela
   .Close
End With

nUfir = RetornaUFIR(2014)
sAgencia = "02345000024"

Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & 2014
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nExpParc = FormatNumber(RdoAux!VALORPARCELA, 2)
nExpUnica = FormatNumber(RdoAux!ValorUnica, 2)

'## ************************************************ ##
'## ***************** G U I A ********************** ##
'## ************************************************ ##

Sql = "TRUNCATE TABLE LASERTMP"
cn.Execute Sql, rdExecDirect

Open sPathBin & "\LASERISSFIXOTL.TXT" For Output As #1

Sql = "SELECT DISTINCT CODREDUZIDO From DEBITOPARCELA "
Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 2 OR CODLANCAMENTO=6 ) "
'Sql = Sql & "WHERE (ANOEXERCICIO = 2014) AND (CODLANCAMENTO = 2 OR CODLANCAMENTO=6 ) AND (DATAVENCIMENTO > '01/10/2014') "
'Sql = Sql & "AND (CODREDUZIDO=113374) "
Sql = Sql & " ORDER BY CODREDUZIDO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
'        If !CODREDUZIDO = 112299 Then MsgBox "TESTE"
        If xId Mod 5 = 0 Then
           CallPb xId, nNumRec
        End If
        Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
        Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,"
        Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
        Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
        Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
        Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then GoTo Proximo
            nCodLogr = !CodLogradouro
            sExercicio = "2014" '1-4
            sCodInscricao = Format(!codigomob, "00000000000000")
            sContribuinte = FillSpace(!RazaoSocial, 40) '5-44
            sFantasia = FillSpace(SubNull(!NOMEFANTASIA), 40) '45-84
            sCodAtiv = Format(!codatividade, "00000000000000")
            sDescAtiv = FillSpace(Left$(!ATIVEXTENSO, 50), 50)
            Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    sEnd = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero), 46) '212-257
                    nNum = RdoAux2!Numero
                Else
                    nNum = 0
                End If
               .Close
            End With
            sCEP = RetornaCEP(nCodLogr, nNum)
            sCompl = FillSpace(SubNull(Left(!Complemento, 20)), 20) '258-277
            sBairro = FillSpace(SubNull(!DescBairro), 30) '278-307
            Sql = "SELECT MOBILIARIOENDENTREGA.CODMOBILIARIO, MOBILIARIOENDENTREGA.CODLOGRADOURO, MOBILIARIOENDENTREGA.NOMELOGRADOURO, "
            Sql = Sql & "MOBILIARIOENDENTREGA.NUMIMOVEL,MOBILIARIOENDENTREGA.COMPLEMENTO, MOBILIARIOENDENTREGA.UF,MOBILIARIOENDENTREGA.CODCIDADE,"
            Sql = Sql & "MOBILIARIOENDENTREGA.CODBAIRRO, MOBILIARIOENDENTREGA.CEP,MOBILIARIOENDENTREGA.DESCBAIRRO, MOBILIARIOENDENTREGA.DESCCIDADE, BAIRRO.DESCBAIRRO AS DESCBAIRRO2,"
            Sql = Sql & "CIDADE.DESCCIDADE AS DESCCIDADE2 FROM BAIRRO INNER JOIN MOBILIARIOENDENTREGA ON BAIRRO.SIGLAUF = MOBILIARIOENDENTREGA.UF AND BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO INNER JOIN "
            Sql = Sql & "CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE Where CODMOBILIARIO = " & !codigomob
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    If !CodLogradouro = 0 Then
                        sEndEntrega = FillSpace(!NomeLogradouro & " Nº " & CStr(!numimovel), 46) '85-130
                        sNumEntrega = SubNull(!numimovel)
                        If !CodBairro = 0 Then
                            sBairroEntrega = FillSpace(SubNull(!DescBairro), 30) '151-180
                        ElseIf !CodBairro = 999 Then
                            sBairroEntrega = FillSpace(" ", 30) '151-180
                        Else
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30)
                        End If
                        If !CodCidade = 0 Then
                            sCidEntrega = FillSpace(SubNull(!desccidade), 20) '181-200
                        Else
                            sCidEntrega = FillSpace(SubNull(!desccidade2), 20)
                        End If
                        If !CodCidade = 413 Then
                            sCepEntrega = Format(!Cep, "00000-000") '201-209
                        Else
                            sCepEntrega = RetornaCEP(!CodLogradouro, Val(sNumEntrega))
                        End If
                        sComplEntrega = FillSpace(SubNull(!Complemento), 20)
                        sUFEntrega = !UF '210-211
                    Else
                        Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !CodLogradouro
                        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux4
                            If .RowCount > 0 Then
                                sEndEntrega = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux3!numimovel), 46) '85-130
                                'sEndEntrega = FillSpace(Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero), 46) '85-130
                                sNumEntrega = SubNull(RdoAux3!numimovel)
                                sCepEntrega = RetornaCEP(RdoAux3!CodLogradouro, Val(sNumEntrega))
                            End If
                            sBairroEntrega = FillSpace(SubNull(RdoAux3!DescBairro2), 30) '151-180
                            sCidEntrega = FillSpace(SubNull(RdoAux3!desccidade2), 20) '181-200
                            sComplEntrega = FillSpace(SubNull(RdoAux2!Complemento), 20)
                            sUFEntrega = RdoAux3!UF '210-211
                           .Close
                        End With
                    End If
                Else
                    nCodLogr = 0
                    sEndEntrega = sEnd '85-130
                    sBairroEntrega = FillSpace(sBairro, 30) '151-180
                    sCidEntrega = FillSpace("JABOTICABAL", 20) '181-200
                    sCepEntrega = "14870-000" '201-209
                    sComplEntrega = FillSpace(" ", 20)
                    sUFEntrega = "SP" '210-211
                End If
               .Close
            End With
            
           .Close
        End With
        tEnd = sEndEntrega
        tNum = Val(sNumEntrega)
        tBairro = sBairroEntrega
        tCidade = sCidEntrega
        If Left(sCepEntrega, 1) = "_" Then sCepEntrega = "         "
        If Trim(sCepEntrega) = "" Then sCepEntrega = "14870-000"
        
        'INSCRICAO
        sInscricao = Format(Val(sCodInscricao), "000000000000000")
'        If nCodLogr > 0 Then
'            Sql = "SELECT DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,"
'            Sql = Sql & "CODLOGR,LI_NUM FROM vwCnsImovel WHERE CODLOGR=" & nCodLogr
'            Sql = Sql & " AND LI_NUM=" & nNum
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                If .RowCount > 0 Then
'                   sInscricao = Format(!Distrito, "00") & Format(!Setor, "00") & Format(!Quadra, "0000") & Format(!Lote, "00000") & Format(!Seq, "00")
'                End If
'               .Close
'            End With
'        Else
'            sInscricao = "000000000000000"
'        End If
        'TIPOS DE CARNE A GERAR
        Sql = "SELECT DISTINCT CODLANCAMENTO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=2014"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            bISSFixo = False
            bTxLic = False
            Do Until .EOF
                Select Case !CodLancamento
                Case 2
                    bISSFixo = True
                Case 6
                    bTxLic = True
                End Select
              .MoveNext
            Loop
           .Close
        End With
        If bTxLic Then 'TAXA DE LICENÇA
            sTipoImposto = FillSpace("TAXA DE LICENÇA", 20)
            nCodLanc = 6
            Sql = "SELECT COUNT(CODREDUZIDO) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=2014 "
            Sql = Sql & " AND CODLANCAMENTO=6"
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If !CONTADOR < 4 Then
                    nCodLanc = 2
                End If
               .Close
            End With
        Else
            sTipoImposto = FillSpace("ISS FIXO/TLL", 20)
            nCodLanc = 2
        End If
        
        sQtdeParc = Format(nQtdeParcF, "0000000000") '427-428
        Sql = "SELECT DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.DATAVENCIMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,"
        Sql = Sql & "NumDocumento.DATADOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento Inner Join "
        Sql = Sql & "DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
        Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND "
        Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
        Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = " & nCodLanc
        'Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & Val(RdoAux!CODREDUZIDO) & " AND DEBITOPARCELA.ANOEXERCICIO >= " & Val(sExercicio) & " AND DEBITOPARCELA.CODLANCAMENTO = " & nCodLanc & "  AND (DATAVENCIMENTO > '01/10/2014')"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nSomaTrib = 0
            For x = 0 To 12
                aMesAno(x) = ""
            Next
            For x = 1 To 10
                aValorTributo(x) = ""
                aValorTributoUnica(x) = ""
            Next
            If .RowCount < 4 Then
                GoTo Proximo
            End If
            Do Until .EOF
               Select Case !NumParcela
                    Case 1
                        sMes = " 01"
                    Case 2
                        sMes = " 02"
                    Case 3
                        sMes = " 03"
                    Case 4
                        sMes = " 04"
                    Case 5
                        sMes = " 05"
                    Case 6
                        sMes = " 06"
                    Case 7
                        sMes = " 07"
                    Case 8
                        sMes = " 08"
                    Case 9
                        sMes = " 09"
                    Case 10
                        sMes = " 10"
                    Case 11
                        sMes = " 11"
                    Case 12
                        sMes = " 12"
               End Select
               If !NumParcela = 0 Then
                   aMesAno(!NumParcela) = "UNICA"
               Else
                   aMesAno(!NumParcela) = sMes & "/03"
               End If
               aNumDoc(!NumParcela) = FillLeft(!NumDocumento, 9)
               If !NumParcela = 0 Then
                  aVencParc(0) = Format(!DataVencimento, "dd/mm/yyyy")
               Else
                  aVencParc(!NumParcela) = Format(!DataVencimento, "dd/mm/yyyy")
               End If
               If !NumParcela = 1 Then
                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                    Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                    Sql = Sql & " ORDER BY DESCTRIBUTO"
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        x = 1
                        Do Until .EOF
                            aDescTrib(x) = FillSpace(!ABREVTRIBUTO, 15)
                            aValorTributo(x) = FillLeft(FormatNumber(!ValorTributo, 2), 17)
                            If !CodTributo = 14 Or !CodTributo = 11 Then
                                aValorTributoUnica(x) = FillLeft(FormatNumber(!ValorTributo * 3, 2), 17)
                                nDesc5Perc = FormatNumber(aValorTributoUnica(x) * 0.05, 2)
                            End If
                            x = x + 1
                           .MoveNext
                        Loop
                       .Close
                    End With
                    For s = x To 10
                        aDescTrib(s) = FillSpace(" ", 15)
                    Next
                    sDescTrib = ""
                    For x = 0 To 10
                        sDescTrib = sDescTrib & aDescTrib(x) '449-598
                    Next
               ElseIf !NumParcela = 0 Then
                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
                    Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & Val(RdoAux!CODREDUZIDO) & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO<>3"
                    Sql = Sql & " ORDER BY DESCTRIBUTO"
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        x = 1
                        Do Until .EOF
                            If !CodTributo <> 14 And !CodTributo <> 11 Then
 '                               aValorTributoUnica(x) = FillLeft(FormatNumber(!VALORTRIBUTO + (!VALORTRIBUTO * 0.05), 2), 17)
 '                           Else
                                aValorTributoUnica(x) = FillLeft(FormatNumber(!ValorTributo, 2), 17)
                            End If
                            x = x + 1
                           .MoveNext
                        Loop
                       .Close
                    End With
               End If
              .MoveNext
            Loop
           .Close
        End With
        aDescParc(0) = "UNICA"
        For x = 1 To Val(sQtdeParc)
            aDescParc(x) = Format(x, "00") & "/03"
        Next
        For x = nQtdeParcF + 1 To 12
            aDescParc(x) = "00/00"
        Next
        sDescParc = ""
        For x = 0 To 12
            sDescParc = sDescParc & aDescParc(x)
        Next
        
        sMesAno = ""
'        aMesAno(0) = FillLeft(" ", 6)
        If aMesAno(0) = "" Then GoTo Proximo
        For x = nQtdeParcF + 1 To 12
            If aMesAno(x) = "" Then
               aMesAno(x) = "      "
            End If
        Next
        For x = 0 To 12
            sMesAno = sMesAno & FillLeft(aMesAno(x), 13)
        Next
        
        For x = nQtdeParcF + 1 To 12
            aVencParc(x) = "00/00/0000"
        Next
        sVencParc = ""
        For x = 0 To 12
            sVencParc = sVencParc & aVencParc(x) '662-791
        Next
                
        For x = Val(sQtdeParc) + 1 To 12
            aNumDoc(x) = "000000000"
        Next
        sNumDoc = ""
        'aNumDoc(0) = "0"
        For x = 0 To 12
            sNumDoc = sNumDoc & Format(aNumDoc(x), "000000000") '662-791
        Next
                
        sValorTributoUnica = ""
        sValorTributo = ""
        For x = 1 To 10
            If aValorTributo(x) = "" Then
               aValorTributo(x) = FillLeft("0,00", 17)
            End If
        Next
        For x = 1 To 10
            If aValorTributoUnica(x) = "" Then
               aValorTributoUnica(x) = FillLeft("0,00", 17)
            End If
        Next
        
        For x = 1 To 10
            sValorTributo = sValorTributo & FillLeft(FormatNumber(aValorTributo(x), 2), 17)
            sValorTributoUnica = sValorTributoUnica & FillLeft(aValorTributoUnica(x), 17)
        Next
        
        nTotalTrib = 0
        nTotalTribUnica = 0
        For x = 0 To 10
            If x > 0 Then
                If aValorTributo(x) = "" Then Exit For
                nTotalTrib = nTotalTrib + CDbl(aValorTributo(x))
            End If
        Next
        sTotalTrib = FillLeft(CStr(FormatNumber(nTotalTrib, 2)), 17)
        For x = 0 To 10
            If x > 0 Then
                If aValorTributoUnica(x) = "" Then Exit For
                nTotalTribUnica = nTotalTribUnica + CDbl(aValorTributoUnica(x))
            End If
        Next
        nTotalTribUnica = nTotalTribUnica - nDesc5Perc
        sTotalTribUnica = FillLeft(CStr(FormatNumber(nTotalTribUnica, 2)), 17)
FIMVS:


        'nCodLanc = 13
        aValorParc(0) = FormatNumber(nTotalTribUnica + nExpUnica, 2)
'        aValorParc(0) = FormatNumber(nTotalTrib - (nTotalTribUnica * 0.05) + nExpUnica, 2)
        aValorParcSEXP(0) = FormatNumber(nTotalTribUnica, 2)
        For x = 1 To nQtdeParcF
            aValorParc(x) = FormatNumber(((nTotalTrib)) + nExpParc, 2)
            aValorParcSEXP(x) = FormatNumber(((nTotalTrib)), 2)
        Next
        For x = nQtdeParcF + 1 To 12
            aValorParc(x) = "0,00"
            aValorParcSEXP(x) = "0,00"
        Next
        sValorParc = ""
        sValorParcSEXP = ""
        For x = 0 To 12
            sValorParc = sValorParc & FillLeft(aValorParc(x), 18) '662-791
            sValorParcSEXP = sValorParcSEXP & FillLeft(aValorParcSEXP(x), 18) '662-791
        Next
        
        sValorEXP = FillLeft(FormatNumber(nExpParc, 2), 17)
        
        sDataProc = Format(Now, "dd/mm/yyyy") '439-448
        sDataDoc = FillLeft(Format(Now, "dd/mm/yyyy"), 15) '561-570
'        sDataProc = "01/01/2014" '439-448
'        sDataDoc = FillLeft("01/01/2014", 15) '561-570
        

'        For x = 0 To nQtdeParcF
'            If Not IsDate(aVencParc(x)) Then Exit For
'            If x = 0 Then
'            '    aCodBarra(x) = "8160" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
 '               aCodBarra(x) = "8160" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(aNumDoc(x)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
 '           Else
'             '   aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'                aCodBarra(x) = "8170" & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "00000000000") & "2177" & Year(aVencParc(x)) & Format(Month(aVencParc(x)), "00") & Format(Day(aVencParc(x)), "00") & Format(Val(aNumDoc(x)), "000000000") & Format(x, "00") & Format(nCodLanc, "00") & "9999"
'            End If
'        Next
        
        dDataBase = "07/10/1997"
        aCodBarra(0) = Format("0", "00000000000000000000000000000000000000000000")
        For x = 0 To nQtdeParcF
            If Not IsDate(aVencParc(x)) Then Exit For
            nFatorVencto = CDate(aVencParc(x)) - dDataBase
            'aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & "023" & Format(aNumDoc(x), "00000000") & Format(Val(Chomp(aNumDoc(x), chomp_right, 1)), "0000000") & "0003300"
            aCodBarra(x) = "03390" & Format(nFatorVencto, "0000") & Format(Val(RetornaNumero(FormatNumber(aValorParc(x), 2))), "0000000000") & sAgencia & Format(Val(aNumDoc(x)), "0000000") & "0003300"
        Next
        For x = 4 To 12
            aCodBarra(x) = Format("0", "00000000000000000000000000000000000000000000")
        Next
        sCodTrib = FillSpace(" ", 150) '?????
        sCodBarra = ""
        For x = 0 To 12
             sCodBarra = sCodBarra & aCodBarra(x) '792-1012
        Next
 
        sTotalTrib = FillLeft("0,00", 17) '????
        ax = sExercicio & sContribuinte & sFantasia & sEndEntrega & sComplEntrega & sBairroEntrega & sCidEntrega & sCepEntrega
        ax = ax & sUFEntrega & sEnd & sCompl & sBairro & sCEP & sTipoImposto & sInscricao & sCodInscricao & sDescAtiv
        ax = ax & sDescParc & sMesAno & FillLeft(sQtdeParc, 10) & sDataDoc & sDataProc & sNumDoc & sCodTrib & sValorTributoUnica & sValorParcSEXP
        ax = ax & sTotalTrib & "          " & sVencParc & sValorParc & sCodBarra & sDescTrib & sValorEXP & FillLeft(FormatNumber("3,51", 2), 17)
        tTipo = "3"
        Print #1, ax


        tDado = ax
        ax = tDado & "," & tEnd & "," & tNum & "," & tTipo
        
        Sql = "INSERT LASERTMP (DADO,CIDADE,BAIRRO,ENDERECO,NUMERO,TIPO) VALUES('" & Mask(tDado) & "','"
        Sql = Sql & Trim(Mask(tBairro)) & "','" & Trim(Mask(tCidade)) & "','" & Trim(Mask(tEnd)) & "','" & tNum & "','" & tTipo & "')"
        cn.Execute Sql, rdExecDirect
        
Proximo:
        xId = xId + 1
       .MoveNext
 '     Exit Do
    Loop
   .Close
End With

Close #1
'Exit Sub
ORDENA:

Open sPathBin & "\LASERISSFIXOTL.TXT" For Output As #1
Sql = "SELECT DADO,ENDERECO,NUMERO,TIPO FROM LASERTMP WHERE TIPO=3 ORDER BY CIDADE,ENDERECO,NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    Do Until .EOF
        Print #1, Trim(!dado) & Format(xId, "000000")
        xId = xId + 1
       .MoveNext
    Loop
End With
Close #1

End Sub

Private Sub CarregaEnderecoContabil()
Dim Sql As String, RdoAux As rdoResultset, nLast As Integer
ReDim aEnd(0)
Sql = "SELECT codigoesc, nomeesc, nomelogradouro, numero, nomebairro, cep, nomecidade, uf, recebecarne "
Sql = Sql & "From escritoriocontabil Where (codigoesc > 0)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aEnd(UBound(aEnd) + 1)
        nLast = UBound(aEnd)
        aEnd(nLast).nCodigo = !codigoesc
        aEnd(nLast).sNome = !NOMEESC
        aEnd(nLast).sLogradouro = !NomeLogradouro
        aEnd(nLast).nNumero = !Numero
        aEnd(nLast).sBairro = SubNull(!NOMEBairro)
        aEnd(nLast).sCEP = !Cep
        aEnd(nLast).sCidade = !NomeCidade
        aEnd(nLast).sUF = !UF
        aEnd(nLast).bRecebe = !RECEBECARNE
       .MoveNext
    Loop
   .Close
End With

End Sub
