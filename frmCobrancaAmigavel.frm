VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCobrancaAmigavel 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobrança Amigável"
   ClientHeight    =   4815
   ClientLeft      =   6210
   ClientTop       =   2460
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   5445
   Begin Tributacao.XP_ProgressBar PB 
      Height          =   240
      Left            =   765
      TabIndex        =   30
      Top             =   4005
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16750899
   End
   Begin VB.ComboBox cmbSetor 
      Height          =   315
      ItemData        =   "frmCobrancaAmigavel.frx":0000
      Left            =   3150
      List            =   "frmCobrancaAmigavel.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   3555
      Width           =   2190
   End
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      ItemData        =   "frmCobrancaAmigavel.frx":0029
      Left            =   1485
      List            =   "frmCobrancaAmigavel.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3105
      Width           =   3135
   End
   Begin VB.CheckBox chkSuj 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEEE&
      Caption         =   "Sujeito a cobrança judicial"
      Height          =   195
      Left            =   90
      TabIndex        =   25
      Top             =   3645
      Width           =   2175
   End
   Begin VB.ComboBox cmbTipoISS 
      Height          =   315
      ItemData        =   "frmCobrancaAmigavel.frx":004C
      Left            =   1470
      List            =   "frmCobrancaAmigavel.frx":005C
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   2715
      Width           =   3135
   End
   Begin VB.ComboBox cmbAjuizado 
      Height          =   315
      ItemData        =   "frmCobrancaAmigavel.frx":00B9
      Left            =   1470
      List            =   "frmCobrancaAmigavel.frx":00C6
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2325
      Width           =   3135
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Cidadão"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   3480
      TabIndex        =   19
      Top             =   150
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtObs2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1920
      Width           =   3825
   End
   Begin VB.TextBox txtObs1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      MaxLength       =   300
      TabIndex        =   7
      Top             =   1560
      Width           =   3825
   End
   Begin VB.TextBox txtResp 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1200
      Width           =   3825
   End
   Begin VB.TextBox txtAno1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      MaxLength       =   4
      TabIndex        =   4
      Top             =   840
      Width           =   1035
   End
   Begin VB.TextBox txtCod1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   2
      Top             =   480
      Width           =   1035
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "ISS"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   2490
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "IPTU"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1500
      TabIndex        =   0
      Top             =   150
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtCod2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4260
      MaxLength       =   6
      TabIndex        =   3
      Top             =   480
      Width           =   1035
   End
   Begin VB.TextBox txtAno2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4260
      MaxLength       =   4
      TabIndex        =   5
      Top             =   840
      Width           =   1035
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   2430
      TabIndex        =   9
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   4350
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
      MICON           =   "frmCobrancaAmigavel.frx":0116
      PICN            =   "frmCobrancaAmigavel.frx":0132
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
      Left            =   3570
      TabIndex        =   10
      ToolTipText     =   "Sair da Tela"
      Top             =   4365
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
      MICON           =   "frmCobrancaAmigavel.frx":028C
      PICN            =   "frmCobrancaAmigavel.frx":02A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdMalaDireta 
      Height          =   345
      Left            =   540
      TabIndex        =   20
      ToolTipText     =   "Etiquetas para Mala Direta"
      Top             =   4365
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Mala Direta"
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
      MICON           =   "frmCobrancaAmigavel.frx":0316
      PICN            =   "frmCobrancaAmigavel.frx":0332
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      Caption         =   "Setor.:"
      Height          =   195
      Index           =   1
      Left            =   2475
      TabIndex        =   29
      Top             =   3615
      Width           =   555
   End
   Begin VB.Label Label8 
      Caption         =   "Status Lancam..:"
      Height          =   195
      Left            =   135
      TabIndex        =   27
      Top             =   3165
      Width           =   1275
   End
   Begin VB.Label Label7 
      Caption         =   "Lançamento........:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2775
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "Ajuizado..............:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   2385
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Vencto.......:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1950
      Width           =   1395
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Observação 1.....:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   1590
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Inicial...........:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   870
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Final.............:"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   870
      Width           =   1395
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Responsável.......:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Inicial......:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   510
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Imposto..:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   150
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Final........:"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   510
      Width           =   1395
   End
   Begin VB.Menu mnuSel 
      Caption         =   "Seleção"
      Visible         =   0   'False
      Begin VB.Menu mnuCarta 
         Caption         =   "Cartas"
      End
      Begin VB.Menu mnuBoletos 
         Caption         =   "Boletos"
      End
   End
End
Attribute VB_Name = "frmCobrancaAmigavel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSituacao As Integer
    sSituacao As String
    sVencto As String
    sDA As String
    sAj As String
    nCodTributo As Double
    nValorTributo As Double
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorAtual As Double
    nCodBanco As Integer
    dDataPag As Date
End Type

Private Sub cmdMalaDireta_Click()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery, RdoAux3 As rdoResultset, i As Long, x As Integer
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String, nPos As Long
Dim sBairro As String, sCidade As String, sUF As String, sCep As String, sEnderecoEntrega As String, nNumEntrega As Integer, sBairroEntrega As String
Dim sCidadeEntrega As String, sUFEntrega As String, sCepEntrega As String, aDebito() As Debito, bAchou As Boolean, sComplEntrega As String
Dim sCampo1 As String, sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String, aCodigos() As Long, k As Integer
Dim xImovel As clsImovel, clsImovel As New clsImovel
Set xImovel = New clsImovel
If MsgBox("Deseja emitir a mala direta para os códigos acima?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

If Val(txtCod1.Text) = 0 Or Val(txtCod2.Text) = 0 Then
    MsgBox "Digite código inicial e final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) = 0 Or Val(txtAno2.Text) = 0 Then
    MsgBox "Digite Ano inicial e final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtCod1.Text) > Val(txtCod2.Text) Then
    MsgBox "Código inicial maior que final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) > Val(txtAno2.Text) Then
    MsgBox "Código inicial maior que final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) < 1990 Or Val(txtAno1.Text) > Year(Now) Then
    MsgBox "Ano inicial fora de intervalo.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno2.Text) < 1990 Or Val(txtAno2.Text) > Year(Now) Then
    MsgBox "Ano final fora de intervalo.", vbCritical, "Atenção"
    Exit Sub
End If

ReDim aCodigos(0)
If Val(txtCod1.Text) >= 100000 Then
    If cmbTipoISS.ListIndex > 0 Then
        Sql = "SELECT * FROM VWDEBITOISS WHERE CODREDUZIDO BETWEEN " & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text) & " AND ANOEXERCICIO "
        Sql = Sql & "BETWEEN " & Val(txtAno1.Text) & " AND " & Val(txtAno2.Text)
        If cmbTipoISS.ListIndex = 1 Then
            Sql = Sql & " AND CODLANCAMENTO=5"
        ElseIf cmbTipoISS.ListIndex = 2 Then
            Sql = Sql & " AND CODLANCAMENTO=3"
        ElseIf cmbTipoISS.ListIndex = 3 Then
            Sql = Sql & " AND (CODLANCAMENTO=2)"
        End If
'        Sql = Sql & " order by codreduzido desc"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aCodigos(UBound(aCodigos) + 1)
                aCodigos(UBound(aCodigos)) = !CODREDUZIDO
               .MoveNext
            Loop
           .Close
        End With
    End If
End If


Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Ocupado
Set qd.ActiveConnection = cn
nTot = Val(txtCod2.Text) - Val(txtCod1.Text) + 1
For i = Val(txtCod1.Text) To Val(txtCod2.Text)
    If i < 100000 Then
        Sql = "select codreduzido,inativo from cadimob where codreduzido=" & i
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If !Inativo Then
                   .Close
                    GoTo NEXTONE
                End If
            End If
           .Close
        End With
    End If

    nPos = nPos + 1
    CallPb nPos, CLng(nTot)
    
    ReDim aDebito(0)
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    'SPEXTRATO 14, 14, 1950, 2005, 1, 99, 0, 9, 1, 99, 0, 9, 3, 3
    qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
    qd(0) = i
    qd(1) = i
    qd(2) = Val(txtAno1.Text)
    qd(3) = Val(txtAno2.Text)
    
    If cmbTipoISS.ListIndex = 0 Then
        qd(4) = 0
        qd(5) = 99
    Else
        qd(4) = cmbTipoISS.ItemData(cmbTipoISS.ListIndex)
        qd(5) = cmbTipoISS.ItemData(cmbTipoISS.ListIndex)
    End If
    qd(6) = 0
    qd(7) = 999 'SEQUENCIA
    qd(8) = 1
    qd(9) = 99 'PARCELA
    qd(10) = 0
    qd(11) = 9 'COMPLEMENTO
    qd(12) = cmbStatus.ItemData(cmbStatus.ListIndex)
    qd(13) = cmbStatus.ItemData(cmbStatus.ListIndex) 'STATUSLANC
    qd(14) = Format(txtObs2.Text, "mm/dd/yyyy")
    qd(15) = NomeDoUsuario
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        Do Until .EOF
'            If !CODREDUZIDO = 105086 Then MsgBox "teste"
               'Ajuizado
                If cmbAjuizado.ListIndex = 1 Then
                    If IsNull(!dataajuiza) Then GoTo proximo
                ElseIf cmbAjuizado.ListIndex = 2 Then
                    If Not IsNull(!dataajuiza) Then GoTo proximo
                End If
                If !ValorTributo = 0 Then GoTo proximo
                If !DataVencimento > Now Then GoTo proximo
'                If !DataVencimento > CDate("07/31/2014") Then GoTo Proximo
                'iss
                If cmbTipoISS.ListIndex > 0 Then
                    bAchou = False
                    If UBound(aCodigos) > 0 Then
                        For k = 1 To UBound(aCodigos)
                            If !CODREDUZIDO = aCodigos(k) Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                    End If
                End If

                'If Year(!DataVencimento) >= Year(Now) Then GoTo proximo
                If !DataVencimento >= Now Then GoTo proximo
                nAno = !AnoExercicio
                bAchou = False
                For x = 1 To UBound(aDebito)
                    If aDebito(x).nCodReduz = !CODREDUZIDO Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If bAchou Then
                    GoTo proximo
                Else
                    ReDim Preserve aDebito(UBound(aDebito) + 1)
                    aDebito(UBound(aDebito)).nCodReduz = !CODREDUZIDO
                End If
                If !CODREDUZIDO < 100000 Then
                    Sql = "select distinct * from vwcobrancaamigavel where codimovel =" & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !nomecidadao
                            sInscricao = !Inscricao
                            If !Ee_TipoEnd = 0 Then
                                xImovel.RetornaEndereco RdoAux!CODREDUZIDO, Imobiliario, Localizacao
                            ElseIf !Ee_TipoEnd = 1 Then
                                xImovel.RetornaEndereco RdoAux!CODREDUZIDO, Imobiliario, cadastrocidadao
                            Else
                                xImovel.RetornaEndereco RdoAux!CODREDUZIDO, Imobiliario, Entrega
                            End If
                            sEnderecoEntrega = xImovel.Endereco
                            nNumEntrega = Val(xImovel.Numero)
                            sComplemento = xImovel.Complemento
                            sBairroEntrega = xImovel.Bairro
                            sCidadeEntrega = xImovel.Cidade
                            sUFEntrega = xImovel.UF
                            sCep = xImovel.Cep
                        End If
                       .Close
                    End With
                ElseIf !CODREDUZIDO >= 100000 And !CODREDUZIDO < 300000 Then
                    Sql = "SELECT mobiliario.codigomob, mobiliario.razaosocial, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,"
                    Sql = Sql & "mobiliario.numero, mobiliario.complemento, bairro.descbairro, cidade.desccidade, mobiliario.siglauf, mobiliario.cep, mobiliario.codlogradouro, "
                    Sql = Sql & "mobiliario.nomelogradouro AS nomelogradouro2 FROM mobiliario INNER JOIN bairro ON mobiliario.siglauf = bairro.siglauf AND mobiliario.codcidade = bairro.codcidade AND "
                    Sql = Sql & "mobiliario.codbairro = bairro.codbairro INNER JOIN cidade ON bairro.siglauf = cidade.siglauf AND bairro.codcidade = cidade.codcidade LEFT OUTER JOIN "
                    Sql = Sql & "vwLOGRADOURO ON mobiliario.codlogradouro = vwLOGRADOURO.CODLOGRADOURO Where mobiliario.codigomob = " & !CODREDUZIDO & " and mobiliario.dataencerramento is null and "
                    Sql = Sql & "MOBILIARIO.CODIGOMOB NOT in (SELECT CODMOBILIARIO FROM vwMOBILIARIOSUSPENSO WHERE CODTIPOEVENTO=2)"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !RazaoSocial
                            sInscricao = !codigomob
                            If Val(SubNull(!CodLogradouro)) > 0 Then
                                sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                            Else
                                sEndereco = SubNull(!NOMELOGRADOURO2)
                            End If
                            nNumero = !Numero
                            sComplemento = SubNull(!Complemento)
                            sBairro = !DescBairro
                            sCidade = !descCidade
                            sUF = !SiglaUF
                            sCep = SubNull(!Cep)
                            sEnderecoEntrega = sEndereco
                            nNumEntrega = nNumero
                            sBairroEntrega = sBairro
                            sCidadeEntrega = sCidade
                            sUFEntrega = sUF
                            sCepEntrega = sCep
                        Else
                            GoTo proximo
                        End If
                       .Close
                    End With
                Else
                    Sql = "select distinct * from vwcobrancaamigavelcidadao where codcidadao =" & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !nomecidadao
                            sInscricao = !CodCidadao
                            If Not IsNull(!NomeLogradouro) Then
                                sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                sCep = RetornaCEP(!CodLogradouro, nNumero)
                            Else
                                sEndereco = SubNull(!NOMELOGRADOURO2)
                                sCep = SubNull(!Cep)
                            End If
                            nNumero = Val(SubNull(!NUMIMOVEL))
                            sComplemento = SubNull(!Complemento)
                            sBairro = SubNull(!DescBairro)
                            sCidade = SubNull(!descCidade)
                            sUF = SubNull(!SiglaUF)
                            
                            sEnderecoEntrega = sEndereco
                            nNumEntrega = nNumero
                            sBairroEntrega = sBairro
                            sCidadeEntrega = sCidade
                            sUFEntrega = sUF
                            sCepEntrega = sCep
                        Else
                            sNome = ""
                            sInscricao = ""
                            sEndereco = ""
                            nNumero = 0
                            sComplemento = ""
                            sBairro = ""
                            sCidade = ""
                            sUF = ""
                            sCep = ""
                            sEnderecoEntrega = ""
                            nNumEntrega = 0
                            sBairroEntrega = ""
                            sCidadeEntrega = ""
                            sUFEntrega = ""
                            sCep = ""
                        End If
                       .Close
                    End With
                End If
                             
                sEnd = sEnderecoEntrega & ", " & CStr(nNumEntrega)
                sCompl = sComplemento
                sBairro = sBairroEntrega
                sCid = sCidadeEntrega
                sUF = sUFEntrega
                             
                Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
                Sql = Sql & NomeDeLogin & "'," & nPos & ",'" & Format(!CODREDUZIDO, "000000") & "','" & Left(Mask(sNome), 200) & "','" & Left(Mask(CStr(sEnd)) & " " & sComplemento, 60) & "','" & Mask(CStr(sBairro) & " - " & sCid) & "/" & sUF & "','" & sCep & "')"
                cn.Execute Sql, rdExecDirect
             
proximo:
            .MoveNext
        Loop
       .Close
    End With
NEXTONE:
Next i
Pb.value = 100
Liberado
frmReport.ShowReport "ETIQUETAPROTOCOLO", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdPrint_Click()


If Not IsDate(txtObs2.Text) Then
    MsgBox "Data de vencimento inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

PopupMenu mnuSel


End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()



Centraliza Me
txtObs2.Text = Format(Now, "dd/mm/yyyy")
cmbAjuizado.ListIndex = 0
cmbTipoISS.ListIndex = 0
'txtObs2.text = RetornaDiaUtil(Format(DateAdd("d", 15, Now), "dd/mm/yyyy"))
cmbTipoISS.Clear
cmbTipoISS.AddItem "(Todos os Débitos)"
Sql = "SELECT CODLANCAMENTO,DESCREDUZ FROM LANCAMENTO ORDER BY DESCREDUZ"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipoISS.AddItem (!descreduz)
        cmbTipoISS.ItemData(cmbTipoISS.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With
cmbTipoISS.ListIndex = 0
cmbStatus.ListIndex = 0
cmbSetor.ListIndex = 0
End Sub

Private Sub mnuBoletos_Click()

'MsgBox "Bloqueado até a regularização dos boletos de cobrança", vbCritical, "Aviso!"
'Exit Sub

If Val(txtCod1.Text) = 0 Or Val(txtCod2.Text) = 0 Then
    MsgBox "Digite código inicial e final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) = 0 Or Val(txtAno2.Text) = 0 Then
    MsgBox "Digite Ano inicial e final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtCod1.Text) > Val(txtCod2.Text) Then
    MsgBox "Código inicial maior que final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) > Val(txtAno2.Text) Then
    MsgBox "Código inicial maior que final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) < 1990 Or Val(txtAno1.Text) > Year(Now) Then
    MsgBox "Ano inicial fora de intervalo.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno2.Text) < 1990 Or Val(txtAno2.Text) > Year(Now) Then
    MsgBox "Ano final fora de intervalo.", vbCritical, "Atenção"
    Exit Sub
End If


Ocupado
FillTable2
Liberado

End Sub

Private Sub mnuCarta_Click()
If Val(txtCod1.Text) = 0 Or Val(txtCod2.Text) = 0 Then
    MsgBox "Digite código inicial e final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) = 0 Or Val(txtAno2.Text) = 0 Then
    MsgBox "Digite Ano inicial e final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtCod1.Text) > Val(txtCod2.Text) Then
    MsgBox "Código inicial maior que final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) > Val(txtAno2.Text) Then
    MsgBox "Código inicial maior que final.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno1.Text) < 1990 Or Val(txtAno1.Text) > Year(Now) Then
    MsgBox "Ano inicial fora de intervalo.", vbCritical, "Atenção"
    Exit Sub
End If
If Val(txtAno2.Text) < 1990 Or Val(txtAno2.Text) > Year(Now) Then
    MsgBox "Ano final fora de intervalo.", vbCritical, "Atenção"
    Exit Sub
End If


Ocupado
FillTable
Liberado

End Sub

Private Sub txtAno1_KeyPress(KeyAscii As Integer)
Tweak txtAno1, KeyAscii, IntegerPositive
End Sub

Private Sub txtAno2_KeyPress(KeyAscii As Integer)
Tweak txtAno2, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod1_KeyPress(KeyAscii As Integer)
Tweak txtCod1, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod2_KeyPress(KeyAscii As Integer)
Tweak txtCod2, KeyAscii, IntegerPositive
End Sub

Private Sub FillTable()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery, RdoAux3 As rdoResultset, i As Long
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nCodTrib As Integer, nValorTrib As Double
Dim nValorPrincipal As Double, nValorMulta As Double, nValorJuros As Double, nValorCorrecao As Double, nValorTotal As Double
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String
Dim sBairro As String, sCidade As String, sUF As String, sCep As String, sEnderecoEntrega As String, nNumEntrega As Integer, sBairroEntrega As String
Dim sCidadeEntrega As String, sUFEntrega As String, sCepEntrega As String, aDebito() As Debito
Dim nTot As Long, nPos As Long, aCodigos() As Long, k As Integer, bAchou As Boolean, sObs As String
Pb.value = 0
ReDim aCodigos(0)
If Val(txtCod1.Text) >= 100000 Then
    If cmbTipoISS.ListIndex > 0 Then
        Sql = "SELECT * FROM VWDEBITOISS WHERE CODREDUZIDO BETWEEN " & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text) & " AND ANOEXERCICIO "
        Sql = Sql & "BETWEEN " & Val(txtAno1.Text) & " AND " & Val(txtAno2.Text)
        If cmbTipoISS.ListIndex = 1 Then
            Sql = Sql & " AND CODLANCAMENTO=5"
        ElseIf cmbTipoISS.ListIndex = 2 Then
            Sql = Sql & " AND CODLANCAMENTO=3"
        ElseIf cmbTipoISS.ListIndex = 3 Then
            Sql = Sql & " AND (CODLANCAMENTO=2)"
        End If
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aCodigos(UBound(aCodigos) + 1)
                aCodigos(UBound(aCodigos)) = !CODREDUZIDO
               .MoveNext
            Loop
           .Close
        End With
    End If
    
End If

Sql = "DELETE FROM COBRANCAAMIGAVEL WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM  COBRANCAAMIGAVELDETALHE WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Ocupado

Set qd.ActiveConnection = cn
nTot = Val(txtCod2.Text) - Val(txtCod1.Text) + 1
For i = Val(txtCod1.Text) To Val(txtCod2.Text)
    If i < 100000 Then
        Sql = "select codreduzido,inativo from cadimob where codreduzido=" & i
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If !Inativo Then
                   .Close
                    GoTo NEXTONE
                End If
            End If
           .Close
        End With
    ElseIf i >= 100000 And i < 500000 Then
        Sql = "select codigomob,dataencerramento from mobiliario where codigomob=" & i
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If Not IsNull(!dataencerramento) Then
                   .Close
                    GoTo NEXTONE
                End If
            End If
           .Close
        End With
        Sql = "SELECT codmobiliario, DataEv, codtipoevento From vwMOBILIARIOSUSPENSO Where (codmobiliario=" & i & ") and (codtipoevento = 2)"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                .Close
                 GoTo NEXTONE
            End If
           .Close
        End With
    
    End If


    DoEvents
    nPos = nPos + 1
    If nPos Mod 20 = 0 Then
       CallPb nPos, CLng(nTot)
    End If
    
    ReDim aDebito(0)
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
    qd(0) = i
    qd(1) = i
    qd(2) = Val(txtAno1.Text)
    qd(3) = Val(txtAno2.Text)
    
    If cmbTipoISS.ListIndex = 0 Then
        qd(4) = 0
        qd(5) = 99
    Else
        qd(4) = cmbTipoISS.ItemData(cmbTipoISS.ListIndex)
        qd(5) = cmbTipoISS.ItemData(cmbTipoISS.ListIndex)
    End If
    qd(6) = 0
    qd(7) = 999 'SEQUENCIA
    qd(8) = 1
    qd(9) = 99 'PARCELA
    qd(10) = 0
    qd(11) = 9 'COMPLEMENTO
    qd(12) = cmbStatus.ItemData(cmbStatus.ListIndex)
    qd(13) = cmbStatus.ItemData(cmbStatus.ListIndex) 'STATUSLANC
    qd(14) = Format(txtObs2.Text, "mm/dd/yyyy")
    qd(15) = NomeDoUsuario
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        
        Do Until .EOF
        
                If (!statuslanc <> 3 And !statuslanc <> 19) Then GoTo proximo
                DoEvents
                'Ajuizado
                If cmbAjuizado.ListIndex = 1 Then
                    If IsNull(!dataajuiza) Then GoTo proximo
                ElseIf cmbAjuizado.ListIndex = 2 Then
                    If Not IsNull(!dataajuiza) Then GoTo proximo
                End If
                
                'iss
                If cmbTipoISS.ListIndex >= 0 Then
                    bAchou = False
                    If UBound(aCodigos) > 0 Then
                        For k = 1 To UBound(aCodigos)
                            If !CODREDUZIDO = aCodigos(k) Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                
                'If Year(!DataVencimento) >= Year(Now) Then GoTo proximo
                If !DataVencimento >= Now Then GoTo proximo
                nAno = !AnoExercicio
                Sql = "SELECT CODREDUZIDO FROM COBRANCAAMIGAVEL WHERE CODREDUZIDO=" & !CODREDUZIDO
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                        If .RowCount > 0 Then
                            .Close
                            GoTo DETALHE
                        End If
                End With
                
                
               
                
                If !CODREDUZIDO < 100000 Then
                    Sql = "select distinct * from vwcobrancaamigavel where codimovel =" & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !nomecidadao
                            sInscricao = !Inscricao
                            sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                            nNumero = !Li_Num
                            sComplemento = SubNull(!Li_Compl)
                            sBairro = SubNull(!DescBairro)
                            sCidade = "JABOTICABAL"
                            sUF = "SP"
                            sCep = RetornaCEP(!CodLogr, nNumero)
                            sEnderecoEntrega = IIf(IsNull(!Ee_NomeLog), sEndereco, !Ee_NomeLog)
                            nNumEntrega = IIf(IsNull(!Ee_NumImovel), nNumero, !Ee_NumImovel)
                            sBairroEntrega = Trim$(IIf(IsNull(!Ee_DESCBairro), IIf(IsNull(!BAIRRO2), "", !BAIRRO2), !Ee_DESCBairro))
                            sCidadeEntrega = IIf(IsNull(!descCidade), sCidade, !descCidade)
                            sUFEntrega = IIf(IsNull(!Ee_Uf), sUF, !Ee_Uf)
                            sCep = IIf(IsNull(!Ee_Cep), SubNull(!Cep2), !Ee_Cep)
                            sCepEntrega = sCep
                            If !Ee_TipoEnd = 1 Then
                                sEnderecoEntrega = SubNull(!Endereco)
                                nNumEntrega = Val(SubNull(!NUMIMOVEL))
                                sBairroEntrega = SubNull(!NOMEBAIRRO2)
                                sCidadeEntrega = SubNull(!descCidade)
                                sUFEntrega = SubNull(!SiglaUF)
                                sCep = SubNull(!Cep2)
                            End If
                        End If
                       .Close
                    End With
             ElseIf !CODREDUZIDO >= 100000 And !CODREDUZIDO < 500000 Then
                    Sql = "SELECT mobiliario.codigomob, mobiliario.razaosocial, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,"
                    Sql = Sql & "mobiliario.numero, mobiliario.complemento, bairro.descbairro, cidade.desccidade, mobiliario.siglauf, mobiliario.cep, mobiliario.codlogradouro, "
                    Sql = Sql & "mobiliario.nomelogradouro AS nomelogradouro2 FROM mobiliario INNER JOIN bairro ON mobiliario.siglauf = bairro.siglauf AND mobiliario.codcidade = bairro.codcidade AND "
                    Sql = Sql & "mobiliario.codbairro = bairro.codbairro INNER JOIN cidade ON bairro.siglauf = cidade.siglauf AND bairro.codcidade = cidade.codcidade LEFT OUTER JOIN "
                    'Sql = Sql & "vwLOGRADOURO ON mobiliario.codlogradouro = vwLOGRADOURO.CODLOGRADOURO Where mobiliario.codigomob = " & !CODREDUZIDO & " and mobiliario.dataencerramento is null and "
                    Sql = Sql & "vwLOGRADOURO ON mobiliario.codlogradouro = vwLOGRADOURO.CODLOGRADOURO Where mobiliario.codigomob = " & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !RazaoSocial
                            sInscricao = !codigomob
                            If Val(SubNull(!CodLogradouro)) > 0 Then
                                sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                            Else
                                sEndereco = SubNull(!NOMELOGRADOURO2)
                            End If
                            nNumero = !Numero
                            sComplemento = SubNull(!Complemento)
                            sBairro = !DescBairro
                            sCidade = !descCidade
                            sUF = !SiglaUF
                            sCep = SubNull(!Cep)
                            sEnderecoEntrega = sEndereco
                            nNumEntrega = nNumero
                            sBairroEntrega = sBairro
                            sCidadeEntrega = sCidade
                            sUFEntrega = sUF
                            sCepEntrega = sCep
                        Else
                            GoTo proximo
                        End If
                       .Close
                    End With
             Else
                    Sql = "select distinct * from vwcobrancaamigavelcidadao where codcidadao =" & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !nomecidadao
                            sInscricao = !CodCidadao
                            If Not IsNull(!NomeLogradouro) Then
                                sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                sCep = RetornaCEP(!CodLogradouro, nNumero)
                            Else
                                sEndereco = SubNull(!NOMELOGRADOURO2)
                                sCep = SubNull(!Cep)
                            End If
                            nNumero = Val(SubNull(!NUMIMOVEL))
                            sComplemento = SubNull(!Complemento)
                            sBairro = SubNull(!DescBairro)
                            sCidade = SubNull(!descCidade)
                            sUF = SubNull(!SiglaUF)
                            
                            sEnderecoEntrega = sEndereco
                            nNumEntrega = nNumero
                            sBairroEntrega = sBairro
                            sCidadeEntrega = sCidade
                            sUFEntrega = sUF
                            sCepEntrega = sCep
                        Else
                            sNome = ""
                            sInscricao = ""
                            sEndereco = ""
                            nNumero = 0
                            sComplemento = ""
                            sBairro = ""
                            sCidade = ""
                            sUF = ""
                            sCep = ""
                            sEnderecoEntrega = ""
                            nNumEntrega = 0
                            sBairroEntrega = ""
                            sCidadeEntrega = ""
                            sUFEntrega = ""
                            sCep = ""
                        End If
                       .Close
                    End With
             End If
             
             Sql = "INSERT COBRANCAAMIGAVEL (USUARIO,CODREDUZIDO,NOMECIDADAO,INSCRICAO,ENDERECO,NUMERO,COMPLEMENTO,"
             Sql = Sql & "BAIRRO,CIDADE,UF,ENDERECOENTREGA,NUMEROENTREGA,BAIRROENTREGA,CIDADEENTREGA,"
             Sql = Sql & "UFENTREGA,CEPENTREGA,DATACALCULO) VALUES('" & NomeDeLogin & "'," & RdoAux!CODREDUZIDO & ",'" & Mask(Left$(sNome, 50)) & "','" & sInscricao & "','" & Left(Mask(sEndereco), 50) & "',"
             Sql = Sql & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','" & Mask(sBairro) & "','" & sCidade & "','" & sUF & "','" & Mask(Left(sEnderecoEntrega, 50)) & "',"
             Sql = Sql & nNumEntrega & ",'" & Mask(sBairroEntrega) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & sCepEntrega & "','" & Format(txtObs2.Text, "mm/dd/yyyy") & "')"
             cn.Execute Sql, rdExecDirect
             
DETALHE:
         
            nEval = UBound(aDebito)
            nEval = UBound(aDebito)
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nCodReduz = !CODREDUZIDO And aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
    
            If Not Achou Then
               ReDim Preserve aDebito(UBound(aDebito) + 1)
               nEval = UBound(aDebito)
               aDebito(nEval).nCodReduz = !CODREDUZIDO
               aDebito(nEval).nAno = !AnoExercicio
               aDebito(nEval).nLanc = !CodLancamento
               aDebito(nEval).sLanc = !DESCLANCAMENTO
               aDebito(nEval).nSeq = !SeqLancamento
               aDebito(nEval).nParc = !NumParcela
               aDebito(nEval).nCompl = !CODCOMPLEMENTO
               aDebito(nEval).nSituacao = !statuslanc
               aDebito(nEval).sSituacao = !Situacao
               aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
               If !statuslanc = 3 Or !statuslanc = 19 Then
                  If Not IsNull(!ValorTotal) Then
                    aDebito(nEval).nValorAtual = !ValorTotal
                  Else
                    aDebito(nEval).nValorAtual = 0
                  End If
               End If
               aDebito(nEval).nValorJuros = !ValorJuros
               aDebito(nEval).nValorMulta = !ValorMulta
               aDebito(nEval).nValorCorrecao = !ValorCorrecao
               aDebito(nEval).nValorTributo = !ValorTributo
            Else
               aDebito(nEval).nValorAtual = aDebito(nEval).nValorAtual + !ValorTotal
               aDebito(nEval).nValorJuros = aDebito(nEval).nValorJuros + !ValorJuros
               aDebito(nEval).nValorMulta = aDebito(nEval).nValorMulta + !ValorMulta
               aDebito(nEval).nValorCorrecao = aDebito(nEval).nValorCorrecao + !ValorCorrecao
               aDebito(nEval).nValorTributo = aDebito(nEval).nValorTributo + !ValorTributo
            End If
proximo:
            .MoveNext
        Loop
       .Close
    End With
    
    For x = 1 To UBound(aDebito)
            With aDebito(x)
                sObs = ""
                Sql = "SELECT * FROM OBSPARCELA WHERE CODREDUZIDO=" & .nCodReduz & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & .nSeq & " AND NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & .nCompl & " ORDER BY SEQ"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sObs = Left(!obs, 200)
                       .Close
                    End If
                End With
                If .nValorAtual > 0 Then
                    Sql = "INSERT COBRANCAAMIGAVELDETALHE(USUARIO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,"
                    Sql = Sql & "CODCOMPLEMENTO,VALORPRINCIPAL,VALORJUROS,VALORMULTA,VALORCORRECAO,VALORTOTAL,DATAVENCIMENTO,"
                    Sql = Sql & "DESCLANCAMENTO,OBS) VALUES('" & NomeDeLogin & "'," & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ","
                    Sql = Sql & Virg2Ponto(CStr(.nValorTributo)) & "," & Virg2Ponto(CStr(.nValorJuros)) & "," & Virg2Ponto(CStr(.nValorMulta)) & "," & Virg2Ponto(CStr(.nValorCorrecao)) & ","
                    Sql = Sql & Virg2Ponto(CStr(.nValorAtual)) & ",'" & Format(.sVencto, "mm/dd/yyyy") & "','" & .sLanc & "','" & sObs & "')"
                    cn.Execute Sql, rdExecDirect
                End If
            End With
    Next
NEXTONE:
Next i
Liberado

Sql = "SELECT * FROM COBRANCAAMIGAVEL WHERE USUARIO='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If frmMdi.frTeste.Visible = True Then
            frmReport.ShowReport "COBRANCAAMIGAVELTMP", frmMdi.hwnd, Me.hwnd '
        Else
            If chkSuj.value = vbUnchecked Then
                If bAnistia Then
                    frmReport.ShowReport "COBRANCAAMIGAVELDA", frmMdi.hwnd, Me.hwnd '
                Else
                    If cmbSetor.ListIndex = 0 Then
                        frmReport.ShowReport "COBRANCAAMIGAVELDA", frmMdi.hwnd, Me.hwnd '
                    Else
                        frmReport.ShowReport "COBRANCAAMIGAVELDA2", frmMdi.hwnd, Me.hwnd '
                    End If
                End If
            Else
                frmReport.ShowReport "COBRANCAAMIGAVELSUJEITO", frmMdi.hwnd, Me.hwnd '
            End If
        End If
        Sql = "DELETE FROM COBRANCAAMIGAVEL WHERE USUARIO='" & NomeDeLogin & "'"
        cn.Execute Sql, rdExecDirect
        Sql = "DELETE FROM  COBRANCAAMIGAVELDETALHE WHERE USUARIO='" & NomeDeLogin & "'"
        cn.Execute Sql, rdExecDirect
   Else
        MsgBox "Não existem débitos com as informações acima solicitadas.", vbExclamation, "Atenção"
   End If
   .Close
End With

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub FillTable2()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery, RdoAux3 As rdoResultset, i As Long
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nCodTrib As Integer, nValorTrib As Double
Dim nValorPrincipal As Double, nValorMulta As Double, nValorJuros As Double, nValorCorrecao As Double, nValorTotal As Double
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String, nCodReduz As Long
Dim sBairro As String, sCidade As String, sUF As String, sCep As String, sEnderecoEntrega As String, nNumEntrega As Integer, sBairroEntrega As String
Dim sCidadeEntrega As String, sUFEntrega As String, sCepEntrega As String, aDebito() As Debito, sQuadras As String, sLotes As String
Dim nTot As Long, nPos As Long, aCodigos() As Long, k As Integer, bAchou As Boolean, sCPFCNPJ As String
Dim nNumDoc As Long, bMulta As Boolean, sNumDoc As String, sLanc As String, sFullTrib As String, sDataVencto As String
Dim nSeq2 As Integer, sAj As String, sDA As String, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, nValorDoc As Double
Dim nSid As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date
Dim sBarra As String, sDigitavel2 As String, nValorDam As Double, nValorPrincDam As Double, nNumGuia As Long, sTipoEnd As String

Pb.value = 0
ReDim aCodigos(0)
If Val(txtCod1.Text) >= 100000 Then
    If cmbTipoISS.ListIndex > 0 Then
        Sql = "SELECT * FROM VWDEBITOISS WHERE CODREDUZIDO BETWEEN " & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text) & " AND ANOEXERCICIO "
        Sql = Sql & "BETWEEN " & Val(txtAno1.Text) & " AND " & Val(txtAno2.Text)
        If cmbTipoISS.ListIndex = 1 Then
            Sql = Sql & " AND CODLANCAMENTO=5"
        ElseIf cmbTipoISS.ListIndex = 2 Then
            Sql = Sql & " AND CODLANCAMENTO=3"
        ElseIf cmbTipoISS.ListIndex = 3 Then
            Sql = Sql & " AND (CODLANCAMENTO=2)"
        End If
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aCodigos(UBound(aCodigos) + 1)
                aCodigos(UBound(aCodigos)) = !CODREDUZIDO
               .MoveNext
            Loop
           .Close
        End With
    End If
    
End If

Sql = "DELETE FROM COBRANCAAMIGAVEL WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM  COBRANCAAMIGAVELDETALHE WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Ocupado

Set qd.ActiveConnection = cn
nTot = Val(txtCod2.Text) - Val(txtCod1.Text) + 1
For i = Val(txtCod1.Text) To Val(txtCod2.Text)
    If i < 100000 Then
        Sql = "select codreduzido,inativo from cadimob where codreduzido=" & i
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If !Inativo Then
                   .Close
                    GoTo NEXTONE
                End If
            End If
           .Close
        End With
    End If


    DoEvents
    nPos = nPos + 1
    If nPos Mod 20 = 0 Then
       CallPb nPos, CLng(nTot)
    End If
    
    ReDim aDebito(0)
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
    qd(0) = i
    qd(1) = i
    qd(2) = Val(txtAno1.Text)
    qd(3) = Val(txtAno2.Text)
    
    If cmbTipoISS.ListIndex = 0 Then
        qd(4) = 0
        qd(5) = 99
    Else
        qd(4) = cmbTipoISS.ItemData(cmbTipoISS.ListIndex)
        qd(5) = cmbTipoISS.ItemData(cmbTipoISS.ListIndex)
    End If
    qd(6) = 0
    qd(7) = 999 'SEQUENCIA
    qd(8) = 1
    qd(9) = 99 'PARCELA
    qd(10) = 0
    qd(11) = 9 'COMPLEMENTO
    qd(12) = cmbStatus.ItemData(cmbStatus.ListIndex)
    qd(13) = cmbStatus.ItemData(cmbStatus.ListIndex) 'STATUSLANC
    qd(14) = Format(txtObs2.Text, "mm/dd/yyyy")
    qd(15) = NomeDoUsuario
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        
        Do Until .EOF
        
                If (!statuslanc <> 3 And !statuslanc <> 19) Then GoTo proximo
                DoEvents
                'Ajuizado
                If cmbAjuizado.ListIndex = 1 Then
                    If IsNull(!dataajuiza) Then GoTo proximo
                ElseIf cmbAjuizado.ListIndex = 2 Then
                    If Not IsNull(!dataajuiza) Then GoTo proximo
                End If
                'If Year(!DataVencimento) >= Year(Now) Then GoTo proximo
                If !DataVencimento >= Now Then GoTo proximo
                'iss
                If cmbTipoISS.ListIndex > 0 Then
                    bAchou = False
                    If UBound(aCodigos) > 0 Then
                        For k = 1 To UBound(aCodigos)
                            If !CODREDUZIDO = aCodigos(k) Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                
                nAno = !AnoExercicio
                Sql = "SELECT CODREDUZIDO FROM COBRANCAAMIGAVEL WHERE CODREDUZIDO=" & !CODREDUZIDO
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                        If .RowCount > 0 Then
                            .Close
                            GoTo DETALHE
                        End If
                End With
                If !CODREDUZIDO < 100000 Then
                    Sql = "select distinct * from vwcobrancaamigavel where codimovel =" & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !nomecidadao
                            sInscricao = !Inscricao
                            sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                            nNumero = !Li_Num
                            sComplemento = SubNull(!Li_Compl)
                            sBairro = SubNull(!DescBairro)
                            sCidade = "JABOTICABAL"
                            sUF = "SP"
                            sCep = RetornaCEP(!CodLogr, nNumero)
                            sEnderecoEntrega = IIf(IsNull(!Ee_NomeLog), sEndereco, !Ee_NomeLog)
                            nNumEntrega = IIf(IsNull(!Ee_NumImovel), nNumero, !Ee_NumImovel)
                            sBairroEntrega = Trim$(IIf(IsNull(!Ee_DESCBairro), IIf(IsNull(!BAIRRO2), "", !BAIRRO2), !Ee_DESCBairro))
                            sCidadeEntrega = IIf(IsNull(!descCidade), sCidade, !descCidade)
                            sUFEntrega = IIf(IsNull(!Ee_Uf), sUF, !Ee_Uf)
                            sCep = IIf(IsNull(!Ee_Cep), SubNull(!Cep2), !Ee_Cep)
                            sCepEntrega = sCep
                            If !Ee_TipoEnd = 1 Then
                                sEnderecoEntrega = SubNull(!Endereco)
                                nNumEntrega = Val(SubNull(!NUMIMOVEL))
                                sBairroEntrega = SubNull(!NOMEBAIRRO2)
                                sCidadeEntrega = SubNull(!descCidade)
                                sUFEntrega = SubNull(!SiglaUF)
                                sCep = SubNull(!Cep2)
                            End If
                        End If
                       .Close
                    End With
             ElseIf !CODREDUZIDO >= 100000 And !CODREDUZIDO < 500000 Then
                    Sql = "SELECT mobiliario.codigomob, mobiliario.razaosocial, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,"
                    Sql = Sql & "mobiliario.numero, mobiliario.complemento, bairro.descbairro, cidade.desccidade, mobiliario.siglauf, mobiliario.cep, mobiliario.codlogradouro, "
                    Sql = Sql & "mobiliario.nomelogradouro AS nomelogradouro2 FROM mobiliario INNER JOIN bairro ON mobiliario.siglauf = bairro.siglauf AND mobiliario.codcidade = bairro.codcidade AND "
                    Sql = Sql & "mobiliario.codbairro = bairro.codbairro INNER JOIN cidade ON bairro.siglauf = cidade.siglauf AND bairro.codcidade = cidade.codcidade LEFT OUTER JOIN "
                    Sql = Sql & "vwLOGRADOURO ON mobiliario.codlogradouro = vwLOGRADOURO.CODLOGRADOURO Where mobiliario.codigomob = " & !CODREDUZIDO & " and mobiliario.dataencerramento is null and "
                    Sql = Sql & "MOBILIARIO.CODIGOMOB NOT in (SELECT CODMOBILIARIO FROM vwMOBILIARIOSUSPENSO WHERE CODTIPOEVENTO=2)"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !RazaoSocial
                            sInscricao = !codigomob
                            If Val(SubNull(!CodLogradouro)) > 0 Then
                                sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                            Else
                                sEndereco = SubNull(!NOMELOGRADOURO2)
                            End If
                            nNumero = !Numero
                            sComplemento = SubNull(!Complemento)
                            sBairro = !DescBairro
                            sCidade = !descCidade
                            sUF = !SiglaUF
                            sCep = SubNull(!Cep)
                            sEnderecoEntrega = sEndereco
                            nNumEntrega = nNumero
                            sBairroEntrega = sBairro
                            sCidadeEntrega = sCidade
                            sUFEntrega = sUF
                            sCepEntrega = sCep
                        Else
                            GoTo proximo
                        End If
                       .Close
                    End With
             Else
                    Sql = "select distinct * from vwcobrancaamigavelcidadao where codcidadao =" & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNome = !nomecidadao
                            sInscricao = !CodCidadao
                            If Not IsNull(!NomeLogradouro) Then
                                sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                sCep = RetornaCEP(!CodLogradouro, nNumero)
                            Else
                                sEndereco = SubNull(!NOMELOGRADOURO2)
                                sCep = SubNull(!Cep)
                            End If
                            nNumero = Val(SubNull(!NUMIMOVEL))
                            sComplemento = SubNull(!Complemento)
                            sBairro = SubNull(!DescBairro)
                            sCidade = SubNull(!descCidade)
                            sUF = SubNull(!SiglaUF)
                            
                            sEnderecoEntrega = sEndereco
                            nNumEntrega = nNumero
                            sBairroEntrega = sBairro
                            sCidadeEntrega = sCidade
                            sUFEntrega = sUF
                            sCepEntrega = sCep
                        Else
                            sNome = ""
                            sInscricao = ""
                            sEndereco = ""
                            nNumero = 0
                            sComplemento = ""
                            sBairro = ""
                            sCidade = ""
                            sUF = ""
                            sCep = ""
                            sEnderecoEntrega = ""
                            nNumEntrega = 0
                            sBairroEntrega = ""
                            sCidadeEntrega = ""
                            sUFEntrega = ""
                            sCep = ""
                        End If
                       .Close
                    End With
             End If
             
             Sql = "INSERT COBRANCAAMIGAVEL (USUARIO,CODREDUZIDO,NOMECIDADAO,INSCRICAO,ENDERECO,NUMERO,COMPLEMENTO,"
             Sql = Sql & "BAIRRO,CIDADE,UF,ENDERECOENTREGA,NUMEROENTREGA,BAIRROENTREGA,CIDADEENTREGA,"
             Sql = Sql & "UFENTREGA,CEPENTREGA,DATACALCULO) VALUES('" & NomeDeLogin & "'," & RdoAux!CODREDUZIDO & ",'" & Mask(Left$(sNome, 50)) & "','" & sInscricao & "','" & Left(Mask(sEndereco), 50) & "',"
             Sql = Sql & nNumero & ",'" & Mask(Left(sComplemento, 50)) & "','" & Mask(sBairro) & "','" & sCidade & "','" & sUF & "','" & Mask(Left(sEnderecoEntrega, 50)) & "',"
             Sql = Sql & nNumEntrega & ",'" & Mask(sBairroEntrega) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & sCepEntrega & "','" & Format(txtObs2.Text, "mm/dd/yyyy") & "')"
             cn.Execute Sql, rdExecDirect
             
DETALHE:
         
            nEval = UBound(aDebito)
            nEval = UBound(aDebito)
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nCodReduz = !CODREDUZIDO And aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
    
            If Not Achou Then
               ReDim Preserve aDebito(UBound(aDebito) + 1)
               nEval = UBound(aDebito)
               aDebito(nEval).nCodReduz = !CODREDUZIDO
               aDebito(nEval).nAno = !AnoExercicio
               aDebito(nEval).nLanc = !CodLancamento
               aDebito(nEval).sLanc = !DESCLANCAMENTO
               aDebito(nEval).nSeq = !SeqLancamento
               aDebito(nEval).nParc = !NumParcela
               aDebito(nEval).nCompl = !CODCOMPLEMENTO
               aDebito(nEval).nSituacao = !statuslanc
               aDebito(nEval).sSituacao = !Situacao
               aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
               If !statuslanc = 3 Or !statuslanc = 19 Then
                  If Not IsNull(!ValorTotal) Then
                    'aDebito(nEval).nValorAtual = !ValorTotal
                    aDebito(nEval).nValorAtual = !ValorTributo + !ValorCorrecao
                  Else
                    aDebito(nEval).nValorAtual = 0
                  End If
               End If
               '**** voltar multa E juros para normal ****
               
               aDebito(nEval).nValorJuros = !ValorJuros - (!ValorJuros * 0.8)
               aDebito(nEval).nValorMulta = !ValorMulta - (!ValorMulta * 0.8)
'               aDebito(nEval).nValorJuros = 0
'               aDebito(nEval).nValorMulta = 0
               aDebito(nEval).nValorCorrecao = !ValorCorrecao
               aDebito(nEval).nValorTributo = !ValorTributo
            Else
 '               aDebito(nEval).nValorAtual = aDebito(nEval).nValorAtual + !ValorTotal
'               aDebito(nEval).nValorJuros = aDebito(nEval).nValorJuros + !ValorJuros
'               aDebito(nEval).nValorMulta = aDebito(nEval).nValorMulta + !ValorMulta
               aDebito(nEval).nValorJuros = 0
               aDebito(nEval).nValorMulta = 0
               aDebito(nEval).nValorCorrecao = aDebito(nEval).nValorCorrecao + !ValorCorrecao
               aDebito(nEval).nValorTributo = aDebito(nEval).nValorTributo + !ValorTributo
               aDebito(nEval).nValorAtual = aDebito(nEval).nValorTributo + aDebito(nEval).nValorCorrecao
            End If
proximo:
            .MoveNext
        Loop
       .Close
    End With
    
    For x = 1 To UBound(aDebito)
        With aDebito(x)
            If .nValorAtual > 0 Then
                Sql = "INSERT COBRANCAAMIGAVELDETALHE(USUARIO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,"
                Sql = Sql & "CODCOMPLEMENTO,VALORPRINCIPAL,VALORJUROS,VALORMULTA,VALORCORRECAO,VALORTOTAL,DATAVENCIMENTO,"
                Sql = Sql & "DESCLANCAMENTO) VALUES('" & NomeDeLogin & "'," & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ","
                Sql = Sql & Virg2Ponto(CStr(.nValorTributo)) & "," & Virg2Ponto(CStr(.nValorJuros)) & "," & Virg2Ponto(CStr(.nValorMulta)) & "," & Virg2Ponto(CStr(.nValorCorrecao)) & ","
                Sql = Sql & Virg2Ponto(CStr(.nValorAtual)) & ",'" & Format(.sVencto, "mm/dd/yyyy") & "','" & .sLanc & "')"
                cn.Execute Sql, rdExecDirect
            End If
        End With
    Next
NEXTONE:
Next i
Liberado


Sql = "SELECT * FROM COBRANCAAMIGAVEL WHERE USUARIO='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        nSid = Int(Rnd(100) * 1000000)
        
        Sql = "delete from boleto where sid=" & nSid
        cn.Execute Sql, rdExecDirect
    
        sDataVencto = txtObs2.Text 'mudar
        nPos = 1
        Do Until .EOF
            Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If IsNull(RdoAux2!maximo) Then
                nNumDoc = 0
            Else
                nNumDoc = RdoAux2!maximo + 1
            End If
            RdoAux2.Close
                        
            Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO,emissor) VALUES("
            Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0,0,0,0,'" & NomeDeLogin & " (COBRANÇA AMIGÁVEL)" & "')"
            cn.Execute Sql, rdExecDirect
            
            sNumDoc = CStr(nNumDoc) & "-" & RetornaDVNumDoc(nNumDoc)
            sNumDoc2 = CStr(nNumDoc) & RetornaDVNumDoc(nNumDoc)
            sNumDoc3 = CStr(nNumDoc) & Modulo11(nNumDoc)
            
            nValorDoc = 0
            nValorPrincDam = 0
            nCodReduz = !CODREDUZIDO
            
            Dim nPlano As Integer
            nPlano = 16
            
            Sql = "SELECT * FROM COBRANCAAMIGAVELDETALHE WHERE USUARIO='" & NomeDeLogin & "' AND CODREDUZIDO=" & !CODREDUZIDO
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                Do Until .EOF
                    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                    Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,VALORJUROS,VALORMULTA,VALORCORRECAO,PLANO) VALUES(" & !CODREDUZIDO & ","
                    Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nNumDoc & ","
                    Sql = Sql & Virg2Ponto(CStr(!ValorJuros)) & "," & Virg2Ponto(CStr(!ValorMulta)) & "," & Virg2Ponto(CStr(!ValorCorrecao)) & "," & nPlano & ")"
                    cn.Execute Sql, rdExecDirect
                    
                    Sql = "insert boleto(usuario,computer,sid,seq,inscricao,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,quadra,lote,numdoc,nomefunc,datadam,fulllanc,fulltrib,"
                    Sql = Sql & "anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datavencto,aj,da,principal,juros,multa,correcao,total,numdoc2,valordam) values('"
                    Sql = Sql & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nPos & ",'" & RdoAux!Inscricao & "'," & !CODREDUZIDO & ",'" & Left(Mask(RdoAux!nomecidadao), 40) & "','" & sCPFCNPJ & "','"
                    Sql = Sql & Left(Mask(RdoAux!Endereco), 40) & "'," & RdoAux!Numero & ",'" & Left(Mask(RdoAux!Complemento), 30) & "','" & Left(Mask(RdoAux!Bairro), 25) & "','" & Mask(RdoAux!Cidade) & "','" & RdoAux!UF & "','" & Mask(sQuadras) & "','"
                    Sql = Sql & Mask(sLotes) & "','" & sNumDoc & "','" & NomeDeLogin & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','BOLETO EMITIDO COM 100% DESCONTO NA MULTA E JUROS - REFIS 2017 LC Nº185/2017','" & !DESCLANCAMENTO & "'," & !AnoExercicio & ","
                    'Sql = Sql & Mask(sLotes) & "','" & sNumDoc & "','" & NomeDeLogin & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','AVISO DE DÉBITO - COBRANÇA AMIGÁVEL','" & !DESCLANCAMENTO & "'," & !AnoExercicio & ","
                    Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','',''," & Virg2Ponto(Format(!ValorPrincipal, "#0.00")) & ","
                    Sql = Sql & Virg2Ponto(Format(!ValorJuros, "#0.00")) & "," & Virg2Ponto(Format(!ValorMulta, "#0.00")) & "," & Virg2Ponto(Format(!ValorCorrecao, "#0.00")) & "," & Virg2Ponto(Format(!ValorTotal, "#0.00")) & ",'" & sNumDoc2
                    Sql = Sql & "'," & 0 & ")"
                    cn.Execute Sql, rdExecDirect
                    
                    nValorDoc = nValorDoc + Format(!ValorTotal, "#0.00")
                    nValorPrincDam = nValorPrincDam + !ValorPrincipal
                    nPos = nPos + 1
                   .MoveNext
                Loop
               .Close
            End With
            
'            sNossoNumero = Format(sNumDoc3, "0000000000000")
'            sDv = Trim(Calculo_DV10("028" & Left(sNossoNumero, 7)))
'            sDigitavel = "0339912354028" & Left(sNossoNumero, 7) & sDv
'
'            sDigitavel = sDigitavel & Right(sNossoNumero, 6) & "0102"
'            sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'            sDigitavel = sDigitavel & sDv
'
'            dDataBase = "07/10/1997"
'            nFatorVencto = CDate(sDataVencto) - dDataBase
'            sQuintoGrupo = Format(nFatorVencto, "0000")
'            sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
'
 '           sBarra = "0339" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "91235028"
'            sBarra = sBarra & sNossoNumero & "0102"
 '           sDv = Trim(Calculo_DV11(sBarra))
'            sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
'
'            sDigitavel = sDigitavel & sDv & sQuintoGrupo
            
'            sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
'            sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
'            sBarra = Gera2of5Str(sBarra)
            
'            Sql = "update boleto set digitavel='" & sDigitavel2 & "',codbarra='" & Mask(sBarra) & "',valordam=" & Virg2Ponto(RemovePonto(Format(nValorDoc, "#0.00"))) & ",valorprincdam=" & Virg2Ponto(RemovePonto(Format(nValorDoc, "#0.00"))) & " where sid=" & nSid & " and codreduzido=" & nCodReduz
'            cn.Execute Sql, rdExecDirect
            
            sNossoNumero = "2678478"
            sDigitavel = "001900000"
            sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
            sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
            sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
            sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
            sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
            sDigitavel = sDigitavel & sDv
            sDataDam = sDataVencto
            dDataBase = "07/10/1997"
            nFatorVencto = CDate(sDataDam) - dDataBase
            sQuintoGrupo = Format(nFatorVencto, "0000")
            sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
            sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "00000026784780"
            sBarra = sBarra & sNumDoc3 & "18"
            sDv = Trim(Calculo_DV11(sBarra))
            sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
            
            sDigitavel = sDigitavel & sDv & sQuintoGrupo
            
            sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
            sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
            sBarra = Gera2of5Str(sBarra)
            
            Dim sValor As String, dDataVencto As Date, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
            '****************
            sValor = nValorDoc
            dDataVencto = CDate(sDataDam)
            NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumDoc, nCodReduz)
            NumBarra2a = Left$(NumBarra2, 13)
            NumBarra2b = Mid$(NumBarra2, 14, 13)
            NumBarra2c = Mid$(NumBarra2, 27, 13)
            NumBarra2d = Right$(NumBarra2, 13)
        
            StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
            
            '***************
            Sql = "update boleto set numbarra2a='" & NumBarra2a & "',numbarra2b='" & NumBarra2b & "',numbarra2c='" & NumBarra2c & "',numbarra2d='" & NumBarra2d & "',digitavel='" & sDigitavel2 & "',codbarra='" & Mask(sBarra) & "',valordam=" & Virg2Ponto(RemovePonto(Format(nValorDoc, "#0.00"))) & " where sid=" & nSid & " and codreduzido=" & nCodReduz
            'Sql = "update boleto set digitavel='" & sDigitavel2 & "',codbarra='" & Mask(sBarra) & "',valordam=" & Virg2Ponto(RemovePonto(Format(nValorDoc, "#0.00"))) & " where sid=" & nSid & " and codreduzido=" & nCodReduz
            cn.Execute Sql, rdExecDirect
           
           .MoveNext 'proximo boleto
        Loop
        
        frmReport.ShowReport2 "BOLETOCOBRANCA_v4", frmMdi.hwnd, Me.hwnd, nSid, nSid

        Sql = "delete from boleto where sid=" & nSid
        cn.Execute Sql, rdExecDirect
        Sql = "delete from boleto where usuario='ROSE'"
        cn.Execute Sql, rdExecDirect
        Sql = "DELETE FROM COBRANCAAMIGAVEL WHERE USUARIO='" & NomeDeLogin & "'"
        cn.Execute Sql, rdExecDirect
        Sql = "DELETE FROM  COBRANCAAMIGAVELDETALHE WHERE USUARIO='" & NomeDeLogin & "'"
        cn.Execute Sql, rdExecDirect
   Else
        MsgBox "Não existem débitos com as informações acima solicitadas.", vbExclamation, "Atenção"
   End If
   .Close
End With

End Sub


