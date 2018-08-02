VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDadosImovel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhes do Imóvel"
   ClientHeight    =   6240
   ClientLeft      =   4365
   ClientTop       =   2340
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   9570
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   8370
      TabIndex        =   12
      ToolTipText     =   "Gravar foto do imóvel"
      Top             =   1350
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
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
      MICON           =   "frmDadosImovel.frx":0000
      PICN            =   "frmDadosImovel.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   5160
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton uFoto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7740
      Picture         =   "frmDadosImovel.frx":03BD
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Próxima Foto"
      Top             =   1950
      Width           =   375
   End
   Begin VB.CommandButton pFoto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6270
      Picture         =   "frmDadosImovel.frx":0507
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Foto Anterior"
      Top             =   1950
      Width           =   375
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6630
      MaxLength       =   6
      TabIndex        =   0
      Top             =   150
      Width           =   1305
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   1350
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmDadosImovel.frx":0651
      PICN            =   "frmDadosImovel.frx":066D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdLoad 
      Height          =   315
      Left            =   6030
      TabIndex        =   2
      ToolTipText     =   "Carrega detalhes do imóvel"
      Top             =   1350
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Carregar"
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
      MICON           =   "frmDadosImovel.frx":06DB
      PICN            =   "frmDadosImovel.frx":06F7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   4860
      TabIndex        =   3
      ToolTipText     =   "ox"
      Top             =   1350
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmDadosImovel.frx":0765
      PICN            =   "frmDadosImovel.frx":0781
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   6105
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   10769
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmDadosImovel.frx":08DB
   End
   Begin VB.Image img 
      DataSource      =   "Msrdc"
      Height          =   3435
      Left            =   4770
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   4755
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "de"
      Height          =   195
      Left            =   7020
      TabIndex        =   10
      Top             =   2010
      Width           =   255
   End
   Begin VB.Label lblFotoAte 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   7290
      TabIndex        =   9
      Top             =   2010
      Width           =   255
   End
   Begin VB.Label lblFotoDe 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   6750
      TabIndex        =   8
      Top             =   2010
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido.:"
      Height          =   195
      Index           =   1
      Left            =   5220
      TabIndex        =   4
      Top             =   210
      Width           =   1320
   End
End
Attribute VB_Name = "frmDadosImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xImovel As clsImovel
Dim sSubPath As String, Rss As String
'PARAMETROS
Dim nUfirCalc As Double
Dim nUfir1999 As Double
Dim nAliquotaPredial As Double
Dim nAliquotaTerritorial As Double
Dim bTemPredial As Boolean
Dim bFracaoIdeal As Boolean
Dim nAreaTerreno As Double
Dim nAreaPrincipal As Double
Dim nCodAgrupamento As Integer
Dim nValorAgrupamento As Double
Dim nValorAgrupamento98 As Double
Dim nNumTestadas As Integer
Dim nTestadaPrincipal As Double
Dim nCodGleba As Integer
Dim nFatorGleba As Double
Dim nFatorGleba98 As Double
Dim nCodProfundidade As Integer
Dim nValorProfundidade As Double
Dim nFatorProfundidade As Double
Dim nFatorProfundidade98 As Double
Dim nCodSituacao As Integer
Dim nFatorSituacao As Double
Dim nFatorSituacao98 As Double
Dim nCodPedologia As Integer
Dim nFatorPedologia As Double
Dim nFatorPedologia98 As Double
Dim nCodTopografia As Integer
Dim nFatorTopografia As Double
Dim nFatorTopografia98 As Double
Dim nFatorDistrito As Double
Dim nFatorDistrito98 As Double
Dim nValorFatores As Double
Dim nValorFatores98 As Double
Dim nFatorCategoria As Double
Dim nFatorCategoria98 As Double
Dim nValorVenalTerritorial As Double
Dim nValorVenalTerritorial98 As Double
Dim nValorVenalPredial As Double
Dim nValorVenalPredial98 As Double
Dim nCodTributo As Integer
Dim nValorVenalImovel As Double
Dim nValorVenalImovel98 As Double
Dim nValorIptu As Double, nValorITU As Double
Dim nValorIPTU98 As Double, nValorITU98 As Double
Dim nTaxaLimpeza As Double, nTaxaConservacao As Double
Dim nValorFinal As Double
'GERAL
Dim nCodReduz As Long
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim Sql As String
Dim nAnoCalculo As Integer
'TIPOS
Private Type PROFUNDIDADE
    Distrito As Integer
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORPROFUN
    Distrito As Integer
    Codigo As Integer
    Fator As Double
End Type
Private Type GLEBA
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORCATEG
    Uso As Integer
    Tipo As Integer
    Categoria As Integer
    Fator As Double
End Type

'MATRIZES
Dim aFatorD() As Double
Dim aFatorD98() As Double
Dim aFatorP() As Double
Dim aFatorP98() As Double
Dim aFatorT() As Double
Dim aFatorT98() As Double
Dim aFatorS() As Double
Dim aFatorS98() As Double
Dim aFatorG() As Double
Dim aFatorG98() As Double
Dim aFatorR() As Double
Dim aFatorR98() As Double
Dim aProf() As PROFUNDIDADE
Dim aFatorF() As FATORPROFUN
Dim aFatorF98() As FATORPROFUN
Dim aFatorC() As FATORCATEG
Dim aFatorC98() As FATORCATEG
Dim aGleba() As GLEBA
Dim sEnd As String


Private Sub cmdGravar_Click()
Dim z As Long
If File1.FileName = "" Then
    MsgBox "Não existe foto para este imóvel.", vbCritical, "Atenção"
Else
    SavePicture img.Picture, sPathBin & "\" & sEnd & ".jpg"
    MsgBox "Foto gravada em " & sPathBin & "\" & sEnd & ".jpg"
End If
End Sub

Private Sub cmdLoad_Click()

Rtb.Text = ""
img.Picture = Nothing
nCodReduz = Val(txtCod.Text)
Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        Rtb.Font.Size = 14
        img.Picture = LoadPicture()
        Rtb.SelAlignment = rtfCenter
        Negrito
        Rtb.SelText = "Imóvel não cadastrado"
    Else
        Ocupado
        Rtb.Font.Size = 9
        Rtb.SelAlignment = rtfLeft
        Escreve
        'Calculo
        CarregaFoto
        Liberado
        
    End If
End With

End Sub

Private Sub CarregaFoto()
    
    
'If UCase(NomeDoComputador) = "TESLA" Or UCase(NomeDoComputador) = "GTI" Then
'    img.Picture = LoadPicture(sPathBin & "\FotoTeste.jpg")
'    Exit Sub
'End If
    
sSubPath = sPathFoto
If xImovel.Setor = 1 Then
    sSubPath = sSubPath & "\FOTOS_S1"
ElseIf xImovel.Setor = 2 Then
    sSubPath = sSubPath & "\FOTOS_S2"
ElseIf xImovel.Setor = 3 Then
    sSubPath = sSubPath & "\FOTOS_S3"
ElseIf xImovel.Setor = 4 Then
    sSubPath = sSubPath & "\FOTOS_S4"
Else
    Exit Sub
End If
On Error Resume Next
If Dir(sSubPath, vbDirectory) = "" Then
   img.Visible = False
   Exit Sub
Else
   Rss = Format(xImovel.Distrito, "00") & "-" & Format(xImovel.Setor, "00") & "-" & Format(xImovel.Quadra, "0000") & "-" & Format(xImovel.Lote, "00000") & "*.jpg"
   File1.Pattern = Rss
   File1.Path = sSubPath
   If File1.ListCount > 0 Then
      img.Visible = True
      File1.ListIndex = 0
      img.Picture = LoadPicture(sSubPath & "\" & File1.FileName)
      lblFotoDe.Caption = "1"
      lblFotoAte.Caption = File1.ListCount
 Else
      img.Visible = False
      lblFotoDe.Caption = "0"
      lblFotoAte.Caption = "0"
   End If
End If

End Sub

Private Sub cmdPrint_Click()
Dim sLinha As String, sFoto As String

If Rtb.Text = "" Then
    MsgBox "Nada a imprimir.", vbCritical, "Atenção"
    Exit Sub
End If


On Error Resume Next
Open App.Path & "\bin\detalhe.bat" For Output As #1
Print #1, "@echo off"
Print #1, "echo Aguarde....gerando relatorio."
Print #1, """" & App.Path & "\bin\DETALHE.html" & """"
Close #1

Open App.Path & "\bin\detalhe.html" For Output As #1

sLinha = "<html><head><title>Detalhes do imóvel</title></head>"
Print #1, sLinha
sLinha = "<body><font size =2>"
Print #1, sLinha
sLinha = Replace(Rtb.Text, vbCrLf, "<BR>") & "<BR><BR>"
Print #1, sLinha
sFoto = File1.FileName


If UCase(NomeDoComputador) = "TESLA" Or UCase(NomeDoComputador) = "GTI" Then
    sLinha = "<img src=""FotoTeste.jpg"" width=500 height=350 ></body></html>"
    Print #1, sLinha
    Close #1
    
    x = Shell(App.Path & "\Bin\detalhe.bat", vbNormalFocus)
    Exit Sub
End If


If sFoto <> "" Then
    If xImovel.Setor = 1 Then
        sLinha = "<img src=""file://///192.168.200.130/FotosGTI/FOTOS_S1/" & sFoto & """ width=500 height=350 ></body></html>"
    ElseIf xImovel.Setor = 2 Then
        sLinha = "<img src=""file://///192.168.200.130/FotosGTI/FOTOS_S2/" & sFoto & """ width=500 height=350 ></body></html>"
    ElseIf xImovel.Setor = 3 Then
        sLinha = "<img src=""file://///192.168.200.130/FotosGTI/FOTOS_S3/" & sFoto & """ width=500 height=350 ></body></html>"
    ElseIf xImovel.Setor = 4 Then
        sLinha = "<img src=""file://///192.168.200.130/FotosGTI/FOTOS_S4/" & sFoto & """ width=500 height=350 ></body></html>"
    End If
    Print #1, sLinha
End If
Close #1

x = Shell(App.Path & "\Bin\detalhe.bat", vbNormalFocus)
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
nAnoCalculo = Year(Now)
LoadMatrix
Centraliza Me
Set xImovel = New clsImovel

End Sub

Private Sub Calculo()
Dim nSomaTestada As Double, nAreaTerrenoReal As Double, qd As New rdoQuery
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer, bVVDeclarado As Boolean
Dim bIsento As Boolean, nTestada1 As Double, x As Integer, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset

nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3

bIsento = False

Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & nAnoCalculo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        If RdoAux3.RowCount = 0 Then
            Rtb.SelText = "Este imóvel esta classificado como: " & RdoAux!DESCTIPO
            Rtb.SelText = "" & vbCrLf: Negrito
            bIsento = True
        End If
        RdoAux3.Close
    End If
   .Close
End With

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = !Li_CodBairro
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    Negrito
    Rtb.SelText = "Tem Predial: ": Normal
    Rtb.SelText = IIf(bTemPredial, "Sim", "Não") & vbCrLf: Negrito
    Rtb.SelText = "Área Construida: ": Normal
    Rtb.SelText = FormatNumber(nAreaPrincipal, 2) & " m²" & vbCrLf: Negrito
    
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 1 Then
            nTestadaPrincipal = !AREATESTADA
            nTestada1 = !AREATESTADA
        Else
            nSomaTestada = 0
            Do Until .EOF
               If !NUMFACE = RdoAux!Seq Then
                  nTestada1 = !AREATESTADA
               End If
               nSomaTestada = nSomaTestada + !AREATESTADA
              .MoveNext
            Loop
            If nNumTestadas = 0 Then
                Rtb.SelText = "O imovel esta sem testada cadastrada"
                Exit Sub
            End If
            nTestadaPrincipal = nSomaTestada / nNumTestadas
        End If
    End With
    Negrito
    Rtb.SelText = "Testada Principal: ": Normal
    Rtb.SelText = FormatNumber(nTestadaPrincipal, 2) & " m" & vbCrLf: Negrito
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    '****************
        'BUSCA ÁREA PRINCIPAL
    'Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 1 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) And RdoAux2!QTDEPAV < 2 And nAreaTerreno < 600 Then
                        If nAnoCalculo > 2006 Then
                            Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
                            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                            If RdoAux4.RowCount = 0 Then
                                bIsento = True
                                MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                            Else
                                If ImovelAreaUnica(RdoAux4!CODPROPRIETARIO) Then
                                    MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                                    bIsento = True
                                End If
                            End If
                            RdoAux4.Close
                        Else
                            bIsento = True
                            MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                        End If
                    End If
                End If
               .Close
            End With
        Else
            bIsento = False
'            Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
'            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
 '           If RdoAux4.RowCount > 0 Then
 '               If ImovelAreaUnica(RdoAux4!CODPROPRIETARIO) Then
 '                   MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
 '                   bIsento = True
 '               End If
 '           End If
  '          RdoAux4.Close
        End If

        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       
        'novo VVP ***********************************
        If nAnoCalculo > 2007 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            If bTemPredial Then
                Do Until .EOF
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                    nFatorCategoria = 0
                    For x = 1 To UBound(aFatorC)
                        If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                           nFatorCategoria = aFatorC(x).Fator
                           Exit For
                        End If
                    Next
                    nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
                   .MoveNext
                Loop
            End If
        Else
            If bTemPredial Then
                 nUso = !USOCONSTR
                 nTipo = !TIPOCONSTR
                 nCat = !CATCONSTR
            End If
        End If
       .Close
    End With
    
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    For x = 1 To UBound(aGleba)
        If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
             Exit For
        ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
             Exit For
        End If
    Next
    nCodGleba = aGleba(x).Codigo
    'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
    nFatorGleba = aFatorG(nCodGleba)
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= Round(aProf(x).Min, 2) And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
     Else
        nFatorProfundidade = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba, 2)
    'CÁLCULO VALOR VENAL TERRITORIAL
    nFatorDistrito = aFatorD(!Distrito)
    nValorFatores = nValorFatores * nFatorDistrito
    nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
'        nFatorDistrito = aFatorD(!Distrito)
'        nValorFatores = nValorFatores * nFatorDistrito
        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        If nAnoCalculo < 2008 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            For x = 1 To UBound(aFatorC)
                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                   nFatorCategoria = aFatorC(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
        End If
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
    Else
        nFatorDistrito = 0
        nFatorCategoria = 0
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nCodTributo = 1
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
        nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)  'reajuste 2004-2005 (TIRADO)
    Else
        nCodTributo = 2
        nValorVenalImovel = nValorVenalTerritorial
        nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)  'reajuste 2004-2005 (TIRADO)
    End If
    
    'VALORVENAL DECLARADO
    Sql = "SELECT VALOR FROM VVDECLARADO WHERE CODREDUZIDO=" & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
        bVVDeclarado = True
        nValorVenalImovel = RdoAux2!Valor
        If bTemPredial Then
            nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)
        Else
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
        End If
    Else
        bVVDeclarado = False
    End If
    
    
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    If bTemPredial Then
       nValorFinal = nValorIptu
    Else
       nValorFinal = nValorITU
    End If
    
    
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoTmp.Close
On Error GoTo 0
qd.Sql = "{ Call spCALCULO(?,?) }"
qd(0) = nCodReduz
qd(1) = 0
'qd(2) = 0
Set RdoAux4 = qd.OpenResultset(rdOpenKeyset)

    
    
    
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    Rtb.SelText = "Agrupamento: ": Normal
    Rtb.SelText = FormatNumber(nValorAgrupamento, 2) & vbCrLf:   Negrito
    Rtb.SelText = "Soma dos Fatores: ": Normal
    Rtb.SelText = FormatNumber(nValorFatores, 2) & vbCrLf:   Negrito
    
    Rtb.SelText = "Valor Venal Teritorial: ": Normal
    'Rtb.SelText = "R$ " & FormatNumber(nValorVenalTerritorial, 2) & vbCrLf:    Negrito
    Rtb.SelText = "R$ " & FormatNumber(RdoAux4!VVT, 2) & vbCrLf:    Negrito
    Rtb.SelText = "Valor Venal Predial: ": Normal
    'Rtb.SelText = "R$ " & FormatNumber(nValorVenalPredial, 2) & vbCrLf:    Negrito
    Rtb.SelText = "R$ " & FormatNumber(RdoAux4!VVP, 2) & vbCrLf:    Negrito
    If bVVDeclarado Then
        Rtb.SelText = "Valor Venal Imóvel (DECLARADO): ": Normal
    Else
        Rtb.SelText = "Valor Venal Imóvel: ": Normal
    End If
    'Rtb.SelText = "R$ " & FormatNumber(nValorVenalImovel, 2) & vbCrLf:     Negrito
    Rtb.SelText = "R$ " & FormatNumber(RdoAux4!VVI, 2) & vbCrLf:     Negrito
    
    Rtb.SelText = "Imóvel Inativo: ":  Normal
    Rtb.SelText = IIf(xImovel.Inativo, "Sim", "Não") & vbCrLf: Negrito
    Rtb.SelText = "Valor do IPTU: ": Normal
    'Rtb.SelText = "R$ " & FormatNumber(nValorFinal, 2)
    If RdoAux4!ValorIptu > 0 Then
        Rtb.SelText = "R$ " & FormatNumber(RdoAux4!ValorIptu, 2)
    Else
        Rtb.SelText = "R$ " & FormatNumber(RdoAux4!ValorITU, 2)
    End If
End With

End Sub

Private Sub CalculoOld()
Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, x As Integer, RdoAux3 As rdoResultset

nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3

bIsento = False

Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Year(Now)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        If RdoAux3.RowCount = 0 Then
            Rtb.SelText = "Este imóvel esta classificado como: " & RdoAux!DESCTIPO
            Rtb.SelText = "" & vbCrLf: Negrito
            bIsento = True
        End If
        RdoAux3.Close
    End If
   .Close
End With

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = !Li_CodBairro
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    Negrito
    Rtb.SelText = "Tem Predial: ": Normal
    Rtb.SelText = IIf(bTemPredial, "Sim", "Não") & vbCrLf: Negrito
    Rtb.SelText = "Área Construida: ": Normal
    Rtb.SelText = FormatNumber(nAreaPrincipal, 2) & " m²" & vbCrLf: Negrito
    
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 1 Then
            nTestadaPrincipal = !AREATESTADA
            nTestada1 = !AREATESTADA
        Else
            nSomaTestada = 0
            Do Until .EOF
               If !NUMFACE = RdoAux!Seq Then
                  nTestada1 = !AREATESTADA
               End If
               nSomaTestada = nSomaTestada + !AREATESTADA
              .MoveNext
            Loop
            If nNumTestadas = 0 Then
                Rtb.SelText = "O imovel esta sem testada cadastrada"
                Exit Sub
            End If
            nTestadaPrincipal = nSomaTestada / nNumTestadas
        End If
    End With
    Negrito
    Rtb.SelText = "Testada Principal: ": Normal
    Rtb.SelText = FormatNumber(nTestadaPrincipal, 2) & " m" & vbCrLf: Negrito
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    'BUSCA ÁREA PRINCIPAL
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If Not IsNull(!soma) Then
                If RdoAux2.RowCount = 0 Then
                    MsgBox "O imóvel informado possue erro nas áreas cadastradas.", vbCritical, "Atenção"
                    Exit Sub
                End If
                Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux3.RowCount = 0 Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 1 And RdoAux2!TIPOCONSTR = 1 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) Then
                        bIsento = True
                        Rtb.SelText = "IMÓVEL ISENTO DE IPTU POR TER ÁREA CONSTRUIDA MENOR QUE 65 m²" & vbCrLf: Negrito
                    End If
                End If
                RdoAux3.Close
            End If
           .Close
        End With
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
        If bTemPredial Then
             nUso = !USOCONSTR
             nTipo = !TIPOCONSTR
             nCat = !CATCONSTR
             Select Case !USOCONSTR
                  Case 0
                     nTaxaLimpeza = 3.78
                  Case 1, 2, 3, 4, 5
                     nTaxaLimpeza = 10.57
                  Case Else
                     nTaxaLimpeza = 3.01
             End Select
        Else
             nTaxaLimpeza = 3.01
        End If
        nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
       '--CÁLCULO DA TAXA DE CONSERVAÇÃO
        If RdoAux!PAVIMENTO = 1 Then
           nTaxaConservacao = 1.35 * nTestadaPrincipal
        Else
           nTaxaConservacao = 0
        End If
        If nCodBairro = 81 Then
           nTaxaLimpeza = 1
           nTaxaConservacao = 1
        End If
       .Close
    End With
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
       nValorAgrupamento98 = aFatorR98(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
       nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
    End If
    
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
'    If !Dt_CodUsoTerreno = 6 Then
        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
        For x = 1 To UBound(aGleba)
            If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                 Exit For
            ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                 Exit For
            End If
        Next
        nCodGleba = aGleba(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
        nFatorGleba = aFatorG(nCodGleba)
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
        nFatorGleba98 = aFatorG98(nCodGleba)
 '   Else
 '       nFatorGleba = 1
 '       nFatorGleba98 = 1
 '   End If
    
    
    
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= Round(aProf(x).Min, 2) And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
        nFatorProfundidade98 = 0
        For x = 1 To UBound(aFatorF98)
            If aFatorF98(x).Distrito = !Distrito And aFatorF98(x).Codigo = nCodProfundidade Then
               nFatorProfundidade98 = aFatorF98(x).Fator
               Exit For
            End If
        Next
     Else
        nFatorProfundidade = 1
        nFatorProfundidade98 = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    'FATOR SITUACAO 98
    nFatorSituacao98 = aFatorS98(nCodSituacao)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    'FATOR PEDOLOGIA 98
    nFatorPedologia98 = aFatorP98(nCodPedologia)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    'FATOR TOPOGRAFIA 98
    nFatorTopografia98 = aFatorT98(nCodTopografia)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
    nValorFatores98 = nFatorTopografia98 * nFatorSituacao98 * nFatorPedologia98 * nFatorProfundidade98 * nFatorGleba98
    'CÁLCULO VALOR VENAL TERRITORIAL
    nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
    nValorVenalTerritorial98 = nAreaTerreno * nValorAgrupamento98 * nValorFatores98
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        'FATOR DISTRITO 98
        nFatorDistrito98 = aFatorD98(!Distrito)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        nValorVenalPredial = 0
        nValorVenalPredial98 = 0
        For x = 1 To UBound(aFatorC)
            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
               nFatorCategoria = aFatorC(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
        
       'FATOR CATEGORIA 98
        nFatorCategoria98 = 0
        For x = 1 To UBound(aFatorC98)
            If aFatorC98(x).Uso = nUso And aFatorC98(x).Tipo = nTipo And aFatorC98(x).Categoria = nCat Then
               nFatorCategoria98 = aFatorC98(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial98 = nValorVenalPredial98 + (nAreaPrincipal * nFatorCategoria98)
        
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        nValorVenalPredial98 = nValorVenalPredial98 * nFatorDistrito98
    Else
        nValorVenalPredial = 0
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nCodTributo = 1
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
        nValorVenalImovel98 = nValorVenalTerritorial98 + nValorVenalPredial98
        nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)
        nValorIPTU98 = nValorVenalImovel98 * (nAliquotaPredial / 100)
        nValorIPTU98 = nValorIPTU98 + nTaxaConservacao + nTaxaLimpeza
        nValorIPTU98 = CDbl(nValorIPTU98) * 1.6916
    Else
        nCodTributo = 2
        nValorVenalImovel = nValorVenalTerritorial
        nValorVenalImovel98 = nValorVenalTerritorial98
        nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
        nValorITU98 = nValorVenalImovel98 * (nAliquotaTerritorial / 100)
        nValorITU98 = nValorITU98 + nTaxaConservacao + nTaxaLimpeza
        nValorITU98 = CDbl(nValorITU98) * 1.6916
    End If
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    Rtb.SelText = "Agrupamento: ": Normal
    Rtb.SelText = FormatNumber(nValorAgrupamento, 2) & vbCrLf:   Negrito
    Rtb.SelText = "Soma dos Fatores: ": Normal
    Rtb.SelText = FormatNumber(nValorFatores, 2) & vbCrLf:   Negrito
    
    Rtb.SelText = "Valor Venal Teritorial: ": Normal
    Rtb.SelText = "R$ " & FormatNumber(nValorVenalTerritorial, 2) & vbCrLf:    Negrito
    Rtb.SelText = "Valor Venal Predial: ": Normal
    Rtb.SelText = "R$ " & FormatNumber(nValorVenalPredial, 2) & vbCrLf:    Negrito
    Rtb.SelText = "Valor Venal Imóvel: ": Normal
    Rtb.SelText = "R$ " & FormatNumber(nValorVenalImovel, 2) & vbCrLf:     Negrito
    
    
    If bTemPredial Then
   '     If nValorIPTU98 > nValorIPTU Then
           nValorFinal = nValorIptu
   '     Else
   '        nValorFinal = nValorIPTU98
   '     End If
    Else
   '     If nValorITU98 > nValorITU Then
           nValorFinal = nValorITU
   '     Else
   '        nValorFinal = nValorITU98
   '     End If
    End If
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub Escreve()
Dim Sql As String, RdoTmp As rdoResultset


xImovel.CarregaImovel nCodReduz
With Rtb
    .SelUnderline = True
     Negrito
    .SelText = "Detalhes do Imóvel nº " & nCodReduz & vbCrLf & vbCrLf
     Negrito
    .SelText = "Ativo ": Normal
    .SelText = IIf(xImovel.Inativo, "Não", "Sim") & vbCrLf: Negrito
    
    .SelText = "Inscrição Cadastral: ":     Normal
    .SelText = xImovel.Inscricao & vbCrLf:     Negrito
    .SelText = "Proprietário Principal: ":     Normal
    .SelText = xImovel.NomePropPrincipal & vbCrLf: Negrito
    .SelText = "Endereço: ": Normal
    .SelText = xImovel.EnderecoCompleto & " " & xImovel.Li_Compl & vbCrLf: Negrito
    sEnd = xImovel.EnderecoCompleto
    .SelText = "Bairro: ": Normal
    .SelText = xImovel.DescBairro & vbCrLf: Negrito
    .SelText = "Cep: ": Normal
    .SelText = RetornaCEP(xImovel.CodLogr, xImovel.Li_Num) & vbCrLf: Negrito
    .SelText = "Quadra Original: ":     Normal
    .SelText = xImovel.Li_Quadras & vbCrLf:    Negrito
    .SelText = "Lote Original: ":     Normal
    .SelText = xImovel.Li_Lotes & vbCrLf:    Negrito
    
    .SelText = "Área do Terreno: ":     Normal
    .SelText = FormatNumber(xImovel.Dt_AreaTerreno, 2) & " m²" & vbCrLf:    Negrito
    .SelText = "Fração Ideal: ":     Normal
    .SelText = FormatNumber(xImovel.Dt_FracaoIdeal, 2) & vbCrLf:    Negrito
    .SelText = "Topografia: ":     Normal
    .SelText = xImovel.DescTopografia & vbCrLf:     Negrito
    .SelText = "Pedologia: ":     Normal
    .SelText = xImovel.DescPedologia & vbCrLf:     Negrito
    .SelText = "Situação: ":     Normal
    .SelText = xImovel.DescSituacao & vbCrLf:     Negrito
    .SelText = "Uso do Terreno: ":     Normal
    .SelText = xImovel.DescUsoTerreno & vbCrLf:     Negrito
    .SelText = "Benfeitoria: ":     Normal
    .SelText = xImovel.DescBenfeitoria & vbCrLf:     Negrito
    .SelText = "Categoria da Propriedade: ":     Normal
    .SelText = xImovel.DescCategProp & vbCrLf:     Negrito
    .SelText = "Testadas: ": Normal
    xImovel.CarregaTestada
    For f = 1 To xImovel.QtdeTestada
       .SelText = "Face: " & Format(xImovel.Testada(f, 1), "00") & " - " & FormatNumber(xImovel.Testada(f, 2), 2) & " m"
       If f <> xImovel.QtdeTestada Then
         .SelText = ", "
       Else
       .SelText = "" & vbCrLf:        Negrito
       End If
    Next
    Calculo
    '.SelText = "Imóvel Inativo: ":  Normal
    '.SelText = IIf(xImovel.Inativo, "Sim", "Não") & vbCrLf: Negrito
    '.SelText = "Valor do IPTU: ": Normal
    
   ' Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & Val(txtCod.text) & " AND ANOEXERCICIO=" & nAnoCalculo & " AND "
   ' Sql = Sql & "CODLANCAMENTO=1 AND NUMPARCELA>0 AND CODTRIBUTO<>3"
   ' Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
    '    If IsNull(!soma) Then
    '        Rtb.SelText = "R$ 0,00" & vbCrLf
    '    Else
    '        Rtb.SelText = "R$ " & FormatNumber(!soma, 2) & vbCrLf
    '    End If
    '   .Close
 '   End With
End With

End Sub

Private Sub Fonte(Alinhamento As RichTextLib.SelAlignmentConstants, Cor As Long, Size As Integer, Negrito As Boolean, Italico As Boolean, Sublinhado As Boolean)

With Rtb
    .SelAlignment = Alinhamento
    .SelColor = Cor
    .SelFontSize = Size
    .SelBold = Negrito
    .SelUnderline = Sublinhado
    .SelItalic = Italico
End With

End Sub

Private Sub Negrito()
Rtb.SelBold = True
End Sub

Private Sub Normal()
Rtb.SelBold = False
End Sub

Private Sub LoadMatrix()

ReDim aFatorD(3)
ReDim aFatorD98(3)
ReDim aFatorP(6)
ReDim aFatorP98(6)
ReDim aFatorT(6)
ReDim aFatorT98(6)
ReDim aFatorS(6)
ReDim aFatorS98(6)
ReDim aFatorG(23)
ReDim aFatorG98(23)
ReDim aFatorR(8)
ReDim aFatorR98(8)

Sql = "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & nAnoCalculo & " ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA= 1998 ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG= 1998 ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO= 1998 ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA= 1998 ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO= 1998 ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR=" & nAnoCalculo & "  AND  CODMOEDA=1; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR= 1998  AND  CODMOEDA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        aFatorP(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorP98(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT98(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS98(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG98(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD98(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR98(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .Close
End With

ReDim aProf(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN FROM PROFUNDIDADE ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aProf(UBound(aProf) + 1)
        aProf(UBound(aProf)).Distrito = !CODDISTRITO
        aProf(UBound(aProf)).Codigo = !CODPROFUN
        aProf(UBound(aProf)).Min = !MINPROFUN
        aProf(UBound(aProf)).Max = !MAXPROFUN
       .MoveNext
     Loop
    .Close
End With

ReDim aFatorF(0)
ReDim aFatorF98(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN; " & _
      "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN= 1998 ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorF(UBound(aFatorF) + 1)
        aFatorF(UBound(aFatorF)).Distrito = !CODDISTRITO
        aFatorF(UBound(aFatorF)).Codigo = !CODPROFUN
        aFatorF(UBound(aFatorF)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        ReDim Preserve aFatorF98(UBound(aFatorF98) + 1)
        aFatorF98(UBound(aFatorF98)).Distrito = !CODDISTRITO
        aFatorF98(UBound(aFatorF98)).Codigo = !CODPROFUN
        aFatorF98(UBound(aFatorF98)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .Close
End With

ReDim aGleba(0)
Sql = "SELECT CODGLEBA,MINGLEBA,MAXGLEBA FROM GLEBA ORDER BY CODGLEBA "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aGleba(UBound(aGleba) + 1)
        aGleba(UBound(aGleba)).Codigo = !CODGLEBA
        aGleba(UBound(aGleba)).Min = !MINGLEBA
        aGleba(UBound(aGleba)).Max = !MAXGLEBA
       .MoveNext
     Loop
    .Close
End With

ReDim aFatorC(0)
ReDim aFatorC98(0)
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1; " & _
      "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=1998 AND CODMOEDA=1 "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorC(UBound(aFatorC) + 1)
        aFatorC(UBound(aFatorC)).Uso = !CODUSO
        aFatorC(UBound(aFatorC)).Tipo = !CodTipo
        aFatorC(UBound(aFatorC)).Categoria = !CODCATEG
        aFatorC(UBound(aFatorC)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        ReDim Preserve aFatorC98(UBound(aFatorC98) + 1)
        aFatorC98(UBound(aFatorC98)).Uso = !CODUSO
        aFatorC98(UBound(aFatorC98)).Tipo = !CodTipo
        aFatorC98(UBound(aFatorC98)).Categoria = !CODCATEG
        aFatorC98(UBound(aFatorC98)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub

Private Sub Label3_Click()
If MsgBox("Continuar", vbQuestion + vbYesNo, "SCHWARTZ") = vbNo Then Exit Sub

CalculoRolIsento

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdLoad_Click
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If

End Sub

Private Sub pFoto_Click()

If File1.ListCount = 0 Then Exit Sub

If Val(lblFotoDe.Caption) = 1 Then
   Exit Sub
Else
   lblFotoDe.Caption = Val(lblFotoDe.Caption) - 1
   lblFotoDe.Refresh
   File1.ListIndex = File1.ListIndex - 1
   Me.MousePointer = vbHourglass
   img.Picture = LoadPicture(sSubPath & "\" & File1.FileName)
   Me.MousePointer = vbDefault
End If

End Sub

Private Sub uFoto_Click()

If File1.ListCount = 0 Then Exit Sub
On Error Resume Next
If Val(lblFotoDe.Caption) = File1.ListCount Then
   Exit Sub
Else
   lblFotoDe.Caption = File1.ListIndex + 2
   lblFotoDe.Refresh
   File1.ListIndex = File1.ListIndex + 1
   Me.MousePointer = vbHourglass
   img.Picture = LoadPicture(sSubPath & "\" & File1.FileName)
   Me.MousePointer = vbDefault
End If

End Sub

Private Sub CalculoRol()

Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, x As Integer

nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3

Open sPathBin & "\ROLIPTU.TXT" For Output As #1

For nCodReduz = 1 To 40000
    If cGetInputState() <> 0 Then DoEvents
    xImovel.CarregaImovel nCodReduz
    
    ax = nCodReduz & "#" & xImovel.Inscricao & "#" & xImovel.NomePropPrincipal & "#"
    ax = ax & xImovel.EnderecoCompleto & "#" & xImovel.DescBairro & "#" & xImovel.Li_Quadras & "#"
    ax = ax & xImovel.Li_Lotes & "#"
    
    bIsento = False
    
    Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
    Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            bIsento = True
        End If
       .Close
    End With
    
    'CÁLCULO
    Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
    Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
    Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
    Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
    Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
    Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
    
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then GoTo Proximo
        'DADOS DO IMOVEL0
        nCodBairro = Val(SubNull(!Li_CodBairro))
    '    lblIC.Caption = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00")
        nAreaTerreno = !Dt_AreaTerreno
        nAreaTerrenoReal = nAreaTerreno
        nCodSituacao = !Dt_CodSituacao
        nCodPedologia = !Dt_CodPedol
        nCodTopografia = !Dt_CodTopog
        nCodAgrupamento = !CODAGRUPA
        bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
        If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
        'TEM ÁREA?
        If Not IsNull(!SOMAAREA) Then
            bTemPredial = True
            nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
        Else
            bTemPredial = False
            nAreaPrincipal = 0
        End If
        
        'TESTADAS
        Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nNumTestadas = .RowCount
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                If nNumTestadas > 0 Then
                    nTestadaPrincipal = nSomaTestada / nNumTestadas
                Else
                    nTestadaPrincipal = 0
                End If
                
            End If
        End With
        
        ax = ax & nTestadaPrincipal & "#" & nAreaTerreno & "#" & nAreaPrincipal & "#"
        
        'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
        '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
        
        'BUSCA ÁREA PRINCIPAL
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                GoTo fimRdoAux2
            End If
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 0 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) Then
                        bIsento = True
                    End If
                End If
               .Close
            End With
            If bFracaoIdeal Then
                nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
            End If
           'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
            If bTemPredial Then
                 nUso = !USOCONSTR
                 nTipo = !TIPOCONSTR
                 nCat = !CATCONSTR
                 Select Case !USOCONSTR
                      Case 0
                         nTaxaLimpeza = 3.78
                      Case 1, 2, 3, 4, 5
                         nTaxaLimpeza = 10.57
                      Case Else
                         nTaxaLimpeza = 3.01
                 End Select
            Else
                 nTaxaLimpeza = 3.01
            End If
            nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
           '--CÁLCULO DA TAXA DE CONSERVAÇÃO
            If RdoAux!PAVIMENTO = 1 Then
               nTaxaConservacao = 1.35 * nTestadaPrincipal
            Else
               nTaxaConservacao = 0
            End If
            If nCodBairro = 81 Then
               nTaxaLimpeza = 1
               nTaxaConservacao = 1
            End If
           .Close
        End With
fimRdoAux2:
        'VALOR DOS AGRUPAMENTOS
        If !Dt_CodUsoTerreno = 6 Then
           nValorAgrupamento = aFatorR(7)
           nValorAgrupamento98 = aFatorR98(7)
        Else
           nValorAgrupamento = aFatorR(nCodAgrupamento)
           nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
        End If
        
        '**************************
        'CÁLCULO DOS FATORES
        '**************************
        '**************************
        '### FATOR GLEBA ###
        '**************************
        If !Dt_CodUsoTerreno = 6 Then
            'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
            For x = 1 To UBound(aGleba)
                If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                     Exit For
                ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                     Exit For
                End If
            Next
            nCodGleba = aGleba(x).Codigo
            'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
            nFatorGleba = aFatorG(nCodGleba)
            'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
            nFatorGleba98 = aFatorG98(nCodGleba)
        Else
            nFatorGleba = 1
            nFatorGleba98 = 1
        End If
        '**************************
        '### FATOR PROFUNDIDADE ###
        '**************************
        If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
            If nTestadaPrincipal > 0 Then
                nValorProfundidade = FormatNumber(nAreaTerreno / nTestadaPrincipal, 2)
            Else
                nValorProfundidade = 1
            End If
             'nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
            'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
            For x = 1 To UBound(aProf)
                If aProf(x).Distrito = !Distrito Then
                   If nValorProfundidade >= aProf(x).Min And nValorProfundidade <= aProf(x).Max Then
                      Exit For
                   ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                      Exit For
                   End If
                End If
            Next
            If x > UBound(aProf) Then x = 1
            nCodProfundidade = aProf(x).Codigo
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
            nFatorProfundidade = 0
            For x = 1 To UBound(aFatorF)
                If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade = aFatorF(x).Fator
                   Exit For
                End If
            Next
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
            nFatorProfundidade98 = 0
            For x = 1 To UBound(aFatorF98)
                If aFatorF98(x).Distrito = !Distrito And aFatorF98(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade98 = aFatorF98(x).Fator
                   Exit For
                End If
            Next
         Else
            nFatorProfundidade = 1
            nFatorProfundidade98 = 1
         End If
        '**************************
        '### FATOR SITUAÇÃO ###
        '**************************
        nFatorSituacao = aFatorS(nCodSituacao)
        'FATOR SITUACAO 98
        nFatorSituacao98 = aFatorS98(nCodSituacao)
        '**************************
        '### FATOR PEDOLOGIA ###
        '**************************
        nFatorPedologia = aFatorP(nCodPedologia)
        'FATOR PEDOLOGIA 98
        nFatorPedologia98 = aFatorP98(nCodPedologia)
        '**************************
        '### FATOR TOPOGRAFIA ###
        '**************************
        nFatorTopografia = aFatorT(nCodTopografia)
        'FATOR TOPOGRAFIA 98
        nFatorTopografia98 = aFatorT98(nCodTopografia)
        '**************************
        'FIM DO CÁLCULO DOS FATORES
        '**************************
        'MULTIPLICA OS FATORES
        nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
        nValorFatores98 = nFatorTopografia98 * nFatorSituacao98 * nFatorPedologia98 * nFatorProfundidade98 * nFatorGleba98
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores
        nValorVenalTerritorial98 = nAreaTerreno * nValorAgrupamento98 * nValorFatores98
        'CÁLCULO VALOR VENAL PREDIAL
        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
        If bTemPredial Then
            '**************************
            '### FATOR DISTRITO ###
            '**************************
            nFatorDistrito = aFatorD(!Distrito)
            'FATOR DISTRITO 98
            nFatorDistrito98 = aFatorD98(!Distrito)
            '**************************
            '### FATOR CATEGORIA ###
            '**************************
            nValorVenalPredial = 0
            nValorVenalPredial98 = 0
            For x = 1 To UBound(aFatorC)
                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                   nFatorCategoria = aFatorC(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
            
           'FATOR CATEGORIA 98
            nFatorCategoria98 = 0
            For x = 1 To UBound(aFatorC98)
                If aFatorC98(x).Uso = nUso And aFatorC98(x).Tipo = nTipo And aFatorC98(x).Categoria = nCat Then
                   nFatorCategoria98 = aFatorC98(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial98 = nValorVenalPredial98 + (nAreaPrincipal * nFatorCategoria98)
            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
            nValorVenalPredial98 = nValorVenalPredial98 * nFatorDistrito98
        Else
            nValorVenalPredial = 0
        End If
        'VALOR ITU/IPTU
        If bTemPredial Then
            nCodTributo = 1
            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
            nValorVenalImovel98 = nValorVenalTerritorial98 + nValorVenalPredial98
            nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)
            nValorIPTU98 = nValorVenalImovel98 * (nAliquotaPredial / 100)
            nValorIPTU98 = nValorIPTU98 + nTaxaConservacao + nTaxaLimpeza
            nValorIPTU98 = CDbl(nValorIPTU98) * 1.6916
        Else
            nCodTributo = 2
            nValorVenalImovel = nValorVenalTerritorial
            nValorVenalImovel98 = nValorVenalTerritorial98
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
            nValorITU98 = nValorVenalImovel98 * (nAliquotaTerritorial / 100)
            nValorITU98 = nValorITU98 + nTaxaConservacao + nTaxaLimpeza
            nValorITU98 = CDbl(nValorITU98) * 1.6916
        End If
        'COMPARAÇÃO ENTRE OS CÁLCULOS
        
        If bTemPredial Then
            If nValorIPTU98 > nValorIptu Then
               nValorFinal = nValorIptu
            Else
               nValorFinal = nValorIPTU98
            End If
        Else
            If nValorITU98 > nValorITU Then
               nValorFinal = nValorITU
            Else
               nValorFinal = nValorITU98
            End If
        End If
        ax = ax & nValorVenalTerritorial & "#" & nValorVenalPredial & "#" & nValorFinal
    End With
    Print #1, ax
Proximo:
Next

Close #1
MsgBox "fim"
End Sub

Private Sub CalculoRolIsento()

Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, bImune As Boolean, sTipoIsencao As String, nTestada1 As Double, x As Integer

nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3

Open sPathBin & "\ROLIPTU.TXT" For Output As #1

For nCodReduz = 1 To 40000
    If cGetInputState() <> 0 Then DoEvents
    xImovel.CarregaImovel nCodReduz
    
    ax = nCodReduz & "#" & xImovel.Inscricao & "#" & xImovel.NomePropPrincipal & "#"
    ax = ax & xImovel.EnderecoCompleto & "#" & xImovel.DescBairro & "#" & xImovel.Li_Quadras & "#"
    ax = ax & xImovel.Li_Lotes & "#"
    
    bImune = False: bIsento = False: sTipoIsencao = ""
    Sql = "SELECT codreduzido, anoisencao FROM isencao WHERE codreduzido = " & nCodReduz & " AND codisencao = 1"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            bImune = True
            sTipoIsencao = "M"
        Else
            Sql = "SELECT codreduzido, anoisencao FROM isencao WHERE codreduzido = " & nCodReduz & " AND ANOISENCAO=" & Year(Now) & " AND (codisencao = 2 or codisencao=3) "
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    bIsento = True
                    sTipoIsencao = "S"
                End If
               .Close
            End With
        End If
       .Close
    End With
    If Not bIsento And Not bImune Then GoTo Proximo
    'CÁLCULO
    Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
    Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
    Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
    Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
    Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
    Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
    
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then GoTo Proximo
        'DADOS DO IMOVEL0
        nCodBairro = Val(SubNull(!Li_CodBairro))
        nAreaTerreno = !Dt_AreaTerreno
        nAreaTerrenoReal = nAreaTerreno
        nCodSituacao = !Dt_CodSituacao
        nCodPedologia = !Dt_CodPedol
        nCodTopografia = !Dt_CodTopog
        nCodAgrupamento = !CODAGRUPA
        bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
        If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
        'TEM ÁREA?
        If Not IsNull(!SOMAAREA) Then
            bTemPredial = True
            nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
        Else
            bTemPredial = False
            nAreaPrincipal = 0
        End If
        
        'TESTADAS
        Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nNumTestadas = .RowCount
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                If nNumTestadas > 0 Then
                    nTestadaPrincipal = nSomaTestada / nNumTestadas
                Else
                    nTestadaPrincipal = 0
                End If
                
            End If
        End With
        
        ax = ax & nTestadaPrincipal & "#" & nAreaTerreno & "#" & nAreaPrincipal & "#"
        
        'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
        '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
        
        'BUSCA ÁREA PRINCIPAL
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                GoTo fimRdoAux2
            End If
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 0 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) Then
                        bIsento = True
                    End If
                End If
               .Close
            End With
            If bFracaoIdeal Then
                nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
            End If
           'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
            If bTemPredial Then
                 nUso = !USOCONSTR
                 nTipo = !TIPOCONSTR
                 nCat = !CATCONSTR
                 Select Case !USOCONSTR
                      Case 0
                         nTaxaLimpeza = 3.78
                      Case 1, 2, 3, 4, 5
                         nTaxaLimpeza = 10.57
                      Case Else
                         nTaxaLimpeza = 3.01
                 End Select
            Else
                 nTaxaLimpeza = 3.01
            End If
            nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
           '--CÁLCULO DA TAXA DE CONSERVAÇÃO
            If RdoAux!PAVIMENTO = 1 Then
               nTaxaConservacao = 1.35 * nTestadaPrincipal
            Else
               nTaxaConservacao = 0
            End If
            If nCodBairro = 81 Then
               nTaxaLimpeza = 1
               nTaxaConservacao = 1
            End If
           .Close
        End With
fimRdoAux2:
        'VALOR DOS AGRUPAMENTOS
        If !Dt_CodUsoTerreno = 6 Then
           nValorAgrupamento = aFatorR(7)
           nValorAgrupamento98 = aFatorR98(7)
        Else
           nValorAgrupamento = aFatorR(nCodAgrupamento)
           nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
        End If
        
        '**************************
        'CÁLCULO DOS FATORES
        '**************************
        '**************************
        '### FATOR GLEBA ###
        '**************************
        If !Dt_CodUsoTerreno = 6 Then
            'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
            For x = 1 To UBound(aGleba)
                If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                     Exit For
                ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                     Exit For
                End If
            Next
            nCodGleba = aGleba(x).Codigo
            'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
            nFatorGleba = aFatorG(nCodGleba)
            'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
            nFatorGleba98 = aFatorG98(nCodGleba)
        Else
            nFatorGleba = 1
            nFatorGleba98 = 1
        End If
        '**************************
        '### FATOR PROFUNDIDADE ###
        '**************************
        If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
            If nTestadaPrincipal > 0 Then
                nValorProfundidade = FormatNumber(nAreaTerreno / nTestadaPrincipal, 2)
            Else
                nValorProfundidade = 1
            End If
             'nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
            'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
            For x = 1 To UBound(aProf)
                If aProf(x).Distrito = !Distrito Then
                   If nValorProfundidade >= aProf(x).Min And nValorProfundidade <= aProf(x).Max Then
                      Exit For
                   ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                      Exit For
                   End If
                End If
            Next
            If x > UBound(aProf) Then x = 1
            nCodProfundidade = aProf(x).Codigo
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
            nFatorProfundidade = 0
            For x = 1 To UBound(aFatorF)
                If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade = aFatorF(x).Fator
                   Exit For
                End If
            Next
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
            nFatorProfundidade98 = 0
            For x = 1 To UBound(aFatorF98)
                If aFatorF98(x).Distrito = !Distrito And aFatorF98(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade98 = aFatorF98(x).Fator
                   Exit For
                End If
            Next
         Else
            nFatorProfundidade = 1
            nFatorProfundidade98 = 1
         End If
        '**************************
        '### FATOR SITUAÇÃO ###
        '**************************
        nFatorSituacao = aFatorS(nCodSituacao)
        'FATOR SITUACAO 98
        nFatorSituacao98 = aFatorS98(nCodSituacao)
        '**************************
        '### FATOR PEDOLOGIA ###
        '**************************
        nFatorPedologia = aFatorP(nCodPedologia)
        'FATOR PEDOLOGIA 98
        nFatorPedologia98 = aFatorP98(nCodPedologia)
        '**************************
        '### FATOR TOPOGRAFIA ###
        '**************************
        nFatorTopografia = aFatorT(nCodTopografia)
        'FATOR TOPOGRAFIA 98
        nFatorTopografia98 = aFatorT98(nCodTopografia)
        '**************************
        'FIM DO CÁLCULO DOS FATORES
        '**************************
        'MULTIPLICA OS FATORES
        nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
        nValorFatores98 = nFatorTopografia98 * nFatorSituacao98 * nFatorPedologia98 * nFatorProfundidade98 * nFatorGleba98
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores
        nValorVenalTerritorial98 = nAreaTerreno * nValorAgrupamento98 * nValorFatores98
        'CÁLCULO VALOR VENAL PREDIAL
        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
        If bTemPredial Then
            '**************************
            '### FATOR DISTRITO ###
            '**************************
            nFatorDistrito = aFatorD(!Distrito)
            'FATOR DISTRITO 98
            nFatorDistrito98 = aFatorD98(!Distrito)
            '**************************
            '### FATOR CATEGORIA ###
            '**************************
            nValorVenalPredial = 0
            nValorVenalPredial98 = 0
            For x = 1 To UBound(aFatorC)
                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                   nFatorCategoria = aFatorC(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
            
           'FATOR CATEGORIA 98
            nFatorCategoria98 = 0
            For x = 1 To UBound(aFatorC98)
                If aFatorC98(x).Uso = nUso And aFatorC98(x).Tipo = nTipo And aFatorC98(x).Categoria = nCat Then
                   nFatorCategoria98 = aFatorC98(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial98 = nValorVenalPredial98 + (nAreaPrincipal * nFatorCategoria98)
            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
            nValorVenalPredial98 = nValorVenalPredial98 * nFatorDistrito98
        Else
            nValorVenalPredial = 0
        End If
        'VALOR ITU/IPTU
        If bTemPredial Then
            nCodTributo = 1
            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
            nValorVenalImovel98 = nValorVenalTerritorial98 + nValorVenalPredial98
            nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)
            nValorIPTU98 = nValorVenalImovel98 * (nAliquotaPredial / 100)
            nValorIPTU98 = nValorIPTU98 + nTaxaConservacao + nTaxaLimpeza
            nValorIPTU98 = CDbl(nValorIPTU98) * 1.6916
        Else
            nCodTributo = 2
            nValorVenalImovel = nValorVenalTerritorial
            nValorVenalImovel98 = nValorVenalTerritorial98
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
            nValorITU98 = nValorVenalImovel98 * (nAliquotaTerritorial / 100)
            nValorITU98 = nValorITU98 + nTaxaConservacao + nTaxaLimpeza
            nValorITU98 = CDbl(nValorITU98) * 1.6916
        End If
        'COMPARAÇÃO ENTRE OS CÁLCULOS
        
        If bTemPredial Then
            If nValorIPTU98 > nValorIptu Then
               nValorFinal = nValorIptu
            Else
               nValorFinal = nValorIPTU98
            End If
        Else
            If nValorITU98 > nValorITU Then
               nValorFinal = nValorITU
            Else
               nValorFinal = nValorITU98
            End If
        End If
        ax = ax & nValorVenalTerritorial & "#" & nValorVenalPredial & "#" & nValorFinal & "#" & sTipoIsencao
    End With
    Print #1, ax
Proximo:
Next

Close #1
MsgBox "fim"
End Sub


