VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmBuscaDoc 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca em arquivos bancários"
   ClientHeight    =   2790
   ClientLeft      =   5865
   ClientTop       =   3405
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   7785
   Begin VB.CheckBox chkDA 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Débito Automático"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   270
      TabIndex        =   22
      Top             =   1620
      Width           =   2850
   End
   Begin VB.ListBox lstPath 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   450
      TabIndex        =   21
      Top             =   4815
      Width           =   4695
   End
   Begin VB.TextBox txtDocumento 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      MaxLength       =   10
      TabIndex        =   18
      Top             =   675
      Width           =   1515
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmBuscaDoc.frx":0000
      Left            =   945
      List            =   "frmBuscaDoc.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   225
      Width           =   1050
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   270
      TabIndex        =   15
      Text            =   "H:\Trabalho\GTI\ArqBanco\"
      Top             =   1215
      Width           =   2985
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   135
      TabIndex        =   13
      Top             =   2025
      Width           =   6360
      _ExtentX        =   11218
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
      Color           =   8421376
   End
   Begin prjChameleon.chameleonButton cmdAbrir 
      Height          =   315
      Left            =   6660
      TabIndex        =   0
      ToolTipText     =   "Abrir arquivo do banco"
      Top             =   2400
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Abrir"
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
      MICON           =   "frmBuscaDoc.frx":0004
      PICN            =   "frmBuscaDoc.frx":0020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1845
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   4245
      Begin VB.TextBox txtDA 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1230
         TabIndex        =   12
         Top             =   1530
         Width           =   2955
      End
      Begin VB.TextBox txtValorPago 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1200
         TabIndex        =   10
         Top             =   1260
         Width           =   2955
      End
      Begin VB.TextBox txtDataCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1200
         TabIndex        =   9
         Top             =   945
         Width           =   2955
      End
      Begin VB.TextBox txtBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1200
         TabIndex        =   8
         Top             =   615
         Width           =   2955
      End
      Begin VB.TextBox txtArq 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1200
         TabIndex        =   7
         Top             =   300
         Width           =   2955
      End
      Begin VB.Label Label2 
         Caption         =   "Retorno DA..:"
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Valor Pago...:"
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Data Crédito.:"
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   940
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Banco..........:"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   615
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Arquivo........:"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
   End
   Begin prjChameleon.chameleonButton cmdRebuild 
      Height          =   315
      Left            =   2070
      TabIndex        =   14
      ToolTipText     =   "Reconstruir a lista de arquivos bancários"
      Top             =   225
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Rebuild"
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
      MICON           =   "frmBuscaDoc.frx":01FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdFind 
      Height          =   315
      Left            =   6660
      TabIndex        =   20
      ToolTipText     =   "Busca documento dentro dos arquivos"
      Top             =   2025
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "B&usca"
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
      MICON           =   "frmBuscaDoc.frx":0216
      PICN            =   "frmBuscaDoc.frx":0232
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento..:"
      Height          =   225
      Index           =   4
      Left            =   360
      TabIndex        =   19
      Top             =   735
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Height          =   240
      Left            =   360
      TabIndex        =   16
      Top             =   270
      Width           =   600
   End
   Begin VB.Label lblMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   2370
      Width           =   6345
   End
End
Attribute VB_Name = "frmBuscaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FebrabanA
    CodigoRegistro As String * 1
    CodigoRemessa As String * 1
    CodigoConvenio As String * 20
    NomeEmpresa As String * 20
    CodigoBanco As String * 3
    NomeBanco As String * 20
    DataGeracao As String * 8
    NumeroSeq As String * 6
    VersaoLayout As String * 2
    Filler As String * 69
End Type
Private Type FebrabanA2
    CodigoBanco As String * 3
    NumeroLote As String * 4
    TipoRegistro As String * 1
    Filler1 As String * 8
    TipoInscricao As String * 1
    NumeroInscricao As String * 15
    Agencia As String * 5
    NumeroConta As String * 10
    Filler2 As String * 5
    CodigoCedente As String * 9
    Filler3 As String * 11
    NomeEmpresa As String * 30
    NomeBanco As String * 30
    Filler4 As String * 10
    CodigoRetorno As String * 1
    DataGeracao As String * 8 'DDMMAAAA
    Filler5 As String * 6
    NumeroSeq As String * 6
    NumeroVersao As String * 3
    Filler6 As String * 74
End Type

Private Type FebrabanA3
    CodigoRegistro As String * 2
    CodigoRetorno As String * 7
    CodigoCobrança As String * 2
    NomeCobrança As String * 8
    CodigoPrefeitura As String * 11
    NomePrefeitura As String * 26
    CodigoBanco As String * 3
    NomeBanco As String * 7
    DataCredito As String * 6
    Filler1 As String * 4
    Filler2 As String * 38
    NumRegistros As String * 6
End Type


Private Type FebrabanF 'RETORNO DO DEBITO AUTOMATICO
   CodigoRegistro As String * 1
   Distrito As String * 2
   Setor As String * 2
   Quadra As String * 4
   Lote As String * 5
   Seq As String * 2
   FillerID As String * 10
   CodAgencia As String * 4
   ContaCliente As String * 14
   DataVencto As String * 8
   ValorDebito As String * 15
   CodRetorno As String * 2
   NumDoc As String * 9
   Filler1 As String * 51
   Filler2 As String * 20
   CodMovimento As String * 1
End Type
Private Type FebrabanG
   CodigoRegistro As String * 1
   ContaPrefeitura As String * 20
   DataPagamento As String * 8
   DataCredito As String * 8
   PreCodBarra As String * 4
   ValorRecebido As String * 11
   CodigoMunic As String * 4
   DataVencto As String * 8
   NumDocumento As String * 9
   NumParcela As String * 2
   SituacaoRetorno As String * 2
   FillerSmar As String * 4
   ValorRetornado As String * 12
   ValorTarifa As String * 7
   NumSeq As String * 8
   CodAgencia As String * 8
   FormaPagamento As String * 1
   NumAutentica As String * 23
   Filler As String * 10
End Type
Private Type FebrabanG2
   CodigoBanco As String * 3
   NumeroLote As String * 4
   TipoRegistro As String * 1
   NumSequencial As String * 5
   CodigoSegmento As String * 1
   Filler1 As String * 1
   CodigoMovimento As String * 2
   Agencia As String * 5
   NumeroConta As String * 10
   Filler2 As String * 8
   NossoNumero As String * 13
   CodigoCarteira As String * 1
   NumDocumento As String * 15
   DataVencto As String * 8 'DDMMAAAA
   ValorTitulo As String * 15
   BancoCobrador As String * 3
   AgenciaCobradora As String * 5
   UsoCedente As String * 25
   CodigoMoeda As String * 2
   TipoInscricao As String * 1 '1=CPF, 2=CNPJ
   NumeroInscricao As String * 15
   NomeSacado As String * 40
   ContaCobranca As String * 10
   ValorTarifa As String * 15
   Custas As String * 10
   Filler3 As String * 22
End Type

Private Type FebrabanG2U
    CodigoBanco As String * 3
    NumeroLote As String * 4
    TipoRegistro As String * 1
    NumeroSeq As String * 5
    CodSegmento As String * 1
    Filler1 As String * 1
    CodigoMov As String * 2
    JurosMulta As String * 15
    ValorDesconto As String * 15
    ValorAbatimento As String * 15
    ValorIOF As String * 15
    ValorPago As String * 15
    ValorCreditado As String * 15
    ValorOutrasDespesas As String * 15
    ValorOutrosCreditos As String * 15
    DataOcorrencia As String * 8 'DDMMAAAA
    DataCredito As String * 8 'DDMMAAAA
    Outros As String * 87
End Type

Private Type FebrabanG3
   CodigoRegistro As String * 1
   ContaPrefeitura As String * 11
   CodAgencia As String * 3
   NumDocumento As String * 8
   Codigo06 As String * 2
   DataPagamento As String * 6
   CodigoBanco As String * 5
   ValorTaxa As String * 13
   Filler1 As String * 26
   ValorPago As String * 13
   ValorSeiLa As String * 13
   CodigoC As String * 1
   Filler2 As String * 12
   NumSeq As String * 6
End Type

Private Type SimplesNacionalHeader
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    CodigoConvenio As String * 20
    DataGeracao As String * 8
    NumeroRemessa As String * 6
    NumeroVersao As String * 2
    Filler1 As String * 22
    Filler2 As String * 8
    CodigoBanco As String * 3
    Filler3 As String * 422
End Type
Private Type SimplesNacionalDetalhe
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    DataArrecada As String * 8
    DataVencimento As String * 8
    Filler1 As String * 12
    Filler2 As String * 37
    Cnpj As String * 14
    Filler3 As String * 11
    Esfera As String * 1
    Competencia As String * 6
    ValorPrincipal As String * 17
    ValorMulta As String * 17
    ValorJuros As String * 17
    Filler4 As String * 47
    ValorAutentica As String * 17
    NumeroAutentica As String * 23
    CodigoBanco As String * 3
    CodigoAgencia As String * 4
    Filler6 As String * 249
End Type

Private Type FebrabanZ3
    CodigoSeiLa As String * 7
    Filler1 As String * 3
    TotalRegistro As String * 6
    ValorTotal As String * 14
    CodigoC As String * 1
    DataCredito As String * 6
    Filler2 As String * 77
    NumSeq As String * 6
End Type

Dim aFebrabanA() As FebrabanA
Dim aFebrabanA2() As FebrabanA2
Dim aFebrabanG() As FebrabanG
Dim aFebrabanG2() As FebrabanG2
Dim aFebrabang2U() As FebrabanG2U
Dim aFebrabanF() As FebrabanF
Dim aSimplesH() As SimplesNacionalHeader
Dim aSimplesD() As SimplesNacionalDetalhe
Dim aFebrabanA3() As FebrabanA3
Dim aFebrabanG3() As FebrabanG3
Dim aFebrabanZ3() As FebrabanZ3

Private cCRCSearch As cFileSearchCRC
Dim sFile As String, nCodReduz As Long

Private Sub cmbAno_Click()
LoadFileBanco
End Sub

Private Sub cmdAbrir_Click()
Dim x As Long, z As Integer

If txtArq.Text <> "" Then
   x = Shell("NOTEPAD" & " " & txtArq.Text, vbNormalFocus)
End If

End Sub

Private Sub cmdFind_Click()
Dim sPathFile As String, strLinha As String, x As Long, Sql As String, RdoAux As rdoResultset
Dim strBuffer As String, lngResult As Long, dDataVencto As Date
Dim nPosicao As Integer, nCount As Integer, sRetorno As String, sDataVencto As String
Dim Header As FebrabanA, RegistroF As FebrabanF

ReDim aFebrabanA(0): ReDim aFebrabanF(0)
'LoadFileBanco
If Len(txtDocumento.Text) < 6 Then
    MsgBox "Número de documento inválido.", vbCritical, "Verifique"
    Exit Sub
End If

If lstPath.ListCount = 0 Then
    LoadFileBanco
End If

If lstPath.ListCount = 0 Then
    MsgBox "Lista " & sFile & " não localizada.", vbCritical, "Verifique"
    Exit Sub
End If

LimpaArq
Ocupado
Me.MousePointer = vbHourglass
lblMsg.Caption = "Procurando o Documento..."
lblMsg.Refresh
DoEvents

If chkDA.value = vbUnchecked Then
    For x = 0 To lstPath.ListCount - 1
        If lstPath.ItemData(x) = 1 Then GoTo PROXIMO
        If x Mod 20 = 0 Then
            CallPb x, CLng(lstPath.ListCount - 1)
        End If
        Set cCRCSearch = New cFileSearchCRC
        sPathFile = lstPath.List(x)
        cCRCSearch.SearchAlgorithm = Asm_BMHA
        strBuffer = sTr$(cCRCSearch.FileMapSearch(sPathFile, txtDocumento.Text))
        Set cCRCSearch = Nothing
        If Val(strBuffer) > 0 Then
            txtArq.Text = sPathFile
            LeArquivo sPathFile
            lblMsg.Caption = "Documento Localizado !!!"
            lblMsg.Refresh
            Exit For
        Else
            LimpaArq
        End If
PROXIMO:
    Next
ElseIf chkDA.value = vbChecked Then
    
    For x = 0 To lstPath.ListCount - 1
        If lstPath.ItemData(x) = 0 Then GoTo PROXIMO2
        If x Mod 20 = 0 Then
            CallPb x, CLng(lstPath.ListCount - 1)
        End If
        sPathFile = lstPath.List(x)
        Sql = "SELECT parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento,"
        Sql = Sql & "parceladocumento.NumParcela , parceladocumento.CODCOMPLEMENTO, parceladocumento.NumDocumento, debitoparcela.DataVencimento "
        Sql = Sql & "FROM parceladocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
        Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND "
        Sql = Sql & "parceladocumento.NumParcela = debitoparcela.NumParcela And parceladocumento.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO "
        Sql = Sql & "Where parceladocumento.NumDocumento = " & Val(txtDocumento.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                nCodReduz = !CODREDUZIDO
                dDataVencto = !DataVencimento
                
'**********************
                Open sPathFile For Binary Access Read As #1
                Get #1, 1, Header
                Posicao = Len(Header) + 3
                Get #1, Posicao, RegistroF
                Close #1

                sDataVencto = ConvDataSerial(RegistroF.DataVencto)
                If IsDate(sDataVencto) Then
                    If Month(CDate(sDataVencto)) = Month(dDataVencto) And Year(CDate(sDataVencto)) = Year(dDataVencto) Then
                        Set cCRCSearch = New cFileSearchCRC
                        cCRCSearch.SearchAlgorithm = Asm_BMHA
                        strBuffer = sTr$(cCRCSearch.FileMapSearch(sPathFile, nCodReduz))
                    End If
                End If
                        
'**********************
                
                If Val(strBuffer) > 0 Then
                    txtArq.Text = sPathFile
                    LeArquivo sPathFile
                    lblMsg.Caption = "Documento Localizado !!!"
                    lblMsg.Refresh
                    Exit For
                Else
                    LimpaArq
                End If
            Else
                LimpaArq
            End If
       .Close
        End With
PROXIMO2:
    Next
End If

If txtArq.Text = "" Then
    lblMsg.Caption = "Documento não encontrado !!!"
    lblMsg.Refresh
End If
PBar.value = 0
Me.MousePointer = vbDefault
Liberado

End Sub

Private Sub LimpaArq()
txtArq.Text = ""
txtBanco.Text = ""
txtDataCredito.Text = ""
txtValorPago.Text = ""
txtDA.Text = ""
End Sub

Private Sub LoadFileBanco()

Dim strLinha As String

sFile = App.Path & "\bin\Files" & cmbAno.Text & ".txt"

If Dir$(sFile) = "" Then
    MsgBox "Lista " & sFile & " não localizada, criando lista.", vbCritical, "Verifique"
    cmdRebuild_Click
    Exit Sub
End If

lstPath.Clear
Ocupado

On Error Resume Next
Close #1
On Error GoTo 0

Open sFile For Input As #1
   Do While Not EOF(1)
        Line Input #1, strLinha
        lstPath.AddItem strLinha
        If Mid(strLinha, Len(txtPath.Text) + 6, 2) = "DA" Then
            lstPath.ItemData(lstPath.NewIndex) = 1
        Else
            lstPath.ItemData(lstPath.NewIndex) = 0
        End If
   Loop
Close #1
Liberado
End Sub

Private Sub cmdRebuild_Click()
    
Dim sDirTemp As String
Dim iFilesFile As Integer
Dim iDirsFile As Integer
Dim iStart As Long
Set gsDirsQueue = Nothing
Set gsDirs = Nothing
Set gsFiles = Nothing

If txtPath.Text = "" Then Exit Sub

If Right(txtPath.Text, 1) <> "\" Then
    txtPath.Text = txtPath.Text & "\"
End If
Ocupado
DoEvents

On Error Resume Next
Kill App.Path & "\Bin\Files" & cmbAno.Text & ".txt"
On Error GoTo 0
iFilesFile = FreeFile
Open App.Path & "\Bin\Files" & cmbAno.Text & ".txt" For Output Access Write As iFilesFile

gsDirsQueue.Add txtPath.Text & cmbAno.Text
gsDirs.Add txtPath.Text & cmbAno.Text

While gsDirsQueue.Count > 0
    sDirTemp = FixPath(gsDirsQueue(1))
    sTemp = ""
    On Error Resume Next
    sTemp = Dir$(sDirTemp & "\*.*", vbNormal + vbReadOnly + vbHidden + vbDirectory)
    If sTemp = "" Then
        Liberado
        MsgBox "Diretório: " & sDirTemp & " não encontrado.", vbCritical, "ERRO"
        Close iFilesFile
        Exit Sub
    End If
    On Error GoTo 0
    While sTemp <> ""
        If sTemp <> "." And sTemp <> ".." Then
            If (GetAttr(sDirTemp & "\" & sTemp) And &H10) = vbDirectory Then
                gsDirsQueue.Add Item:=sDirTemp & "\" & sTemp, After:=DirSearchB(gsDirsQueue, sDirTemp & "\" & sTemp) ' Comes and goes here.
                gsDirs.Add Item:=sDirTemp & "\" & sTemp, After:=DirSearchB(gsDirs, sDirTemp & "\" & sTemp) ' Comes and goes here.
            Else  ' We have a file.
                gsFiles.Add sDirTemp & "\" & sTemp
                Print #iFilesFile, sDirTemp & "\" & sTemp
            End If
        End If
        sTemp = Dir$()
    Wend

    gsDirsQueue.Remove (1)
Wend

Close
Liberado
MsgBox "Lista dos arquivos bancários de " & cmbAno.Text & ", foi reconstruida.", vbInformation, "Informação"
    
End Sub

Private Sub cmdStop_Click()
bContinueRun = False
End Sub

Private Sub Form_Load()
Dim x As Integer

Centraliza Me
'sPathArqBanco = GetSetting("GTI", "PATH", "ARQUIVOBANCO")
txtPath.Text = sPathArqBanco

For x = 2003 To Year(Now)
    cmbAno.AddItem x
Next
cmbAno.ListIndex = cmbAno.ListCount - 1

End Sub

Private Function Valida() As Boolean
Valida = False
If Val(cmbDe.Text) > Val(cmbAte.Text) Then
    MsgBox "Ano inicial maior que final.", vbCritical, "Atenção"
    Exit Function
End If

If Val(txtNumDoc.Text) = 0 Then
    MsgBox "Digite o número do documento.", vbCritical, "Atenção"
    Exit Function
End If

Valida = True
End Function

Private Sub CallPb(nVal As Long, nTot As Long)

If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveSetting "GTI", "PATH", "ARQUIVOBANCO", txtPath.Text

End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
Tweak txtDocumento, KeyAscii, IntegerPositive
End Sub

Private Sub txtPath_LostFocus()
LoadFileBanco
End Sub

Private Sub LeArquivo(sFullPath As String)
Dim sPath As String, sAno As String, aAno() As String, x As Integer, m As Integer, f As Integer, d As Integer
Dim sDia As String, sMes As String, nMesDe As Integer, nMesAte As Integer, sArquivo As String
Dim bDA As Boolean, sRetorno As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String
Dim Header As FebrabanA
Dim Header2 As FebrabanA2
Dim Registro As FebrabanG
Dim Registro2 As FebrabanG2
Dim RegistroF As FebrabanF
Dim Registro2U As FebrabanG2U
Dim Posicao  As Long, nTot As Long, nCount As Long
Dim sNumDoc As String, nNumDocBusca As Long, bAchou As Boolean, k As Long
Dim Header3 As FebrabanA3, Registro3 As FebrabanG3, Footer3 As FebrabanZ3

nNumDocBusca = Val(txtDocumento.Text)

Ocupado
    ReDim aFebrabanA(0): ReDim aFebrabanG(0): ReDim aFebrabanF(0): ReDim aFebrabanA2(0): ReDim aFebrabanG2(0)
    ReDim aFebrabanA3(0): ReDim aFebrabanG3(0): ReDim aFebrabanZ3(0)
    
    On Error Resume Next
    Close #1
   On Error GoTo 0
    
    Open sFullPath For Binary Access Read As #1
        Get #1, 1, Header
        aFebrabanA(0).CodigoRegistro = Trim$(Header.CodigoRegistro)
        If aFebrabanA(0).CodigoRegistro <> "A" Then
            Close #1
            GoTo BANESPA2
        End If
        aFebrabanA(0).CodigoBanco = Trim$(Header.CodigoBanco)
        aFebrabanA(0).NomeBanco = Trim$(Header.NomeBanco)
        aFebrabanA(0).DataGeracao = Trim$(Header.DataGeracao)
        aFebrabanA(0).Filler = Trim$(Header.Filler)
        If Left$(aFebrabanA(0).Filler, 10) = "DEBITO AUT" Or Left$(aFebrabanA(0).Filler, 10) = "DÉBITO AUT" Then
            Close #1
            bDA = True
            GoTo DEBITOAUTO
        Else
            bDA = False
        End If
        Posicao = Len(Header) + 3
        nCount = 0
        Do While Not EOF(1)
             Get #1, Posicao, Registro
             If Registro.CodigoRegistro <> "Z" Then
                aFebrabanG(nCount).DataPagamento = Registro.DataPagamento
                aFebrabanG(nCount).DataCredito = Registro.DataCredito
                aFebrabanG(nCount).DataVencto = Registro.DataVencto
                aFebrabanG(nCount).NumDocumento = Registro.NumDocumento
                aFebrabanG(nCount).NumParcela = Registro.NumParcela
                aFebrabanG(nCount).SituacaoRetorno = Registro.SituacaoRetorno
                aFebrabanG(nCount).ValorRetornado = Registro.ValorRetornado
                aFebrabanG(nCount).ValorTarifa = Registro.ValorTarifa
                aFebrabanG(nCount).CodAgencia = Registro.CodAgencia
                If Val(aFebrabanG(nCount).NumDocumento) = nNumDocBusca Then
                    bAchou = True
                    PBar.value = 0
                    lblMsg.Caption = "Documento Localizado !!!"
                    lblMsg.Refresh
                    'txtArq.Text = RetornaArquivo(lstArq.List(x))
                    txtBanco.Text = aFebrabanA(0).CodigoBanco & "-" & aFebrabanA(0).NomeBanco
                    txtDataCredito.Text = ConvDataSerial(aFebrabanG(nCount).DataCredito)
                    txtValorPago.Text = FormatNumber(CDbl(aFebrabanG(nCount).ValorRetornado) / 100, 2)
                    GoTo fim3
                    
                End If
             End If
             Posicao = Posicao + Len(Registro) + 2
             nCount = nCount + 1
             ReDim Preserve aFebrabanG(nCount)
        Loop
        
     Close #1
    
    
BANESPA2:
Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header2
    aFebrabanA2(0).TipoRegistro = Trim$(Header2.TipoRegistro)
    aFebrabanA2(0).DataGeracao = Trim$(Header2.DataGeracao)
    
    Posicao = Len(Header2) + 3
    nCount = 0: nQtde = 0: nValor = 0
    On Error Resume Next
    Do While Not EOF(1)
         Get #1, Posicao, Registro2
        If Registro2.TipoRegistro = 1 Or Registro2.TipoRegistro = 9 Then GoTo PROXIMO2
         If Registro2.TipoRegistro = 5 Then GoTo Rodape
         
              aFebrabanG2(nCount).TipoRegistro = Registro2.TipoRegistro
              aFebrabanG2(nCount).ContaCobranca = Registro2.NumeroConta
              aFebrabanG2(nCount).DataVencto = Registro2.DataVencto
              aFebrabanG2(nCount).ValorTitulo = (CDbl(FormatNumber(Registro2.ValorTitulo, 2))) / 100
              aFebrabanG2(nCount).NossoNumero = Registro2.NossoNumero
              aFebrabanG2(nCount).ValorTitulo = (CDbl(FormatNumber(Registro2.ValorTitulo, 2))) / 100
              aFebrabanG2(nCount).ValorTarifa = Registro2.ValorTarifa
              aFebrabanG2(nCount).NumSequencial = Registro2.NumSequencial
              aFebrabanG2(nCount).AgenciaCobradora = Registro2.AgenciaCobradora
              If Val(aFebrabanG2(nCount).NossoNumero) = nNumDocBusca Then
                    bAchou = True
                    PBar.value = 0
                    lblMsg.Caption = "Documento Localizado !!!"
                    lblMsg.Refresh
                    'txtArq.Text = RetornaArquivo(lstArq.List(x))
                    Posicao = Posicao + Len(Registro2) + 2
                    Get #1, Posicao, Registro2U
                    aFebrabang2U(nCount).DataCredito = Registro2U.DataCredito
                    aFebrabang2U(nCount).ValorCreditado = (CDbl(FormatNumber(Registro2U.ValorCreditado, 2))) / 100
                    txtDataCredito.Text = Left(aFebrabang2U(nCount).DataCredito, 2) & "/" & Mid(aFebrabang2U(nCount).DataCredito, 3, 2) & "/" & Right(aFebrabang2U(nCount).DataCredito, 4)
                    txtValorPago.Text = FormatNumber(aFebrabang2U(nCount).ValorCreditado, 2)
                    Close #1
                    GoTo fim3
              End If
              Posicao = Posicao + Len(Registro2) + 2
              Get #1, Posicao, Registro2U
              aFebrabang2U(nCount).DataCredito = Registro2U.DataCredito
              aFebrabang2U(nCount).ValorCreditado = (CDbl(FormatNumber(Registro2U.ValorCreditado, 2))) / 100
              nQtde = nQtde + 1
              nValor = nValor + CDbl(aFebrabang2U(nCount).ValorCreditado)
         nCount = nCount + 1
         ReDim Preserve aFebrabanG2(nCount)
         ReDim Preserve aFebrabang2U(nCount)
PROXIMO2:
         Posicao = Posicao + Len(Registro2) + 2
    Loop
Rodape:

 Close #1
 
BANESPA3:
Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header3
    aFebrabanA3(0).CodigoRegistro = Trim$(Header3.CodigoRegistro)
    aFebrabanA3(0).CodigoRetorno = Trim$(Header3.CodigoRetorno)
    aFebrabanA3(0).CodigoCobrança = Trim$(Header3.CodigoCobrança)
    aFebrabanA3(0).NomeCobrança = Trim$(Header3.NomeCobrança)
    aFebrabanA3(0).CodigoPrefeitura = Trim$(Header3.CodigoPrefeitura)
    aFebrabanA3(0).CodigoBanco = Trim$(Header3.CodigoBanco)
    aFebrabanA3(0).NomeBanco = Trim$(Header3.NomeBanco)
    aFebrabanA3(0).DataCredito = Trim$(Header3.DataCredito)
    aFebrabanA3(0).Filler1 = Trim$(Header3.Filler1)
    aFebrabanA3(0).Filler2 = Trim$(Header3.Filler2)
    aFebrabanA3(0).NumRegistros = Trim$(Header3.NumRegistros)
    
    Posicao = Len(Header3) + 3
    nCount = 0
    On Error Resume Next
    Do While Not EOF(1)
         Get #1, Posicao, Registro3
         If Registro3.CodigoRegistro <> "9" Then
              aFebrabanG3(nCount).CodigoRegistro = Registro3.CodigoRegistro
              aFebrabanG3(nCount).ContaPrefeitura = Registro3.ContaPrefeitura
              aFebrabanG3(nCount).DataPagamento = Registro3.DataPagamento
              aFebrabanG3(nCount).ValorPago = (CDbl(FormatNumber(Registro3.ValorPago, 2))) / 100
              aFebrabanG3(nCount).NumDocumento = Registro3.NumDocumento
              aFebrabanG3(nCount).ValorPago = (CDbl(FormatNumber(Registro3.ValorPago, 2))) / 100
              aFebrabanG3(nCount).ValorTaxa = Registro3.ValorTaxa
              aFebrabanG3(nCount).NumSeq = Registro3.NumSeq
              aFebrabanG3(nCount).CodAgencia = Registro3.CodAgencia
 
              If Val(Left(aFebrabanG3(nCount).NumDocumento, Len(aFebrabanG3(nCount).NumDocumento) - 1)) = nNumDocBusca Then
                    bAchou = True
                    PBar.value = 0
                    lblMsg.Caption = "Documento Localizado !!!"
                    lblMsg.Refresh
                    'txtArq.Text = RetornaArquivo(lstArq.List(x))
                    txtBanco.Text = aFebrabanA3(0).CodigoBanco & "-" & aFebrabanA3(0).NomeBanco
                    txtDataCredito.Text = ConvDataSerial(aFebrabanA3(0).DataCredito)
                    txtValorPago.Text = FormatNumber(aFebrabanG3(nCount).ValorPago, 2)
                    Close #1
                    GoTo fim3
                
              End If
 
         Else
              Get #1, Posicao, Footer3
              txtBanco.Text = aFebrabanA3(0).CodigoBanco & "-" & aFebrabanA3(0).NomeBanco
              txtDataCredito.Text = ConvDataSerial(aFebrabanA3(0).DataCredito)
              txtValorPago.Text = FormatNumber(Footer3.ValorTotal / 100, 2)
              aFebrabanZ3(0).CodigoSeiLa = Footer3.CodigoSeiLa
              aFebrabanZ3(0).Filler1 = Footer3.Filler1
              aFebrabanZ3(0).TotalRegistro = Footer3.TotalRegistro
              aFebrabanZ3(0).ValorTotal = Footer3.ValorTotal
              Exit Do
         End If
         Posicao = Posicao + Len(Registro3) + 2
         nCount = nCount + 1
         ReDim Preserve aFebrabanG3(nCount)
    
    Loop
 Close #1

Fim2:
    
DEBITOAUTO:
    Open sFullPath For Binary Access Read As #1
       Get #1, 1, Header
       aFebrabanA(0).CodigoBanco = Header.CodigoBanco
       aFebrabanA(0).NomeBanco = Header.NomeBanco
       aFebrabanA(0).DataGeracao = Header.DataGeracao
       Posicao = Len(Header) + 3
       nCount = 0
       Do While Not EOF(1)
            Get #1, Posicao, RegistroF
            If Left$(RegistroF.CodigoRegistro, 1) = "F" Or Left$(RegistroF.CodigoRegistro, 1) = "E" Then
                 Get #1, Posicao, RegistroF
                 aFebrabanF(nCount).Distrito = RegistroF.Distrito
                 aFebrabanF(nCount).Setor = RegistroF.Setor
                 aFebrabanF(nCount).Quadra = RegistroF.Quadra
                 aFebrabanF(nCount).Lote = RegistroF.Lote
                 aFebrabanF(nCount).Seq = RegistroF.Seq
                 aFebrabanF(nCount).ContaCliente = RegistroF.ContaCliente
                 aFebrabanF(nCount).DataVencto = RegistroF.DataVencto
                 aFebrabanF(nCount).ValorDebito = RegistroF.ValorDebito
                 aFebrabanF(nCount).CodRetorno = RegistroF.CodRetorno
                 aFebrabanF(nCount).NumDoc = RegistroF.NumDoc
                 aFebrabanF(nCount).CodMovimento = RegistroF.CodMovimento
                 
                 With aFebrabanF(nCount)
                       Select Case .CodRetorno
                               Case "00"
                                       sRetorno = "Débito Efetuado"
                               Case "01"
                                       sRetorno = "Insuficiência de Fundos"
                               Case "02"
                                       sRetorno = "Conta Corrente não Cadastrada"
                               Case "04"
                                       sRetorno = "Outras Restrições"
                               Case "10"
                                       sRetorno = "Agência em Regime de Encerramento"
                               Case "12"
                                       sRetorno = "Valor Inválido"
                               Case "13"
                                       sRetorno = "Data de Lançamento inválida"
                               Case "14"
                                       sRetorno = "Agência Inválida"
                               Case "15"
                                       sRetorno = "DAC da conta corrente inválido"
                               Case "18"
                                       sRetorno = "Data do Débito anterior ao do processamento"
                               Case "30"
                                       sRetorno = "Sem contrato de débito automático"
                               Case "96"
                                       sRetorno = "Manutenção do Cadastro"
                               Case "97"
                                       sRetorno = "Cancelamento - Não Encontrado"
                               Case "98"
                                       sRetorno = "Cancelamento - não efetuado, fora de tempo habil"
                               Case "99"
                                       sRetorno = "Cancelamento - cancelado conforme solicitado"
                               Case Else
                                      sRetorno = "Erro Indefinido"
                       End Select
                       
                        If aFebrabanF(nCount).NumDoc = nCodReduz Then
                            bAchou = True: PBar.value = 0
                            lblMsg.Caption = "Documento Localizado !!!"
                            lblMsg.Refresh
                            'txtArq.Text = RetornaArquivo(lstArq.List(x))
                            txtBanco.Text = aFebrabanA(0).CodigoBanco & "-" & aFebrabanA(0).NomeBanco
                            txtDataCredito.Text = ConvDataSerial(aFebrabanF(nCount).DataVencto)
                            txtValorPago.Text = FormatNumber(CDbl(aFebrabanF(nCount).ValorDebito) / 100, 2)
                            txtDA.Text = aFebrabanF(nCount).CodRetorno & " - " & sRetorno
                            Close #1
                            GoTo fim3
                        End If
                 End With
                 ReDim Preserve aFebrabanF(nCount)
            End If
            Posicao = Posicao + Len(Registro) + 2
            nCount = nCount + 1
            ReDim Preserve aFebrabanF(nCount)
       Loop
    Close #1
'PROXIMO:
'Next x
lblMsg.Caption = "Documento não encontrado !!!"
lblMsg.Refresh

fim3:

Liberado
Exit Sub
Erro:
 MsgBox Err.Description
On Error Resume Next
Close #1
Liberado

End Sub
