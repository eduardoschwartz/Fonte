VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmFichaCompensacaoLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de arquivos para registro bancário"
   ClientHeight    =   5880
   ClientLeft      =   6570
   ClientTop       =   3045
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   12030
   Begin MSComctlLib.ListView lvOrigem 
      Height          =   1785
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   3149
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Documento"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt.Vencto"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Geração"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CPF"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Endereço"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bairro"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cep"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Cidade"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "UF"
         Object.Width           =   953
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdGerarArquivo 
      Height          =   345
      Left            =   150
      TabIndex        =   2
      Top             =   2220
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Gerar arquivo de lote para registro"
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
      BCOL            =   16777152
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmFichaCompensacaoLote.frx":0000
      PICN            =   "frmFichaCompensacaoLote.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvLote 
      Height          =   2295
      Left            =   60
      TabIndex        =   3
      Top             =   3030
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Seq"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Arquivo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Comprovante"
         Object.Width           =   2820
      EndProperty
   End
   Begin MSComctlLib.ListView lvConteudo 
      Height          =   2295
      Left            =   5670
      TabIndex        =   6
      Top             =   3030
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Documento"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt.Vencto"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Geração"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CPF"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Endereço"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bairro"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cep"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Cidade"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "UF"
         Object.Width           =   1058
      EndProperty
   End
   Begin prjChameleon.chameleonButton btAnexar 
      Height          =   345
      Left            =   2400
      TabIndex        =   7
      ToolTipText     =   "Anexar ao lote o comprovante de transmissão"
      Top             =   5430
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Anexar comprovante"
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
      BCOL            =   16777152
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmFichaCompensacaoLote.frx":0098
      PICN            =   "frmFichaCompensacaoLote.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btGravar 
      Height          =   345
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Gravar o arquivo de lote em um local para ser transmitido para o banco"
      Top             =   5430
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Exportar arquivo"
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
      BCOL            =   16777152
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFichaCompensacaoLote.frx":020E
      PICN            =   "frmFichaCompensacaoLote.frx":022A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btVisualizar 
      Height          =   345
      Left            =   4680
      TabIndex        =   9
      ToolTipText     =   "Visualizar o comprovante de transmissão"
      Top             =   5430
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Visualizar comprovante"
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
      BCOL            =   16777152
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmFichaCompensacaoLote.frx":05CF
      PICN            =   "frmFichaCompensacaoLote.frx":05EB
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
      Caption         =   "Documentos contidos no lote"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   5730
      TabIndex        =   5
      Top             =   2790
      Width           =   2445
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lotes cadastrados"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2790
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos a serem enviados para registro"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   6135
   End
End
Attribute VB_Name = "frmFichaCompensacaoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type HeaderArquivo
    nCodigoBanco As String
    nLote As String
    nTipoRegistro As String
    sUsoFebraban1 As String
    nTipoInscricao As String
    nNumInscricao As String
    sCodigoConvenio As String
    nAgencia As String
    sDvAgencia As String
    nNumeroConta As String
    sDvConta As String
    sDvAgenciaConta As String
    sNomeEmpresa As String
    sNomeBanco As String
    sUsoFebraban2 As String
    nCodigoRemessa As String
    nDataGeracao As String
    nHoraGeracao As String
    nNumSeqArquivo As String
    nNumVersaoLayout As String
    nDensidade As String
    sUsoBanco As String
    sUsoEmpresa As String
    sUsoFebraban3 As String
End Type

Private Type TrailerArquivo
    nCodigoBanco As String
    nLote As String
    nTipo As String
    sUsoFebraban1 As String
    nQtdeLote As String
    nQtdeRegistro As String
    nQtdeContas As String
    sUsoFebraban2 As String
End Type

Private Type HeaderLote
    nCodigoBanco As String
    nLote As String
    nTipoRegistro As String
    sTipoOperacao As String
    nTipoServico As String
    sUsoFebraban1 As String
    nNumVersao As String
    sUsoFebraban2 As String
    nTipoInscricao As String
    nNumInscricao As String
    sCodConvenio As String
    nAgencia As String
    sDvAgencia As String
    nNumeroConta As String
    sDvConta As String
    sDvAgenciaConta As String
    sNomeEmpresa As String
    sMensagem1 As String
    sMensagem2 As String
    nNumeroRemessa As String
    sDataGeracao As String
    sDataCredito As String
    sUsoFebraban3 As String
End Type

Private Type TrailerLote
    nCodigoBanco As String
    nLote As String
    nTipo As String
    sUsoFebraban1 As String
    nQtdeRegistro As String
    sUsoFebraban2 As String
End Type

Private Type SegmentoP
    nCodigoBanco As String
    nLote As String
    nTipo As String
    nSeqReg As String
    sCodSegmento As String
    sUsoFebraban1 As String
    nCodMovimento As String
    nAgencia As String
    sDvAgencia As String
    nConta As String
    sDvConta As String
    sNossoNumero As String
    nCodCarteira As String
    nFormaCadastro As String
    sTipoDocumento As String
    nIdentificacaoEmissao As String
    sIdentificacaoDistribuicao As String
    sNumeroDocumento As String
    nDataVencimento As String
    nValorNominal As String
    nAgenciaCobranca As String
    sDvAgenciaCobranca As String
    nEspecieTitulo As String
    sAceite As String
    nDataEmissao As String
    nCodigoJuros As String
    nDataJuros As String
    nJurosMora As String
    nCodigoDesconto1 As String
    nDataDesconto1 As String
    nValorConcedido As String
    nValorIOF As String
    nValorAbatimento As String
    sIdentificaTitulo As String
    nCodigoProtesto As String
    nNumDiasProtesto As String
    nCodigoBaixa As String
    sNumDiasBaixa As String
    nCodigoMoeda As String
    nNumeroContrato As String
    sUsoLivre As String
End Type

Private Type SegmentoQ
    nCodigoBanco As String
    nLote As String
    nTipo As String
    nSeqReg As String
    sCodSegmento As String
    sUsoFebraban1 As String
    nCodMovimento As String
    nTipoInscricao As String
    nNumeroInscricao As String
    sNome As String
    sEndereco As String
    sBairro As String
    nCep As String
    nCepsufixo As String
    sCidade As String
    sUF As String
    nipoInscricaoSacado As String
    nNumeroInscricaoSacado As String
    sNomeSacado As String
    nBancoCorresponde As String
    sNossoNumeroBancoCorr As String
    sUsoFebraban2 As String
End Type

Private Type Boletos
    sNossoNumero As String
    sNumDocumento As String
    sDataVencimento As String
    sValorNominal As String
    sDataBase As String
    nTipoInscricao As String
    nNumeroInscricao As String
    sNome As String
    sEndereco As String
    sBairro As String
    sCep As String
    sSufixoCep As String
    sCidade As String
    sUF As String
End Type

Dim nNumRemessa As Long, sArquivo As String, sDataArquivo As String, aBoletos() As Boletos

Private Sub btAnexar_Click()
Dim sPathDestino As String, sPathSandBox As String, fso As New FileSystemObject, fName As String, cc As cCommonDlg, sFileName As String, nNumero_Lote As Integer, sData_Lote As String

If lvLote.ListItems.Count = 0 Then Exit Sub
If lvLote.SelectedItem.Index < -0 Then
    MsgBox "Selecione um lote para exportar.", vbOKOnly, "Erro"
    Exit Sub
End If

sData_Lote = lvLote.SelectedItem.Text
nNumero_Lote = Val(lvLote.SelectedItem.SubItems(1))
sFileName = Left(lvLote.SelectedItem.SubItems(3), 3) & "12" & Right(lvLote.SelectedItem.SubItems(3), 8)

sPathDestino = sPathAnexo & "12\" & Right(lvLote.SelectedItem.SubItems(3), 4) & "\" & Mid(lvLote.SelectedItem.SubItems(3), 8, 2) & "\" & sFileName

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Comprovante (*.pdf)|*.pdf", , sPathBin, "Selecione o comprovante para anexar", , Me.HWND, OFN_HIDEREADONLY, False
If fName <> "" Then
    fso.CopyFile fName, sPathDestino, True
    Sql = "update ficha_compensacao_lote set comprovante='" & sFileName & "' where data_lote='" & Format(sData_Lote, "mm/dd/yyyy") & "' and numero_lote=" & nNumero_Lote
    cn.Execute Sql, rdExecDirect
    lvLote.SelectedItem.SubItems(4) = sFileName
Else
    MsgBox "Operação cancelada.", vbInformation, "Atenção"
End If

End Sub

Private Sub btGravar_Click()
Dim sPathOrigem As String, sPathSandBox As String, fso As New FileSystemObject, fName As String, cc As cCommonDlg

If lvLote.ListItems.Count = 0 Then Exit Sub
If lvLote.SelectedItem.Index < -0 Then
    MsgBox "Selecione um lote para exportar.", vbOKOnly, "Erro"
    Exit Sub
End If

sPathOrigem = sPathAnexo & "11\" & Right(lvLote.SelectedItem.SubItems(3), 4) & "\" & Mid(lvLote.SelectedItem.SubItems(3), 8, 2) & "\" & lvLote.SelectedItem.SubItems(3)

Set cc = New cCommonDlg
If cc.VBGetSaveFileName(fName, "", , "*.txt", , sPathBin, "Local para salvar o arquivo", "txt", , 0) Then
    sPathSandBox = fName
    fso.CopyFile sPathOrigem, fName, True
End If

End Sub

Private Sub btVisualizar_Click()
Dim sPathOrigem As String, sFileName As String, sPathSandBox As String, fso As New FileSystemObject

If lvLote.ListItems.Count = 0 Then Exit Sub
If lvLote.SelectedItem.Index < -0 Then
    MsgBox "Selecione um lote para exportar.", vbOKOnly, "Erro"
    Exit Sub
End If

If lvLote.SelectedItem.SubItems(4) = "" Then
    MsgBox "Não existe comprovante anexado a este lote.", vbCritical, "Atenção"
Else
    sFileName = lvLote.SelectedItem.SubItems(4)
    sPathOrigem = sPathAnexo & "12\" & Right(lvLote.SelectedItem.SubItems(3), 4) & "\" & Mid(lvLote.SelectedItem.SubItems(3), 8, 2) & "\" & sFileName
    sPathSandBox = App.Path & "\" & "Sandbox"
    If fso.FolderExists(sPathSandBox) = False Then
        fso.CreateFolder (sPathSandBox)
    End If
    
    If fso.FileExists(sPathOrigem) Then
        fso.CopyFile sPathOrigem, sPathSandBox & "\" & sFileName & ".pdf", True
        ShellExecute 0&, "open", sPathSandBox & "\" & sFileName & ".pdf", vbNullString, vbNullString, conSwNormal
    Else
        MsgBox "Arquivo não localizado.", vbCritical, "Erro"
    End If
End If

End Sub

Private Sub Form_Load()
Centraliza Me
Fill
End Sub

Private Sub Fill()
Ocupado
CarregaOrigem
CarregaLote
If lvLote.ListItems.Count > 0 Then
    lvLote.ListItems(1).Selected = True
    lvLote_Click
    On Error Resume Next
    lvLote.SetFocus
End If
Liberado
End Sub

Private Sub CarregaOrigem()
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem

lvOrigem.ListItems.Clear
Sql = "select numero_documento,data_vencimento,valor_documento,datadocumento,nome,cpf,endereco,bairro,cep,cidade,uf "
Sql = Sql & "from ficha_compensacao_documento inner join numdocumento on ficha_compensacao_documento.numero_documento=numdocumento.numdocumento "
Sql = Sql & "where data_lote is null order by numero_documento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvOrigem.ListItems.Add(, , !numero_documento)
        itmX.SubItems(1) = Format(!Data_Vencimento, "dd/mm/yyyy")
        itmX.SubItems(2) = Format(!Datadocumento, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(!valor_documento, "#0.00")
        itmX.SubItems(4) = !Nome
        itmX.SubItems(5) = !CPF
        itmX.SubItems(6) = !Endereco
        itmX.SubItems(7) = !Bairro
        itmX.SubItems(8) = !Cep
        itmX.SubItems(9) = SubNull(!Cidade)
        itmX.SubItems(10) = SubNull(!UF)
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaLote()

Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem

lvLote.ListItems.Clear
Sql = "select data_lote,numero_lote,quantidade,arquivo,comprovante from ficha_compensacao_lote order by data_lote desc,numero_lote desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvLote.ListItems.Add(, , Format(!Data_lote, "dd/mm/yyyy"))
        itmX.SubItems(1) = Format(!numero_lote, "000")
        itmX.SubItems(2) = Format(!quantidade, "0000")
        itmX.SubItems(3) = SubNull(!Arquivo)
        itmX.SubItems(4) = SubNull(!comprovante)
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub lvLote_Click()
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem, nNumero_Lote As Integer, sData_Lote As String
Ocupado
If lvLote.ListItems.Count = 0 Then Exit Sub
sData_Lote = lvLote.SelectedItem.Text
nNumero_Lote = Val(lvLote.SelectedItem.SubItems(1))

lvConteudo.ListItems.Clear
Sql = "select numero_documento,data_vencimento,valor_documento,datadocumento,nome,cpf,endereco,bairro,cep,cidade,uf "
Sql = Sql & "from ficha_compensacao_documento inner join numdocumento on ficha_compensacao_documento.numero_documento=numdocumento.numdocumento "
Sql = Sql & "where data_lote='" & Format(sData_Lote, "mm/dd/yyyy") & "' and numero_lote=" & nNumero_Lote & " order by numero_documento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvConteudo.ListItems.Add(, , !numero_documento)
        itmX.SubItems(1) = Format(!Data_Vencimento, "dd/mm/yyyy")
        itmX.SubItems(2) = Format(!Datadocumento, "dd/mm/yyyy")
        itmX.SubItems(3) = Format(!valor_documento, "#0.00")
        itmX.SubItems(4) = !Nome
        itmX.SubItems(5) = !CPF
        itmX.SubItems(6) = !Endereco
        itmX.SubItems(7) = !Bairro
        itmX.SubItems(8) = !Cep
        itmX.SubItems(9) = SubNull(!Cidade)
        itmX.SubItems(10) = SubNull(!UF)
       .MoveNext
    Loop
   .Close
End With
Liberado
End Sub

Private Sub cmdGerarArquivo_Click()
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem, nSeq As Integer, nNumdoc As Long, x As Integer, sArquivo As String
Dim sCPFCNPJ As String, dDataVencto As Date, nValorGuia As Double, nTipoDoc As Integer, sNome As String, sEndereco As String
Dim sBairro As String, sCep As String, sCidade As String, sUF As String, nPosReg As Long, nContador As Long, nInicio As Integer
Dim aHeaderArquivo() As HeaderArquivo, FF1 As Integer, sHeaderArquivo As String, fso As New FileSystemObject, sPath As String
Dim aTrailerArquivo() As TrailerArquivo, sTrailerArquivo As String, nQtdeRegistroArquivo As Long, nQtdeRegistroLote As Long, aHeaderLote() As HeaderLote, sHeaderLote As String
Dim aTrailerLote() As TrailerLote, sTrailerLote As String, aSegmentoP() As SegmentoP, sSegmentoP As String, aSegmentoQ() As SegmentoQ, sSegmentoQ As String, sFileName As String

If MsgBox("Gerar o arquivo com estes documentos?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

nSeq = 1
Sql = "select max(numero_lote) as maximo from ficha_compensacao_lote where data_lote='" & Format(Now, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If Not IsNull(RdoAux!maximo) Then
    nSeq = RdoAux!maximo + 1
End If
RdoAux.Close

sPath = sPathAnexo & "12"
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If
sPath = sPathAnexo & "12\" & Format(Year(Now), "0000")
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If
sPath = sPathAnexo & "12\" & Format(Year(Now), "0000") & "\" & Format(Month(Now), "00")
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If

sPath = sPathAnexo & "11"
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If
sPath = sPathAnexo & "11\" & Format(Year(Now), "0000")
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If
sPath = sPathAnexo & "11\" & Format(Year(Now), "0000") & "\" & Format(Month(Now), "00")
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If


sFileName = Format(nSeq, "000") & "11" & RetornaNumero(Format(Now, "dd/mm/yyyy"))
sArquivo = sPath & "\" & sFileName
ReDim aBoletos(0)

Sql = "insert ficha_compensacao_lote(data_lote,numero_lote,quantidade,arquivo,userid) values('" & Format(Now, "mm/dd/yyyy") & "',"
Sql = Sql & nSeq & "," & lvOrigem.ListItems.Count & ",'" & sFileName & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect

For x = 1 To lvOrigem.ListItems.Count
    nNumdoc = CLng(lvOrigem.ListItems(x).Text)
    dDataVencto = CDate(lvOrigem.ListItems(x).SubItems(1))
    nValorGuia = CDbl(lvOrigem.ListItems(x).SubItems(3))
    sNome = lvOrigem.ListItems(x).SubItems(4)
    sCPFCNPJ = lvOrigem.ListItems(x).SubItems(5)
    sEndereco = lvOrigem.ListItems(x).SubItems(6)
    sBairro = lvOrigem.ListItems(x).SubItems(7)
    sCep = lvOrigem.ListItems(x).SubItems(8)
    sCidade = lvOrigem.ListItems(x).SubItems(9)
    sUF = lvOrigem.ListItems(x).SubItems(10)
    
    Sql = "update ficha_compensacao_documento set data_lote='" & Format(Now, "mm/dd/yyyy") & "',numero_lote=" & nSeq & " where numero_documento=" & nNumdoc
    cn.Execute Sql, rdExecDirect
    
    If Len(sCPFCNPJ) = 11 Then
        nTipoDoc = 1
    Else
        nTipoDoc = 2
    End If
    
    ReDim Preserve aBoletos(UBound(aBoletos) + 1)
    aBoletos(UBound(aBoletos)).sNossoNumero = "287353200" & CStr(nNumdoc)
    aBoletos(UBound(aBoletos)).sNumDocumento = Format(nNumdoc, "00000000")
    aBoletos(UBound(aBoletos)).sDataVencimento = Format(dDataVencto, "dd") & Format(dDataVencto, "mm") & Format(dDataVencto, "yyyy")
    aBoletos(UBound(aBoletos)).sValorNominal = FillLeft(RetornaNumero(CStr(nValorGuia * 100)), 15)
    If Val(aBoletos(UBound(aBoletos)).sValorNominal) <> RetornaNumero(FormatNumber(nValorGuia, 2)) Then
        MsgBox "Erro na geração do arquivo", vbCritical, "Alerta"
        Exit Sub
    End If
    aBoletos(UBound(aBoletos)).sDataBase = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
    aBoletos(UBound(aBoletos)).nTipoInscricao = CStr(nTipoDoc)
    aBoletos(UBound(aBoletos)).nNumeroInscricao = sCPFCNPJ
    aBoletos(UBound(aBoletos)).sNome = Left(sNome, 40)
    aBoletos(UBound(aBoletos)).sEndereco = Left(sEndereco, 40)
    aBoletos(UBound(aBoletos)).sBairro = Left(sBairro, 15)
    aBoletos(UBound(aBoletos)).sCep = Left(sCep, 5)
    aBoletos(UBound(aBoletos)).sSufixoCep = Right(sCep, 3)
    aBoletos(UBound(aBoletos)).sCidade = sCidade
    aBoletos(UBound(aBoletos)).sUF = sUF
Next

FF1 = FreeFile()
Open sArquivo For Output As FF1

'*********************************
'******** Header Arquivo *********
'*********************************
ReDim aHeaderArquivo(1)
With aHeaderArquivo(1)
    .nCodigoBanco = "001"
    .nLote = "0000"
    .nTipoRegistro = "0"
    .sUsoFebraban1 = FillSpace(" ", 9)
    .nTipoInscricao = "2"
    .nNumInscricao = "50387844000105"
    .sCodigoConvenio = FillLeft("2873532", 9) & "001417019  "
    .nNumeroConta = FillLeft("74000", 12)
    .sDvConta = "4 "
    .nAgencia = "00269"
    .sDvAgencia = "0"
    .sNomeEmpresa = FillSpace("PREFEITURA MUN. DE JABOTICABAL", 30)
    .sNomeBanco = FillSpace("BANCO DO BRASIL S.A.", 30)
    .sUsoFebraban2 = FillSpace(" ", 10)
    .nCodigoRemessa = "1"
    .nDataGeracao = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
    .nHoraGeracao = Format(Now, "hhmmss")
    .nNumSeqArquivo = "000000"
    .nNumVersaoLayout = "000"
    .nDensidade = "00000"
    .sUsoBanco = FillSpace(" ", 20)
    .sUsoEmpresa = FillSpace(" ", 20)
    .sUsoFebraban3 = FillSpace(" ", 29)
    
    sHeaderArquivo = .nCodigoBanco & .nLote & .nTipoRegistro & .sUsoFebraban1 & .nTipoInscricao & .nNumInscricao & .sCodigoConvenio & .nAgencia & .sDvAgencia & .nNumeroConta & .sDvConta & .sNomeEmpresa
    sHeaderArquivo = sHeaderArquivo & .sNomeBanco & .sUsoFebraban2 & .nCodigoRemessa & .nDataGeracao & .nHoraGeracao & .nNumSeqArquivo & .nNumVersaoLayout & .nDensidade & .sUsoBanco & .sUsoEmpresa & .sUsoFebraban3
End With

Print #FF1, sHeaderArquivo
'****** Fim Header Arquivo ************

'*********************************
'******** Header do Lote *********
'*********************************
ReDim aHeaderLote(1)
With aHeaderLote(1)
    .nCodigoBanco = "001"
    .nLote = "0001"
    .nTipoRegistro = "1"
    .sTipoOperacao = "R"
    .nTipoServico = "01"
    .sUsoFebraban1 = "  "
    .nNumVersao = "000"
    .sUsoFebraban2 = " "
    .nTipoInscricao = "2"
    .nNumInscricao = "050387844000105"
    .sCodConvenio = FillLeft("2873532", 9) & "001417019  " 'IPTU/ISS/TXLIC
    .nNumeroConta = FillLeft("74000", 12)
    .sDvConta = "4 "
    .nAgencia = "00269"
    .sDvAgencia = "0"
    .sNomeEmpresa = FillSpace("PREFEITURA MUN. DE JABOTICABAL", 30)
    .sMensagem1 = FillSpace(" ", 40)
    .sMensagem2 = FillSpace(" ", 40)
    .nNumeroRemessa = FillLeft(CStr(nNumRemessa), 8)
    .sDataGeracao = Format(Now, "dd") & Format(Now, "mm") & Format(Now, "yyyy")
    .sDataCredito = "00000000"
    .sUsoFebraban3 = FillSpace(" ", 33)
    
    sHeaderLote = .nCodigoBanco & .nLote & .nTipoRegistro & .sTipoOperacao & .nTipoServico & .sUsoFebraban1 & .nNumVersao & .sUsoFebraban2 & .nTipoInscricao & .nNumInscricao
    sHeaderLote = sHeaderLote & .sCodConvenio & .nAgencia & .sDvAgencia & .nNumeroConta & .sDvConta & .sNomeEmpresa & .sMensagem1 & .sMensagem2 & .nNumeroRemessa & .sDataGeracao & .sDataCredito & .sUsoFebraban3
End With

Print #FF1, sHeaderLote
'****** Fim Header do Lote********

'*********************************
'******** Segmento P e Q *********
'*********************************

nContador = 1
'nInicio = 49999
For nPosReg = 1 To UBound(aBoletos)
    ReDim aSegmentoP(1): ReDim aSegmentoQ(1)
    With aSegmentoP(1)
    
        .nCodigoBanco = "001"
        .nLote = "0001"
        .nTipo = "3"
        .nSeqReg = FillLeft(CStr(nContador), 5)
        .sCodSegmento = "P"
        .sUsoFebraban1 = " "
        .nCodMovimento = "01"
        .nAgencia = "00269"
        .sDvAgencia = "0"
        .nConta = FillLeft("74000", 12)
        .sDvConta = "4 "
        .sNossoNumero = FillSpace(aBoletos(nPosReg).sNossoNumero, 20)
        .nCodCarteira = "7"
        .nFormaCadastro = "1"
        .sTipoDocumento = "1"
        .nIdentificacaoEmissao = "2"
        .sIdentificacaoDistribuicao = "2"
        .sNumeroDocumento = FillSpace(aBoletos(nPosReg).sNumDocumento, 15)
        .nDataVencimento = aBoletos(nPosReg).sDataVencimento
        .nValorNominal = FillLeft(aBoletos(nPosReg).sValorNominal, 15)
        .nAgenciaCobranca = "00000"
        .sDvAgenciaCobranca = "0"
        .nEspecieTitulo = "01"
        .sAceite = "N"
        .nDataEmissao = aBoletos(nPosReg).sDataBase
        .nCodigoJuros = "0"
        .nDataJuros = FillLeft("0", 8)
        .nJurosMora = FillLeft("0", 15)
        .nCodigoDesconto1 = "0"
        .nDataDesconto1 = FillLeft("0", 8)
        .nValorConcedido = FillLeft("0", 15)
        .nValorIOF = FillLeft("0", 15)
        .nValorAbatimento = FillLeft("0", 15)
        .sIdentificaTitulo = FillSpace(aBoletos(nPosReg).sNumDocumento, 25)
        .nCodigoProtesto = "3"
        .nNumDiasProtesto = "00"
        .nCodigoBaixa = "0"
        .sNumDiasBaixa = "000"
        .nCodigoMoeda = "09"
        .nNumeroContrato = FillLeft("19663033", 10)
        .sUsoLivre = " "
        
        sSegmentoP = .nCodigoBanco & .nLote & .nTipo & .nSeqReg & .sCodSegmento & .sUsoFebraban1 & .nCodMovimento & .nAgencia & .sDvAgencia & .nConta & .sDvConta & .sNossoNumero
        sSegmentoP = sSegmentoP & .nCodCarteira & .nFormaCadastro & .sTipoDocumento & .nIdentificacaoEmissao & .sIdentificacaoDistribuicao & .sNumeroDocumento & .nDataVencimento
        sSegmentoP = sSegmentoP & .nValorNominal & .nAgenciaCobranca & .sDvAgenciaCobranca & .nEspecieTitulo & .sAceite & .nDataEmissao & .nCodigoJuros & .nDataJuros & .nJurosMora
        sSegmentoP = sSegmentoP & .nCodigoDesconto1 & .nDataDesconto1 & .nValorConcedido & .nValorIOF & .nValorAbatimento & .sIdentificaTitulo & .nCodigoProtesto & .nNumDiasProtesto & .nCodigoBaixa
        sSegmentoP = sSegmentoP & .sNumDiasBaixa & .nCodigoMoeda & .nNumeroContrato & .sUsoLivre
        On Error Resume Next
        Sql = "insert registro_cobranca (numdocumento,dataregistro) values(" & CLng(Left(Trim(.sNumeroDocumento), 8)) & ",'" & Format(Now, "mm/dd/yyyy hh:mm") & "')"
        cn.Execute Sql, rdExecDirect
        
    End With
    
    nContador = nContador + 1
    With aSegmentoQ(1)
        .nCodigoBanco = "001"
        .nLote = "0001"
        .nTipo = "3"
        .nSeqReg = FillLeft(CStr(nContador), 5)
        .sCodSegmento = "Q"
        .sUsoFebraban1 = " "
        .nCodMovimento = "01"
        .nTipoInscricao = aBoletos(nPosReg).nTipoInscricao
        .nNumeroInscricao = FillLeft(aBoletos(nPosReg).nNumeroInscricao, 15)
        .sNome = FillSpace(aBoletos(nPosReg).sNome, 40)
        .sEndereco = FillSpace(aBoletos(nPosReg).sEndereco, 40)
        .sBairro = FillSpace(aBoletos(nPosReg).sBairro, 15)
        If Len(aBoletos(nPosReg).sCep) < 5 Then
            .nCep = "00000"
            .nCepsufixo = "000"
        Else
            .nCep = aBoletos(nPosReg).sCep
            .nCepsufixo = aBoletos(nPosReg).sSufixoCep
        End If
        
        .sCidade = FillSpace(aBoletos(nPosReg).sCidade, 15)
        .sUF = aBoletos(nPosReg).sUF
        .nipoInscricaoSacado = "0"
        .nNumeroInscricaoSacado = FillLeft("0", 15)
        .sNomeSacado = FillSpace(" ", 40)
        .nBancoCorresponde = "000"
        .sNossoNumeroBancoCorr = FillSpace(" ", 20)
        .sUsoFebraban2 = FillSpace(" ", 8)
        
        sSegmentoQ = .nCodigoBanco & .nLote & .nTipo & .nSeqReg & .sCodSegmento & .sUsoFebraban1 & .nCodMovimento & .nTipoInscricao & .nNumeroInscricao & .sNome & .sEndereco & .sBairro
        sSegmentoQ = sSegmentoQ & .nCep & .nCepsufixo & .sCidade & .sUF & .nipoInscricaoSacado & .nNumeroInscricaoSacado & .sNomeSacado & .nBancoCorresponde & .sNossoNumeroBancoCorr & .sUsoFebraban2
    End With
    nContador = nContador + 1
    Print #FF1, sSegmentoP
    Print #FF1, sSegmentoQ
    DoEvents
    
Next

'****** Fim do Segmento P *******

'***************************************
'********** Trailer do Lote ************
'***************************************
nQtdeRegistroLote = (UBound(aBoletos) * 2) + 2
ReDim aTrailerLote(1)
With aTrailerLote(1)
    .nCodigoBanco = "001"
    .nLote = "0001"
    .nTipo = "5"
    .sUsoFebraban1 = FillSpace(" ", 9)
    .nQtdeRegistro = FillLeft(CStr(nQtdeRegistroLote), 6)
    .sUsoFebraban2 = FillSpace(" ", 217)
    
    sTrailerLote = .nCodigoBanco & .nLote & .nTipo & .sUsoFebraban1 & .nQtdeRegistro & .sUsoFebraban2
End With

Print #FF1, sTrailerLote
'****** Fim Trailer do Lote ************


'***************************************
'********** Trailer Arquivo ************
'***************************************
nQtdeRegistroArquivo = nQtdeRegistroLote + 2
ReDim aTrailerArquivo(1)
With aTrailerArquivo(1)
    .nCodigoBanco = "001"
    .nLote = "9999"
    .nTipo = "9"
    .sUsoFebraban1 = FillSpace(" ", 9)
    .nQtdeLote = FillLeft("1", 6)
    .nQtdeRegistro = FillLeft(CStr(nQtdeRegistroArquivo), 6)
    .nQtdeContas = FillLeft("0", 6)
    .sUsoFebraban2 = FillSpace(" ", 205)
    
    sTrailerArquivo = .nCodigoBanco & .nLote & .nTipo & .sUsoFebraban1 & .nQtdeLote & .nQtdeRegistro & .nQtdeContas & .sUsoFebraban2
End With

Print #FF1, sTrailerArquivo
'****** Fim Trailer Arquivo ************

Close #FF1
Fill

End Sub


Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Format(sTexto, String(nTamanho, "0"))

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

