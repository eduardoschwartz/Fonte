VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSimplesCnpj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simples Nacional - CNPJ não cadastrado"
   ClientHeight    =   3375
   ClientLeft      =   4755
   ClientTop       =   3300
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   7590
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1035
      MaxLength       =   50
      TabIndex        =   1
      Top             =   585
      Width           =   6405
   End
   Begin VB.ComboBox cmbCNPJ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   3525
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   5850
      TabIndex        =   3
      ToolTipText     =   "Gravar os Dados e dar baixa nos débitos"
      Top             =   2925
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar CNPJ"
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
      MICON           =   "frmSimplesCnpj.frx":0000
      PICN            =   "frmSimplesCnpj.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvTmp 
      Height          =   1785
      Left            =   90
      TabIndex        =   2
      Top             =   1035
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   3149
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dt.Crédito"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Dt.Vencto"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Banco"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Ag."
         Object.Width           =   1146
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Ano"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Mes"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Vl.Pago"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Arquivo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblQtde 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   6660
      TabIndex        =   7
      Top             =   225
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "QTDE....:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   5850
      TabIndex        =   6
      Top             =   225
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "NOME..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   630
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "CNPJ....:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   225
      Width           =   780
   End
End
Attribute VB_Name = "frmSimplesCnpj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCNPJ_Click()
Dim Sql As String, RdoAux As rdoResultset, z As Long

If cmbCNPJ.ListIndex = -1 Then Exit Sub
z = SendMessage(lvTmp.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT * FROM SIMPLESCNPJ WHERE CNPJ='" & RetornaNumero(cmbCNPJ.Text) & "' ORDER BY DATAARRECADA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
        Set itmX = lvTmp.ListItems.Add(, "C" & Format(.AbsolutePosition, "0000"), Format(!DataArrecada, "dd/mm/yyyy"))
        itmX.SubItems(1) = Format(!DataVencto, "dd/mm/yyyy")
        itmX.SubItems(2) = !Banco
        itmX.SubItems(3) = !Agencia
        itmX.SubItems(4) = !AnoComp
        itmX.SubItems(5) = !MesComp
        itmX.SubItems(6) = FormatNumber(!principal + !Juros + !Multa, 2)
        itmX.SubItems(7) = !ArquivoShort
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdGravar_Click()
Dim nCodReduz As Long, x As Integer, sDataCredito As String, sDataVencimento As String, nValorPago As Double, nBanco As Integer, sAgencia As String, sArquivo As String
Dim nParc As Integer, nCompl As Integer, nSeq As Integer, nAno As Integer, nLanc As Integer, RdoAux2 As rdoResultset, nNumDoc As Long, nSeqAdd As Integer

If Trim(txtNome.Text) = "" Then
    MsgBox "Digite um nome.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Nome: " & txtNome.Text & vbCrLf & "CNPJ: " & cmbCNPJ.Text & vbCrLf & vbCrLf & "Cadastrar e efetuar a baixa?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Ocupado
Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO WHERE CODCIDADAO<700000"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    nCodReduz = !maximo + 1
   .Close
End With

Sql = "INSERT CIDADAO (CODCIDADAO,NOMECIDADAO,CNPJ) VALUES(" & nCodReduz & ",'"
Sql = Sql & txtNome.Text & "','" & cmbCNPJ.Text & "')"
cn.Execute Sql, rdExecDirect
 
'Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCodReduz & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'Sql = Sql & "Cidadão criado através da tela SimplesCNPJ')"
Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCodReduz & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
Sql = Sql & "Cidadão criado através da tela SimplesCNPJ')"
cn.Execute Sql, rdExecDirect
 
 
For x = 1 To lvTmp.ListItems.Count
    sDataCredito = lvTmp.ListItems(x).Text
    sDataVencimento = lvTmp.ListItems(x).SubItems(1)
    nBanco = lvTmp.ListItems(x).SubItems(2)
    sAgencia = lvTmp.ListItems(x).SubItems(3)
    nAno = lvTmp.ListItems(x).SubItems(4)
    nValorPago = lvTmp.ListItems(x).SubItems(6)
    sArquivo = lvTmp.ListItems(x).SubItems(7)
    
    '** TROCA OS BANCOS PELOS BANCOS VIRTUAIS **
    If nBanco = 1 Then
        nBanco = 91
    ElseIf nBanco = 33 Then
        nBanco = 92
    ElseIf nBanco = 237 Then
        nBanco = 93
    ElseIf nBanco = 341 Then
        nBanco = 94
    ElseIf nBanco = 409 Then
        nBanco = 95
    ElseIf nBanco = 151 Then
        nBanco = 96
    ElseIf nBanco = 104 Then
        nBanco = 97
    ElseIf nBanco = 399 Then
        nBanco = 98
    Else
        nBanco = 91
    End If
    
    'SE JA HOUVER UM LANCAMENTO COM ESTE VENCIMENTO PULA
    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=5 AND DATAVENCIMENTO='" & Format(sDataVencimento, "mm/dd/yyyy") & "'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
        nParc = RdoAux2!NumParcela
        nCompl = RdoAux2!CODCOMPLEMENTO
        nSeq = RdoAux2!SeqLancamento
        GoTo proximo
    End If
        
    'O NÚMERO DA PARCELA A SER CRIADA SERÁ O ÚLTIMO NÚMERO DE PARCELA DO ANO
    Sql = "SELECT MAX(NUMPARCELA) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & nAno & ")"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If IsNull(!maximo) Then
            nParc = 1
        Else
            nParc = !maximo + 1
        End If
       .Close
    End With
    nCompl = 0
    nSeq = 0
    'CRIAR PARCELA DE ISS VARIAVEL NESTE MES E ANO COM O VENCIMENTO QUE VEIO DO BANCO
'    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'    Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
'    Sql = Sql & nParc & "," & nCompl & ",2,'" & Format(sDataVencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0,'" & Left$(NomeDeLogin, 25) & "')"
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
    Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
    Sql = Sql & nParc & "," & nCompl & ",2,'" & Format(sDataVencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0," & RetornaUsuarioID(NomeDeLogin) & ")"
    cn.Execute Sql, rdExecDirect
    'CRIAR O TRIBUTO PARA ELA (13 - iss variavel)
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
    Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
    Sql = Sql & nParc & "," & nCompl & "," & 13 & "," & Virg2Ponto(CStr(nValorPago)) & ")"
    cn.Execute Sql, rdExecDirect
    'CRIAR O DOCUMENTO PARA ELA
    Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
         nNumDoc = !maximo + 1
        .Close
    End With
    
    
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,emissor) VALUES(" & nNumDoc & ",'"
    Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & nBanco & "," & Val(RetornaNumero(sAgencia)) & "," & Virg2Ponto(CStr(nValorPago)) & ",'" & NomeDeLogin & " (SIMPLES CNPJ)" & "')"
    cn.Execute Sql, rdExecDirect
    'CRIAR A PARCELADOCUMENTO
    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
    Sql = Sql & nParc & "," & nCompl & "," & nNumDoc & ")"
    cn.Execute Sql, rdExecDirect
    'ULTIMA SEQ DE PAGTO
    Sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND "
    Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND CODCOMPLEMENTO=" & nCompl
    Sql = Sql & " AND NUMPARCELA=" & nParc
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
         If IsNull(!maximo) Then
            nSeqAdd = 0
         Else
            If .RowCount = 0 Then
               nSeqAdd = 0
           Else
              nSeqAdd = !maximo + 1
           End If
        End If
       .Close
    End With
    On Error Resume Next
    'CRIAR DEBITOPAGO
    Sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQPAG,"
    Sql = Sql & "DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL,INTACTO,VALORTARIFA,ARQUIVOBANCO,VALORDIF) VALUES("
    Sql = Sql & nCodReduz & "," & nAno & ",5," & nSeq & "," & nParc & "," & nCompl & "," & nSeqAdd & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','"
    Sql = Sql & Format(sDataCredito, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & "," & nBanco & "," & Val(sAgencia) & "," & nNumDoc & ","
    Sql = Sql & Virg2Ponto(CStr(nValorPago)) & ",0,0" & ",'" & sArquivo & "'," & 0 & ")"
    cn.Execute Sql, rdExecDirect
    On Error GoTo 0
proximo:
    'CRIAR COMPLEMENTOSIMPLES
    On Error Resume Next
    Sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR) VALUES(" & nCodReduz & ","
    Sql = Sql & nAno & ",5," & nSeq & "," & nParc & "," & nCompl & ",'" & sArquivo & "','" & Format(sDataCredito, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & ")"
    cn.Execute Sql, rdExecDirect
    On Error GoTo 0
    
    'REMOVE DO SIMPLES CNPJ
    Sql = "DELETE FROM SIMPLESCNPJ WHERE CNPJ='" & RetornaNumero(CStr(cmbCNPJ.Text)) & "'"
    cn.Execute Sql, rdExecDirect
    
Next
Liberado

MsgBox "Criado código: " & nCodReduz & " e efetuada a baixa.", vbInformation, "Informação"
txtNome.Text = ""
CarregaCNPJ
End Sub

Private Sub Form_Load()
Centraliza Me
CarregaCNPJ
End Sub

Private Sub CarregaCNPJ()
Dim Sql As String, RdoAux As rdoResultset

cmbCNPJ.Clear
Sql = "SELECT DISTINCT CNPJ FROM SIMPLESCNPJ ORDER BY CNPJ"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbCNPJ.AddItem Format(Trim(!Cnpj), "00\.000\.000/0000-00")
       .MoveNext
    Loop
   .Close
End With
lblQtde.Caption = cmbCNPJ.ListCount
If cmbCNPJ.ListCount > 0 Then cmbCNPJ.ListIndex = 0

End Sub
