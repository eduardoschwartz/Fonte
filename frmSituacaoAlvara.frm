VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSituacaoAlvara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situação dos Alvarás Emitidos"
   ClientHeight    =   5760
   ClientLeft      =   7380
   ClientTop       =   7770
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   11415
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmSituacaoAlvara.frx":0000
      Left            =   3570
      List            =   "frmSituacaoAlvara.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   5055
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razão Social"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CPF/CNPJ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nome do Logradouro"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nº Log."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Atividade por Extenso"
         Object.Width           =   4304
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdExcel 
      Height          =   345
      Left            =   8160
      TabIndex        =   5
      ToolTipText     =   "Enviar dados para o Excel"
      Top             =   90
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Gerar em Excel"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmSituacaoAlvara.frx":006C
      PICN            =   "frmSituacaoAlvara.frx":0088
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   345
      Left            =   6810
      TabIndex        =   6
      ToolTipText     =   "Consultar relatório selecionado"
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmSituacaoAlvara.frx":0115
      PICN            =   "frmSituacaoAlvara.frx":0131
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblQtde 
      Caption         =   "0000"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   10620
      TabIndex        =   8
      Top             =   180
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Qtde..:"
      Height          =   195
      Index           =   1
      Left            =   10050
      TabIndex        =   7
      Top             =   180
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Relatório..:"
      Height          =   195
      Left            =   2070
      TabIndex        =   3
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label lblAno 
      Caption         =   "0000"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1230
      TabIndex        =   2
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Ano Alvará..:"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmSituacaoAlvara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultar_Click()
Dim nTipo As Integer, Sql As String, RdoAux As rdoResultset, itmX As ListItem, sDoc As String


nTipo = cmbTipo.ListIndex
If nTipo = -1 Then
    MsgBox "Selecione um tipo de relatório.", vbCritical, "Erro"
    Exit Sub
End If
Ocupado
DoEvents
lvMain.ListItems.Clear
lblQtde.Caption = "0000"

If nTipo = 0 Then 'Alvara internet
    Sql = "SELECT DISTINCT codigomob, razaosocial, cnpj, cpf, LOGRADOURO, numero, ativextenso From vwFULLEMPRESA WHERE dataencerramento IS NULL AND codigomob IN "
    Sql = Sql & "(SELECT DISTINCT CODREDUZIDO From alvara WHERE (YEAR(DATA) = " & Val(lblAno.Caption) & ")) ORDER BY codigomob"
ElseIf nTipo = 1 Then 'empresa com alvará em atraso
    Sql = "SELECT DISTINCT codigomob, razaosocial, cnpj, cpf, LOGRADOURO, numero, ativextenso From vwFULLEMPRESA WHERE (dataencerramento IS NULL) AND (codigomob IN "
    Sql = Sql & "(SELECT debitoparcela.codreduzido FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
    Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And "
    Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitoparcela.anoexercicio = " & Val(lblAno.Caption) & ") and debitoparcela.codlancamento<>20 AND (debitoparcela.statuslanc = 3) AND (debitoparcela.datavencimento < GETDATE()) AND "
    Sql = Sql & "(debitotributo.codtributo = 23) AND (debitoparcela.codreduzido BETWEEN 100000 AND 300000)))  ORDER BY codigomob"
ElseIf nTipo = 2 Then 'empresa com alvará regular
    Sql = "SELECT DISTINCT codigomob, razaosocial, cnpj, cpf, LOGRADOURO, numero, ativextenso From vwFULLEMPRESA WHERE (dataencerramento IS NULL) AND (codigomob NOT IN "
    Sql = Sql & "(SELECT debitoparcela.codreduzido FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
    Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And "
    Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitoparcela.anoexercicio = " & Val(lblAno.Caption) & ") and debitoparcela.codlancamento<>20 AND (debitoparcela.statuslanc = 3) AND (debitoparcela.datavencimento < GETDATE()) AND "
    Sql = Sql & "(debitotributo.codtributo = 23) AND (debitoparcela.codreduzido BETWEEN 100000 AND 300000))) AND (codigomob IN (SELECT DISTINCT codreduzido FROM  debitotributo AS debitotributo_1 "
    Sql = Sql & "WHERE (codreduzido BETWEEN 100000 AND 300000) AND (anoexercicio = " & Val(lblAno.Caption) & ") AND (codtributo = 23) AND (numparcela = 1) AND (codlancamento <> 20))) ORDER BY codigomob"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblQtde.Caption = Format(.RowCount, "0000")
    Do Until .EOF
        sDoc = Trim(SubNull(!Cnpj))
        If sDoc = "" Then
            sDoc = Trim(SubNull(!CPF))
        End If
    
        Set itmX = lvMain.ListItems.Add(, "C" & Format(!codigomob, "0000000"), Format(!codigomob, "0000000"))
        itmX.SubItems(1) = SubNull(!RazaoSocial)
        itmX.SubItems(2) = sDoc
        itmX.SubItems(3) = !Logradouro
        itmX.SubItems(4) = SubNull(!Numero)
        itmX.SubItems(5) = !ativextenso
       .MoveNext
    Loop
   .Close
End With
Liberado

End Sub

Private Sub cmdExcel_Click()
If lvMain.ListItems.Count = 0 Then
    MsgBox "Nada a imprimir.", vbCritical, "Erro"
Else
    PrintExcel
End If
End Sub

Private Sub Form_Load()

Centraliza Me
lblAno.Caption = Year(Now)

End Sub

Private Sub PrintExcel()

If lvMain.ListItems.Count = 0 Then Exit Sub

Dim x As Long, Y As Long, ax As String, Scr_hdc As Long, Z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For Y = 1 To lvMain.ColumnHeaders.Count
    ax = ax & RemoveSpace(lvMain.ColumnHeaders(Y).Text) & " char(255), "
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "[Table1$]", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For x = 1 To lvMain.ListItems.Count
    Rs.AddNew
    nCont = 0
    Rs.Fields(nCont).value = lvMain.ListItems(x).Text
    nCont = nCont + 1
    For Y = 2 To lvMain.ColumnHeaders.Count
         
         Rs.Fields(nCont).value = Left(lvMain.ListItems(x).SubItems(Y - 1), 100)
         nCont = nCont + 1
    
        
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

Z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)


End Sub

