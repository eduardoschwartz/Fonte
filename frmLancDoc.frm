VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLancDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descrição de Lançamentos por Documento"
   ClientHeight    =   4440
   ClientLeft      =   12285
   ClientTop       =   3870
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   45
      TabIndex        =   6
      Top             =   2610
      Width           =   7215
      Begin VB.TextBox txtDesc 
         Height          =   330
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   2
         Top             =   225
         Width           =   5865
      End
      Begin VB.ComboBox cmbAssunto 
         Height          =   315
         ItemData        =   "frmLancDoc.frx":0000
         Left            =   1260
         List            =   "frmLancDoc.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   945
         Width           =   5865
      End
      Begin VB.ComboBox cmbLanc 
         Height          =   315
         ItemData        =   "frmLancDoc.frx":0004
         Left            =   1260
         List            =   "frmLancDoc.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   585
         Width           =   5865
      End
      Begin VB.Label Label1 
         Caption         =   "Assunto........:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   9
         Top             =   990
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Lançamento.:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Descrição....:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   1050
      End
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmLancDoc.frx":0008
      Left            =   1890
      List            =   "frmLancDoc.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   2985
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   1965
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   3466
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descrição"
         Object.Width           =   4058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "CodLanc"
         Object.Width           =   1552
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lançamento"
         Object.Width           =   4058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "CodAssunto"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Assunto Protocolo"
         Object.Width           =   4057
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6225
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   4095
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmLancDoc.frx":000C
      PICN            =   "frmLancDoc.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   45
      TabIndex        =   11
      ToolTipText     =   "Novo Registro"
      Top             =   4095
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmLancDoc.frx":0182
      PICN            =   "frmLancDoc.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1140
      TabIndex        =   12
      ToolTipText     =   "Editar Registro"
      Top             =   4095
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmLancDoc.frx":02F8
      PICN            =   "frmLancDoc.frx":0314
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   2235
      TabIndex        =   13
      ToolTipText     =   "Excluir Registro"
      Top             =   4095
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmLancDoc.frx":046E
      PICN            =   "frmLancDoc.frx":048A
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
      Height          =   315
      Left            =   5130
      TabIndex        =   14
      ToolTipText     =   "Gravar os Dados"
      Top             =   4095
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmLancDoc.frx":052C
      PICN            =   "frmLancDoc.frx":0548
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
      Caption         =   "Tipo de Documento...:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   225
      Width           =   1635
   End
End
Attribute VB_Name = "frmLancDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sEvento As String

Private Sub cmbTipo_Click()
Dim z As Long, nCodTipo As Integer, Sql As String, rdoAux As rdoResultset, itmX As ListItem
Limpa
If cmbTipo.ListIndex = -1 Then Exit Sub

z = SendMessage(lvMain.HWND, LVM_DELETEALLITEMS, 0, 0)
nCodTipo = cmbTipo.ItemData(cmbTipo.ListIndex)

Sql = "SELECT tipodocumentodesc.seq, tipodocumentodesc.codlanc, tipodocumentodesc.codassunto, tipodocumentodesc.descricao, tributo.desctributo, assunto.NOME "
Sql = Sql & "FROM tipodocumentodesc INNER JOIN tributo ON tipodocumentodesc.codlanc = tributo.codtributo INNER JOIN "
Sql = Sql & "assunto ON tipodocumentodesc.codassunto = assunto.CODIGO WHERE CODTIPO=" & nCodTipo
Set rdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, , !descricao)
        itmX.SubItems(1) = !CODLANC
        itmX.SubItems(2) = !desctributo
        itmX.SubItems(3) = !CODASSUNTO
        itmX.SubItems(4) = !Nome
       .MoveNext
    Loop
   .Close
End With

lvMain_Click

End Sub

Private Sub cmdAlterar_Click()
sEvento = "Alterar"
Eventos "INCLUIR"
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
End Sub

Private Sub cmdGravar_Click()
Dim itmX As ListItem, x As Integer, nCodTipo As Integer, nCodLanc As Integer, nCodAss As Integer
nCodTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
nCodLanc = cmbLanc.ItemData(cmbLanc.ListIndex)
nCodAss = cmbAssunto.ItemData(cmbAssunto.ListIndex)

If Trim(txtDesc.Text) = "" Then
    MsgBox "Digite uma descrição", vbExclamation, "Atenção"
    Exit Sub
End If

If cmbLanc.ListIndex = -1 Then
    MsgBox "Selecione um lançamento", vbExclamation, "Atenção"
    Exit Sub
End If

If cmbAssunto.ListIndex = -1 Then
    MsgBox "Selecione um assunto", vbExclamation, "Atenção"
    Exit Sub
End If

If sEvento = "Novo" Then
    Set itmX = lvMain.ListItems.Add(, , txtDesc.Text)
    itmX.SubItems(1) = nCodLanc
    itmX.SubItems(2) = cmbLanc.Text
    itmX.SubItems(3) = nCodAss
    itmX.SubItems(4) = cmbAssunto.Text
Else
    lvMain.SelectedItem.Text = txtDesc.Text
    lvMain.SelectedItem.SubItems(1) = nCodLanc
    lvMain.SelectedItem.SubItems(2) = cmbLanc.Text
    lvMain.SelectedItem.SubItems(3) = nCodAss
    lvMain.SelectedItem.SubItems(4) = cmbAssunto.Text
End If

Sql = "DELETE FROM TIPODOCUMENTODESC WHERE CODTIPO=" & nCodTipo
cn.Execute Sql, rdExecDirect

For x = 1 To lvMain.ListItems.Count
    Sql = "INSERT TIPODOCUMENTODESC(CODTIPO,SEQ,DESCRICAO,CODLANC,CODASSUNTO) VALUES("
    Sql = Sql & nCodTipo & "," & x & ",'" & Trim(Mask(lvMain.ListItems(x).Text)) & "'," & lvMain.ListItems(x).SubItems(1) & "," & lvMain.ListItems(x).SubItems(3) & ")"
    cn.Execute Sql, rdExecDirect
Next
cmbTipo_Click
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
sEvento = "Novo"
Eventos "INCLUIR"
End Sub

Private Sub Form_Load()
Dim Sql As String, rdoAux As rdoResultset
Centraliza Me
Sql = "select codtributo,desctributo from tributo order by desctributo"
Set rdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        cmbLanc.AddItem !desctributo
        cmbLanc.ItemData(cmbLanc.NewIndex) = !CodTributo
       .MoveNext
    Loop
   .Close
End With

Sql = "select codigo,nome from assunto where ativo=1 order by nome"
Set rdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        cmbAssunto.AddItem !Nome
        cmbAssunto.ItemData(cmbAssunto.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

Sql = "select codigo,nome from tipolancdoc order by codigo"
Set rdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With rdoAux
    Do Until .EOF
        cmbTipo.AddItem !Nome
        cmbTipo.ItemData(cmbTipo.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0
lvMain.ColumnHeaders(2).Width = 0
lvMain.ColumnHeaders(4).Width = 0


Eventos "INICIAR"
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmLancDoc
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   cmbTipo.Enabled = True
   cmbTipo.BackColor = vbWhite
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmLancDoc
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   cmbTipo.Enabled = False
   cmbTipo.BackColor = Kde

End If

End Sub

Private Sub Limpa()
    txtDesc.Text = ""
    cmbAssunto.ListIndex = -1
    cmbLanc.ListIndex = -1
End Sub

Private Sub lvMain_Click()

Limpa

With lvMain
    If .ListItems.Count = 0 Then Exit Sub
    txtDesc.Text = .SelectedItem.Text
    cmbLanc.Text = .SelectedItem.SubItems(2)
    cmbAssunto.Text = .SelectedItem.SubItems(4)
End With

End Sub

Private Sub lvMain_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then lvMain_Click
End Sub
