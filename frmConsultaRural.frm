VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmConsultaRural 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Propriedade Rural"
   ClientHeight    =   5220
   ClientLeft      =   1830
   ClientTop       =   2625
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   6750
   Begin VB.ComboBox cmbCriterio 
      Height          =   315
      ItemData        =   "frmConsultaRural.frx":0000
      Left            =   1020
      List            =   "frmConsultaRural.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   4305
   End
   Begin VB.TextBox txtPesq 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   105
      Width           =   4305
   End
   Begin MSComctlLib.ListView lvRural 
      Height          =   3855
      Left            =   60
      TabIndex        =   3
      Top             =   885
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   6800
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Insc.Incra"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cód.Rec.Fed."
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CPF"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nome do Proprietário"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Nome da Propriedade"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CNPJ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Insc.Est."
         Object.Width           =   2540
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdPesq 
      Default         =   -1  'True
      Height          =   345
      Left            =   5460
      TabIndex        =   1
      ToolTipText     =   "Pesquisar"
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmConsultaRural.frx":00B0
      PICN            =   "frmConsultaRural.frx":00CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4170
      TabIndex        =   4
      ToolTipText     =   "Cancelar Seleção"
      Top             =   4830
      Width           =   1215
      _ExtentX        =   2143
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConsultaRural.frx":0226
      PICN            =   "frmConsultaRural.frx":0242
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
      Height          =   315
      Left            =   5460
      TabIndex        =   5
      ToolTipText     =   "Retorna Propriedade Selecionada"
      Top             =   4830
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Selecionar"
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
      MICON           =   "frmConsultaRural.frx":039C
      PICN            =   "frmConsultaRural.frx":03B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Critério.....:"
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   510
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa.:"
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   165
      Width           =   795
   End
End
Attribute VB_Name = "frmConsultaRural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String

Private Sub cmdCancel_Click()
CodIncra = 0
Unload Me
End Sub

Private Sub cmdConsultar_Click()
If lvRural.ListItems.Count = 0 Then
   MsgBox "Selecione uma Propriedade.", vbExclamation, "Atenção"
   Exit Sub
End If
CodRural = Right$(lvRural.SelectedItem.Key, 6)

If CodRural > 0 Then
    frmCadastroRural.lblCodReduzido.Caption = CodRural
    Unload frmConsultaRural
    frmCadastroRural.SetFocus
End If

Unload Me
End Sub

Private Sub cmdPesq_Click()
Ocupado
If txtPesq.Text = "" Then
   MsgBox "Digite o início da pesquisa.", vbExclamation, "Atenção"
   txtPesq.SetFocus
   Liberado
   Exit Sub
End If

If cmbCriterio.ListIndex = -1 Then
   MsgBox "Selecione um critério.", vbExclamation, "Atenção"
   cmbCriterio.SetFocus
   Liberado
   Exit Sub
End If

Screen.MousePointer = vbHourglass
CarregaLista txtPesq.Text
Screen.MousePointer = vbDefault

If lvRural.ListItems.Count = 0 Then
   MsgBox "Nenhum registro coincidente.", vbInformation, "Atenção"
End If
Liberado
End Sub

Private Sub Form_Load()
Ocupado
Add3DBorder lvRural
Centraliza Me
Liberado
cmbCriterio.ListIndex = 0
End Sub

Private Sub CarregaLista(Letra As String)

Dim itmX As ListItem
Dim z As Long
z = SendMessage(lvRural.hwnd, LVM_DELETEALLITEMS, 0, 0)
Ocupado
Sql = "SELECT codreduzido,incra,recfed,CADASTRORURAL.CPF,proprietario,propriedade,nomecidadao,CADASTRORURAL.cnpj,CADASTRORURAL.ie FROM  cadastrorural LEFT OUTER JOIN "
Sql = Sql & "cidadao ON cadastrorural.proprietario = cidadao.codcidadao where "
If cmbCriterio.ListIndex = -1 Then
    MsgBox "Selecione o critério.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbCriterio.ListIndex = 0 Then
   Sql = Sql & "CONVERT(VARCHAR(13),CODREDUZIDO) LIKE '" & Mask(Letra) & "%' ORDER BY CODREDUZIDO"
ElseIf cmbCriterio.ListIndex = 1 Then
   Sql = Sql & "CONVERT(VARCHAR(13),INCRA) LIKE '" & Mask(Letra) & "%' ORDER BY INCRA"
ElseIf cmbCriterio.ListIndex = 2 Then
   Sql = Sql & "CONVERT(VARCHAR(15),RECFED) LIKE '" & Letra & "%' ORDER BY RECFED"
ElseIf cmbCriterio.ListIndex = 3 Then
   Sql = Sql & "CADASTRORURAL.CPF LIKE '" & Letra & "%' ORDER BY CADASTRORURAL.CPF"
ElseIf cmbCriterio.ListIndex = 4 Then
   Sql = Sql & "NOMECIDADAO LIKE '" & Letra & "%' ORDER BY NOMECIDADAO"
ElseIf cmbCriterio.ListIndex = 5 Then
   Sql = Sql & "PROPRIEDADE LIKE '" & Letra & "%' ORDER BY PROPRIEDADE"
ElseIf cmbCriterio.ListIndex = 6 Then
   Sql = Sql & "CADASTRORURAL.CNPJ LIKE '" & Letra & "%' ORDER BY CADASTRORURAL.CNPJ"
ElseIf cmbCriterio.ListIndex = 7 Then
   Sql = Sql & "CADASTRORURAL.IE LIKE '" & Letra & "%' ORDER BY CADASTRORURAL.IE"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Set itmX = lvRural.ListItems.Add(, "R" & !CODREDUZIDO, !CODREDUZIDO)
       itmX.SubItems(1) = !INCRA
       itmX.SubItems(2) = SubNull(!RECFED)
       itmX.SubItems(3) = SubNull(!CPF)
       itmX.SubItems(4) = SubNull(!NOMECIDADAO)
       itmX.SubItems(5) = SubNull(!PROPRIEDADE)
       itmX.SubItems(6) = SubNull(!Cnpj)
       itmX.SubItems(7) = SubNull(!IE)

      .MoveNext
    Loop
   .Close
End With
Liberado
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub lvCid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvCid.SortKey = ColumnHeader.Position - 1
lvCid.Sorted = True
lvCid.SortOrder = lvwAscending
End Sub

Private Sub lvCid_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sNomeLogr As String, nCodCidadao As Long

nCodCidadao = Item.SubItems(5)

If nCodCidadao < 100000 Then
    Sql = "SELECT  dbo.vwCnsImovel.ABREVTIPOLOG, dbo.vwCnsImovel.ABREVTITLOG, dbo.vwCnsImovel.NOMELOGRADOURO, dbo.vwCnsImovel.LI_NUM "
    Sql = Sql & "FROM dbo.CIDADAO INNER JOIN dbo.PROPRIETARIO ON dbo.CIDADAO.CODCIDADAO = dbo.PROPRIETARIO.CODCIDADAO LEFT OUTER JOIN "
    Sql = Sql & "dbo.vwCnsImovel ON dbo.PROPRIETARIO.CODREDUZIDO = dbo.vwCnsImovel.CODREDUZIDO "
    Sql = Sql & "Where dbo.CIDADAO.CodCidadao =" & nCodCidadao
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
         If IsNull(RdoAux2!NomeLogradouro) Then
            sNomeLogradouro = ""
         Else
            sNomeLogr = Trim$(SubNull(RdoAux2!AbrevTipoLog)) & " " & Trim$(SubNull(RdoAux2!AbrevTitLog)) & " " & RdoAux2!NomeLogradouro & " Nº " & RdoAux2!Li_Num
         End If
    Else
         sNomeLogr = ""
    End If
    Item.SubItems(1) = sNomeLogr
End If

End Sub

Private Sub txtPesq_KeyPress(KeyAscii As Integer)
Ocupado
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If txtPesq.Text = "" Then
      MsgBox "Digite o início da pesquisa.", vbExclamation, "Atenção"
      txtPesq.SetFocus
      Exit Sub
   Else
      CarregaLista txtPesq.Text
   End If
End If
Liberado
End Sub


