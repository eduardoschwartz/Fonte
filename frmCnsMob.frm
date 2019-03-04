VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCnsMob 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Empresas"
   ClientHeight    =   5925
   ClientLeft      =   4830
   ClientTop       =   2970
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   8760
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Selecione um ou mais Parâmetros"
      ForeColor       =   &H00000080&
      Height          =   2745
      Left            =   0
      TabIndex        =   9
      Top             =   30
      Width           =   8745
      Begin VB.TextBox txtIE 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5220
         MaxLength       =   15
         TabIndex        =   3
         Top             =   990
         Width           =   1905
      End
      Begin VB.TextBox txtCodAtiv 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1680
         Width           =   765
      End
      Begin VB.TextBox txtAtiv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1680
         Width           =   4170
      End
      Begin VB.TextBox txtRazao 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2130
         MaxLength       =   40
         TabIndex        =   1
         Top             =   630
         Width           =   4995
      End
      Begin VB.TextBox txtCNPJ 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2130
         MaxLength       =   20
         TabIndex        =   2
         Top             =   990
         Width           =   2145
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1350
         Width           =   4170
      End
      Begin VB.TextBox txtNumImovel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         TabIndex        =   7
         Top             =   2340
         Width           =   705
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   0
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtNomeLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2025
         Width           =   4170
      End
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2010
         Width           =   765
      End
      Begin prjChameleon.chameleonButton cmdBuscaProp 
         Height          =   285
         Left            =   6360
         TabIndex        =   4
         ToolTipText     =   "Busca Proprietário"
         Top             =   1350
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnsMob.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelProp 
         Height          =   285
         Left            =   6720
         TabIndex        =   13
         ToolTipText     =   "Limpa Campo Proprietário"
         Top             =   1350
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnsMob.frx":001C
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
         Left            =   7380
         TabIndex        =   22
         ToolTipText     =   "Seleciona o Imóvel"
         Top             =   930
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCnsMob.frx":0038
         PICN            =   "frmCnsMob.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdPesq 
         Height          =   345
         Left            =   7380
         TabIndex        =   23
         ToolTipText     =   "Pesquisar"
         Top             =   540
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
         MICON           =   "frmCnsMob.frx":00C2
         PICN            =   "frmCnsMob.frx":00DE
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
         Left            =   7380
         TabIndex        =   24
         ToolTipText     =   "Cancelar Edição"
         Top             =   1290
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCnsMob.frx":0238
         PICN            =   "frmCnsMob.frx":0254
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdLimpar 
         Height          =   315
         Left            =   7380
         TabIndex        =   25
         ToolTipText     =   "Limpar campos de pesquisa"
         Top             =   1665
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Limpar"
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
         MICON           =   "frmCnsMob.frx":03AE
         PICN            =   "frmCnsMob.frx":03CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox lstNomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   1785
         ItemData        =   "frmCnsMob.frx":0524
         Left            =   2130
         List            =   "frmCnsMob.frx":0526
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de I.E..:"
         Height          =   225
         Index           =   7
         Left            =   4380
         TabIndex        =   26
         Top             =   1050
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade Principal...........:"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   21
         Top             =   1710
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social...................:"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   19
         Top             =   675
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de CNPJ.....................:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   18
         Top             =   1035
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário/Sócio............:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   1380
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Logradouro.......:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   2040
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código da Empresa.........:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   330
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Imóvel....................:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   2370
         Width           =   1905
      End
   End
   Begin MSComctlLib.ListView lvEmpresa 
      Height          =   3075
      Left            =   0
      TabIndex        =   8
      Top             =   2790
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
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
         Text            =   "Nº de CNPJ"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Insc.Est."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Proprietário/Sócios"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Nome do Logradouro"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nº Log."
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmCnsMob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, sSeiLa As String
Dim Sql As String, bCod As Boolean, bInsc As Boolean
Dim bRazao As Boolean, bAtiv As Boolean, bSeq As Boolean, bUnidade As Boolean, bSubUnidade As Boolean
Dim bNome As Boolean, bLog As Boolean, bNum As Boolean, bIE As Boolean

Private Sub cmdBuscaProp_Click()

Dim frm As Object
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frmCnsCidadao.show

End Sub

Private Sub cmdCancel_Click()
   'Unload frmCnsMob
   frmCnsMob.Hide
End Sub

Private Sub cmdConsultar_Click()
Dim Achou As Boolean

If lvEmpresa.ListItems.Count > 0 Then
   CodEmpresa = Val(Left(lvEmpresa.SelectedItem.Text, 7))
'   modLg "Consulta de empresa: " & CodEmpresa & " - " & lvEmpresa.SelectedItem.SubItems(1)
   If sFormMob = "CM" Or sFormMob = "" Then
        Achou = False
        For x = 0 To Forms.Count - 1
            If Forms(x).Name = "frmCadMob" Then
                 Achou = True
                 Exit For
            End If
        Next
        If Not Achou Then
             frmCadMob.show
             frmCadMob.ZOrder 0
             frmCadMob.SetFocus
        End If
    ElseIf sFormMob = "DI2" Then
        frmDebitoImob.SetFocus
    ElseIf sFormMob = "EI" Then
        frm2ViaLaser.SetFocus
    ElseIf sFormMob = "ALUGUEL" Then
        frmManAluguel.SetFocus
    ElseIf sFormMob = "2VIA" Then
        frmEmissao2Via.SetFocus
    ElseIf sFormMob = "2VIAE" Then
        frmEmissao2ViaEspecial.SetFocus
    End If
    'Unload frmCnsMob
    frmCnsMob.Hide
Else
   MsgBox "Selecione a Empresa que deseja consultar.", vbExclamation, "Atenção"
End If

End Sub

Private Sub cmdDelProp_Click()
txtProp.Text = ""
End Sub

Private Sub cmdLimpar_Click()
txtCod.Text = ""
txtRazao.Text = ""
txtCNPJ.Text = ""
txtProp.Text = ""
txtAtiv.Text = ""
txtNomeLogr.Text = ""
txtNumImovel.Text = ""
txtCodAtiv.Text = ""
txtCodLogr.Text = ""
End Sub

Private Sub cmdPesq_Click()
'On Error Resume Next
Dim itmX As ListItem
Dim z As Long, x As Integer

If Val(txtCod.Text) = 0 And Trim$(txtProp.Text) = "" And Val(txtCodLogr.Text) = 0 And Val(txtNumImovel.Text) = 0 And Trim$(txtCNPJ.Text) = "" And Trim$(txtRazao.Text) = "" And Val(txtCodAtiv.Text) = 0 And Val(txtIE.Text) = 0 Then
    MsgBox "Favor selecionar ao menos um critério para busca.", vbExclamation, "Atenção"
    Exit Sub
End If

Screen.MousePointer = vbHourglass
Ocupado

z = SendMessage(lvEmpresa.HWND, LVM_DELETEALLITEMS, 0, 0)

bCod = False
bInsc = False
bRazao = False
bQuadra = False
bAtiv = False
bSeq = False
bUnidade = False
bSubUnidade = False
bNome = False
bLog = False
bNum = False
bIE = False

If Val(txtCod.Text) > 0 Then bCod = True
If Trim$(txtCNPJ.Text) <> "" Then bInsc = True
If Trim$(txtRazao.Text) <> "" Then bRazao = True
If Trim$(txtProp.Text) <> "" Then bNome = True
If Val(txtCodLogr.Text) > 0 Then bLog = True
If Val(txtCodAtiv.Text) > 0 Then bAtiv = True
If txtNumImovel.Text <> "" Then bNum = True
If Val(txtIE.Text) > 0 Then bIE = True

If bNome Then
    Sql = "SELECT CODIGOMOB,DVMOB,RAZAOSOCIAL,NOMEFANTASIA,CODLOGRADOURO,ABREVTIPOLOG,ABREVTITLOG,"
    Sql = Sql & "NOMELOGRADOURO,NUMERO,CNPJ,INSCESTADUAL,CODATIVIDADE,DESCATIVIDADE,CODCIDADAO,NOMECIDADAO FROM VWCNSMOBILIARIOPROP WHERE "
Else
    Sql = "SELECT CODIGOMOB,DVMOB,RAZAOSOCIAL,NOMEFANTASIA,CODLOGRADOURO,ABREVTIPOLOG,ABREVTITLOG,"
    Sql = Sql & "NOMELOGRADOURO,NUMERO,CNPJ,INSCESTADUAL,CODATIVIDADE,DESCATIVIDADE,CODCIDADAO,NOMECIDADAO FROM VWCNSMOBILIARIO WHERE "
End If
If bCod Then
   Sql = Sql & "CODIGOMOB=" & Val(txtCod.Text) & " AND "
End If
If bInsc Then
   Sql = Sql & "CNPJ LIKE '%" & Trim$(txtCNPJ.Text) & "%' AND "
End If
If bIE Then
   Sql = Sql & "INSCESTADUAL LIKE '%" & Trim$(txtIE.Text) & "%' AND "
End If
If bNum Then
   Sql = Sql & "NUMERO=" & Val(txtNumImovel.Text) & " AND "
End If
If bRazao Then
   Sql = Sql & "RAZAOSOCIAL LIKE '%" & Trim$(Mask(txtRazao.Text)) & "%' AND "
End If
If bLog Then
   Sql = Sql & "CODLOGRADOURO=" & Val(txtCodLogr.Text) & " AND "
End If
If bAtiv Then
   Sql = Sql & "CODATIVIDADE=" & Val(txtCodAtiv.Text) & " AND "
End If
If bNome Then
   Sql = Sql & "CODCIDADAO=" & Val(Left(txtProp.Text, 6)) & " AND "
End If
Sql = Left$(Sql, Len(Sql) - 5)

x = 0
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
   With RdoAux
       Do Until .EOF
          x = x + 1
          Set itmX = lvEmpresa.ListItems.Add(, "C" & sTr(x) & Format(!codigomob, "0000000"), Format(!codigomob, "0000000"))
          itmX.SubItems(1) = SubNull(!razaosocial)
          itmX.SubItems(2) = SubNull(!Cnpj)
          itmX.SubItems(3) = SubNull(!INSCESTADUAL)
          itmX.SubItems(4) = SubNull(!nomecidadao)
          itmX.SubItems(5) = Trim$(SubNull(!AbrevTipoLog)) & " " & IIf(IsNull(!AbrevTitLog), "", Trim$(SubNull(!AbrevTitLog)) & " ") & Trim$(SubNull(!NomeLogradouro))
          itmX.SubItems(6) = SubNull(!Numero)
         .MoveNext
       Loop
      .Close
   End With
Else
   Liberado
   MsgBox "Não existem Empresas com estes parâmetros.", vbExclamation, "Atenção"
End If
Liberado
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

Ocupado
CodEmpresa = ""

Centraliza Me
Liberado
End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub

Private Sub lvEmpresa_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvEmpresa.SortKey = ColumnHeader.Position - 1
lvEmpresa.Sorted = True
lvEmpresa.SortOrder = lvwAscending
End Sub

Private Sub txtAtiv_Change()
If Trim$(txtAtiv) = "" Then
   txtCodAtiv.Text = 0
End If
End Sub

Private Sub txtAtiv_GotFocus()
txtAtiv.SelStart = 0
txtAtiv.SelLength = Len(txtAtiv)
End Sub

Private Sub txtAtiv_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   sSeiLa = "A"
   lstNomeLog.Clear
   If txtAtiv.Text <> "" Then
      Sql = "SELECT CODATIVIDADE,DESCATIVIDADE FROM ATIVIDADE "
      Sql = Sql & "WHERE DESCATIVIDADE LIKE '%" & Trim$(txtAtiv) & "%' "
      Sql = Sql & "ORDER BY DESCATIVIDADE"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem !descatividade
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !codatividade
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ListIndex = 0
             lstNomeLog.ZOrder 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Atividade não encontrada.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtAtiv.SetFocus
          End If
      End With
   End If
Else
   txtCodAtiv.Text = 0
End If
End Sub

Private Sub txtCod_GotFocus()
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)
End Sub


Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdPesq_Click
    cmdConsultar_Click
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If


End Sub

Private Sub txtNomeLogr_Change()
If Trim$(txtNomeLogr) = "" Then
   txtCodLogr.Text = 0
End If
End Sub

Private Sub txtNomeLogr_GotFocus()
txtNomeLogr.SelStart = 0
txtNomeLogr.SelLength = Len(txtNomeLogr)
End Sub

Private Sub txtNomeLogr_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   sSeiLa = "L"
   lstNomeLog.Clear
   If txtNomeLogr.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLogr) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder 0
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro não encontrado.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtNomeLogr.SetFocus
          End If
      End With
   End If
Else
   txtCodLogr.Text = 0
End If

End Sub

Private Sub lstNomeLog_DblClick()

If sSeiLa = "L" Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
       txtNumImovel.SetFocus
    End If
Else
    If lstNomeLog.ListIndex > -1 Then
       txtCodAtiv.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtAtiv.Text = lstNomeLog.Text
       lstNomeLog.Visible = False
       txtNomeLogr.SetFocus
    End If
End If

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)

If sSeiLa = "L" Then
    If KeyAscii = vbKeyReturn Then
        If lstNomeLog.ListIndex > -1 Then
           txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
           txtCodLogr_LostFocus
           lstNomeLog.Visible = False
           txtNumImovel.SetFocus
        End If
    ElseIf KeyAscii = vbKeyEscape Then
       lstNomeLog.Visible = False
       txtNomeLogr.SetFocus
    End If
Else
    If KeyAscii = vbKeyReturn Then
        If lstNomeLog.ListIndex > -1 Then
           txtCodAtiv.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
           txtAtiv.Text = lstNomeLog.Text
           lstNomeLog.Visible = False
           txtNomeLogr.SetFocus
        End If
    ElseIf KeyAscii = vbKeyEscape Then
       lstNomeLog.Visible = False
       txtAtiv.SetFocus
    End If
End If

End Sub

Private Sub txtCodLogr_LostFocus()
If Val(txtCodLogr.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLogr.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          txtNomeLogr.Text = ""
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtCodLogr.SetFocus
       End If
      .Close
   End With
End If

End Sub

Private Sub txtRazao_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdPesq_Click
End If
End Sub
