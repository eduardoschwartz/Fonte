VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCnsRegAtend 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Registros de Atendimento"
   ClientHeight    =   5835
   ClientLeft      =   6570
   ClientTop       =   3210
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   8745
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Selecione um ou mais Parâmetros"
      ForeColor       =   &H00000080&
      Height          =   2745
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8745
      Begin VB.ComboBox cmbBairro 
         Height          =   315
         Left            =   2115
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2385
         Width           =   5010
      End
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2115
         TabIndex        =   2
         Top             =   630
         Width           =   1020
      End
      Begin VB.TextBox txtAnoProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4965
         TabIndex        =   3
         Top             =   630
         Width           =   1020
      End
      Begin VB.TextBox txtAnoReg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4965
         TabIndex        =   1
         Top             =   295
         Width           =   1020
      End
      Begin VB.TextBox txtNumReg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2115
         TabIndex        =   0
         Top             =   295
         Width           =   1020
      End
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2010
         Width           =   765
      End
      Begin VB.TextBox txtNomeLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2025
         Width           =   4170
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1305
         Width           =   4170
      End
      Begin VB.TextBox txtAssunto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   100
         TabIndex        =   4
         Top             =   955
         Width           =   5070
      End
      Begin VB.TextBox txtCCusto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1680
         Width           =   4980
      End
      Begin prjChameleon.chameleonButton cmdBuscaProp 
         Height          =   285
         Left            =   6360
         TabIndex        =   6
         ToolTipText     =   "Busca Proprietário"
         Top             =   1305
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
         MICON           =   "frmCnsRegAtend.frx":0000
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
         TabIndex        =   7
         ToolTipText     =   "Limpa Campo Proprietário"
         Top             =   1305
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
         MICON           =   "frmCnsRegAtend.frx":001C
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
         TabIndex        =   13
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
         MICON           =   "frmCnsRegAtend.frx":0038
         PICN            =   "frmCnsRegAtend.frx":0054
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
         Default         =   -1  'True
         Height          =   345
         Left            =   7380
         TabIndex        =   14
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
         MICON           =   "frmCnsRegAtend.frx":00C2
         PICN            =   "frmCnsRegAtend.frx":00DE
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
         TabIndex        =   15
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
         MICON           =   "frmCnsRegAtend.frx":0238
         PICN            =   "frmCnsRegAtend.frx":0254
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
         TabIndex        =   16
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
         MICON           =   "frmCnsRegAtend.frx":03AE
         PICN            =   "frmCnsRegAtend.frx":03CA
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
         ItemData        =   "frmCnsRegAtend.frx":0524
         Left            =   2115
         List            =   "frmCnsRegAtend.frx":0526
         TabIndex        =   17
         Top             =   945
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número do Processo.......:"
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   27
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano do Processo.....:"
         Height          =   225
         Index           =   3
         Left            =   3405
         TabIndex        =   26
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano do Registro......:"
         Height          =   225
         Index           =   7
         Left            =   3420
         TabIndex        =   25
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Bairro................:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   23
         Top             =   2415
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número do Registro........:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   22
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Logradouro.......:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   21
         Top             =   2040
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Cidadão............:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   20
         Top             =   1335
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto...........................:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   19
         Top             =   990
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Custos.............:"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   18
         Top             =   1710
         Width           =   1905
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   3075
      Left            =   45
      TabIndex        =   24
      Top             =   2745
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
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° Reg."
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Ano Reg."
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Nº Proc."
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Ano Proc."
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Assunto"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cidadão"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Centro Custo"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Endereço"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Num"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Bairro"
         Object.Width           =   4304
      EndProperty
   End
End
Attribute VB_Name = "frmCnsRegAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscaProp_Click()
Dim frm As Object
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frmCnsCidadao.show

End Sub

Private Sub cmdConsultar_Click()
Dim Achou As Boolean

If lvMain.ListItems.Count > 0 Then
   NumRegAtend = lvMain.SelectedItem.Text
   AnoRegAtend = lvMain.SelectedItem.SubItems(1)
   frmRegistroAtendimento.SetFocus
   frmCnsRegAtend.Hide
Else
   MsgBox "Selecione o registro que deseja consultar.", vbExclamation, "Atenção"
End If

End Sub

Private Sub cmdDelProp_Click()
txtProp.Text = ""
End Sub

Private Sub cmdLimpar_Click()
txtNumReg.Text = ""
txtAnoReg.Text = ""
txtNumProc.Text = ""
txtAnoProc.Text = ""
txtAssunto.Text = ""
txtProp.Text = ""
txtCCusto.Text = ""
txtNomeLogr.Text = ""
txtCodLogr.Text = ""
cmbBairro.ListIndex = -1
End Sub

Private Sub cmdPesq_Click()
Dim Sql As String, RdoAux As rdoResultset, bCid As Boolean, bCC As Boolean, bLogr As Boolean, sLogr As String, bBairro As Boolean
Dim itmX As ListItem, z As Long, bNumReg As Boolean, bAnoReg As Boolean, bNumProc As Boolean, bAnoProc As Boolean, bAssunto As Boolean

If Val(txtNumReg.Text) = 0 And Val(txtAnoReg.Text) = 0 And Val(txtNumProc.Text) = 0 And Val(txtAnoProc.Text) = 0 And Trim(txtAssunto.Text) = "" And _
    Trim(txtProp.Text) = "" And Trim(txtCCusto.Text) = "" And Val(txtCodLogr) = 0 And cmbBairro.ListIndex = -1 Then
    MsgBox "Favor selecionar ao menos um critério para busca.", vbExclamation, "Atenção"
    Exit Sub
End If

Screen.MousePointer = vbHourglass
Ocupado

z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)

bNumReg = False
bAnoReg = False
bNumProc = False
bAnoProc = False
bAssunto = False
bLogr = False
bCid = False
bCC = False

If Val(txtNumReg.Text) > 0 Then bNumReg = True
If Val(txtAnoReg.Text) > 0 Then bAnoReg = True
If cmbBairro.ListIndex > 0 Then bBairro = True
If Val(txtNumProc.Text) > 0 Then bNumProc = True
If Val(txtCodLogr.Text) > 0 Then bLogr = True
If Val(txtAnoProc.Text) > 0 Then bAnoProc = True
If Trim(txtAssunto.Text) <> "" Then bAssunto = True
If Trim(txtProp.Text) <> "" Then bCid = True
If Trim(txtCCusto.Text) <> "" Then bCC = True

Sql = "SELECT registroatendimento.numreg, registroatendimento.anoreg, registroatendimento.atendente, registroatendimento.tipoatendimento, registroatendimento.obstipo, "
Sql = Sql & "registroatendimento.data, registroatendimento.numproc, registroatendimento.anoproc, registroatendimento.urgente, registroatendimento.assunto,"
Sql = Sql & "registroatendimento.aguardo, registroatendimento.deferido, registroatendimento.indeferido, registroatendimento.dataexec, registroatendimento.dataend,"
Sql = Sql & "registroatendimento.solucao, registroatendimento.cidadao, registroatendimento.ccusto, registroatendimento.codlogr, registroatendimento.codbairro,"
Sql = Sql & "vwFULLCIDADAO.nomecidadao,  vwFULLCIDADAO.descbairro, vwFULLCIDADAO.desccidade,"
Sql = Sql & "vwFULLCIDADAO.Endereco , vwFULLCIDADAO.NomeLogradouro, vwFULLCIDADAO.numimovel, vwFULLCIDADAO.NOMELOGRADOURO2 FROM registroatendimento LEFT OUTER JOIN "
Sql = Sql & "vwFULLCIDADAO ON registroatendimento.cidadao = vwFULLCIDADAO.codcidadao Where 1 = 1 And "
If bNumReg Then
    Sql = Sql & "NUMREG=" & Val(txtNumReg.Text) & " AND "
End If
If bAnoReg Then
   Sql = Sql & "ANOREG=" & Val(txtAnoReg.Text) & " AND "
End If
If bLogr Then
   Sql = Sql & "CODLOGR=" & Val(txtCodLogr.Text) & " AND "
End If
If bNumProc Then
    Sql = Sql & "NUMPROC=" & Val(Left(txtNumProc.Text, Len(txtNumProc.Text) - 1)) & " AND "
End If
If bAnoProc Then
   Sql = Sql & "ANOPROC=" & Val(txtAnoProc.Text) & " AND "
End If
If bAssunto Then
   Sql = Sql & "ASSUNTO LIKE '%" & Mask(txtAssunto.Text) & "%' AND "
End If
If bBairro Then
   Sql = Sql & "REGISTROATENDIMENTO.CODBAIRRO=" & cmbBairro.ItemData(cmbBairro.ListIndex) & " AND "
End If
If bCid Then
   Sql = Sql & "CODCIDADAO=" & Val(Left(txtProp.Text, 6)) & " AND "
End If
If bCC Then
   Sql = Sql & "CCUSTO LIKE '%" & Mask(txtAssunto.Text) & "%' AND "
End If
Sql = Left$(Sql, Len(Sql) - 5)

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
    With RdoAux
        Do Until .EOF
            Set itmX = lvMain.ListItems.Add(, "C" & Format(!NUMREG, "00000") & Format(!ANOREG, "0000"), Format(!NUMREG, "00000"))
            itmX.SubItems(1) = !ANOREG
            itmX.SubItems(2) = Format(!NumProc, "00000") & "-" & RetornaDVProcesso(!NumProc)
            itmX.SubItems(3) = !AnoProc
            itmX.SubItems(4) = SubNull(!assunto)
            itmX.SubItems(5) = SubNull(!nomecidadao)
            itmX.SubItems(6) = SubNull(!ccusto)
            sLogr = ""
            If IsNull(!ccusto) Then
                If Not IsNull(!NomeLogradouro) Then
                    sLogr = !NomeLogradouro
                Else
                    If Not IsNull(!NOMELOGRADOURO2) Then
                        sLogr = !NOMELOGRADOURO2
                    End If
                End If
            Else
            End If
            itmX.SubItems(7) = sLogr
            itmX.SubItems(8) = Val(SubNull(!NUMIMOVEL))
            itmX.SubItems(9) = SubNull(!DescBairro)
           .MoveNext
        Loop
       .Close
    End With
Else
   MsgBox "Não existem registros com estes parâmetros.", vbExclamation, "Atenção"
End If
Liberado
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

cmbBairro.AddItem ""
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND DESCBAIRRO<>'' ORDER BY DESCBAIRRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbBairro.AddItem !DescBairro
        cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub txtAnoProc_KeyPress(KeyAscii As Integer)
Tweak txtAnoProc, KeyAscii, IntegerPositive
End Sub

Private Sub txtAnoReg_KeyPress(KeyAscii As Integer)
Tweak txtAnoReg, KeyAscii, IntegerPositive
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

If lstNomeLog.ListIndex > -1 Then
   txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   
   txtCodLogr_LostFocus
   lstNomeLog.Visible = False
   cmbBairro.SetFocus
End If
End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
       cmbBairro.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
   txtNomeLogr.SetFocus
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

Private Sub txtNumProc_KeyPress(KeyAscii As Integer)
Tweak txtNumProc, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumReg_KeyPress(KeyAscii As Integer)
Tweak txtNumReg, KeyAscii, IntegerPositive
End Sub
