VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCnsProcesso 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta rápida a números de processo"
   ClientHeight    =   5340
   ClientLeft      =   2070
   ClientTop       =   2475
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7605
   Begin VB.TextBox txtAssunto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1980
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox txtAno2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2970
      MaxLength       =   4
      TabIndex        =   3
      Top             =   810
      Width           =   795
   End
   Begin VB.TextBox txtAno1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   2
      Top             =   810
      Width           =   795
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1980
      TabIndex        =   0
      Top             =   150
      Width           =   5535
   End
   Begin MSComctlLib.ListView lvProc 
      Height          =   3645
      Left            =   60
      TabIndex        =   4
      Top             =   1230
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6429
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome do Contribuinte"
         Object.Width           =   5715
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ano"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nº Proc."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Assunto"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dt.Entrada"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dt.Cancel"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Dt.Arquiv"
         Object.Width           =   2117
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdPesq 
      Default         =   -1  'True
      Height          =   345
      Left            =   5130
      TabIndex        =   8
      ToolTipText     =   "Pesquisar"
      Top             =   4920
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCnsProcesso.frx":0000
      PICN            =   "frmCnsProcesso.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdObs 
      Height          =   345
      Left            =   3750
      TabIndex        =   10
      ToolTipText     =   "Observação"
      Top             =   4920
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Observação"
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
      MICON           =   "frmCnsProcesso.frx":0176
      PICN            =   "frmCnsProcesso.frx":0192
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
      Cancel          =   -1  'True
      Height          =   345
      Left            =   6330
      TabIndex        =   11
      ToolTipText     =   "Sair da Tela"
      Top             =   4920
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCnsProcesso.frx":059C
      PICN            =   "frmCnsProcesso.frx":05B8
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
      Caption         =   "Assunto Parcial.........:"
      Height          =   225
      Index           =   3
      Left            =   210
      TabIndex        =   9
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Final:"
      Height          =   225
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   870
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Inicial:"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   870
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Contribuinte:"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   5
      Top             =   210
      Width           =   1635
   End
End
Attribute VB_Name = "frmCnsProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String

Private Sub cmdObs_Click()
Dim nAno As Integer, nNumProc As Long, sObs As String

If lvProc.ListItems.Count = 0 Then Exit Sub

nAno = lvProc.SelectedItem.SubItems(1)
nNumProc = lvProc.SelectedItem.SubItems(2)


Sql = "SELECT ANOPROCESS,NUMEROPROC,OBSERVACAO FROM PROCESSO "
Sql = Sql & "WHERE PROCESSO.AnoProcess =" & nAno & " AND PROCESSO.NUMEROPROC=" & nNumProc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If Not IsNull(!OBSERVACAO) Then
        sObs = Trim$(!OBSERVACAO)
        MsgBox sObs, vbInformation, "OBSERVAÇÃO"
    Else
        MsgBox "SEM OBSERVAÇÃO", vbInformation, "OBSERVAÇÃO"
    End If
   .Close
End With

End Sub

Private Sub cmdPesq_Click()
MsgBox "Desativado"
Exit Sub
If Trim$(txtNome.text) = "" And Trim$(txtAssunto.text) = "" Then
    MsgBox "Digite o Nome do Contribuinte e/ou Assunto.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno1.text) < 1920 Or Val(txtAno1.text) > Year(Now) Then
    MsgBox "Ano inicial inválido", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno2.text) < 1920 Or Val(txtAno2.text) > Year(Now) + 1 Then
    MsgBox "Ano Final inválido", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno1.text) > Val(txtAno2.text) Then
    MsgBox "Ano inicial maior que ano final inválido", vbExclamation, "Atenção"
    Exit Sub
End If

CarregaLista
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmMdi.AddWindow Me.Name, Me.Caption
Centraliza Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
End Sub

Private Sub lvProc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvProc.SortKey = ColumnHeader.Position - 1
lvProc.Sorted = True
lvProc.SortOrder = lvwAscending

End Sub

Private Sub CarregaLista()

Dim itmX As ListItem
Dim z As Long
z = SendMessage(lvProc.hwnd, LVM_DELETEALLITEMS, 0, 0)
Ocupado
Sql = "SELECT CIDADAO.NOMECIDADAO,PROCESSO.* FROM PROCESSO INNER JOIN "
Sql = Sql & "CIDADAO ON PROCESSO.CODCIDAPro = CIDADAO.CODCIDADAO "
Sql = Sql & "WHERE (PROCESSO.AnoProcess BETWEEN " & Val(txtAno1.text) & " AND " & Val(txtAno2.text) & ") AND (CIDADAO.NOMECIDADAO LIKE '" & Trim$(txtNome.text) & "%') "
Sql = Sql & "AND (protocolo.dbo.PROCESSO.COMPLEASSU LIKE '%" & Trim$(txtAssunto.text) & "%')"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       Set itmX = lvProc.ListItems.Add(, , !NOMECIDADAO)
       itmX.SubItems(1) = !ANOPROCESS
       itmX.SubItems(2) = !NUMEROPROC
       itmX.SubItems(3) = !compleassu
       itmX.SubItems(4) = Format(!DATAENTRAD, "dd/mm/yyyy")
       If Year(!DATACANCEL) > 1900 Then
          itmX.SubItems(5) = Format(!DATACANCEL, "dd/mm/yyyy")
       Else
          itmX.SubItems(5) = "- - - -"
       End If
       If Year(!DATAARQUIV) > 1900 Then
          itmX.SubItems(6) = Format(!DATAARQUIV, "dd/mm/yyyy")
       Else
          itmX.SubItems(6) = "- - - -"
       End If
      .MoveNext
    Loop
   .Close
End With
If lvProc.ListItems.Count = 0 Then MsgBox "Nenhum item coincidente.", vbInformation, "Atenção"

Liberado
End Sub

