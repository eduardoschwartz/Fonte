VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProdutividadeEvento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtividade - Cadastro de eventos"
   ClientHeight    =   4035
   ClientLeft      =   4650
   ClientTop       =   3405
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   5250
   Begin MSComCtl2.DTPicker mskDataIni 
      Height          =   315
      Left            =   900
      TabIndex        =   3
      Top             =   3180
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   133234689
      CurrentDate     =   40968
      MaxDate         =   43830
      MinDate         =   40544
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   345
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Editar Registro"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProdutividadeEvento.frx":0000
      PICN            =   "frmProdutividadeEvento.frx":001C
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
      Height          =   345
      Left            =   4020
      TabIndex        =   7
      ToolTipText     =   "Excluir Registro"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProdutividadeEvento.frx":0176
      PICN            =   "frmProdutividadeEvento.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbEvento 
      Height          =   315
      ItemData        =   "frmProdutividadeEvento.frx":0234
      Left            =   900
      List            =   "frmProdutividadeEvento.frx":0236
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2790
      Width           =   2625
   End
   Begin VB.ComboBox cmbFiscal 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   4305
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   1995
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   3519
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Seq"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Evento"
         Object.Width           =   4411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Inicio"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Final"
         Object.Width           =   1941
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   345
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Gravar os Dados"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProdutividadeEvento.frx":0238
      PICN            =   "frmProdutividadeEvento.frx":0254
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
      Height          =   345
      Left            =   4020
      TabIndex        =   9
      ToolTipText     =   "Cancelar Edição"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmProdutividadeEvento.frx":05F9
      PICN            =   "frmProdutividadeEvento.frx":0615
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
      Height          =   345
      Left            =   1740
      TabIndex        =   5
      ToolTipText     =   "Novo Registro"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProdutividadeEvento.frx":076F
      PICN            =   "frmProdutividadeEvento.frx":078B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker mskDataFim 
      Height          =   315
      Left            =   2940
      TabIndex        =   4
      Top             =   3180
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   133234689
      CurrentDate     =   40968
      MaxDate         =   43830
      MinDate         =   40544
   End
   Begin VB.Label lblSeq 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3960
      TabIndex        =   14
      Top             =   2820
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Evento..:"
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   13
      Top             =   2850
      Width           =   705
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Até...:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2430
      TabIndex        =   12
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "De..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   450
      TabIndex        =   11
      Top             =   3210
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal....:"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   210
      Width           =   705
   End
End
Attribute VB_Name = "frmProdutividadeEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bNovo As Boolean

Private Sub cmbFiscal_Click()

If cmbFiscal.ListIndex = -1 Then Exit Sub
CarregaLista

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset, z As Long, itmX As ListItem

lvMain.ListItems.Clear
Sql = "select seq,dataini,datafim,nome from produtividadefiscalevento inner join produtividadeevento "
Sql = Sql & " ON produtividadefiscalevento.codevento=produtividadeevento.codigo "
Sql = Sql & " where codfiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex) & " order by dataini desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Set itmX = lvMain.ListItems.Add(, , Format(!Seq, "000"))
       itmX.SubItems(1) = !Nome
       itmX.SubItems(2) = Format(!dataini, "dd/mm/yyyy")
       itmX.SubItems(3) = Format(!datafim, "dd/mm/yyyy")
       .MoveNext
    Loop
   .Close
End With

If lvMain.ListItems.Count > 0 Then
    lvMain.ListItems(1).Selected = True
    lvMain.SetFocus
    Le
End If
End Sub

Private Sub cmdAlterar_Click()
If lvMain.ListItems.Count = 0 Then
    MsgBox "Selecione um item", vbCritical, "Atenção"
    Exit Sub
End If
bNovo = False
ControlBehaviour (False)
End Sub

Private Sub cmdCancel_Click()
ControlBehaviour (True)
End Sub

Private Sub cmdExcluir_Click()
If lvMain.ListItems.Count = 0 Then
    MsgBox "Selecione um item", vbCritical, "Atenção"
    Exit Sub
End If

If MsgBox("Excluir o período selecionado?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Sql = "delete from produtividadefiscalevento where codfiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex) & " and seq=" & Val(lblSeq.Caption)
cn.Execute Sql, rdExecDirect
Limpa
CarregaLista

End Sub

Private Sub cmdGravar_Click()
Dim itmX As ListItem, nSeq As Integer

If cmbEvento.ListIndex = -1 Then
    MsgBox "Selecione um evento", vbCritical, "Atenção"
    Exit Sub
End If

If CDate(mskDataIni.value) > CDate(mskDataFim.value) Then
    MsgBox "Data inicial tem que ser menor que a data final", vbCritical, "Atenção"
    Exit Sub
End If

If cmbEvento.ItemData(cmbEvento.ListIndex) = 1 And IsTheBoss Then
    MsgBox "Fiscal com cargo de chefia não pode escolher o evento chefia", vbCritical, "Atenção"
    Exit Sub
End If


If Not ValidaPeriodo Then Exit Sub

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM produtividadefiscalevento where codfiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeq = 1
Else
    nSeq = RdoAux!maximo + 1
End If
RdoAux.Close

If bNovo Then
    Sql = "insert produtividadefiscalevento(codfiscal,seq,codevento,dataini,datafim) values("
    Sql = Sql & cmbFiscal.ItemData(cmbFiscal.ListIndex) & "," & nSeq & "," & cmbEvento.ItemData(cmbEvento.ListIndex) & ",'"
    Sql = Sql & Format(mskDataIni.value, "mm/dd/yyyy") & "','" & Format(mskDataFim.value, "mm/dd/yyyy") & "')"
Else
    Sql = "update produtividadefiscalevento set codevento=" & cmbEvento.ItemData(cmbEvento.ListIndex) & ",dataini='" & Format(mskDataIni.value, "mm/dd/yyyy") & "',"
    Sql = Sql & "datafim='" & Format(mskDataFim.value, "mm/dd/yyyy") & "' where codfiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex)
    Sql = Sql & " and seq=" & Val(lblSeq.Caption)
End If
cn.Execute Sql, rdExecDirect

ControlBehaviour (True)
CarregaLista
End Sub

Private Function IsTheBoss()
Dim Sql As String, RdoAux As rdoResultset, bRet As Boolean, sNomeDeLogin As String

Sql = "select codigo,chefe from produtividadefiscal where nome='" & RetornaUsuarioLoginName(cmbFiscal.Text) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!Chefe) Then
    bRet = False
Else
    bRet = RdoAux!Chefe
End If
RdoAux.Close

IsTheBoss = bRet
End Function

Private Function ValidaPeriodo() As Boolean
Dim bRet As Boolean, dDataIni1 As Date, dDataIni2 As Date, dDataFim1 As Date, dDataFim2 As Date, x As Integer
Dim bAchou As Boolean

bAchou = False

dDataIni1 = CDate(mskDataIni.value)
dDataFim1 = CDate(mskDataFim.value)

For x = 1 To lvMain.ListItems.Count
    If Not bNovo Then
        If lvMain.ListItems(x).Text = lblSeq.Caption Then
            GoTo proximo
        End If
    End If
            
    dDataIni2 = CDate(lvMain.ListItems(x).SubItems(2))
    dDataFim2 = CDate(lvMain.ListItems(x).SubItems(3))
    
    If dDataIni1 >= dDataIni2 And dDataIni1 <= dDataFim2 Then
        bAchou = True
        Exit For
    End If
    
    
    If dDataIni1 >= dDataIni2 And dDataIni1 <= dDataFim2 Then
        bAchou = True
        Exit For
    End If
    
    If dDataFim1 >= dDataIni2 And dDataFim1 <= dDataFim2 Then
        bAchou = True
        Exit For
    End If
proximo:
Next

'pesquisa inversa
dDataIni2 = CDate(mskDataIni.value)
dDataFim2 = CDate(mskDataFim.value)

For x = 1 To lvMain.ListItems.Count
    If Not bNovo Then
        If lvMain.ListItems(x).Text = lblSeq.Caption Then
            GoTo PROXIMO2
        End If
    End If
            
    dDataIni1 = CDate(lvMain.ListItems(x).SubItems(2))
    dDataFim1 = CDate(lvMain.ListItems(x).SubItems(3))
    
    If dDataIni1 >= dDataIni2 And dDataIni1 <= dDataFim2 Then
        bAchou = True
        Exit For
    End If
    
    
    If dDataIni1 >= dDataIni2 And dDataIni1 <= dDataFim2 Then
        bAchou = True
        Exit For
    End If
    
    If dDataFim1 >= dDataIni2 And dDataFim1 <= dDataFim2 Then
        bAchou = True
        Exit For
    End If
PROXIMO2:
Next

If bAchou Then
    MsgBox "O período digitado esta em conflito com outro período já cadastrado. Verifique.", vbCritical, "Atenção"
    bRet = False
Else
    bRet = True
End If

ValidaPeriodo = bRet
End Function


Private Sub cmdNovo_Click()
If cmbFiscal.ListIndex = -1 Then
    MsgBox "Selecione o fiscal", vbCritical, "Atenção"
    Exit Sub
End If

bNovo = True
Limpa
ControlBehaviour (False)
End Sub

Private Sub Form_Load()
Dim RdoAux As rdoResultset, Sql As String

bNovo = False
Centraliza Me

Sql = "select codigo,nome,nomecompleto from produtividadefiscal inner join "
Sql = Sql & "usuario on produtividadefiscal.nome = usuario.nomelogin order by nomecompleto; "
Sql = Sql & "select codigo,nome from produtividadeevento where codigo>0 order by nome"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbFiscal.AddItem !NomeCompleto
        cmbFiscal.ItemData(cmbFiscal.NewIndex) = !Codigo
       .MoveNext
    Loop
    .MoreResults
    Do Until .EOF
        cmbEvento.AddItem !Nome
        cmbEvento.ItemData(cmbEvento.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

ControlBehaviour (True)
End Sub

Private Sub ControlBehaviour(bStart As Boolean)
cmdNovo.Visible = bStart
cmdAlterar.Visible = bStart
cmdExcluir.Visible = bStart
cmdGravar.Visible = Not bStart
cmdCancel.Visible = Not bStart
lvMain.Enabled = bStart
cmbFiscal.Enabled = bStart
cmbEvento.Enabled = Not bStart
mskDataIni.Enabled = Not bStart
mskDataFim.Enabled = Not bStart
If bStart Then
    lvMain.BackColor = Branco
    cmbFiscal.BackColor = Branco
    cmbEvento.BackColor = Me.BackColor
    mskDataIni.CalendarTitleBackColor = Me.BackColor
    mskDataFim.CalendarTitleBackColor = Me.BackColor
Else
    lvMain.BackColor = Me.BackColor
    cmbFiscal.BackColor = Me.BackColor
    cmbEvento.BackColor = Branco
    mskDataIni.CalendarTitleBackColor = Branco
    mskDataFim.CalendarTitleBackColor = Branco
End If

End Sub

Private Sub Limpa()
cmbEvento.ListIndex = -1
mskDataIni.value = Format(Now, "dd/mm/yyyy")
mskDataFim.value = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If cmdNovo.Visible = False Then
    MsgBox "Grave os dados ou cancele a operação antes de fechar a tela.", vbCritical, "Atenção"
    Cancel = 1
End If
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
Le
End Sub

Private Sub Le()
Dim x As Integer, sEvento As String

sEvento = lvMain.SelectedItem.SubItems(1)

For x = 0 To cmbEvento.ListCount - 1
    If cmbEvento.List(x) = sEvento Then
        Exit For
    End If
Next
cmbEvento.ListIndex = x
lblSeq.Caption = lvMain.SelectedItem.Text
mskDataIni.value = lvMain.SelectedItem.SubItems(2)
mskDataFim.value = lvMain.SelectedItem.SubItems(3)

End Sub
