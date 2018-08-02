VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSegmentoLogradouro 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segmentos de Logradouros"
   ClientHeight    =   4560
   ClientLeft      =   2595
   ClientTop       =   2580
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1455
      Left            =   45
      TabIndex        =   13
      Top             =   2610
      Width           =   6000
      Begin VB.ComboBox cmbZona 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   990
         Width           =   3435
      End
      Begin VB.ComboBox cmbBairro 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   585
         Width           =   4830
      End
      Begin VB.TextBox txtNumFim 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4455
         MaxLength       =   4
         TabIndex        =   3
         Top             =   225
         Width           =   735
      End
      Begin VB.TextBox txtNumIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2700
         MaxLength       =   4
         TabIndex        =   2
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Zona........:"
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro........:"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblSegmento 
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1125
         TabIndex        =   17
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Final.:"
         Height          =   240
         Index           =   2
         Left            =   3690
         TabIndex        =   16
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Inicial.:"
         Height          =   240
         Index           =   1
         Left            =   1890
         TabIndex        =   15
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Segmento.:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.ComboBox lstRua 
      Height          =   315
      Left            =   90
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   6000
   End
   Begin MSComctlLib.ListView lvSeg 
      Height          =   1695
      Left            =   45
      TabIndex        =   1
      Top             =   900
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
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
         Text            =   "Seg"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nº Inicial"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nº Final"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Zona"
         Object.Width           =   1307
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CodZona"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Bairro"
         Object.Width           =   4657
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CodBairro"
         Object.Width           =   0
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1353
      TabIndex        =   7
      ToolTipText     =   "Alterar o segmento"
      Top             =   4140
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmSegmentoLogradouro.frx":0000
      PICN            =   "frmSegmentoLogradouro.frx":001C
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
      Left            =   2616
      TabIndex        =   8
      ToolTipText     =   "Remover o segmento selecionado"
      Top             =   4140
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "E&xcluir"
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
      MICON           =   "frmSegmentoLogradouro.frx":0176
      PICN            =   "frmSegmentoLogradouro.frx":0192
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
      Left            =   90
      TabIndex        =   6
      ToolTipText     =   "Incluir novo segmento"
      Top             =   4140
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmSegmentoLogradouro.frx":0234
      PICN            =   "frmSegmentoLogradouro.frx":0250
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
      Left            =   3879
      TabIndex        =   9
      ToolTipText     =   "Gravar os Dados"
      Top             =   4140
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
      MICON           =   "frmSegmentoLogradouro.frx":03AA
      PICN            =   "frmSegmentoLogradouro.frx":03C6
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
      Height          =   315
      Left            =   4995
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   4140
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
      MICON           =   "frmSegmentoLogradouro.frx":076B
      PICN            =   "frmSegmentoLogradouro.frx":0787
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
      Height          =   315
      Left            =   4995
      TabIndex        =   20
      ToolTipText     =   "Sair da Tela"
      Top             =   4140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmSegmentoLogradouro.frx":08E1
      PICN            =   "frmSegmentoLogradouro.frx":08FD
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
      Caption         =   "Segmentos do Logradouro:"
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
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   12
      Top             =   675
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Logradouro:"
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
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   90
      Width           =   1995
   End
End
Attribute VB_Name = "frmSegmentoLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sEvento As String, Evento As String
Dim bExec As Boolean

Private Sub cmdAlterar_Click()
Evento = "Alterar"
Eventos "INCLUIR"
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
End Sub

Private Sub cmdExcluir_Click()
Dim nSeq As Integer, nCodLogr As Integer

If lvSeg.ListItems.Count = 0 Then Exit Sub
nSeq = Val(lblSegmento.Caption)
nCodLogr = lstRua.ItemData(lstRua.ListIndex)

If MsgBox("Excluir este segmento?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "delete from segmentologradouro where codlogradouro=" & nCodLogr & " and segmento=" & nSeq
    cn.Execute Sql, rdExecDirect
    lstRua_Click
End If

End Sub

Private Sub cmdGravar_Click()

If cmbZona.ListIndex = -1 Then
    MsgBox "Selecione uma zona.", vbExclamation, "Atenção"
    Exit Sub
End If

If Trim$(txtNumIni.Text) = "" Then txtNumIni.Text = 0
If Trim$(txtNumFim.Text) = "" Then txtNumFim.Text = 0

If Val(txtNumFim.Text) > 0 Then
     If Val(txtNumIni.Text) > Val(txtNumFim.Text) Then
          MsgBox "O valor final tem que ser maior que o valor inicial.", vbExclamation, "Atenção"
          Exit Sub
     End If
End If

Grava

Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()

If lstRua.ListIndex = -1 Then
    MsgBox "Selecione uma rua.", vbExclamation, "Atenção"
    Exit Sub
End If

Evento = "Novo"
Eventos "INCLUIR"
Limpa
lblSegmento.Caption = Format(lvSeg.ListItems.Count + 1, "000")

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

bExec = False
Sql = "SELECT CODLOGRADOURO,LOGRADOURO FROM vwLOGRADOURO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly)
With RdoAux
    Do Until .EOF
        lstRua.AddItem !Logradouro
        lstRua.ItemData(lstRua.NewIndex) = !CODLOGRADOURO
       .MoveNext
    Loop
   .Close
End With
bExec = True

Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 ORDER BY DESCBAIRRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly)
With RdoAux
    Do Until .EOF
        cmbBairro.AddItem !DescBairro
        cmbBairro.ItemData(cmbBairro.NewIndex) = !codbairro
       .MoveNext
    Loop
   .Close
End With
cmbBairro.ListIndex = 0

Sql = "SELECT CODIGO,SIGLAZONA,DESCZONA FROM ZONA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly)
With RdoAux
    Do Until .EOF
        cmbZona.AddItem !siglazona & " - " & !DESCZONA
        cmbZona.ItemData(cmbZona.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
Eventos "INICIAR"
End Sub

Private Sub lstRua_Click()
If Not bExec Then Exit Sub
If lstRua.ListIndex > -1 Then
    Limpa
    CarregaLista
End If
End Sub

Private Sub lvSeg_ItemClick(ByVal Item As MSComctlLib.ListItem)
Le
End Sub

Private Sub txtNumFim_GotFocus()
txtNumFim.SelStart = 0
txtNumFim.SelLength = Len(txtNumFim.Text)
End Sub

Private Sub txtNumFim_KeyPress(KeyAscii As Integer)
Tweak txtNumFim, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumIni_GotFocus()
txtNumIni.SelStart = 0
txtNumIni.SelLength = Len(txtNumIni.Text)
End Sub

Private Sub txtNumIni_KeyPress(KeyAscii As Integer)
Tweak txtNumIni, KeyAscii, IntegerPositive
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmdSair.Visible = True
   lstRua.Enabled = True
   lstRua.BackColor = Branco
   lvSeg.Enabled = True
   lvSeg.BackColor = Branco
   txtNumIni.Locked = True
   txtNumIni.BackColor = Kde
   txtNumFim.Locked = True
   txtNumFim.BackColor = Kde
   cmbBairro.Enabled = False
   cmbBairro.BackColor = Kde
   cmbZona.Enabled = False
   cmbZona.BackColor = Kde
   
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   lstRua.Enabled = False
   lstRua.BackColor = Kde
   lvSeg.Enabled = False
   lvSeg.BackColor = Kde
   txtNumIni.Locked = False
   txtNumIni.BackColor = Branco
   txtNumFim.Locked = False
   txtNumFim.BackColor = Branco
   cmbBairro.Enabled = True
   cmbBairro.BackColor = Branco
   cmbZona.Enabled = True
   cmbZona.BackColor = Branco
   
End If

End Sub

Private Sub Limpa()

lblSegmento.Caption = "000"
txtNumIni.Text = "0"
txtNumFim.Text = "0"
cmbBairro.ListIndex = 0
cmbZona.ListIndex = -1
End Sub

Private Sub Grava()
Dim Sql As String, nCodLogr As Integer, nCodBairro As Integer, nCodZona As Integer, nSeq As Integer, itmX As ListItem

nSeq = Val(lblSegmento.Caption)
nCodLogr = lstRua.ItemData(lstRua.ListIndex)
nCodBairro = cmbBairro.ItemData(cmbBairro.ListIndex)
nCodZona = cmbZona.ItemData(cmbZona.ListIndex)

If Evento = "Novo" Then
    Sql = "insert segmentologradouro(codlogradouro,segmento,numini,numfim,zona,codbairro) values(" & nCodLogr & "," & nSeq & ","
    Sql = Sql & Val(txtNumIni.Text) & "," & Val(txtNumFim.Text) & "," & nCodZona & "," & nCodBairro & ")"
    cn.Execute Sql, rdExecDirect
    
    Set itmX = lvSeg.ListItems.Add(, , lblSegmento.Caption)
    itmX.SubItems(1) = Format(txtNumIni.Text, "0000")
    itmX.SubItems(2) = Format(txtNumFim.Text, "0000")
    itmX.SubItems(3) = Left(cmbZona.Text, 2)
    itmX.SubItems(4) = nCodZona
    itmX.SubItems(5) = cmbBairro.Text
    itmX.SubItems(6) = nCodBairro
Else
    Sql = "update segmentologradouro set numini=" & Val(txtNumIni.Text) & ",numfim=" & Val(txtNumFim.Text) & ",zona=" & nCodZona & ","
    Sql = Sql & "codbairro=" & nCodBairro & " where codlogradouro=" & nCodLogr & " and segmento=" & nSeq
    cn.Execute Sql, rdExecDirect
        
    lvSeg.SelectedItem.SubItems(1) = Format(txtNumIni.Text, "0000")
    lvSeg.SelectedItem.SubItems(2) = Format(txtNumFim.Text, "0000")
    lvSeg.SelectedItem.SubItems(3) = Left(cmbZona.Text, 2)
    lvSeg.SelectedItem.SubItems(4) = nCodZona
    lvSeg.SelectedItem.SubItems(5) = cmbBairro.Text
    lvSeg.SelectedItem.SubItems(6) = nCodBairro
End If

End Sub

Private Sub CarregaLista()
Dim RdoAux As rdoResultset, Sql As String, itmX As ListItem
Dim z As Long
z = SendMessage(lvSeg.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "select * from vwzoneamento where codlogradouro=" & lstRua.ItemData(lstRua.ListIndex) & " order by segmento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvSeg.ListItems.Add(, , Format(!Segmento, "000"))
        itmX.SubItems(1) = Format(!numini, "0000")
        itmX.SubItems(2) = Format(!numfim, "0000")
        itmX.SubItems(3) = !siglazona
        itmX.SubItems(4) = !zona
        itmX.SubItems(5) = !DescBairro
        itmX.SubItems(6) = !codbairro
        
       .MoveNext
    Loop
   .Close
End With

If lvSeg.ListItems.Count > 0 Then
    Le
End If

End Sub

Private Sub Le()
Dim x As Integer, sZona As String

If lvSeg.ListItems.Count = 0 Then Exit Sub

lblSegmento.Caption = lvSeg.SelectedItem.Text
txtNumIni.Text = lvSeg.SelectedItem.SubItems(1)
txtNumFim.Text = lvSeg.SelectedItem.SubItems(2)
cmbBairro.Text = lvSeg.SelectedItem.SubItems(5)
sZona = lvSeg.SelectedItem.SubItems(3)

For x = 0 To cmbZona.ListCount - 1
    If Left(cmbZona.List(x), 2) = sZona Then
        cmbZona.ListIndex = x
        Exit For
    End If
Next

End Sub
