VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCidade 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cidades"
   ClientHeight    =   4170
   ClientLeft      =   17385
   ClientTop       =   2655
   ClientWidth     =   6105
   Icon            =   "frmCidade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstCidade 
      Height          =   2205
      Left            =   0
      TabIndex        =   6
      Top             =   540
      Width           =   6105
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   945
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   6105
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   2
         Top             =   525
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição...........:"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   4
         Top             =   585
         Width           =   1275
      End
   End
   Begin VB.ComboBox cmbUF 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   4275
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   60
      TabIndex        =   9
      ToolTipText     =   "Novo Registro"
      Top             =   3780
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
      MICON           =   "frmCidade.frx":014A
      PICN            =   "frmCidade.frx":0166
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
      Left            =   1110
      TabIndex        =   10
      ToolTipText     =   "Editar Registro"
      Top             =   3780
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
      MICON           =   "frmCidade.frx":02C0
      PICN            =   "frmCidade.frx":02DC
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
      Left            =   2160
      TabIndex        =   11
      ToolTipText     =   "Excluir Registro"
      Top             =   3780
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
      MICON           =   "frmCidade.frx":0436
      PICN            =   "frmCidade.frx":0452
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
      Left            =   3930
      TabIndex        =   12
      ToolTipText     =   "Gravar os Dados"
      Top             =   3780
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
      MICON           =   "frmCidade.frx":04F4
      PICN            =   "frmCidade.frx":0510
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
      Left            =   4950
      TabIndex        =   13
      ToolTipText     =   "Sair da Tela"
      Top             =   3780
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
      MICON           =   "frmCidade.frx":08B5
      PICN            =   "frmCidade.frx":08D1
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
      Left            =   4980
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   3780
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
      MICON           =   "frmCidade.frx":093F
      PICN            =   "frmCidade.frx":095B
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
      Caption         =   "UF:"
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "frmCidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String, bExec As Boolean
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Private Sub cmbUF_Click()
Ocupado
Limpa
CarregaLista
le
Liberado
End Sub

Private Sub cmdAlterar_Click()
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldDesc = txtDesc.Text
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()

On Error GoTo Erro
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir esta Cidade ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "DELETE FROM CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & txtCod.Text
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & txtDesc.Text
       Limpa
       CarregaLista
       le
    End If
    Exit Sub
Erro:

MsgBox "Não é possivel excluir esta cidade pois existem imóveis cadastrados nela.", vbExclamation, "Atenção"

End Sub

Private Sub cmdGravar_Click()
    If txtDesc.Text = "" Then
       MsgBox "Favor digitar a Descrição.", vbExclamation, "Atenção"
       txtDesc.SetFocus
       Exit Sub
    End If
    Grava
    Eventos "INICIAR"
End Sub


Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
Liberado
bResize = True
End Sub

Private Sub Form_Load()
Dim RdoAux2 As rdoResultset
Ocupado
Centraliza Me
bExec = True
sRet = RetEventUserForm(Me.Name)
Sql = "Select SIGLAUF,DESCUF From UF"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do Until .EOF
       cmbUF.AddItem !SiglaUF & "-" & !DESCUF
      .MoveNext
    Loop
   .Close
End With
If cmbUF.ListCount > 0 Then cmbUF.ListIndex = 0

lstCidade.Clear
CarregaLista
le

Eventos "INICIAR"

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmCidade
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
         Ct.Enabled = False
       End If
   Next
   lstCidade.Enabled = True
   cmbUF.Enabled = True
   cmbUF.BackColor = vbWhite
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmCidade
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   txtCod.BackColor = Kde
   txtCod.Locked = True
   lstCidade.Enabled = False
   cmbUF.Enabled = False
   cmbUF.BackColor = Kde
   txtDesc.SetFocus
End If

FormHagana

End Sub

Private Sub le()

'If lstCidade.ListCount > 0 Then lstCidade.ListIndex = 0
If lstCidade.ListIndex = -1 Then Exit Sub
txtCod.Text = lstCidade.ItemData(lstCidade.ListIndex)
txtDesc.Text = lstCidade.Text

End Sub

Private Sub Limpa()
txtCod.Text = ""
txtDesc.Text = ""
End Sub

Private Sub CarregaLista()
Dim lRet As Long
Dim s As String, n As Long
lstCidade.Clear

Sql = "Select CODCIDADE,DESCCIDADE From CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE<>999 ORDER BY DESCCIDADE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
'        lstCidade.AddItem !DESCCidade
'        lstCidade.ItemData(lstCidade.NewIndex) = !CODCIDADE
       s = !descCidade: n = !CodCidade
       lRet = SendMessage(lstCidade.HWND, LB_ADDSTRING, 0, ByVal s)
       SendMessage lstCidade.HWND, LB_SETITEMDATA, lRet, ByVal n
      .MoveNext
    Loop
   .Close
End With

If lstCidade.ListCount > 0 Then
    lstCidade.ListIndex = 0
End If

End Sub

Private Sub Grava()
Dim qd As New rdoQuery
Dim MaxCod As Integer

On Error Resume Next
RdoAux.Close
On Error GoTo 0
Set qd.ActiveConnection = cn

qd.Sql = "{ Call spGRAVACIDADE(?,?,?,?) }"
If Evento = "Novo" Then
   qd(0) = "S"
   qd(2) = 0
Else
   qd(0) = "N"
   qd(2) = txtCod.Text
End If
qd(1) = Left$(cmbUF.Text, 2)
qd(3) = txtDesc.Text
Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)

MaxCod = RdoAux.rdoColumns(0).value

If Evento = "Novo" Then
   txtCod.Text = MaxCod
   Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & txtDesc.Text
 ElseIf Evento = "Alterar" Then
   MaxCod = txtCod.Text
   Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.Text, "000") & " de " & sOldDesc & " para " & txtDesc.Text
End If

s = txtDesc.Text
cmbUF_Click
lstCidade.Text = s
le
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If cmdNovo.Visible = True Then
      cmdNovo_Click
   Else
      cmdGravar_Click
   End If
End If
End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub

Private Sub lstCidade_Click()
If bExec Then le
End Sub
