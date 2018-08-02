VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProfissao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Profissões"
   ClientHeight    =   5655
   ClientLeft      =   15405
   ClientTop       =   3660
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstMain 
      Height          =   4740
      Left            =   60
      TabIndex        =   1
      Top             =   390
      Width           =   5025
   End
   Begin VB.TextBox txtBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5055
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Novo Registro"
      Top             =   5250
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
      MICON           =   "frmProfissao.frx":0000
      PICN            =   "frmProfissao.frx":001C
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
      Left            =   1170
      TabIndex        =   3
      ToolTipText     =   "Editar Registro"
      Top             =   5250
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
      MICON           =   "frmProfissao.frx":0176
      PICN            =   "frmProfissao.frx":0192
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
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Excluir Registro"
      Top             =   5250
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
      MICON           =   "frmProfissao.frx":02EC
      PICN            =   "frmProfissao.frx":0308
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
      Left            =   3750
      TabIndex        =   5
      ToolTipText     =   "Sair da Tela"
      Top             =   5250
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frmProfissao.frx":03AA
      PICN            =   "frmProfissao.frx":03C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmProfissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodProf As Integer

Public Property Let nCodProfissao(nValue As Integer)
    nCodProf = nValue
End Property

Private Sub cmdAlterar_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nCodAtual As Integer, x As Integer

If lstMain.ListIndex = -1 Then
    MsgBox "Selecione um item.", vbCritical, "Atenção"
    Exit Sub
End If

nCodAtual = lstMain.ItemData(lstMain.ListIndex)
If nCodAtual = 1 Then
    MsgBox "Esta profissão não pode ser alterada.", vbCritical, "ERRO"
    Exit Sub
End If


z = InputBox("Digite o nome da nova profissão.", "Alterar profissão")
If z <> "" Then
    z = UCase(z)
    Sql = "select * from profissao where nome='" & Mask(CStr(z)) & "' and codigo <> " & nCodAtual
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        MsgBox "Profissão já cadastrada.", vbCritical, "ERRO!"
        Exit Sub
    Else
        If MsgBox("Deseja alterar a profissão """ & lstMain.Text & """ para """ & z & """?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            Sql = "update profissao set nome ='" & Mask(CStr(z)) & "' where codigo=" & nCodAtual
            cn.Execute Sql, rdExecDirect
            txtBusca.Text = ""
            FillLista
            For x = 0 To lstMain.ListCount - 1
                If lstMain.ItemData(x) = nCodAtual Then
                    lstMain.ListIndex = x
                    Exit For
                End If
            Next

        End If
    End If
End If

End Sub

Private Sub cmdExcluir_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nCodAtual As Integer, x As Integer

If lstMain.ListIndex = -1 Then
    MsgBox "Selecione um item.", vbCritical, "Atenção"
    Exit Sub
End If
nCodAtual = lstMain.ItemData(lstMain.ListIndex)
If nCodAtual = 1 Then
    MsgBox "Esta profissão não pode ser excluida.", vbCritical, "ERRO"
    Exit Sub
End If

Sql = "select * from cidadao where codprofissao = " & nCodAtual
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Esta profissão esta em uso e não pode ser excluída.", vbCritical, "ERRO!"
    RdoAux.Close
    Exit Sub
End If

If MsgBox("Deseja excluir a profissão """ & lstMain.Text & """?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "delete from profissao where codigo=" & nCodAtual
    cn.Execute Sql, rdExecDirect
    txtBusca.Text = ""
    FillLista
End If

End Sub

Private Sub cmdNovo_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nMax As Integer, x As Integer

z = InputBox("Digite o nome da nova profissão.", "Nova profissão")
If z <> "" Then
    z = UCase(z)
    Sql = "select * from profissao where nome='" & Mask(CStr(z)) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        MsgBox "Profissão já cadastrada.", vbCritical, "ERRO!"
        Exit Sub
    Else
        If MsgBox("Deseja incluir a profissão """ & z & """?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            Sql = "select max(codigo) as maximo from profissao"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            nMax = RdoAux!maximo + 1
            RdoAux.Close
            Sql = "insert profissao(codigo,nome) values(" & nMax & ",'" & Mask(CStr(z)) & "')"
            cn.Execute Sql, rdExecDirect
            lstMain.AddItem z
            lstMain.ItemData(lstMain.NewIndex) = nMax
            For x = 0 To lstMain.ListCount - 1
                If lstMain.ItemData(x) = nMax Then
                    lstMain.ListIndex = x
                    Exit For
                End If
            Next
        End If
    End If
End If

End Sub

Private Sub cmdSair_Click()
If lstMain.ListIndex = -1 Then
    frmCidadao.txtProfissao.Text = "(Não especificado)"
    frmCidadao.txtProfissao.Tag = "1"
Else
    frmCidadao.txtProfissao.Text = lstMain.Text
    frmCidadao.txtProfissao.Tag = lstMain.ItemData(lstMain.ListIndex)
End If
Unload Me

End Sub

Private Sub Form_Load()
FillLista
End Sub

Private Sub txtBusca_Change()
FillLista
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
Tweak txtBusca, KeyAscii, AllLettersAllCaps

End Sub

Private Sub FillLista()
Dim Sql As String, RdoAux As rdoResultset, x As Integer
lstMain.Clear
Sql = "select * from profissao where  nome like '%" & Mask(txtBusca.Text) & "%' order by nome"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMain.AddItem !Nome
        lstMain.ItemData(lstMain.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

For x = 0 To lstMain.ListCount - 1
    If lstMain.ItemData(x) = nCodProf Then
        lstMain.ListIndex = x
        Exit For
    End If
Next

End Sub
