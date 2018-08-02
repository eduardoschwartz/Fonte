VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmPaises 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Países"
   ClientHeight    =   5685
   ClientLeft      =   6375
   ClientTop       =   5280
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   5055
   End
   Begin VB.ListBox lstMain 
      Height          =   4740
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   5025
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   90
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
      MICON           =   "frmPais.frx":0000
      PICN            =   "frmPais.frx":001C
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
      Left            =   1200
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
      MICON           =   "frmPais.frx":0176
      PICN            =   "frmPais.frx":0192
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
      Left            =   2310
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
      MICON           =   "frmPais.frx":02EC
      PICN            =   "frmPais.frx":0308
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
      Left            =   3780
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
      MICON           =   "frmPais.frx":03AA
      PICN            =   "frmPais.frx":03C6
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
Attribute VB_Name = "frmPaises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodPais As Integer, nCodPais2 As Integer, bResidencial As Boolean

Public Property Let residencial(bValue As Boolean)
    bResidencial = bValue
End Property

Public Property Let CodPais(nValue As Integer)
    nCodPais = nValue
End Property


Private Sub cmdAlterar_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nCodAtual As Integer, x As Integer

If lstMain.ListIndex = -1 Then
    MsgBox "Selecione um item.", vbCritical, "Atenção"
    Exit Sub
End If

nCodAtual = lstMain.ItemData(lstMain.ListIndex)
If nCodAtual = 1 Then
    MsgBox "Este país não pode ser alterado.", vbCritical, "ERRO"
    Exit Sub
End If


z = InputBox("Digite o nome do novo país.", "Alterar país")
If z <> "" Then
    z = UCase(z)
    Sql = "select * from pais where nome_pais='" & Mask(CStr(z)) & "' and id_pais <> " & nCodAtual
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        MsgBox "País já cadastrado.", vbCritical, "ERRO!"
        Exit Sub
    Else
        If MsgBox("Deseja alterar o país """ & lstMain.Text & """ para """ & z & """?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            Sql = "update pais set nome_pais ='" & Mask(CStr(z)) & "' where id_pais=" & nCodAtual
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
    MsgBox "Este país não pode ser excluido.", vbCritical, "ERRO"
    Exit Sub
End If

Sql = "select * from cidadao where codpais = " & nCodAtual & " or codpais2=" & nCodAtual
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este país esta em uso e não pode ser excluído.", vbCritical, "ERRO!"
    RdoAux.Close
    Exit Sub
End If

If MsgBox("Deseja excluir o país """ & lstMain.Text & """?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "delete from país where id_pais=" & nCodAtual
    cn.Execute Sql, rdExecDirect
    txtBusca.Text = ""
    FillLista
End If

End Sub

Private Sub cmdNovo_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nMax As Integer, x As Integer

z = InputBox("Digite o nome do novo país.", "Novo país")
If z <> "" Then
    z = UCase(z)
    Sql = "select * from país where nome_pais='" & Mask(CStr(z)) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        MsgBox "País já cadastrado.", vbCritical, "ERRO!"
        Exit Sub
    Else
        If MsgBox("Deseja incluir o país """ & z & """?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            Sql = "select max(id_pais) as maximo from pais"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            nMax = RdoAux!maximo + 1
            RdoAux.Close
            Sql = "insert pais(id_pais,nome_pais) values(" & nMax & ",'" & Mask(CStr(z)) & "')"
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

If bResidencial Then
    If lstMain.ListIndex = -1 Then
        frmCidadao.txtPais.Text = "(Não especificado)"
        frmCidadao.txtPais.Tag = "1"
    Else
        frmCidadao.txtPais.Text = lstMain.Text
        frmCidadao.txtPais.Tag = lstMain.ItemData(lstMain.ListIndex)
    End If
Else
    If lstMain.ListIndex = -1 Then
        frmCidadao.txtPais2.Text = "(Não especificado)"
        frmCidadao.txtPais2.Tag = "1"
    Else
        frmCidadao.txtPais2.Text = lstMain.Text
        frmCidadao.txtPais2.Tag = lstMain.ItemData(lstMain.ListIndex)
    End If
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
Sql = "select * from pais where  nome_pais like '%" & Mask(txtBusca.Text) & "%' order by nome_pais"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMain.AddItem !nome_pais
        lstMain.ItemData(lstMain.NewIndex) = !id_pais
       .MoveNext
    Loop
   .Close
End With

For x = 0 To lstMain.ListCount - 1
    If lstMain.ItemData(x) = nCodPais Then
        lstMain.ListIndex = x
        Exit For
    End If
Next

End Sub

