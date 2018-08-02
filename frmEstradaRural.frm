VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEstradaRural 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estradas rurais"
   ClientHeight    =   5100
   ClientLeft      =   3840
   ClientTop       =   2640
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   5700
   Begin VB.ListBox lstProd 
      Height          =   3570
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   45
      Width           =   5625
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1005
      Left            =   0
      TabIndex        =   5
      Top             =   3630
      Width           =   5655
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   7
         Top             =   555
         Width           =   4005
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição...........:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   9
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   255
         Width           =   1275
      End
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4530
      TabIndex        =   0
      ToolTipText     =   "Cancelar Edição"
      Top             =   4725
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
      MICON           =   "frmEstradaRural.frx":0000
      PICN            =   "frmEstradaRural.frx":001C
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
      TabIndex        =   1
      ToolTipText     =   "Novo Registro"
      Top             =   4725
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
      MICON           =   "frmEstradaRural.frx":0176
      PICN            =   "frmEstradaRural.frx":0192
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
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   4725
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
      MICON           =   "frmEstradaRural.frx":02EC
      PICN            =   "frmEstradaRural.frx":0308
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
      Left            =   2190
      TabIndex        =   3
      ToolTipText     =   "Excluir Registro"
      Top             =   4725
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
      MICON           =   "frmEstradaRural.frx":0462
      PICN            =   "frmEstradaRural.frx":047E
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
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "Gravar os Dados"
      Top             =   4725
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmEstradaRural.frx":0520
      PICN            =   "frmEstradaRural.frx":053C
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
      Left            =   4530
      TabIndex        =   11
      ToolTipText     =   "Sair da Tela"
      Top             =   4725
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
      MICON           =   "frmEstradaRural.frx":08E1
      PICN            =   "frmEstradaRural.frx":08FD
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
Attribute VB_Name = "frmEstradaRural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Private Sub cmdAlterar_Click()
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldDesc = UCase$(txtDesc.Text)
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir a estrada " & txtDesc.Text & " ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "SELECT * FROM CADASTRORURAL WHERE CODESTRADA=" & Val(txtCod.Text)
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       If RdoAux.RowCount = 0 Then
            Sql = "DELETE FROM ESTRADARURAL WHERE CODIGO=" & txtCod.Text
            cn.Execute Sql, rdExecDirect
            Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & UCase$(txtDesc.Text)
            Limpa
            CarregaLista
            Le
       Else
            MsgBox "Esta estrada esta sendo utilizada e não pode ser excluida.", vbExclamation, "Atenção"
       End If
    End If
End Sub

Private Sub cmdGravar_Click()
    If UCase$(txtDesc.Text) = "" Then
       MsgBox "Favor digitar o nome da estrada.", vbExclamation, "Atenção"
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
End Sub

Private Sub Form_Load()

Centraliza Me
sRet = RetEventUserForm(Me.Name)
CarregaLista
Le

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
   For Each Ct In frmEstradaRural
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   lstProd.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmEstradaRural
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   txtCod.BackColor = Kde
   txtCod.Locked = True
   lstProd.Enabled = False
   txtDesc.SetFocus
End If

End Sub

Private Sub Le()

If lstProd.ListCount > 0 Then
    If lstProd.ListIndex = -1 Then lstProd.ListIndex = 0
    txtCod.Text = lstProd.ItemData(lstProd.ListIndex)
    txtDesc.Text = lstProd.Text
End If

End Sub

Private Sub Limpa()
txtCod.Text = ""
txtDesc.Text = ""

End Sub

Private Sub CarregaLista()

Sql = "Select CODIGO,NOME FROM ESTRADARURAL ORDER BY NOME"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
lstProd.Clear
With RdoAux
    Do Until .EOF
       lstProd.AddItem !NOME
       lstProd.ItemData(lstProd.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Grava()
Dim nCodNovo As Integer

If Evento = "Novo" Then
    Sql = "SELECT MAX(CODIGO) AS MAXIMO FROM ESTRADARURAL"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!MAXIMO) Then
            nCodNovo = 1
        Else
            nCodNovo = !MAXIMO + 1
        End If
       .Close
    End With
    lstProd.AddItem txtDesc.Text
    lstProd.ItemData(lstProd.NewIndex) = nCodNovo
    txtCod.Text = nCodNovo
    Sql = "INSERT ESTRADARURAL (CODIGO,NOME) VALUES("
    Sql = Sql & nCodNovo & ",'" & UCase$(Mask(txtDesc.Text)) & "')"
Else
    lstProd.List(lstProd.ListIndex) = UCase$(Trim$(txtDesc.Text))
    Sql = "UPDATE ESTRADARURAL SET NOME='" & UCase$(Mask(txtDesc.Text)) & "'"
    Sql = Sql & " WHERE CODIGO=" & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect
      
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


Private Sub lstprod_Click()
Limpa
Le
End Sub


