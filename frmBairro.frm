VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmBairro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Bairros"
   ClientHeight    =   4590
   ClientLeft      =   19485
   ClientTop       =   3990
   ClientWidth     =   6120
   Icon            =   "frmBairro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstBairro 
      Height          =   2205
      Left            =   0
      TabIndex        =   10
      Top             =   930
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   945
      Left            =   0
      TabIndex        =   4
      Top             =   3150
      Width           =   6105
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   6
         Top             =   525
         Width           =   4095
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição...........:"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   8
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   255
         Width           =   1275
      End
   End
   Begin VB.ComboBox cmbUF 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   4275
   End
   Begin VB.ComboBox cmbCidade 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   4275
   End
   Begin prjChameleon.chameleonButton cmdRefresh 
      Height          =   315
      Left            =   5220
      TabIndex        =   9
      ToolTipText     =   "Atualizar Lista"
      Top             =   480
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "!"
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBairro.frx":014A
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
      TabIndex        =   11
      ToolTipText     =   "Cancelar Edição"
      Top             =   4200
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
      MICON           =   "frmBairro.frx":0166
      PICN            =   "frmBairro.frx":0182
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
      Left            =   60
      TabIndex        =   12
      ToolTipText     =   "Novo Registro"
      Top             =   4200
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
      MICON           =   "frmBairro.frx":02DC
      PICN            =   "frmBairro.frx":02F8
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
      TabIndex        =   13
      ToolTipText     =   "Editar Registro"
      Top             =   4200
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
      MICON           =   "frmBairro.frx":0452
      PICN            =   "frmBairro.frx":046E
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
      TabIndex        =   14
      ToolTipText     =   "Excluir Registro"
      Top             =   4200
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
      MICON           =   "frmBairro.frx":05C8
      PICN            =   "frmBairro.frx":05E4
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
      TabIndex        =   15
      ToolTipText     =   "Gravar os Dados"
      Top             =   4200
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
      MICON           =   "frmBairro.frx":0686
      PICN            =   "frmBairro.frx":06A2
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
      Left            =   4980
      TabIndex        =   16
      ToolTipText     =   "Sair da Tela"
      Top             =   4200
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
      MICON           =   "frmBairro.frx":0A47
      PICN            =   "frmBairro.frx":0A63
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frQuadra 
      Height          =   2175
      Left            =   4230
      TabIndex        =   17
      Top             =   945
      Visible         =   0   'False
      Width           =   1860
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   225
         TabIndex        =   19
         Top             =   1215
         Width           =   1410
      End
      Begin VB.ComboBox cmbQuadras 
         Height          =   315
         Left            =   225
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   450
         Width           =   1455
      End
      Begin prjChameleon.chameleonButton cmdGravarQuadra 
         Default         =   -1  'True
         Height          =   315
         Left            =   405
         TabIndex        =   20
         ToolTipText     =   "Alterar quadra"
         Top             =   1665
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
         MICON           =   "frmBairro.frx":0AD1
         PICN            =   "frmBairro.frx":0AED
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblquadra 
         Caption         =   "Alterar para:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   22
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lblquadra 
         Caption         =   "Quadras"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   21
         Top             =   225
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "UF:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   645
   End
End
Attribute VB_Name = "frmBairro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset
Dim Sql As String, bExec As Boolean
Dim Evento As String, sFormCall As String
Dim sRet As String, sUF As String, nCodCidade As Integer
Dim evEdit As Integer, evNew As Integer, evDel As Integer, evEsp As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean, bEsp As Boolean

Public Property Let SiglaUF(sValue As String)
    sUF = sValue
End Property

Public Property Let CodCidade(nValue As Integer)
    nCodCidade = nValue
End Property

Public Property Let FormCall(sValue As String)
    sFormCall = sValue
End Property

Private Sub cmbCidade_Click()

If Not bExec Then Exit Sub
Limpa
Ocupado
CarregaLista
le
'If cmbCidade.Text = "JABOTICABAL" Then
If cmbCidade.Text = "JABOTICABAL" And Not bEsp Then
    cmdNovo.Enabled = False: cmdAlterar.Enabled = False: cmdExcluir.Enabled = False
Else
    cmdNovo.Enabled = True: cmdAlterar.Enabled = True: cmdExcluir.Enabled = True
End If

Liberado

End Sub

Private Sub cmbUF_Click()

If Not bExec Then Exit Sub
cmbCidade.Clear
Sql = "Select CODCIDADE,DESCCIDADE From CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' ORDER BY DESCCIDADE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbCidade.AddItem !descCidade
       cmbCidade.ItemData(cmbCidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With
If cmbCidade.ListCount > 0 Then
   cmbCidade.ListIndex = 0
Else
   Limpa
   lstBairro.Clear
End If

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
    
    If MsgBox("Excluir este Bairro ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "DELETE FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex) & " AND CODBAIRRO=" & txtCod.Text
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & txtDesc.Text & " na UF " & Left$(cmbUF.Text, 2) & " e cidade " & cmbCidade.Text
       Limpa
       CarregaLista
       le
    End If
    
Exit Sub
Erro:
For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next
    
End Sub

Private Sub cmdGravar_Click()
    If txtDesc.Text = "" Then
       MsgBox "Favor digitar a Descrição.", vbExclamation, "Atenção"
       txtDesc.SetFocus
       Exit Sub
    End If
    txtDesc.Text = UCase(txtDesc.Text)
    Grava
    Evento = ""
    Eventos "INICIAR"
End Sub

Private Sub cmdGravarQuadra_Click()
Dim Sql As String, RdoAux As rdoResultset

If cmbQuadras.ListIndex = -1 Then
    MsgBox "Selecione uma quadra", vbCritical, "erro"
    Exit Sub
End If

If MsgBox("Deseja alterar todas as quadras do bairro\loteamento (" & txtDesc.Text & ") de " & cmbQuadras.Text & " para " & txtQuadra.Text & "?") = vbNo Then Exit Sub

Sql = "UPDATE CADIMOB SET LI_QUADRAS='" & txtQuadra.Text & "' WHERE LI_CODBAIRRO=" & txtCod.Text & " AND LI_QUADRAS='" & cmbQuadras.Text & "'"
cn.Execute Sql, rdExecDirect

txtQuadra.Text = ""
le
End Sub

Private Sub cmdNovo_Click()

If cmbCidade.ListIndex = -1 Then
   MsgBox "Selecione uma Cidade.", vbCritical, "Atenção"
   cmbCidade.SetFocus
Else
   Limpa
   Eventos "INCLUIR"
   Evento = "Novo"
End If

End Sub

Private Sub cmdRefresh_Click()

cmbCidade.Clear
Sql = "Select CODCIDADE,DESCCIDADE From CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' ORDER BY DESCCIDADE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbCidade.AddItem !descCidade
       cmbCidade.ItemData(cmbCidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With
If cmbCidade.ListCount > 0 Then
   cmbCidade.ListIndex = 0
Else
   Limpa
   lstBairro.Clear
End If

End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
Dim x As Integer

If sUF <> "" Then
    For x = 0 To cmbUF.ListCount
        If cmbUF.List(x) = sUF Then
            cmbUF.ListIndex = x
            Exit For
        End If
    Next
    
    For x = 0 To cmbCidade.ListCount
        If cmbCidade.ItemData(x) = nCodCidade Then
            cmbCidade.ListIndex = x
            Exit For
        End If
    Next
    Evento = "Novo"
    Eventos "INCLUIR"
End If

Liberado
End Sub

Private Sub Form_Load()
 
If NomeDeLogin <> "FACTORE" And NomeDeLogin <> "HELOISA" And NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "MARIELA" And NomeDeLogin <> "REGINA" And NomeDeLogin <> "TICYANNE.OKIMASU" Then
    frQuadra.Enabled = False
Else
cmdNovo.Enabled = True
    cmdAlterar.Enabled = True
End If
 
Ocupado
Centraliza Me
sRet = RetEventUserForm(Me.Name)

Eventos "INICIAR"

bExec = False
Sql = "Select SIGLAUF,DESCUF From UF ORDER BY DESCUF"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbUF.AddItem !SiglaUF & "-" & !DESCUF
      .MoveNext
    Loop
   .Close
End With
bExec = True
If cmbUF.ListCount > 0 Then
   cmbUF.ListIndex = 24
   cmbCidade.ListIndex = 413
End If

lstBairro.Clear
CarregaLista
le
Liberado

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
   For Each Ct In frmBairro
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          'Ct.Locked = True
         Ct.Enabled = False
       End If
   Next
   lstBairro.Enabled = True
   cmbUF.Enabled = True
   cmbUF.BackColor = vbWhite
   cmbCidade.Enabled = True
   cmbCidade.BackColor = vbWhite
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmBairro
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = vbWhite
          'Ct.Locked = False
         Ct.Enabled = True
       End If
   Next
   txtCod.BackColor = Kde
   txtCod.Locked = True
   lstBairro.Enabled = False
   cmbUF.Enabled = False
   cmbUF.BackColor = Kde
   cmbCidade.Enabled = False
   cmbCidade.BackColor = Kde
End If

txtQuadra.Enabled = True
txtQuadra.BackColor = Branco
FormHagana

End Sub

Private Sub le()
Dim Sql As String, RdoAux As rdoResultset

If lstBairro.ListIndex = -1 Then Exit Sub
txtCod.Text = lstBairro.ItemData(lstBairro.ListIndex)
txtDesc.Text = lstBairro.Text


cmbQuadras.Clear
Sql = "SELECT DISTINCT LI_QUADRAS FROM CADIMOB WHERE li_codbairro=" & txtCod.Text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbQuadras.AddItem SubNull(!Li_Quadras)
       .MoveNext
    Loop
   .Close
End With

If cmbQuadras.ListCount > 0 Then cmbQuadras.ListIndex = 0
End Sub

Private Sub Limpa()
txtCod.Text = ""
txtDesc.Text = ""
End Sub

Private Sub CarregaLista()
lstBairro.Clear

If cmbCidade.ListIndex = -1 Then Exit Sub

Sql = "Select CODBAIRRO,DESCBAIRRO From BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex) & " AND CODBAIRRO<>999 ORDER BY DESCBAIRRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       lstBairro.AddItem !DescBairro
       lstBairro.ItemData(lstBairro.NewIndex) = !CodBairro
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub lstbairro_Click()
le
End Sub

Private Sub Grava()
Dim MaxCod As Integer, x As Integer
Dim OldCidade As String, OldBairro As String

OldCidade = cmbCidade.Text
OldBairro = txtDesc.Text

If cmbCidade.ItemData(cmbCidade.ListIndex) = 413 And NomeDeLogin <> "FACTORE" And NomeDeLogin <> "HELOISA" And NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Não é permitido inserir/alterar bairros em Jaboticabal.", vbCritical, "Atenção"
    Exit Sub
End If

Sql = "SELECT MAX(CODBAIRRO) AS MAXIMO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND  CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex) & "  AND CODBAIRRO<>999"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

If Evento = "Novo" Then
    Sql = "INSERT BAIRRO (SIGLAUF,CODCIDADE,CODBAIRRO,DESCBAIRRO) VALUES('"
    Sql = Sql & Left$(cmbUF.Text, 2) & "'," & cmbCidade.ItemData(cmbCidade.ListIndex) & ","
    Sql = Sql & MaxCod & ",'" & Mask(txtDesc.Text) & "')"
Else
    Sql = "UPDATE BAIRRO SET DESCBAIRRO='" & Mask(txtDesc.Text) & "' WHERE "
    Sql = Sql & "SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex)
    Sql = Sql & " AND CODBAIRRO=" & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect


If Evento = "Novo" Then
   txtCod.Text = MaxCod
   Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & txtDesc.Text & " na UF " & Left$(cmbUF.Text, 2) & " e cidade " & cmbCidade.Text
 ElseIf Evento = "Alterar" Then
   Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.Text, "000") & " de " & sOldDesc & " para " & txtDesc.Text & " na UF " & Left$(cmbUF.Text, 2) & " e cidade " & cmbCidade.Text
End If

cmbCidade.Text = OldCidade
cmbCidade_Click
lstBairro.Text = OldBairro
le

If sUF <> "" Then
    If sFormCall = "frmCadImob" Then
        frmCadImob.cmbCidade_Click
        For x = 0 To frmCadImob.cmbBairro.ListCount
            If UCase(frmCadImob.cmbBairro.List(x)) = UCase(txtDesc.Text) Then
                frmCadImob.cmbBairro.ListIndex = x
                Exit For
            End If
        Next
    ElseIf sFormCall = "frmCidadaoR" Then
        frmCidadao.cmbCidade_Click
        For x = 0 To frmCidadao.cmbBairro.ListCount
            If UCase(frmCidadao.cmbBairro.List(x)) = UCase(txtDesc.Text) Then
                frmCidadao.cmbBairro.ListIndex = x
                Exit For
            End If
        Next
    ElseIf sFormCall = "frmCidadaoC" Then
        frmCidadao.cmbCidade2_Click
        For x = 0 To frmCidadao.cmbBairro2.ListCount
            If UCase(frmCidadao.cmbBairro2.List(x)) = UCase(txtDesc.Text) Then
                frmCidadao.cmbBairro2.ListIndex = x
                Exit For
            End If
        Next
    End If
End If

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
evEsp = 11
bEsp = False

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True
If InStr(1, sRet, Format(evEsp, "000"), vbBinaryCompare) > 0 Then bEsp = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
' 'Tweak txtDesc, KeyAscii, AllLettersAllCaps
End Sub
