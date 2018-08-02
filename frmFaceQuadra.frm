VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmFaceQuadra 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Faces de Quadra"
   ClientHeight    =   4845
   ClientLeft      =   6510
   ClientTop       =   3945
   ClientWidth     =   6105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5010
      TabIndex        =   26
      ToolTipText     =   "Sair da Tela"
      Top             =   4455
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmFaceQuadra.frx":0000
      PICN            =   "frmFaceQuadra.frx":001C
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
      TabIndex        =   22
      ToolTipText     =   "Novo Registro"
      Top             =   4410
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
      MICON           =   "frmFaceQuadra.frx":008A
      PICN            =   "frmFaceQuadra.frx":00A6
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
      TabIndex        =   23
      ToolTipText     =   "Editar Registro"
      Top             =   4440
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
      MICON           =   "frmFaceQuadra.frx":0200
      PICN            =   "frmFaceQuadra.frx":021C
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
      TabIndex        =   24
      ToolTipText     =   "Excluir Registro"
      Top             =   4440
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
      MICON           =   "frmFaceQuadra.frx":0376
      PICN            =   "frmFaceQuadra.frx":0392
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
      Left            =   5010
      TabIndex        =   25
      ToolTipText     =   "Gravar os Dados"
      Top             =   4440
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
      MICON           =   "frmFaceQuadra.frx":0434
      PICN            =   "frmFaceQuadra.frx":0450
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1605
      Left            =   60
      TabIndex        =   8
      Top             =   2730
      Width           =   6015
      Begin VB.ListBox lstNomeLog 
         BackColor       =   &H00C0FFFF&
         Height          =   1425
         ItemData        =   "frmFaceQuadra.frx":07F5
         Left            =   1410
         List            =   "frmFaceQuadra.frx":07F7
         TabIndex        =   20
         Top             =   90
         Visible         =   0   'False
         Width           =   4260
      End
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   14
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   13
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox txtFace 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3930
         MaxLength       =   50
         TabIndex        =   12
         Top             =   210
         Width           =   1005
      End
      Begin VB.ComboBox cmbAgrupa 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmFaceQuadra.frx":07F9
         Left            =   1410
         List            =   "frmFaceQuadra.frx":07FB
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   765
      End
      Begin VB.TextBox txtNomeLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   9
         Top             =   870
         Width           =   4260
      End
      Begin VB.TextBox txtQuadra2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1230
         Width           =   2070
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra Planta..:"
         Height          =   195
         Index           =   3
         Left            =   2250
         TabIndex        =   21
         Top             =   1290
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Logradouro:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra...............:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Face..................:"
         Height          =   195
         Index           =   11
         Left            =   2610
         TabIndex        =   17
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agrupamento.....:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Logradour:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   15
         Top             =   930
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   6120
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   307
      TabIndex        =   7
      Top             =   90
      Width           =   4665
   End
   Begin VB.ComboBox cmbQuadra 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmFaceQuadra.frx":07FD
      Left            =   4290
      List            =   "frmFaceQuadra.frx":07FF
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1005
   End
   Begin VB.ComboBox cmbSetor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2580
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   885
   End
   Begin VB.ComboBox cmbDist 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid grdFace 
      Height          =   2220
      Left            =   30
      TabIndex        =   3
      Top             =   510
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   3916
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   15658734
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Quadra  |^Face  |<Logradouro                                                         |^Agrup. "
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3915
      TabIndex        =   27
      ToolTipText     =   "Cancelar Edição"
      Top             =   4440
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmFaceQuadra.frx":0801
      PICN            =   "frmFaceQuadra.frx":081D
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
      Caption         =   "Quadra:"
      Height          =   225
      Index           =   2
      Left            =   3540
      TabIndex        =   6
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Setor:"
      Height          =   225
      Index           =   1
      Left            =   1860
      TabIndex        =   5
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Distrito:"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   90
      Width           =   645
   End
End
Attribute VB_Name = "frmFaceQuadra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String, bExec As Boolean
Dim aRua(8) As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean, nFace As Integer

Private Sub cmbDist_Click()
If cmbDist.ListIndex = -1 Or Not bExec Then Exit Sub
cmbSetor.Clear
Sql = "SELECT CODSETOR FROM SETOR WHERE CODDISTRITO=" & Val(cmbDist.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
   bExec = False
   Do Until .EOF
      cmbSetor.AddItem !CODSETOR
     .MoveNext
   Loop
  .Close
   bExec = True
   cmbSetor.ListIndex = 0
End With

End Sub

Private Sub cmbQuadra_Click()

Ocupado

If cmbQuadra.ListIndex = 0 Then
    cmdNovo_Click
    Liberado
    Exit Sub
Else
    CarregaLista
    If grdFace.Rows > 1 Then
       For x = 1 To grdFace.Rows - 1
           If x > 8 Then Exit For
           aRua(x) = grdFace.TextMatrix(x, 2)
       Next
       If grdFace.Rows - 1 <= 8 Then
          MontaFace grdFace.Rows - 1, cmbQuadra.text
       Else
          If cmbQuadra.ListIndex > 0 Then
             MontaFace 8, Val(cmbQuadra.text)
          Else
             Picture1.Cls
          End If
       End If
    Else
       Picture1.Cls
    End If
End If
Liberado

End Sub

Private Sub cmbSetor_Click()
If cmbSetor.ListIndex = -1 Or Not bExec Then Exit Sub
cmbQuadra.Clear
Screen.MousePointer = vbHourglass

Sql = "SELECT DISTINCT CODQUADRA FROM FACEQUADRA WHERE "
Sql = Sql & "CODDISTRITO=" & Val(cmbDist.text) & " AND "
Sql = Sql & "CODSETOR=" & Val(cmbSetor.text) & " ORDER BY CODQUADRA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
   bExec = False
   cmbQuadra.AddItem "(Nova)"
   Do Until .EOF
      cmbQuadra.AddItem !CODQUADRA
     .MoveNext
   Loop
  .Close
   bExec = True
   cmbQuadra.ListIndex = 1
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAlterar_Click()
    If grdFace.Rows = 1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo Erro
Dim OldD As Integer, OldS As Integer, OldQ As Integer

    If grdFace.Rows = 1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir esta Face de Quadra ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       OldD = cmbDist.text
       OldS = cmbSetor.text
       OldQ = txtQuadra.text
       Sql = "DELETE FROM FACEQUADRA WHERE "
       Sql = Sql & "CODDISTRITO=" & Val(cmbDist.text) & " AND "
       Sql = Sql & "CODSETOR=" & Val(cmbSetor.text) & " AND "
       Sql = Sql & "CODQUADRA=" & Val(cmbQuadra.text) & " AND "
       Sql = Sql & "CODFACE=" & Val(grdFace.TextMatrix(grdFace.Row, 1))
       cn.Execute Sql, rdExecDirect
       'revisar
       Log Form, Me.Caption, Exclusão, "Excluído registro "
       cmbDist.text = OldD
       cmbSetor.text = OldS
       cmbQuadra.text = OldQ
    End If
    
Exit Sub
Erro:
If rdoErrors(1).Number = 547 Then
   MsgBox "Não é possível excluir esta Face de Quadra, pois ela esta ligada a um Imóvel.", vbExclamation, "Atenção"
Else
   Resume Next
End If


End Sub

Private Sub cmdGravar_Click()
    If Val(txtQuadra.text) = 0 Then
       MsgBox "Digite o Nº da Quadra.", vbExclamation, "Atenção"
       txtQuadra.SetFocus
       Exit Sub
    End If
    If Val(txtFace.text) = 0 Then
       MsgBox "Digite o Nº da Face.", vbExclamation, "Atenção"
       txtFace.SetFocus
       Exit Sub
    End If
    If Val(txtCodLogr.text) = 0 Then
       MsgBox "Logradouro inválido.", vbExclamation, "Atenção"
       txtCodLogr.SetFocus
       Exit Sub
    End If
    If cmbAgrupa.ListIndex = -1 Then
       MsgBox "Selecione um Agrupamento.", vbExclamation, "Atenção"
       cmbAgrupa.SetFocus
       Exit Sub
    End If
    If Evento = "Novo" Then
       Sql = "SELECT CODAGRUPA FROM FACEQUADRA WHERE "
       Sql = Sql & "CODDISTRITO=" & Val(cmbDist.text) & " AND "
       Sql = Sql & "CODSETOR=" & Val(cmbSetor.text) & " AND "
       Sql = Sql & "CODQUADRA=" & Val(cmbQuadra.text) & " AND "
       Sql = Sql & "CODFACE=" & Val(txtFace.text)
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
       If RdoAux.RowCount > 0 Then
          MsgBox "Nº de Face existente para esta quadra.", vbExclamation, "Atenção"
          txtFace.SetFocus
          RdoAux.Close
          Exit Sub
       End If
    End If
    Grava
    Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()

    Limpa
    Evento = "Novo"
    Eventos "INCLUIR"
    If cmbQuadra.ListIndex = 0 Then
        grdFace.Rows = 1
        Sql = "SELECT MAX(CODQUADRA) AS MAXIMO FROM FACEQUADRA WHERE  "
        Sql = Sql & "CODDISTRITO=" & cmbDist.text & " AND CODSETOR=" & cmbSetor.text & " AND CODQUADRA < 1910"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!MAXIMO) Then
                txtQuadra.text = "1"
            Else
                txtQuadra.text = !MAXIMO + 1
            End If
            txtFace.text = 1
           .Close
        End With
    Else
        Sql = "SELECT MAX(CODFACE) AS MAXIMO FROM FACEQUADRA WHERE  "
        Sql = Sql & "CODDISTRITO=" & cmbDist.text & " AND CODSETOR=" & cmbSetor.text & " AND CODQUADRA=" & cmbQuadra.text
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!MAXIMO) Then
                txtFace.text = "1"
            Else
                txtFace.text = !MAXIMO + 1
            End If
           .Close
        End With
    End If
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
bExec = False
LoadCombo
Le
Eventos "INICIAR"

End Sub

Private Sub LoadCombo()

Sql = "SELECT CODDISTRITO FROM DISTRITO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
   Do Until .EOF
      cmbDist.AddItem !CODDISTRITO
     .MoveNext
   Loop
  .Close
   bExec = True
   cmbDist.ListIndex = 0
End With

For x = 1 To 8
    cmbAgrupa.AddItem x
Next

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
   For Each Ct In frmFaceQuadra
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   cmbDist.BackColor = Branco
   cmbDist.Enabled = True
   cmbSetor.BackColor = Branco
   cmbSetor.Enabled = True
   cmbQuadra.BackColor = Branco
   cmbQuadra.Enabled = True
   grdFace.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmFaceQuadra
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   If Evento = "Novo" Then
      If cmbQuadra.ListIndex = 0 Then
'         txtQuadra.BackColor = Kde
'         txtQuadra.Enabled = False
         'txtFace.SetFocus
         'txtFace.BackColor = Kde
         'txtFace.Enabled = False
      Else
         txtQuadra.BackColor = Kde
         txtQuadra.Enabled = False
         cmbQuadra.BackColor = Kde
         cmbQuadra.Enabled = False
         txtQuadra.text = cmbQuadra.text
         'txtFace.BackColor = Kde
         'txtFace.Enabled = False
         txtCodLogr.SetFocus
      End If
   Else
      txtQuadra.text = cmbQuadra.text
      txtQuadra.BackColor = Kde
      txtQuadra.Enabled = False
      cmbQuadra.BackColor = Kde
      cmbQuadra.Enabled = False
      'txtFace.BackColor = Kde
      'txtFace.Enabled = False
      txtCodLogr.SetFocus
   End If
   grdFace.Enabled = False
   cmbDist.BackColor = Kde
   cmbDist.Enabled = False
   cmbSetor.BackColor = Kde
   cmbSetor.Enabled = False
   cmbQuadra.BackColor = Kde
   cmbQuadra.Enabled = False
End If

FormHagana

End Sub

Private Sub Le()
If grdFace.Row = 0 Then Exit Sub
Sql = "SELECT CODDISTRITO,CODSETOR,CODQUADRA,CODFACE,"
Sql = Sql & "CODLOGR,CODAGRUPA,CODTIPOLOG,NOMETIPOLOG,"
Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,ABREVTITLOG,"
Sql = Sql & "NOMELOGRADOURO,QUADRAS FROM vwFACEQUADRA WHERE "
Sql = Sql & "CODDISTRITO=" & Val(cmbDist.text) & " AND "
Sql = Sql & "CODSETOR=" & Val(cmbSetor.text) & " AND "
Sql = Sql & "CODQUADRA=" & Val(cmbQuadra.text) & " AND "
Sql = Sql & "CODFACE=" & Val(grdFace.TextMatrix(grdFace.Row, 1))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtQuadra.text = !CODQUADRA
        nFace = !CODFACE
        txtFace.text = !CODFACE
        txtCodLogr.text = Format(!CodLogr, "0000")
        txtNomeLogr.text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
        If cmbAgrupa.ListCount > 0 And Not IsNull(!CODAGRUPA) Then
            cmbAgrupa.text = !CODAGRUPA
        End If
        txtQuadra2.text = SubNull(!Quadras)
    Else
        Limpa
    End If
End With

End Sub

Private Sub Limpa()

txtQuadra.text = ""
txtFace.text = ""
txtCodLogr.text = ""
txtNomeLogr.text = ""
cmbAgrupa.ListIndex = -1
txtQuadra2.text = ""

End Sub

Private Sub CarregaLista()
Dim sNomeLog As String
If cmbQuadra.ListIndex = -1 Then cmbQuadra.ListIndex = 0
Limpa
Screen.MousePointer = vbHourglass
Sql = "SELECT CODDISTRITO,CODSETOR,CODQUADRA,CODFACE,"
Sql = Sql & "CODLOGR,CODAGRUPA,CODTIPOLOG,NOMETIPOLOG,"
Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,ABREVTITLOG,"
Sql = Sql & "NOMELOGRADOURO,QUADRAS FROM vwFACEQUADRA WHERE "
Sql = Sql & "CODDISTRITO=" & Val(cmbDist.text) & " AND "
Sql = Sql & "CODSETOR=" & Val(cmbSetor.text)
If cmbQuadra.ListIndex > 0 Then
   Sql = Sql & " AND CODQUADRA=" & Val(cmbQuadra.text)
End If
Sql = Sql & " ORDER BY CODDISTRITO,CODSETOR,CODQUADRA,CODFACE,CODLOGR"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
grdFace.Rows = 1
With RdoAux
    Do Until .EOF
       sNomeLog = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       grdFace.AddItem Format(!CODQUADRA, "0000") & Chr(9) & !CODFACE & Chr(9) & sNomeLog & Chr(9) & !CODAGRUPA
      .MoveNext
    Loop
   .Close
    Le
End With
Screen.MousePointer = vbDefault

End Sub

Private Sub Grava()

If Evento = "Novo" Then
    Sql = "INSERT FACEQUADRA(CODDISTRITO,CODSETOR,CODQUADRA,CODFACE,CODLOGR,CODAGRUPA,QUADRAS) values("
    Sql = Sql & Val(cmbDist.text) & "," & Val(cmbSetor.text) & "," & Val(txtQuadra.text) & ","
    Sql = Sql & Val(txtFace.text) & "," & Val(txtCodLogr.text) & "," & Val(cmbAgrupa.text) & ",'" & Mask(txtQuadra2.text) & "')"
Else
    Sql = "UPDATE FACEQUADRA SET CODLOGR=" & Val(txtCodLogr.text) & ",CODAGRUPA=" & Val(cmbAgrupa.text) & ",QUADRAS='" & Mask(txtQuadra2.text) & "',CODFACE=" & Val(txtFace.text)
    Sql = Sql & " WHERE CODDISTRITO = " & Val(cmbDist.text) & " AND CODSETOR=" & Val(cmbSetor.text) & " AND "
    Sql = Sql & "CODQUADRA=" & Val(txtQuadra.text) & " AND CODFACE=" & nFace
End If
cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   grdFace.AddItem txtQuadra.text & Chr(9) & txtFace.text & Chr(9) & txtNomeLogr.text & Chr(9) & cmbAgrupa.text
   grdFace.Row = grdFace.Rows - 1
   grdFace.ColSel = 3
   cmbQuadra.AddItem Val(txtQuadra.text)
   cmbQuadra.text = Val(txtQuadra.text)
   cmbQuadra_Click
   'REVISAR
   Log Form, Me.Caption, Inclusão, "Inserido registro " & txtNomeLogr.text
 ElseIf Evento = "Alterar" Then
   grdFace.TextMatrix(grdFace.Row, 1) = txtFace.text
   grdFace.TextMatrix(grdFace.Row, 2) = txtNomeLogr.text
   grdFace.TextMatrix(grdFace.Row, 3) = cmbAgrupa.text
   'REVISAR
   Log Form, Me.Caption, Alteração, "Alterado registro " & txtNomeLogr.text
End If
cmbQuadra_Click
      
End Sub


Private Sub MontaFace(Lados As Integer, NumeroQuadra As Integer)
Dim Larg As Integer, ax() As Integer, aY() As Integer
Exit Sub
gtiObj.LikroTmuna ArqBinImg, "MARM", ArqBinImgTmp
Picture1.Picture = LoadPicture(ArqBinImgTmp)
Picture1.ForeColor = vbBlack

Select Case Lados
    Case 1
        ReDim ax(4)
        ReDim aY(4)
        Larg = 150
        MoveTo Picture1.ScaleWidth / 2 + 55, Picture1.ScaleHeight / 2 + 70
        DrawLine Larg, 90
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        DrawLine Larg, 0
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, -90
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg, 180
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1) - 15, aY(3), 900, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 30, aY(3), 900, vbBlue)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 110, 130, 0, Roxo)
    Case 2
        ReDim ax(4)
        ReDim aY(4)
        Larg = 150
        MoveTo Picture1.ScaleWidth / 2 + 55, Picture1.ScaleHeight / 2 + 70
        DrawLine Larg, 90
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        DrawLine Larg, 0
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, -90
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg, 180
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1) - 15, aY(3), 900, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 30, aY(3), 900, vbBlue)
        Call TextoFace("Face 2", Picture1, ax(4), aY(2) - 15, 0, Marrom&)
        Call TextoFace(aRua(2), Picture1, ax(4), aY(2) - 30, 0, vbRed)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 110, 130, 0, Roxo)
    Case 3
        ReDim ax(4)
        ReDim aY(4)
        Larg = 150
        MoveTo Picture1.ScaleWidth / 2 + 55, Picture1.ScaleHeight / 2 + 70
        DrawLine Larg, 90
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        DrawLine Larg, 0
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, -90
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg, 180
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1) - 15, aY(3), 900, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 30, aY(3), 900, vbBlue)
        Call TextoFace("Face 2", Picture1, ax(4), aY(2) - 15, 0, Marrom&)
        Call TextoFace(aRua(2), Picture1, ax(4), aY(2) - 30, 0, vbRed)
        Call TextoFace("Face 3", Picture1, ax(3) + 3, aY(3), 900, Marrom&)
        Call TextoFace(aRua(3), Picture1, ax(3) + 15, aY(3), 900, vbBlack)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 110, 130, 0, Roxo)
    Case 4
        ReDim ax(4)
        ReDim aY(4)
        Larg = 150
        MoveTo Picture1.ScaleWidth / 2 + 55, Picture1.ScaleHeight / 2 + 70
        DrawLine Larg, 90
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        DrawLine Larg, 0
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, -90
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg, 180
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1) - 15, aY(3), 900, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 30, aY(3), 900, vbBlue)
        Call TextoFace("Face 2", Picture1, ax(4), aY(2) - 15, 0, Marrom&)
        Call TextoFace(aRua(2), Picture1, ax(4), aY(2) - 30, 0, vbRed)
        Call TextoFace("Face 3", Picture1, ax(3) + 3, aY(3), 900, Marrom&)
        Call TextoFace(aRua(3), Picture1, ax(3) + 15, aY(3), 900, vbBlack)
        Call TextoFace("Face 4", Picture1, ax(4), aY(4), 0, Marrom&)
        Call TextoFace(aRua(4), Picture1, ax(4), aY(4) + 15, 0, VerdeEscuro)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 110, 130, 0, Roxo)
    Case 5
        ReDim ax(5)
        ReDim aY(5)
        Larg = 110
        MoveTo Picture1.ScaleWidth / 2 + 55, Picture1.ScaleHeight / 2 + 90
        DrawLine Larg + 44, 90
        ax(5) = Picture1.CurrentX
        aY(5) = Picture1.CurrentY
        DrawLine Larg, 0
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, -45
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg - 1, -135
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        DrawLine Larg, 180
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1) - 15, aY(5), 900, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 30, aY(5), 900, vbBlue)
        Call TextoFace("Face 2", Picture1, ax(1) - 10, aY(1) - 15, 450, Marrom&)
        Call TextoFace(aRua(2), Picture1, ax(1) - 20, aY(1) - 25, 450, vbRed)
        Call TextoFace("Face 3", Picture1, ax(2) + 15, aY(2) - 15, -450, Marrom&)
        Call TextoFace(aRua(3), Picture1, ax(2) + 25, aY(2) - 25, -450, vbBlack)
        Call TextoFace("Face 4", Picture1, ax(4), aY(4), 900, Marrom&)
        Call TextoFace(aRua(4), Picture1, ax(4) + 15, aY(4), 900, vbBlue)
        Call TextoFace("Face 5", Picture1, ax(5), aY(5), 0, Marrom&)
        Call TextoFace(aRua(5), Picture1, ax(5), aY(5) + 15, 0, VerdeEscuro)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 110, 150, 0, Roxo)
    Case 6
        ReDim ax(6)
        ReDim aY(6)
        Larg = 100
        MoveTo Picture1.ScaleWidth / 2 + 40, Picture1.ScaleHeight / 2 + 70
        DrawLine Larg + 10, 90
        ax(6) = Picture1.CurrentX
        aY(6) = Picture1.CurrentY
        DrawLine Larg, 30
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, -30
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg + 10, -90
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        DrawLine Larg, -150
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        DrawLine Larg, 150
        ax(5) = Picture1.CurrentX
        aY(5) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1), aY(1), -600, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 10, aY(1) + 5, -600, vbBlue)
        Call TextoFace("Face 2", Picture1, ax(1) - 12, aY(1) - 9, 600, Marrom&)
        Call TextoFace(aRua(2), Picture1, ax(1) - 25, aY(1) - 15, 600, vbRed)
        Call TextoFace("Face 3", Picture1, ax(2), aY(2) - 15, 0, Marrom&)
        Call TextoFace(aRua(3), Picture1, ax(2), aY(2) - 30, 0, vbBlack)
        Call TextoFace("Face 4", Picture1, ax(3) + 15, aY(3) - 5, -600, Marrom&)
        Call TextoFace(aRua(4), Picture1, ax(3) + 25, aY(3) - 12, -600, vbBlue)
        Call TextoFace("Face 5", Picture1, ax(5), aY(5), 600, Marrom&)
        Call TextoFace(aRua(5), Picture1, ax(5) + 12, aY(5) + 5, 600, VerdeEscuro)
        Call TextoFace("Face 6", Picture1, ax(6), aY(6), 0, Marrom&)
        Call TextoFace(aRua(6), Picture1, ax(6), aY(6) + 15, 0, vbRed)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 110, 130, 0, Roxo)
    Case 7
        ReDim ax(7)
        ReDim aY(7)
        Larg = 90
        MoveTo Picture1.ScaleWidth / 2 + 40, Picture1.ScaleHeight / 2 + 110
        DrawLine Larg + 20, 90
        ax(7) = Picture1.CurrentX
        aY(7) = Picture1.CurrentY
        DrawLine Larg, 20
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, 0
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg + 10, -60
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        DrawLine Larg + 10, -120
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        DrawLine Larg, 180
        ax(5) = Picture1.CurrentX
        aY(5) = Picture1.CurrentY
        DrawLine Larg + 1, 160
        ax(6) = Picture1.CurrentX
        aY(6) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1), aY(1), -700, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 13, aY(1) + 2, -700, vbBlue)
        Call TextoFace("Face 2", Picture1, ax(1) - 15, aY(1) - 7, 900, Marrom&)
        Call TextoFace(aRua(2), Picture1, ax(1) - 28, aY(1) - 7, 900, vbRed)
        Call TextoFace("Face 3", Picture1, ax(2) - 8, aY(2) - 13, 300, Marrom&)
        Call TextoFace(aRua(3), Picture1, ax(2) - 15, aY(2) - 25, 300, vbBlack)
        Call TextoFace("Face 4", Picture1, ax(3) + 10, aY(3) - 15, -300, Marrom&)
        Call TextoFace(aRua(4), Picture1, ax(3) + 15, aY(3) - 27, -300, vbBlue)
        Call TextoFace("Face 5", Picture1, ax(4) + 17, aY(4), -900, Marrom&)
        Call TextoFace(aRua(5), Picture1, ax(4) + 30, aY(4), -900, VerdeEscuro)
        Call TextoFace("Face 6", Picture1, ax(6) + 2, aY(6), 700, Marrom&)
        Call TextoFace(aRua(6), Picture1, ax(6) + 14, aY(6) + 3, 700, vbRed)
        Call TextoFace("Face 7", Picture1, ax(7), aY(7) + 2, 0, Marrom&)
        Call TextoFace(aRua(7), Picture1, ax(7), aY(7) + 15, 0, vbBlack)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 110, 150, 0, Roxo)
    Case 8
        ReDim ax(8)
        ReDim aY(8)
        Larg = 85
        MoveTo Picture1.ScaleWidth / 2 + 45, Picture1.ScaleHeight / 2 + 115
        DrawLine Larg + 20, 90
        ax(8) = Picture1.CurrentX
        aY(8) = Picture1.CurrentY
        DrawLine Larg, 30
        ax(1) = Picture1.CurrentX
        aY(1) = Picture1.CurrentY
        DrawLine Larg, 0
        ax(2) = Picture1.CurrentX
        aY(2) = Picture1.CurrentY
        DrawLine Larg, -30
        ax(3) = Picture1.CurrentX
        aY(3) = Picture1.CurrentY
        DrawLine Larg + 20, -90
        ax(4) = Picture1.CurrentX
        aY(4) = Picture1.CurrentY
        DrawLine Larg, -150
        ax(5) = Picture1.CurrentX
        aY(5) = Picture1.CurrentY
        DrawLine Larg, 180
        ax(6) = Picture1.CurrentX
        aY(6) = Picture1.CurrentY
        DrawLine Larg, 150
        ax(7) = Picture1.CurrentX
        aY(7) = Picture1.CurrentY
        Call TextoFace("Face 1", Picture1, ax(1) - 2, aY(1), -600, Marrom&)
        Call TextoFace(aRua(1), Picture1, ax(1) - 15, aY(1) + 5, -600, vbBlue)
        Call TextoFace("Face 2", Picture1, ax(1) - 17, aY(1), 900, Marrom&)
        Call TextoFace(aRua(2), Picture1, ax(1) - 30, aY(1), 900, vbRed)
        Call TextoFace("Face 3", Picture1, ax(2) - 12, aY(2) - 12, 600, Marrom&)
        Call TextoFace(aRua(3), Picture1, ax(2) - 22, aY(2) - 20, 600, vbBlack)
        Call TextoFace("Face 4", Picture1, ax(3), aY(3) - 17, 0, Marrom&)
        Call TextoFace(aRua(4), Picture1, ax(3), aY(3) - 30, 0, vbBlue)
        Call TextoFace("Face 5", Picture1, ax(4) + 17, aY(4) - 7, -600, Marrom&)
        Call TextoFace(aRua(5), Picture1, ax(4) + 28, aY(4) - 15, -600, VerdeEscuro)
        Call TextoFace("Face 6", Picture1, ax(6) + 2, aY(6), 900, Marrom&)
        Call TextoFace(aRua(6), Picture1, ax(6) + 15, aY(6), 900, vbRed)
        Call TextoFace("Face 7", Picture1, ax(7) + 2, aY(7), 600, Marrom&)
        Call TextoFace(aRua(7), Picture1, ax(7) + 15, aY(7) + 5, 600, vbBlack)
        Call TextoFace("Face 8", Picture1, ax(8), aY(8) + 2, 0, Marrom&)
        Call TextoFace(aRua(8), Picture1, ax(8), aY(8) + 15, 0, VerdeEscuro)
        Call TextoFace("Quadra " & Format(NumeroQuadra, "000"), Picture1, 120, 150, 0, Roxo)
End Select

End Sub

Private Sub DrawLine(ByVal Length As Single, ByVal Angle As Single)

pi = 4 * Atn(1)
cx = Picture1.CurrentX
cy = Picture1.CurrentY

'Angle is in Degrees
Angle = Angle Mod 360
Angle = Angle * pi / 180
'Xp = 0
yp = Abs(Length)
Rx = Xp * cOS(Angle) - yp * Sin(Angle)
Ry = Xp * Sin(Angle) + yp * cOS(Angle)
rxg = cx + Rx
ryg = cy - Ry

Picture1.DrawStyle = 2
Picture1.Line (cx, cy)-(rxg, ryg)

' if negative length go back to start position
If Length < 0 Then
    Picture1.CurrentX = cx
    Picture1.CurrentY = cy
End If
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Function TextoFace(text As String, picturebox As picturebox, LadoX As Integer, LadoY As Integer, Angulo As Integer, Cor As Long)
    Dim Font As LOGFONT
    Dim prevFont As Long, hFont As Long
    
'    Const FontSize = 6 ' Desired point size of font
    Font.lfEscapement = Angulo    ' 180-degree rotation
'    Font.lfFaceName = picturebox.Font & Chr$(0)  'Null character at end

    ' Windows expects the font size to be in pixels and to
    ' be negative if you are specifying the character height
    ' you want.
    Font.lfHeight = -10
    If UCase$(Left$(text, 4)) = "QUAD" Then
       Font.lfUnderline = True
    Else
       Font.lfUnderline = False
    End If
    hFont = CreateFontIndirect(Font)
    prevFont = SelectObject(picturebox.hdc, hFont)
    picturebox.CurrentX = LadoX
    picturebox.CurrentY = LadoY
    picturebox.ForeColor = Cor
    picturebox.Print text
    hFont = SelectObject(Picture1.hdc, prevFont)
    DeleteObject hFont
    If cGetInputState() <> 0 Then DoEvents
    If cGetInputState() <> 0 Then DoEvents
End Function

Private Sub MoveTo(x As Single, Y As Single)
Picture1.CurrentX = x
Picture1.CurrentY = Y
End Sub

Private Sub grdFace_Click()
Le
End Sub

Private Sub grdFace_RowColChange()
Le
End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
   txtCodLogr.text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   txtCodLogr_LostFocus
   lstNomeLog.Visible = False
   cmbAgrupa.SetFocus
End If

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
       cmbAgrupa.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
   txtNomeLogr.SetFocus
End If

End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub

Private Sub txtCodLogr_GotFocus()
txtCodLogr.SelStart = 0
txtCodLogr.SelLength = Len(txtCodLogr)

End Sub

Private Sub txtCodLogr_KeyPress(KeyAscii As Integer)
Tweak txtCodLogr, KeyAscii, IntegerPositive
End Sub

Private Sub txtCodLogr_LostFocus()
If Val(txtCodLogr.text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLogr.text = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          txtNomeLogr.text = ""
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtCodLogr.SetFocus
       End If
      .Close
   End With
End If

End Sub

Private Sub txtFace_KeyPress(KeyAscii As Integer)
Tweak txtFace, KeyAscii, IntegerPositive
End Sub

Private Sub txtNomeLogr_Change()
If Trim$(txtNomeLogr) = "" Then
   txtCodLogr.text = 0
End If
End Sub

Private Sub txtNomeLogr_GotFocus()
txtNomeLogr.SelStart = 0
txtNomeLogr.SelLength = Len(txtNomeLogr)
End Sub

Private Sub txtNomeLogr_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   lstNomeLog.Clear
   If txtNomeLogr.text <> "" Then
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
                lstNomeLog.AddItem Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CODLOGRADOURO
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Digite o nome do logradouro a ser pesquisado, sem especificar o tipo e o título.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtNomeLogr.SetFocus
          End If
      End With
   End If
Else
   txtCodLogr.text = 0
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

