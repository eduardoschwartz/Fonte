VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTributoAliquota 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de Preços Públicos"
   ClientHeight    =   5115
   ClientLeft      =   4230
   ClientTop       =   1995
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6120
   Begin VB.ComboBox cmbAno 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTributoAliquota.frx":0000
      Left            =   780
      List            =   "frmTributoAliquota.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1005
      Left            =   15
      TabIndex        =   4
      Top             =   3630
      Width           =   6090
      Begin VB.ComboBox cmbTrib 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   4515
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   3
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Aliquota.....:"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tributo................:"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   5
         Top             =   285
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   2910
      Left            =   30
      TabIndex        =   1
      Top             =   690
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   5133
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Código |<Descrição do Tributo                                                |>Valor Aliquota      "
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4830
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   4710
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
      MICON           =   "frmTributoAliquota.frx":0004
      PICN            =   "frmTributoAliquota.frx":0020
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
      TabIndex        =   9
      ToolTipText     =   "Novo Registro"
      Top             =   4710
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
      MICON           =   "frmTributoAliquota.frx":017A
      PICN            =   "frmTributoAliquota.frx":0196
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
      Top             =   4710
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
      MICON           =   "frmTributoAliquota.frx":02F0
      PICN            =   "frmTributoAliquota.frx":030C
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
      Top             =   4710
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
      MICON           =   "frmTributoAliquota.frx":0466
      PICN            =   "frmTributoAliquota.frx":0482
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
      Height          =   345
      Left            =   4980
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   4680
      Width           =   885
      _ExtentX        =   1561
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmTributoAliquota.frx":0524
      PICN            =   "frmTributoAliquota.frx":0540
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
      Left            =   3720
      TabIndex        =   13
      ToolTipText     =   "Gravar os Dados"
      Top             =   4710
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
      MICON           =   "frmTributoAliquota.frx":05AE
      PICN            =   "frmTributoAliquota.frx":05CA
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
      Caption         =   "Ano.....:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   255
      Width           =   585
   End
End
Attribute VB_Name = "frmTributoAliquota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String

Private Sub cmbAno_Click()
Limpa
CarregaLista
Le
End Sub

Private Sub cmdAlterar_Click()
    If cmbTrib.ListIndex = -1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    Evento = "Alterar"
    Eventos "INCLUIR"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()

On Error GoTo Erro
If cmbTrib.ListIndex = -1 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

If MsgBox("Excluir este Tributo da Tabela de Preços Públicos ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
   Sql = "DELETE FROM TRIBUTOALIQUOTA WHERE ANO=" & Val(cmbAno.Text) & " AND CODTRIBUTO=" & cmbTrib.ItemData(cmbTrib.ListIndex)
   cn.Execute Sql, rdExecDirect
   Log Form, Me.Caption, Exclusão, "Excluído registro " & cmbAno.Text & " - " & cmbTrib.Text
   Limpa
   CarregaLista
   Le
End If
    
Exit Sub
Erro:

For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next

End Sub

Private Sub cmdGravar_Click()
    
If cmbTrib.ListIndex = -1 Then
   MsgBox "Selecione o Tributo.", vbExclamation, "Atenção"
   cmbLanc.SetFocus
   Exit Sub
End If

If Trim(txtValor.Text) = "" Then txtValor.Text = 0

If Evento = "Novo" Then
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & cmbAno.Text & " AND CODTRIBUTO=" & cmbTrib.ItemData(cmbTrib.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Tributo já incluído na Tabela de Preços Públicos", vbExclamation, "Atenção"
        Exit Sub
       .Close
    End If
End With
End If

Grava
Eventos "INICIAR"
Evento = ""

End Sub


Private Sub cmdNovo_Click()
    Limpa
    Evento = "Novo"
    Eventos "INCLUIR"
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
Liberado
End Sub

Private Sub Form_Load()
Dim x As Integer

For x = 2001 To Year(Now) + 1
    cmbAno.AddItem x
Next
Ocupado

Centraliza Me
sRet = RetEventUserForm(Me.Name)

grdTrib.Rows = 1

Sql = "SELECT CODTRIBUTO,DESCTRIBUTO FROM TRIBUTO ORDER BY DESCTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
Do Until RdoAux.EOF
     cmbTrib.AddItem RdoAux!desctributo
     cmbTrib.ItemData(cmbTrib.NewIndex) = RdoAux!CodTributo
     RdoAux.MoveNext
Loop

cmbAno.Text = Year(Now)
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
   For Each Ct In frmTributoAliquota
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
         Ct.Enabled = False
       End If
   Next
   cmbAno.BackColor = Branco
   cmbAno.Enabled = True
   grdTrib.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmTributoAliquota
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   cmbAno.BackColor = Branco
   cmbAno.Enabled = True
   If Evento <> "Novo" Then
        cmbTrib.BackColor = Kde
        cmbTrib.Enabled = False
   End If
End If

End Sub

Private Sub Le()
Dim x As Integer

If grdTrib.Row = 0 Then Exit Sub

For x = 0 To cmbTrib.ListCount - 1
      cmbTrib.ListIndex = x
      If cmbTrib.ItemData(cmbTrib.ListIndex) = Val(grdTrib.TextMatrix(grdTrib.Row, 0)) Then
           Exit For
      End If
Next
txtValor.Text = grdTrib.TextMatrix(grdTrib.Row, 2)

End Sub

Private Sub Limpa()
cmbTrib.ListIndex = -1
txtValor.Text = ""
End Sub

Private Sub CarregaLista()

grdTrib.Rows = 1

Sql = "SELECT TRIBUTOALIQUOTA.CODTRIBUTO,TRIBUTO.DESCTRIBUTO,TRIBUTOALIQUOTA.VALORALIQ "
Sql = Sql & "FROM TRIBUTOALIQUOTA INNER JOIN TRIBUTO ON TRIBUTOALIQUOTA.CODTRIBUTO = TRIBUTO.CODTRIBUTO "
Sql = Sql & "WHERE ANO =" & cmbAno.Text & " ORDER BY DESCTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       grdTrib.AddItem Format(!CodTributo, "000") & Chr(9) & !desctributo & Chr(9) & FormatNumber(!VALORALIQ, 4)
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub grdTrib_Click()
Le
End Sub

Private Sub Grava()

If Evento = "Novo" Then
   Sql = "INSERT TRIBUTOALIQUOTA (ANO,CODTRIBUTO,VALORALIQ) VALUES(" & cmbAno.Text & ","
   Sql = Sql & cmbTrib.ItemData(cmbTrib.ListIndex) & "," & Virg2Ponto(RemovePonto(txtValor.Text)) & ")"
Else
   Sql = "UPDATE TRIBUTOALIQUOTA SET VALORALIQ=" & Virg2Ponto(RemovePonto(txtValor.Text))
   Sql = Sql & " WHERE ANO=" & cmbAno.Text & " AND CODTRIBUTO=" & cmbTrib.ItemData(cmbTrib.ListIndex)
End If
cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   grdTrib.AddItem Format(cmbTrib.ItemData(cmbTrib.ListIndex), "000") & Chr(9) & cmbTrib.Text & Chr(9) & FormatNumber(txtValor.Text, 4)
ElseIf Evento = "Alterar" Then
   grdTrib.TextMatrix(grdTrib.Row, 2) = FormatNumber(txtValor.Text, 4)
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

Private Sub grdTrib_SelChange()
grdTrib_Click
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive, 4
End Sub
