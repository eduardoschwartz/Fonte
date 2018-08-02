VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLancamento 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Lançamentos"
   ClientHeight    =   5295
   ClientLeft      =   2460
   ClientTop       =   3840
   ClientWidth     =   6120
   Icon            =   "frmLancamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4980
      TabIndex        =   13
      ToolTipText     =   "Sair da Tela"
      Top             =   4860
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLancamento.frx":014A
      PICN            =   "frmLancamento.frx":0166
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
      Height          =   1245
      Left            =   0
      TabIndex        =   1
      Top             =   3510
      Width           =   6105
      Begin VB.TextBox txtDescC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   3
         Top             =   870
         Width           =   4605
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox txtDescR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   2
         Top             =   540
         Width           =   4605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.Completa...:"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   7
         Top             =   930
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.Resumida..:"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   5
         Top             =   585
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdLanc 
      Height          =   3450
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6085
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "Código |<Descricão  Completa                               |<Descricão  Resumida            "
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4980
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   4860
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLancamento.frx":01D4
      PICN            =   "frmLancamento.frx":01F0
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
      Top             =   4860
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
      MICON           =   "frmLancamento.frx":034A
      PICN            =   "frmLancamento.frx":0366
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
      Top             =   4860
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
      MICON           =   "frmLancamento.frx":04C0
      PICN            =   "frmLancamento.frx":04DC
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
      Top             =   4860
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLancamento.frx":0636
      PICN            =   "frmLancamento.frx":0652
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
      Top             =   4860
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
      MICON           =   "frmLancamento.frx":06F4
      PICN            =   "frmLancamento.frx":0710
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
Attribute VB_Name = "frmLancamento"
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
    If Val(txtCod.Text) = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    If Val(txtCod) < 6 Then
       MsgBox "Lançamentos 1 a 5 não podem ser alterados.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldDesc = txtDescR.Text
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
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
        
    If MsgBox("Excluir este Lancamento ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
    Ocupado
       Sql = "SELECT CODREDUZIDO FROM DEBITOPARCELA WHERE CODLANCAMENTO=" & Val(txtCod.Text)
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
          If .RowCount > 0 Then
              MsgBox "Não é possível excluir este lançamento.", vbCritical, "atenção"
              Liberado
              Exit Sub
          End If
         .Close
       End With
       Sql = "SELECT CODLANCAMENTO FROM TRIBUTOLANCAMENTO WHERE CODLANCAMENTO=" & Val(txtCod.Text)
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
          If .RowCount > 0 Then
              MsgBox "Você deve excluir antes todos os tributos deste lançamento (tela de Tributo-Lançamento).", vbCritical, "Atenção"
              Liberado
              Exit Sub
          End If
         .Close
       End With
       Sql = "DELETE FROM LANCAMENTO WHERE CODLANCAMENTO=" & txtCod.Text
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & txtDescR.Text
       Limpa
       CarregaLista
       Le
       Liberado
    End If
Exit Sub
Erro:
MsgBox Err.Description

End Sub

Private Sub cmdGravar_Click()
    If txtDescR.Text = "" Then
       MsgBox "Favor digitar a Descrição Resumida.", vbExclamation, "Atenção"
       txtDescR.SetFocus
       Exit Sub
    End If
    If txtDescC.Text = "" Then
       MsgBox "Favor digitar a Descrição Completa.", vbExclamation, "Atenção"
       txtDescC.SetFocus
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
grdLanc.Rows = 1
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
   For Each Ct In frmLancamento
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   grdLanc.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmLancamento
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   txtCod.BackColor = Kde
   txtCod.Locked = True
   grdLanc.Enabled = False
   txtDescR.SetFocus
End If

FormHagana

End Sub

Private Sub Le()
If grdLanc.Row = 0 Then Exit Sub
txtCod.Text = grdLanc.TextMatrix(grdLanc.Row, 0)
txtDescR.Text = grdLanc.TextMatrix(grdLanc.Row, 2)
txtDescC.Text = grdLanc.TextMatrix(grdLanc.Row, 1)
End Sub

Private Sub Limpa()
txtCod.Text = ""
txtDescR.Text = ""
txtDescC.Text = ""
End Sub

Private Sub CarregaLista()

Sql = "Select CODLANCAMENTO,DESCREDUZ,DESCFULL From LANCAMENTO ORDER BY codlancamento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

grdLanc.Rows = 1
With RdoAux
   .MoveFirst
    Do Until .EOF
       grdLanc.AddItem !CodLancamento & Chr(9) & !DESCFULL & Chr(9) & !DESCREDUZ
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Grava()
Dim MaxCod As Integer

Sql = "SELECT MAX(CODLANCAMENTO) AS MAXIMO FROM LANCAMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

If Evento = "Novo" Then
    Sql = "INSERT LANCAMENTO (CODLANCAMENTO,DESCREDUZ,DESCFULL) VALUES("
    Sql = Sql & MaxCod & ",'" & Mask(txtDescR.Text) & "','" & Mask(txtDescC.Text) & "')"
Else
    Sql = "UPDATE LANCAMENTO SET DESCREDUZ='" & Mask(txtDescR.Text) & "',DESCFULL='" & Mask(txtDescC.Text) & "' WHERE "
    Sql = Sql & "CODLANCAMENTO=" & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   grdLanc.AddItem MaxCod & Chr(9) & txtDescC.Text & Chr(9) & txtDescR.Text
   txtCod.Text = MaxCod
   grdLanc.Row = grdLanc.Rows - 1
   grdLanc.ColSel = 1
   Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & txtDescR.Text
 ElseIf Evento = "Alterar" Then
   grdLanc.TextMatrix(grdLanc.Row, 2) = txtDescR.Text
   grdLanc.TextMatrix(grdLanc.Row, 1) = txtDescC.Text
   Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.Text, "000") & " de " & sOldDesc & " para " & txtDescR.Text
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

Private Sub grdLanc_RowColChange()
Le
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

