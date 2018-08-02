VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTipoLogradouro 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Logradouro"
   ClientHeight    =   4050
   ClientLeft      =   5625
   ClientTop       =   4800
   ClientWidth     =   6120
   Icon            =   "frmTipoLogradouro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4980
      TabIndex        =   13
      ToolTipText     =   "Sair da Tela"
      Top             =   3660
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
      MICON           =   "frmTipoLogradouro.frx":014A
      PICN            =   "frmTipoLogradouro.frx":0166
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
      Top             =   3630
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
      MICON           =   "frmTipoLogradouro.frx":01D4
      PICN            =   "frmTipoLogradouro.frx":01F0
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
      Top             =   3630
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
      MICON           =   "frmTipoLogradouro.frx":034A
      PICN            =   "frmTipoLogradouro.frx":0366
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
      Top             =   3630
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
      MICON           =   "frmTipoLogradouro.frx":04C0
      PICN            =   "frmTipoLogradouro.frx":04DC
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
      Top             =   3630
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
      MICON           =   "frmTipoLogradouro.frx":0636
      PICN            =   "frmTipoLogradouro.frx":0652
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
      Top             =   3630
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
      MICON           =   "frmTipoLogradouro.frx":06F4
      PICN            =   "frmTipoLogradouro.frx":0710
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTipoLog 
      Height          =   2160
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3810
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "Código  |<Descricão                                                               |<Abrev     "
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1305
      Left            =   0
      TabIndex        =   1
      Top             =   2250
      Width           =   6105
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   4
         Top             =   555
         Width           =   4095
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox txtAbrev 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   915
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição...........:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Abreviatura.........:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   975
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmTipoLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String, sOldAbrev As String
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Private Sub cmdAlterar_Click()
    If txtCod.text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldDesc = txtDesc.text
    sOldAbrev = txtAbrev.text
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()
    If txtCod.text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir este Título de Logradouro ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "DELETE FROM TIPOLOGRADOURO WHERE CODTIpoLOG=" & txtCod.text
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.text, "000") & "-" & txtDesc.text
       Limpa
       CarregaLista
       Le
    End If
End Sub

Private Sub cmdGravar_Click()
    If txtDesc.text = "" Then
       MsgBox "Favor digitar o Nome do Logradouro.", vbExclamation, "Atenção"
       txtDesc.SetFocus
       Exit Sub
    End If
    If txtAbrev.text = "" Then
       MsgBox "Favor digitar a Abreviatura do Logradouro.", vbExclamation, "Atenção"
       txtAbrev.SetFocus
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

grdTipoLog.Rows = 1
grdTipoLog.ColWidth(0) = 0
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
   For Each Ct In frmTipoLogradouro
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   grdTipoLog.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmTipoLogradouro
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   txtCod.BackColor = Kde
   txtCod.Locked = True
   grdTipoLog.Enabled = False
   txtDesc.SetFocus
End If

FormHagana

End Sub

Private Sub Le()
If grdTipoLog.Row = 0 Then Exit Sub
txtCod.text = grdTipoLog.TextMatrix(grdTipoLog.Row, 0)
txtDesc.text = grdTipoLog.TextMatrix(grdTipoLog.Row, 1)
txtAbrev.text = grdTipoLog.TextMatrix(grdTipoLog.Row, 2)

End Sub

Private Sub Limpa()
txtCod.text = ""
txtDesc.text = ""
txtAbrev.text = ""
End Sub

Private Sub CarregaLista()

Sql = "Select CODTIPOLOG,NOMETIPOLOG,ABREVTIPOLOG From TIPOLOGRADOURO WHERE CODTIPOLOG<>9999 "
Sql = Sql & "ORDER BY NOMETIPOLOG"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
grdTipoLog.Rows = 1
With RdoAux
   .MoveFirst
    Do Until .EOF
       grdTipoLog.AddItem !CODTIPOLOG & Chr(9) & !NomeTipoLog & Chr(9) & !AbrevTipoLog
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub grdTipoLog_Click()
Le
End Sub

Private Sub Grava()

Dim MaxCod As Integer
Sql = "SELECT MAX(CODTIPOLOG) AS MAXIMO FROM TIPOLOGRADOURO WHERE CODTIPOLOG<9999"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!MAXIMO) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!MAXIMO + 1
End If
RdoAux.Close

If Evento = "Novo" Then
    Sql = "INSERT TIPOLOGRADOURO (CODTIPOLOG,NOMETIPOLOG,ABREVTIPOLOG) VALUES("
    Sql = Sql & MaxCod & ",'" & Mask(txtDesc.text) & "','" & Mask(txtAbrev.text) & "')"
Else
    Sql = "UPDATE TIPOLOGRADOURO SET NOMETIPOLOG='" & Mask(txtDesc.text) & "',ABREVTIPOLOG='" & Mask(txtAbrev.text) & "' WHERE "
    Sql = Sql & "CODTIPOLOG=" & Val(txtCod.text)
End If
cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   grdTipoLog.AddItem MaxCod & Chr(9) & txtDesc.text & Chr(9) & txtAbrev.text
   txtCod.text = MaxCod
   grdTipoLog.Row = grdTipoLog.Rows - 1
   grdTipoLog.ColSel = 2
   Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & txtDesc.text & "-" & txtAbrev.text
 ElseIf Evento = "Alterar" Then
   grdTipoLog.TextMatrix(grdTipoLog.Row, 1) = txtDesc.text
   grdTipoLog.TextMatrix(grdTipoLog.Row, 2) = txtAbrev.text
   Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.text, "000") & " de " & sOldDesc & "-" & sOldAbrev & " para " & txtDesc.text & "-" & txtAbrev.text
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

Private Sub grdTipoLog_RowColChange()
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

