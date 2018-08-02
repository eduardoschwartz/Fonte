VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTaxaSelic 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela Selic"
   ClientHeight    =   3345
   ClientLeft      =   8640
   ClientTop       =   4395
   ClientWidth     =   3660
   Icon            =   "frmTaxaSelic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   1980
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmTaxaSelic.frx":000C
      PICN            =   "frmTaxaSelic.frx":0028
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
      Height          =   615
      Left            =   30
      TabIndex        =   4
      Top             =   2730
      Width           =   3615
      Begin VB.TextBox txtMes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   2
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtAno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   660
         MaxLength       =   4
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   555
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2940
         MaxLength       =   8
         TabIndex        =   3
         Top             =   210
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mês..:"
         Height          =   195
         Index           =   1
         Left            =   1290
         TabIndex        =   13
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano..:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Perc.:"
         Height          =   195
         Index           =   11
         Left            =   2430
         TabIndex        =   5
         Top             =   240
         Width           =   465
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTaxa 
      Height          =   2670
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4710
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Ano          |^Mês    |>Perc. %     "
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      ToolTipText     =   "Cancelar Edição"
      Top             =   2010
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmTaxaSelic.frx":0096
      PICN            =   "frmTaxaSelic.frx":00B2
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
      Left            =   3000
      TabIndex        =   8
      ToolTipText     =   "Novo Registro"
      Top             =   390
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmTaxaSelic.frx":020C
      PICN            =   "frmTaxaSelic.frx":0228
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
      Left            =   3000
      TabIndex        =   9
      ToolTipText     =   "Editar Registro"
      Top             =   780
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmTaxaSelic.frx":0382
      PICN            =   "frmTaxaSelic.frx":039E
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
      Left            =   3000
      TabIndex        =   10
      ToolTipText     =   "Excluir Registro"
      Top             =   1170
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmTaxaSelic.frx":04F8
      PICN            =   "frmTaxaSelic.frx":0514
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
      Left            =   3000
      TabIndex        =   11
      ToolTipText     =   "Gravar os Dados"
      Top             =   1560
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmTaxaSelic.frx":05B6
      PICN            =   "frmTaxaSelic.frx":05D2
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
Attribute VB_Name = "frmTaxaSelic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String
Dim Evento As String

Private Sub cmdAlterar_Click()
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
    If txtAno.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
        
    If MsgBox("Excluir esta taxa ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "DELETE FROM TAXASELIC WHERE ANO=" & Val(txtAno.Text) & " AND MES=" & Val(txtMes.Text)
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído registro " & txtAno.Text & "-" & txtMes.Text
       Limpa
       CarregaLista
       Le
    End If
Exit Sub
Erro:
MsgBox Err.Description

End Sub

Private Sub cmdGravar_Click()
    If Not IsNumeric(txtValor.Text) Then txtValor.Text = "0"
    If Val(txtAno.Text) < 2007 Or Val(txtAno.Text) > Year(Now) Then
       MsgBox "Ano Inválido.", vbCritical, "Atenção"
       Exit Sub
    End If
    If Val(txtMes.Text) < 1 Or Val(txtMes.Text) > 12 Then
       MsgBox "Mês inválido.", vbCritical, "Atenção"
       Exit Sub
    End If
    If CDbl(txtValor.Text) = 0 Then
       MsgBox "Digite o valor.", vbCritical, "Atenção"
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
grdTaxa.Rows = 1
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
   For Each Ct In frmTaxaSelic
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   grdTaxa.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmTaxaSelic
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   grdTaxa.Enabled = False
   If Evento = "Alterar" Then
        txtAno.Enabled = False
        txtMes.Enabled = False
        txtValor.SetFocus
   Else
        txtAno.SetFocus
   End If
   
   
   
End If

End Sub

Private Sub Le()
If grdTaxa.Row = 0 Then Exit Sub
txtAno.Text = grdTaxa.TextMatrix(grdTaxa.Row, 0)
txtMes.Text = grdTaxa.TextMatrix(grdTaxa.Row, 1)
txtValor.Text = grdTaxa.TextMatrix(grdTaxa.Row, 2)
End Sub

Private Sub Limpa()
txtAno.Text = ""
txtMes.Text = ""
txtValor.Text = ""
End Sub

Private Sub CarregaLista()

Sql = "Select * From TAXASELIC ORDER BY ANO,MES"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

grdTaxa.Rows = 1
With RdoAux
   .MoveFirst
    Do Until .EOF
       grdTaxa.AddItem !Ano & Chr(9) & !Mes & Chr(9) & FormatNumber(!VALOR, 2)
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Grava()
Dim MaxCod As Integer

If Evento = "Novo" Then
    Sql = "INSERT TAXASELIC (ANO,MES,VALOR) VALUES("
    Sql = Sql & txtAno.Text & ",'" & txtMes.Text & "','" & Virg2Ponto(txtValor.Text) & "')"
Else
    Sql = "UPDATE TAXASELIC SET VALOR=" & Virg2Ponto(txtValor.Text) & " WHERE "
    Sql = Sql & "ANO=" & txtAno.Text & " AND MES=" & txtMes.Text
End If
cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   grdTaxa.AddItem txtAno.Text & Chr(9) & txtMes.Text & Chr(9) & txtValor.Text
   grdTaxa.Row = grdTaxa.Rows - 1
   grdTaxa.ColSel = 2
 ElseIf Evento = "Alterar" Then
   grdTaxa.TextMatrix(grdTaxa.Row, 2) = txtValor
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

Private Sub grdTaxa_RowColChange()
Le
End Sub
