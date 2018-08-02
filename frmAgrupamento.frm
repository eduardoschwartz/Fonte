VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAgrupamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agrupamentos"
   ClientHeight    =   4230
   ClientLeft      =   16185
   ClientTop       =   4755
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   3300
   Begin VB.TextBox txtRedutor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3300
      Width           =   1065
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2970
      Width           =   1065
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2640
      Width           =   675
   End
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   3165
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   345
      Left            =   180
      TabIndex        =   5
      ToolTipText     =   "Novo Registro"
      Top             =   3780
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
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
      MICON           =   "frmAgrupamento.frx":0000
      PICN            =   "frmAgrupamento.frx":001C
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
      Height          =   345
      Left            =   660
      TabIndex        =   6
      ToolTipText     =   "Editar Registro"
      Top             =   3780
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
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
      MICON           =   "frmAgrupamento.frx":0176
      PICN            =   "frmAgrupamento.frx":0192
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
      Height          =   345
      Left            =   1140
      TabIndex        =   7
      ToolTipText     =   "Excluir Registro"
      Top             =   3780
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAgrupamento.frx":02EC
      PICN            =   "frmAgrupamento.frx":0308
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
      Height          =   345
      Left            =   2670
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   3780
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAgrupamento.frx":03AA
      PICN            =   "frmAgrupamento.frx":03C6
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
      Height          =   345
      Left            =   2190
      TabIndex        =   9
      ToolTipText     =   "Gravar os Dados"
      Top             =   3780
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAgrupamento.frx":0520
      PICN            =   "frmAgrupamento.frx":053C
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
      Caption         =   "Redutor..:"
      Height          =   225
      Index           =   2
      Left            =   210
      TabIndex        =   11
      Top             =   3360
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor......:"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   3030
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código...:"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   2700
      Width           =   885
   End
End
Attribute VB_Name = "frmAgrupamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String

Private Sub cmdAlterar_Click()
If Val(txtCod.Text) = 0 Then
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
Dim Sql As String

If txtCod.Text = "" Then
    MsgBox "Não existem Registros.", vbCritical, "Atenção"
    Exit Sub
End If
 
If MsgBox("Excluir este registro ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
    Sql = "delete from agrupamento where codigo=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
    CarregaLista
    Le
End If

End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset

If Val(txtCod.Text) = 0 Then
    MsgBox "Digite um código de agrupamento.", vbCritical, "Erro"
    Exit Sub
End If

If Val(txtValor.Text) = 0 Then
    MsgBox "Digite um valor para o agrupamento.", vbCritical, "Erro"
    Exit Sub
End If

If Evento = "Novo" Then
    Sql = "select valor from agrupamento where codigo=" & Val(txtCod.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        MsgBox "Código já cadastrado.", vbCritical, "Erro"
        Exit Sub
    End If
End If

Grava
Eventos "INICIAR"

End Sub

Private Sub cmdNovo_Click()
txtCod.Text = ""
txtValor.Text = ""
Eventos "INCLUIR"
Evento = "Novo"
End Sub

Private Sub Form_Load()
Centraliza Me
lstMain.Clear
CarregaLista
Le
Eventos "INICIAR"

End Sub

Private Sub lstMain_Click()
Le
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtRedutor_KeyPress(KeyAscii As Integer)
Tweak txtRedutor, KeyAscii, DecimalPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmAgrupamento
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   lstMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmAgrupamento
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   lstMain.Enabled = False
   txtCod.SetFocus
End If

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset
lstMain.Clear
Sql = "select codigo,valor,redutor from agrupamento where ano=" & Year(Now) & " order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMain.AddItem !Codigo & " - R$" & Format(!Valor, "#0.00") & " - %" & Format(!Redutor, "#0.00")
        lstMain.ItemData(lstMain.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
lstMain.ListIndex = 0


End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset

If lstMain.ListIndex = -1 Then Exit Sub

Sql = "select * from agrupamento where ano=" & Year(Now) & " and codigo=" & lstMain.ItemData(lstMain.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    txtCod.Text = lstMain.ItemData(lstMain.ListIndex)
    txtValor.Text = FormatNumber(!Valor, 2)
    txtRedutor.Text = FormatNumber(!Redutor, 2)
   .Close
End With
    
End Sub

Private Sub Grava()
Dim MaxCod As Integer, Sql As String, RdoAux As rdoResultset

Sql = "SELECT MAX(codigo) AS MAXIMO FROM agrupamento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

If txtValor.Text = "" Then txtValor.Text = 0
If txtRedutor.Text = "" Then txtRedutor.Text = 0

If Evento = "Novo" Then
    Sql = "INSERT agrupamento (ano,codigo,valor,redutor) VALUES(" & Year(Now) & "," & MaxCod & "," & Virg2Ponto(txtValor.Text) & "," & Virg2Ponto(txtValor.Text) & ")"
Else
    Sql = "UPDATE agrupamento SET valor=" & Virg2Ponto(txtValor.Text) & ",redutor=" & Virg2Ponto(txtRedutor.Text) & " where ano=" & Year(Now) & " and codigo=" & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect

CarregaLista
Le
End Sub

