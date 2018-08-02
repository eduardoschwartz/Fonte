VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCategConstr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categoria da cosntrução"
   ClientHeight    =   3735
   ClientLeft      =   12225
   ClientTop       =   4935
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4575
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4080
      Width           =   675
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1065
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   345
      Left            =   270
      TabIndex        =   3
      ToolTipText     =   "Novo Registro"
      Top             =   4320
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmCategConstr.frx":0000
      PICN            =   "frmCategConstr.frx":001C
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
      Left            =   2250
      TabIndex        =   4
      ToolTipText     =   "Editar Registro"
      Top             =   3330
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
      MICON           =   "frmCategConstr.frx":0176
      PICN            =   "frmCategConstr.frx":0192
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
      Left            =   1920
      TabIndex        =   5
      ToolTipText     =   "Excluir Registro"
      Top             =   4290
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmCategConstr.frx":02EC
      PICN            =   "frmCategConstr.frx":0308
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
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Cancelar Edição"
      Top             =   3330
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
      MICON           =   "frmCategConstr.frx":03AA
      PICN            =   "frmCategConstr.frx":03C6
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
      Left            =   2760
      TabIndex        =   7
      ToolTipText     =   "Gravar os Dados"
      Top             =   3330
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
      MICON           =   "frmCategConstr.frx":0520
      PICN            =   "frmCategConstr.frx":053C
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
      Caption         =   "Código...:"
      Height          =   225
      Index           =   0
      Left            =   330
      TabIndex        =   9
      Top             =   4140
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor......:"
      Height          =   225
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   3420
      Width           =   795
   End
End
Attribute VB_Name = "frmCategConstr"
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

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset



If Val(txtValor.Text) = 0 Then
    MsgBox "Digite um valor para a categoria.", vbCritical, "Erro"
    Exit Sub
End If


Grava
Eventos "INICIAR"

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
   For Each Ct In frmCategConstr
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
   For Each Ct In frmCategConstr
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
Dim Sql As String, RdoAux As rdoResultset, sCod As String
lstMain.Clear
Sql = "SELECT DISTINCT fatorcateg.coduso, usoconstr.descusoconstr, fatorcateg.codcateg, categconstr.desccategconstr, fatorcateg.fatorcateg2 "
Sql = Sql & "FROM fatorcateg INNER JOIN usoconstr ON fatorcateg.coduso = usoconstr.codusoconstr INNER JOIN categconstr ON fatorcateg.codcateg = categconstr.codcategconstr "
Sql = Sql & "Where (FATORCATEG.anocateg = 2017) ORDER BY fatorcateg.coduso, fatorcateg.codcateg"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sCod = CStr(!coduso) & CStr(!codcateg)
        lstMain.AddItem !descusoconstr & " - " & !desccategconstr & " - R$" & Format(!fatorcateg2, "#0.00")
        lstMain.ItemData(lstMain.NewIndex) = Val(sCod)
       .MoveNext
    Loop
   .Close
End With
lstMain.ListIndex = 0


End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset

If lstMain.ListIndex = -1 Then Exit Sub

Sql = "select fatorcateg2 from fatorcateg where anocateg=" & Year(Now) & " and coduso=" & Left(CStr(lstMain.ItemData(lstMain.ListIndex)), 1) & " and codcateg=" & Right(CStr(lstMain.ItemData(lstMain.ListIndex)), 1)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    txtCod.Text = lstMain.ItemData(lstMain.ListIndex)
    txtValor.Text = FormatNumber(!fatorcateg2, 2)
   .Close
End With
    
End Sub

Private Sub Grava()
Dim MaxCod As Integer, Sql As String, RdoAux As rdoResultset

If txtValor.Text = "" Then txtValor.Text = 0

Sql = "UPDATE fatorcateg SET fatorcateg2=" & Virg2Ponto(RemovePonto(txtValor.Text)) & " where anocateg=" & Year(Now) & " and coduso=" & Left(CStr(lstMain.ItemData(lstMain.ListIndex)), 1) & " and codcateg=" & Right(CStr(lstMain.ItemData(lstMain.ListIndex)), 1)
cn.Execute Sql, rdExecDirect

CarregaLista
Le
End Sub


