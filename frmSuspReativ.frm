VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmSuspReativ 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico de Suspenção/Reativação"
   ClientHeight    =   4245
   ClientLeft      =   2955
   ClientTop       =   3000
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   6855
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   90
      Width           =   3750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Calcular pela Data de:"
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   3600
      TabIndex        =   16
      Top             =   3045
      Width           =   2625
      Begin VB.OptionButton Opt 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Suspenção"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   270
         Width           =   1155
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Reativação"
         Height          =   210
         Index           =   1
         Left            =   1350
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2565
      Width           =   390
   End
   Begin VB.TextBox txtNumProc 
      Height          =   285
      Left            =   1785
      TabIndex        =   4
      Top             =   3015
      Width           =   1275
   End
   Begin VB.ComboBox cmbEvento 
      Height          =   315
      ItemData        =   "frmSuspReativ.frx":0000
      Left            =   2130
      List            =   "frmSuspReativ.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2565
      Width           =   2250
   End
   Begin esMaskEdit.esMaskedEdit mskDataProc 
      Height          =   285
      Left            =   1785
      TabIndex        =   5
      Top             =   3330
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmSuspReativ.frx":0025
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskData 
      Height          =   285
      Left            =   5490
      TabIndex        =   3
      Top             =   2550
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmSuspReativ.frx":0041
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5730
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   3810
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
      MICON           =   "frmSuspReativ.frx":005D
      PICN            =   "frmSuspReativ.frx":0079
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
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Novo Registro"
      Top             =   3810
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
      MICON           =   "frmSuspReativ.frx":01D3
      PICN            =   "frmSuspReativ.frx":01EF
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
      Left            =   1170
      TabIndex        =   12
      ToolTipText     =   "Editar Registro"
      Top             =   3810
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
      MICON           =   "frmSuspReativ.frx":0349
      PICN            =   "frmSuspReativ.frx":0365
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
      Left            =   2220
      TabIndex        =   13
      ToolTipText     =   "Excluir Registro"
      Top             =   3810
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
      MICON           =   "frmSuspReativ.frx":04BF
      PICN            =   "frmSuspReativ.frx":04DB
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
      Left            =   4680
      TabIndex        =   14
      ToolTipText     =   "Gravar os Dados"
      Top             =   3780
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
      MICON           =   "frmSuspReativ.frx":057D
      PICN            =   "frmSuspReativ.frx":0599
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1935
      MaxLength       =   6
      TabIndex        =   0
      Top             =   90
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid grdSusp 
      Height          =   1995
      Left            =   45
      TabIndex        =   8
      Top             =   405
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   3519
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "^Seq   |<Evento                               |^Data               |<Nº do Processo     |^Data Processo |^C  "
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5760
      TabIndex        =   15
      ToolTipText     =   "Sair da Tela"
      Top             =   3810
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
      MICON           =   "frmSuspReativ.frx":093E
      PICN            =   "frmSuspReativ.frx":095A
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
      Caption         =   "Data do Processo......:"
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   21
      Top             =   3375
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sequência..:"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   20
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo..........:"
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   19
      Top             =   3060
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Evento.:"
      Height          =   225
      Index           =   3
      Left            =   4440
      TabIndex        =   18
      Top             =   2610
      Width           =   990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Evento.:"
      Height          =   225
      Index           =   4
      Left            =   1485
      TabIndex        =   17
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código da Empresa....:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   105
      Width           =   1740
   End
End
Attribute VB_Name = "frmSuspReativ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Dim Evento As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Private Sub cmdAlterar_Click()
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()
Dim nTipo As Integer
Dim nSeq As Integer

If grdSusp.Rows = 1 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

If grdSusp.Row <> grdSusp.Rows - 1 Then
   MsgBox "Apenas o último evento pode ser excluído.", vbCritical, "Atenção"
   Exit Sub
End If

If MsgBox("Excluir este evento ?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
   nSeq = Val(grdSusp.TextMatrix(grdSusp.Rows - 1, 0))
   nTipo = Val(Left$(grdSusp.TextMatrix(grdSusp.Rows - 1, 1), 1))
   Sql = "DELETE FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & Val(txtCod.Text) & " AND "
   Sql = Sql & "CODTIPOEVENTO=" & nTipo & " AND SEQEVENTO=" & nSeq
   cn.Execute Sql, rdExecDirect
   CarregaLista
End If

End Sub

Private Sub cmdGravar_Click()

Dim nTipo As Integer

If cmbEvento.ListIndex = -1 Then
   MsgBox "Selecione o evento.", vbExclamation, "Atenção"
   Exit Sub
End If
If Not IsDate(mskData.Text) Then
   MsgBox "Digite a data do evento.", vbExclamation, "Atenção"
   Exit Sub
End If

If grdSusp.Rows > 1 Then
    If CDate(mskData.Text) <= CDate(grdSusp.TextMatrix(grdSusp.Rows - 1, 2)) Then
       MsgBox "Data do evento inválida.", vbExclamation, "Atenção"
       Exit Sub
    End If
End If

If Not IsDate(mskDataProc.Text) Then
   MsgBox "Digite a data do processo.", vbExclamation, "Atenção"
   Exit Sub
End If
If txtNumProc.Text = "" Then
   MsgBox "Digite o nº do processo.", vbExclamation, "Atenção"
   Exit Sub
End If

If grdSusp.Rows > 1 Then
   nTipo = Val(Left(grdSusp.TextMatrix(grdSusp.Rows - 1, 1), 1))
   If nTipo = 2 And cmbEvento.ListIndex = 0 Then   'suspenao
      MsgBox "Esta empresa já esta suspensa.", vbExclamation, "Atenção"
      Exit Sub
   ElseIf nTipo = 3 And cmbEvento.ListIndex = 1 Then 'reativacao
      MsgBox "Esta empresa já foi reativada.", vbExclamation, "Atenção"
      Exit Sub
   End If
End If

Grava
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
If Val(txtCod.Text) > 0 Then
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
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
   For Each Ct In frmSuspReativ
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Or TypeOf Ct Is OptionButton Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   grdSusp.Enabled = True
   txtCod.Enabled = True
   txtCod.BackColor = Branco
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmSuspReativ
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Or TypeOf Ct Is OptionButton Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   Opt(0).BackColor = Kde
   Opt(1).BackColor = Kde
   txtSeq.Locked = True
   txtSeq.BackColor = Kde
   grdSusp.Enabled = False
End If

FormHagana

End Sub

Private Sub Le()

With grdSusp
    If .Rows = 1 Then Exit Sub
    txtSeq.Text = .TextMatrix(.Row, 0)
    cmbEvento.ListIndex = IIf(Left(.TextMatrix(.Row, 1), 1) = 2, 0, 1)
    mskData.Text = .TextMatrix(.Row, 2)
    If .TextMatrix(.Row, 4) <> "" Then
       mskDataProc.Text = .TextMatrix(.Row, 4)
    Else
       LimpaMascara mskDataProc
    End If
    txtNumProc.Text = .TextMatrix(.Row, 3)
    If Val(.TextMatrix(.Row, 5)) = 0 Then
       Opt(0).value = True
    Else
       Opt(1).value = True
    End If
End With

End Sub

Private Sub Limpa()
With grdSusp
    txtSeq.Text = ""
    cmbEvento.ListIndex = -1
    LimpaMascara mskData
    LimpaMascara mskDataProc
    txtNumProc.Text = ""
    Opt(1).value = True
End With
End Sub

Private Sub CarregaLista()
txtNome.Text = ""
Sql = "Select CODIGOMOB,RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

grdSusp.Rows = 1
Limpa
With RdoAux
    If .RowCount = 0 Then
       MsgBox "Código não cadastrado.", vbExclamation, "Atenção"
      .Close
       Exit Sub
    Else
        txtNome.Text = !RazaoSocial
    End If
End With

Sql = "SELECT MOBILIARIOEVENTO.CODMOBILIARIO,MOBILIARIOEVENTO.CODTIPOEVENTO,TIPOEVENTOMOBILIARIO.DESCTIPOEVENTO,MOBILIARIOEVENTO.SEQEVENTO,"
Sql = Sql & "MOBILIARIOEVENTO.DATAEVENTO,MOBILIARIOEVENTO.NUMPROCEVENTO,MOBILIARIOEVENTO.DATAPROCEVENTO,TIPOCALCULO FROM MOBILIARIOEVENTO INNER JOIN "
Sql = Sql & "TIPOEVENTOMOBILIARIO ON MOBILIARIOEVENTO.CODTIPOEVENTO = TIPOEVENTOMOBILIARIO.CODTIPOEVENTO "
Sql = Sql & "Where MOBILIARIOEVENTO.CODTIPOEVENTO > 1 AND CODMOBILIARIO=" & Val(txtCod.Text)
Sql = Sql & " ORDER BY MOBILIARIOEVENTO.SEQEVENTO,MOBILIARIOEVENTO.CODTIPOEVENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       grdSusp.AddItem !SEQEVENTO & Chr(9) & !CODTIPOEVENTO & " - " & !DESCTIPOEVENTO & Chr(9) & Format(!DATAEVENTO, "dd/mm/yyyy") & Chr(9) & !NUMPROCEVENTO & Chr(9) & Format(!DATAPROCEVENTO, "dd/mm/yyyy") & Chr(9) & !TIPOCALCULO
      .MoveNext
    Loop
End With
If grdSusp.Rows > 1 Then Le

cmdNovo.Enabled = True
'cmdAlterar.Enabled = True
cmdExcluir.Enabled = True


End Sub

Private Sub Grava()
Dim nTipo As Integer
Dim nSeq As Integer
Dim sTexto1 As String

If cmbEvento.ListIndex = 0 Then
    nTipo = 2
Else
    nTipo = 3
End If

Sql = "SELECT MAX(SEQEVENTO) AS MAXIMO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & Val(txtCod.Text)
Sql = Sql & " AND CODTIPOEVENTO=" & nTipo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 0
    Else
        nSeq = !maximo + 1
    End If
End With

Sql = "INSERT MOBILIARIOEVENTO (CODMOBILIARIO,CODTIPOEVENTO,SEQEVENTO,DATAEVENTO,NUMPROCEVENTO,DATAPROCEVENTO,TIPOCALCULO) VALUES("
Sql = Sql & Val(txtCod.Text) & "," & nTipo & "," & nSeq & ",'" & Format(mskData.Text, "mm/dd/yyyy") & "','" & txtNumProc.Text & "','" & Format(mskDataProc.Text, "mm/dd/yyyy") & "'," & IIf(Opt(0).value = True, 0, 1) & ")"
cn.Execute Sql, rdExecDirect


Sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeq = 0
Else
    nSeq = RdoAux!maximo + 1
End If
            
sTexto1 = "A Empresa foi " & IIf(cmbEvento.ListIndex = 0, "Suspensa", "Reativada") & " através do processo nº " & txtNumProc.Text & " em " & mskData.Text & " por " & RetornaUsuarioFullName & "."
            
'Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USUARIO) VALUES("
'Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sTexto1) & "','GTI')"
Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sTexto1) & "',236)"
cn.Execute Sql, rdExecDirect


CarregaLista


'Integração_Eicon
Sql = "select codigo from eicon_suspensao where codigo=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    Sql = "insert eicon_suspensao(codigo) values(" & Val(txtCod.Text) & ")"
    cn.Execute Sql, rdExecDirect
End If
RdoAux.Close


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

Private Sub grdSusp_Click()

If grdSusp.Rows = 1 Then Exit Sub
If grdSusp.Row > 0 Then Le

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    CarregaLista
Else
    cmdNovo.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    Tweak txtCod, KeyAscii, IntegerPositive
End If

End Sub

Private Sub txtNumProc_LostFocus()
Dim sValidaProc As String
If Trim(txtNumProc.Text) = "" Then Exit Sub
sValidaProc = ValidaProcesso(txtNumProc.Text)
If sValidaProc = "Processo não Cadastrado." Then
    MsgBox sValidaProc, vbCritical, "Atenção"
    LimpaMascara mskDataProc
Else
    mskDataProc.Text = Format(RetornaDataProcesso(Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 5)), Val(Right$(txtNumProc.Text, 4))), "dd/mm/yyyy")
End If

End Sub
