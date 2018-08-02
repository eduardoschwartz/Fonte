VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{B60B1875-E5CA-11D2-BC3D-78A407C10000}#1.0#0"; "ksdpanel.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTxExpe 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Taxa de Expediente"
   ClientHeight    =   3795
   ClientLeft      =   4665
   ClientTop       =   2640
   ClientWidth     =   5610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   5610
   Begin VB.ComboBox cmbAno 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTxExpe.frx":0000
      Left            =   825
      List            =   "frmTxExpe.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   1245
   End
   Begin KSDPanel.Panel Panel2 
      Height          =   465
      Left            =   30
      TabIndex        =   14
      Top             =   3300
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   820
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   6
      BackColor       =   15658734
      Begin prjChameleon.chameleonButton cmdNovo 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         ToolTipText     =   "Novo Registro"
         Top             =   60
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTxExpe.frx":0050
         PICN            =   "frmTxExpe.frx":006C
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
         TabIndex        =   7
         ToolTipText     =   "Editar Registro"
         Top             =   60
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTxExpe.frx":01C6
         PICN            =   "frmTxExpe.frx":01E2
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
         TabIndex        =   8
         ToolTipText     =   "Excluir Registro"
         Top             =   60
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
         MICON           =   "frmTxExpe.frx":033C
         PICN            =   "frmTxExpe.frx":0358
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdHelp 
         Height          =   315
         Left            =   3240
         TabIndex        =   16
         ToolTipText     =   "Ajuda desta Tela"
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Ajuda"
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
         MICON           =   "frmTxExpe.frx":03FA
         PICN            =   "frmTxExpe.frx":0416
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
         Left            =   4305
         TabIndex        =   10
         ToolTipText     =   "Sair da Tela"
         Top             =   60
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
         MICON           =   "frmTxExpe.frx":0570
         PICN            =   "frmTxExpe.frx":058C
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
         Left            =   3240
         TabIndex        =   9
         ToolTipText     =   "Gravar os Dados"
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTxExpe.frx":05FA
         PICN            =   "frmTxExpe.frx":0616
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
         Left            =   4320
         TabIndex        =   15
         ToolTipText     =   "Cancelar Edição"
         Top             =   60
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
         MICON           =   "frmTxExpe.frx":09BB
         PICN            =   "frmTxExpe.frx":09D7
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1320
      Left            =   15
      TabIndex        =   11
      Top             =   1965
      Width           =   5565
      Begin VB.TextBox txtValorDAM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4335
         MaxLength       =   50
         TabIndex        =   5
         Top             =   555
         Width           =   1005
      End
      Begin VB.ComboBox cmbLanc 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmTxExpe.frx":0B31
         Left            =   1215
         List            =   "frmTxExpe.frx":0B33
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   4170
      End
      Begin VB.TextBox txtNormal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2010
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   555
         Width           =   1005
      End
      Begin VB.TextBox txtUnica 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2010
         MaxLength       =   50
         TabIndex        =   4
         Top             =   885
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor DAM...:"
         Height          =   195
         Index           =   2
         Left            =   3285
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançamento.:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Parcela Normal.....:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   615
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Parcela Única.......:"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   12
         Top             =   945
         Width           =   1845
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTaxa 
      Height          =   1440
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   2540
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "<Lançamento                                             |>Normal  |>Única    |>Dam        "
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano.....:"
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   18
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "frmTxExpe"
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

Private Sub cmbAno_Click()
Limpa
CarregaLista
Le
End Sub

Private Sub cmdAlterar_Click()
    If cmbLanc.ListIndex = -1 Then
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
If cmbLanc.ListIndex = -1 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

If MsgBox("Excluir esta Taxa de Expediente ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
   Sql = "DELETE FROM EXPEDIENTE WHERE ANOEXPED=" & Val(cmbAno.text) & " AND CODLANCAMENTO=" & cmbLanc.ItemData(cmbLanc.ListIndex)
   cn.Execute Sql, rdExecDirect
   Log Form, Me.Caption, Exclusão, "Excluído registro " & cmbAno.text & " - " & cmbLanc.text
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
    
If cmbLanc.ListIndex = -1 Then
   MsgBox "Selecione o lançamento.", vbExclamation, "Atenção"
   cmbLanc.SetFocus
   Exit Sub
End If

If txtNormal.text = "" Then txtNormal.text = 0
If txtUnica.text = "" Then txtUnica.text = 0
If txtValorDAM.text = "" Then txtValorDAM.text = 0

Grava
Eventos "INICIAR"
Evento = ""

End Sub

Private Sub cmdHelp_Click()
Exit Sub
  With hHelp
    .CHMFile = sPathHelp & "\Tribut.chm"
    .HHTopicID = 1390
    .HHWindow = "Main"
    .HHDisplayTopicID
  End With

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
Dim RdoAux2 As rdoResultset
Dim y As Long
Dim CTL As Object

Ocupado

Centraliza Me
frmMdi.AddWindow Me.Name, Me.Caption
sRet = RetEventUserForm(Me.Name)

grdTaxa.Rows = 1

Sql = "SELECT CODLANCAMENTO,DESCREDUZ FROM LANCAMENTO ORDER BY DESCREDUZ"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
Do Until RdoAux.EOF
     cmbLanc.AddItem RdoAux!DESCREDUZ
     cmbLanc.ItemData(cmbLanc.NewIndex) = RdoAux!CodLancamento
     RdoAux.MoveNext
Loop

cmbAno.text = Year(Now)
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
   cmdHelp.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmTxExpe
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
         Ct.Enabled = False
       End If
   Next
   cmbAno.BackColor = Branco
   cmbAno.Enabled = True
   grdTaxa.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdHelp.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmTxExpe
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   cmbAno.BackColor = Branco
   cmbAno.Enabled = True
   If Evento <> "Novo" Then
        cmbLanc.BackColor = Kde
        cmbLanc.Enabled = False
   End If
End If

'FormHagana

End Sub

Private Sub Le()
Dim x As Integer

If grdTaxa.Row = 0 Then Exit Sub

For x = 0 To cmbLanc.ListCount - 1
      cmbLanc.ListIndex = x
      If cmbLanc.ItemData(cmbLanc.ListIndex) = Val(Left$(grdTaxa.TextMatrix(grdTaxa.Row, 0), 3)) Then
           Exit For
      End If
Next
txtNormal.text = grdTaxa.TextMatrix(grdTaxa.Row, 1)
txtUnica.text = grdTaxa.TextMatrix(grdTaxa.Row, 2)
txtValorDAM.text = grdTaxa.TextMatrix(grdTaxa.Row, 3)

End Sub

Private Sub Limpa()
cmbLanc.ListIndex = -1
txtNormal.text = ""
txtUnica.text = ""
End Sub

Private Sub CarregaLista()

grdTaxa.Rows = 1

Sql = "SELECT LANCAMENTO.CODLANCAMENTO, DESCREDUZ,ANOEXPED, EXPEDIENTE.VALORPARCELA, VALORUNICA,VALORDAM "
Sql = Sql & "FROM LANCAMENTO INNER JOIN EXPEDIENTE ON LANCAMENTO.CODLANCAMENTO = EXPEDIENTE.CODLANCAMENTO "
Sql = Sql & "WHERE ANOEXPED=" & Val(cmbAno.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       grdTaxa.AddItem Format(!CodLancamento, "000") & " - " & !DESCREDUZ & Chr(9) & !VALORPARCELA & Chr(9) & !VALORUNICA & Chr(9) & !VALORDAM
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
End Sub

Private Sub grdTaxa_Click()
Le
End Sub

Private Sub Grava()
Dim qd As New rdoQuery
Dim MaxCod As Integer

On Error Resume Next
RdoAux.Close
On Error GoTo 0
Set qd.ActiveConnection = cn

qd.Sql = "{ Call spGRAVAEXPEDIENTE(?,?,?,?,?,?) }"
If Evento = "Novo" Then
   qd(0) = "S"
Else
   qd(0) = "N"
End If
qd(1) = cmbAno.text
qd(2) = cmbLanc.ItemData(cmbLanc.ListIndex)
qd(3) = Virg2Ponto(txtNormal.text)
qd(4) = Virg2Ponto(txtUnica.text)
qd(5) = Virg2Ponto(txtValorDAM.text)
Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)

If Evento = "Novo" Then
   grdTaxa.AddItem cmbLanc.text & Chr(9) & txtNormal.text & Chr(9) & txtUnica.text & Chr(9) & txtValorDAM.text
 ElseIf Evento = "Alterar" Then
   grdTaxa.TextMatrix(grdTaxa.Row, 1) = txtNormal.text
   grdTaxa.TextMatrix(grdTaxa.Row, 2) = txtUnica.text
   grdTaxa.TextMatrix(grdTaxa.Row, 3) = txtValorDAM.text
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

Private Sub grdTaxa_SelChange()
grdTaxa_Click
End Sub

Private Sub FormHagana()

evNew = 69
evEdit = 70
evDel = 71

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub

Private Sub txtNormal_KeyPress(KeyAscii As Integer)
Tweak txtNormal, KeyAscii, DecimalPositive
End Sub

Private Sub txtUnica_KeyPress(KeyAscii As Integer)
Tweak txtUnica, KeyAscii, DecimalPositive
End Sub

Private Sub txtValorDAM_KeyPress(KeyAscii As Integer)
Tweak txtValorDAM, KeyAscii, DecimalPositive
End Sub
