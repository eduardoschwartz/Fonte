VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{B60B1875-E5CA-11D2-BC3D-78A407C10000}#1.0#0"; "KSDPANEL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmIndCor 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Índice de Correção"
   ClientHeight    =   3060
   ClientLeft      =   5070
   ClientTop       =   1800
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   3915
   Begin KSDPanel.Panel Panel2 
      Height          =   525
      Left            =   0
      TabIndex        =   7
      Top             =   2550
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   926
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
         TabIndex        =   8
         ToolTipText     =   "Novo Registro"
         Top             =   90
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIndCor.frx":0000
         PICN            =   "frmIndCor.frx":001C
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
         Left            =   510
         TabIndex        =   9
         ToolTipText     =   "Editar Registro"
         Top             =   90
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIndCor.frx":0176
         PICN            =   "frmIndCor.frx":0192
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
         Left            =   930
         TabIndex        =   10
         ToolTipText     =   "Excluir Registro"
         Top             =   90
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIndCor.frx":02EC
         PICN            =   "frmIndCor.frx":0308
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
         Left            =   1350
         TabIndex        =   11
         ToolTipText     =   "Ajuda desta Tela"
         Top             =   90
         Width           =   370
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIndCor.frx":03AA
         PICN            =   "frmIndCor.frx":03C6
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
         Left            =   2790
         TabIndex        =   12
         ToolTipText     =   "Cancelar Edição"
         Top             =   90
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIndCor.frx":0520
         PICN            =   "frmIndCor.frx":053C
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
         Left            =   3210
         TabIndex        =   13
         ToolTipText     =   "Sair da Tela"
         Top             =   90
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIndCor.frx":0696
         PICN            =   "frmIndCor.frx":06B2
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
         Left            =   3210
         TabIndex        =   14
         ToolTipText     =   "Gravar os Dados"
         Top             =   90
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIndCor.frx":0720
         PICN            =   "frmIndCor.frx":073C
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
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   1500
      Width           =   3885
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   592
         Width           =   885
      End
      Begin VB.TextBox txtAno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   6
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Índice Correção...:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   645
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano Exercício......:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   285
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTaxa 
      Height          =   1470
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   2593
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Ano                |>Valor Índice             "
   End
End
Attribute VB_Name = "frmIndCor"
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
    If Val(txtAno.text) = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldDesc = txtAno.text
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
If Val(txtAno.text) = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

If MsgBox("Excluir este índice de correção ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
   Sql = "DELETE FROM INDICECORRECAO WHERE ANOCORRECAO=" & Val(txtAno.text)
   cn.Execute Sql, rdExecDirect
   Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtAno.text, "000")
   Limpa
   CarregaLista
   Le
End If
    
Exit Sub
Erro:

For X = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(X).Description
Next

End Sub

Private Sub cmdGravar_Click()
    If Val(txtAno.text) = 0 Then
       MsgBox "Favor digitar o Ano.", vbExclamation, "Atenção"
       txtAno.SetFocus
       Exit Sub
    End If
    
    If Val(txtValor.text) = 0 Then txtValor.text = 0
    
    If Evento = "Novo" Then
            Sql = "SELECT VALORCORRECAO FROM INDICECORRECAO WHERE ANOCORRECAO=" & txtAno.text
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                   If .RowCount > 0 Then
                        MsgBox "Índice ja cadastrado para este Ano.", vbExclamation, "Atenção"
                       .Close
                        Exit Sub
                   End If
            End With
    End If
    
    Grava
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdHelp_Click()
Exit Sub
  With hHelp
    .CHMFile = sPathHelp & "\Tribut.chm"
    .HHTopicID = 1050
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
Dim Y As Long
Dim CTL As Object

Ocupado
Centraliza Me

sRet = RetEventUserForm(Me.Name)

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
   For Each Ct In frmIndCor
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
   For Each Ct In frmIndCor
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   If Evento <> "Novo" Then
        txtAno.BackColor = Kde
        txtAno.Enabled = False
   End If
End If

FormHagana

End Sub

Private Sub Le()
Dim X As Integer

If grdTaxa.Row = 0 Then Exit Sub
txtAno.text = grdTaxa.TextMatrix(grdTaxa.Row, 0)
txtValor.text = grdTaxa.TextMatrix(grdTaxa.Row, 1)

End Sub

Private Sub Limpa()
txtAno.text = ""
txtValor.text = ""
End Sub

Private Sub CarregaLista()

grdTaxa.Rows = 1

Sql = "Select ANOCORRECAO,VALORCORRECAO FROM INDICECORRECAO ORDER BY ANOCORRECAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       grdTaxa.AddItem !ANOCORRECAO & Chr(9) & !VALORCORRECAO
      .MoveNext
    Loop
   .Close
End With

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

qd.Sql = "{ Call spGRAVAINDCORRECAO(?,?,?) }"
If Evento = "Novo" Then
   qd(0) = "S"
Else
   qd(0) = "N"
End If
qd(1) = txtAno.text
qd(2) = Virg2Ponto(txtValor.text)
Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)

If Evento = "Novo" Then
   grdTaxa.AddItem txtAno.text & Chr(9) & txtValor.text
   Log Form, Me.Caption, Inclusão, "Inserido registro " & txtAno.text
 ElseIf Evento = "Alterar" Then
   grdTaxa.TextMatrix(grdTaxa.Row, 1) = txtValor.text
   Log Form, Me.Caption, Alteração, "Alterado registro " & txtAno.text
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

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, DecimalPositive
End Sub
