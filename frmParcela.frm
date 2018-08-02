VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmParcela 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manutenção de Parcela"
   ClientHeight    =   5280
   ClientLeft      =   4455
   ClientTop       =   3495
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbStatus 
      Height          =   315
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   810
      Width           =   3075
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4620
      Width           =   6795
   End
   Begin VB.TextBox txtValor 
      Height          =   315
      Left            =   750
      TabIndex        =   5
      Top             =   3930
      Width           =   1515
   End
   Begin VB.ComboBox cmbTrib 
      Height          =   315
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3570
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   2190
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   2249
      _Version        =   393216
      Rows            =   10
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "<Tributo                                                                  |>Valor                    "
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   5490
      TabIndex        =   1
      ToolTipText     =   "Novo Registro"
      Top             =   2400
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
      MICON           =   "frmParcela.frx":0000
      PICN            =   "frmParcela.frx":001C
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
      Left            =   5490
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   2730
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
      MICON           =   "frmParcela.frx":0176
      PICN            =   "frmParcela.frx":0192
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
      Left            =   5490
      TabIndex        =   3
      ToolTipText     =   "Excluir Registro"
      Top             =   3060
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
      MICON           =   "frmParcela.frx":02EC
      PICN            =   "frmParcela.frx":0308
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
      Left            =   5490
      TabIndex        =   6
      ToolTipText     =   "Gravar os Dados"
      Top             =   2730
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
      MICON           =   "frmParcela.frx":03AA
      PICN            =   "frmParcela.frx":03C6
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
      Left            =   5490
      TabIndex        =   7
      ToolTipText     =   "Cancelar Edição"
      Top             =   3060
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
      MICON           =   "frmParcela.frx":076B
      PICN            =   "frmParcela.frx":0787
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtVencto 
      Height          =   315
      Left            =   3660
      TabIndex        =   14
      Top             =   1140
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94175233
      CurrentDate     =   38187
   End
   Begin VB.Label lblCompl 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3660
      TabIndex        =   28
      Top             =   525
      Width           =   1125
   End
   Begin VB.Label lblParc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3660
      TabIndex        =   27
      Top             =   210
      Width           =   1125
   End
   Begin VB.Label lblSeq 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1260
      TabIndex        =   26
      Top             =   1140
      Width           =   1125
   End
   Begin VB.Label lblLanc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1260
      TabIndex        =   25
      Top             =   825
      Width           =   1125
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1260
      TabIndex        =   24
      Top             =   525
      Width           =   1125
   End
   Begin VB.Label lblCod 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1260
      TabIndex        =   23
      Top             =   210
      Width           =   1125
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
      Height          =   255
      Left            =   2490
      TabIndex        =   22
      Top             =   1170
      Width           =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcela"
      Height          =   255
      Left            =   2490
      TabIndex        =   21
      Top             =   210
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   2490
      TabIndex        =   20
      Top             =   855
      Width           =   1125
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
      Height          =   255
      Left            =   2490
      TabIndex        =   19
      Top             =   525
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sequencia"
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   1140
      Width           =   1125
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lancamento"
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   825
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Exercício"
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   525
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   255
      Left            =   90
      TabIndex        =   15
      Top             =   210
      Width           =   1125
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2940
      TabIndex        =   12
      Top             =   4380
      Width           =   3795
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Alterações Registradas para o usuário:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   4380
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   4005
      Width           =   705
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tributo:"
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   3660
      Width           =   705
   End
End
Attribute VB_Name = "frmParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nLinha As Integer, Sql As String, RdoAux As rdoResultset
Dim Evento As String
Dim doData As Date
Dim soStatus As String
Dim soTributo As String
Dim soValor As Double

Private Sub cmbStatus_Click()
If cmbStatus.ListIndex = -1 Then Exit Sub
If cmbStatus.Text <> soStatus Then
    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & Val(Left$(cmbStatus.Text, 2))
    Sql = Sql & " WHERE CODREDUZIDO = " & Val(lblCod.Caption) & " AND ANOEXERCICIO = " & Val(lblAno.Caption) & " AND CODLANCAMENTO = " & Val(lbllanc.Caption) & " AND "
    Sql = Sql & "SEQLANCAMENTO = " & Val(lblSeq.Caption) & " AND NUMPARCELA = " & lblParc.Caption & " AND CODCOMPLEMENTO = " & lblCompl.Caption
    cn.Execute Sql, rdExecDirect
    frmDebitoImob.grdExtrato.CellText(nLinha, 6) = cmbStatus.Text
    GravaLog "Status alterado de " & soStatus & " para " & cmbStatus.Text
    soStatus = cmbStatus.Text
End If
End Sub

Private Sub cmdAlterar_Click()
Evento = "Alterar"
Eventos Evento
txtValor.SetFocus
End Sub

Private Sub cmdCancel_Click()
Eventos ""
End Sub

Private Sub cmdExcluir_Click()
soTributo = grdTrib.TextMatrix(grdTrib.Row, 0)
soValor = grdTrib.TextMatrix(grdTrib.Row, 1)

If MsgBox("Excluir este Tributo ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    If grdTrib.Rows > 2 Then
       Sql = "DELETE FROM DEBITOTRIBUTO "
       Sql = Sql & " WHERE CODREDUZIDO = " & Val(lblCod.Caption) & " AND ANOEXERCICIO = " & Val(lblAno.Caption) & " AND CODLANCAMENTO = " & Val(lbllanc.Caption) & " AND "
       Sql = Sql & "SEQLANCAMENTO = " & Val(lblSeq.Caption) & " AND NUMPARCELA = " & lblParc.Caption & " AND CODCOMPLEMENTO = " & lblCompl.Caption & " AND CODTRIBUTO=" & Val(Left$(grdTrib.TextMatrix(grdTrib.Row, 0), 3))
       cn.Execute Sql, rdExecDirect
       grdTrib.RemoveItem (grdTrib.Row)
    Else
       grdTrib.Rows = 1
    End If
    GravaLog "Excluido tributo " & soTributo & " com valor de " & FormatNumber(soValor, 2)
    If grdTrib.Rows > 1 Then
        grdTrib.Row = 1
        grdTrib_Click
    End If
End If

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, bAchou As Boolean
If Trim(txtValor.Text) = "" Then txtValor.Text = 0


If Evento = "Novo" Then
    soTributo = Val(Left(cmbTrib.Text, 3))
    soValor = txtValor.Text
    bAchou = False
    For x = 1 To grdTrib.Rows - 1
        If grdTrib.TextMatrix(grdTrib.Row, 0) = cmbTrib.Text Then
            bAchou = True
        End If
    Next
    If bAchou Then
        MsgBox "Tributo já cadastrado.", vbExclamation, "Atenção"
        Exit Sub
    Else
        Sql = "INSERT DEBITOTRIBUTO(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
        Sql = Sql & Val(lblCod.Caption) & "," & Val(lblAno.Caption) & "," & Val(lbllanc.Caption) & ","
        Sql = Sql & Val(lblSeq.Caption) & "," & Val(lblParc.Caption) & "," & Val(lblCompl.Caption) & ","
        Sql = Sql & Left$(cmbTrib.Text, 3) & "," & Virg2Ponto(RemovePonto(txtValor.Text)) & ")"
        cn.Execute Sql, rdExecDirect
        
        grdTrib.AddItem cmbTrib.Text & Chr(9) & FormatNumber(txtValor.Text, 2)
        GravaLog "Adicionado tributo " & cmbTrib.Text & " Valor: " & FormatNumber(txtValor.Text, 2)
    End If
Else
soTributo = grdTrib.TextMatrix(grdTrib.Row, 0)
soValor = grdTrib.TextMatrix(grdTrib.Row, 1)
    Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=" & Virg2Ponto(RemovePonto(txtValor.Text))
    Sql = Sql & " WHERE CODREDUZIDO = " & Val(lblCod.Caption) & " AND ANOEXERCICIO = " & Val(lblAno.Caption) & " AND CODLANCAMENTO = " & Val(lbllanc.Caption) & " AND "
    Sql = Sql & "SEQLANCAMENTO = " & Val(lblSeq.Caption) & " AND NUMPARCELA = " & lblParc.Caption & " AND CODCOMPLEMENTO = " & lblCompl.Caption & " AND CODTRIBUTO=" & Left$(cmbTrib.Text, 3)
    cn.Execute Sql, rdExecDirect
    grdTrib.TextMatrix(grdTrib.Row, 1) = FormatNumber(txtValor.Text, 2)
    GravaLog "Alterado tributo " & soTributo & " Valor de " & soValor & " para " & FormatNumber(txtValor.Text, 2)
End If

Eventos ""
End Sub

Private Sub cmdNovo_Click()
Evento = "Novo"
Eventos Evento
cmbTrib.SetFocus
End Sub

Private Sub dtVencto_CloseUp()
If dtVencto.value <> doData Then
    Sql = "UPDATE DEBITOPARCELA SET DATAVENCIMENTO='" & Format(dtVencto.value, "mm/dd/yyyy")
    Sql = Sql & "' WHERE CODREDUZIDO = " & Val(lblCod.Caption) & " AND ANOEXERCICIO = " & Val(lblAno.Caption) & " AND CODLANCAMENTO = " & Val(lbllanc.Caption) & " AND "
    Sql = Sql & "SEQLANCAMENTO = " & Val(lblSeq.Caption) & " AND NUMPARCELA = " & lblParc.Caption & " AND CODCOMPLEMENTO = " & lblCompl.Caption
    cn.Execute Sql, rdExecDirect
    frmDebitoImob.grdExtrato.CellText(nLinha, 7) = Format(dtVencto.value, "dd/mm/yyyy")
    GravaLog "Vencto alterado de " & doData & " para " & dtVencto.value
    doData = dtVencto.value
End If
End Sub

Private Sub Form_Load()

Sql = "SELECT CODSITUACAO,DESCSITUACAO FROM SITUACAOLANCAMENTO ORDER BY DESCSITUACAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbStatus.AddItem Format(!Codsituacao, "00") & "-" & !DescSituacao
       .MoveNext
    Loop
   .Close
End With

With frmDebitoImob.grdExtrato
    nLinha = .SelectedRow
    lblCod.Caption = frmDebitoImob.txtCod.Text
    lblAno.Caption = .CellText(nLinha, 1)
    lbllanc.Caption = Left$(.CellText(nLinha, 2), 3)
    lblSeq.Caption = .CellText(nLinha, 3)
    lblParc.Caption = .CellText(nLinha, 4)
    lblCompl.Caption = .CellText(nLinha, 5)
'    If Left(.CellText(nLinha, 6), 2) <> "32" Then
        cmbStatus.Text = .CellText(nLinha, 6)
 '   Else
  '      cmbStatus.Text = "32-COMPENSAÇÃO DE CRÉD TRIB-PERMU"
  '  End If
    
    dtVencto.value = .CellText(nLinha, 7)
End With

lblUser.Caption = NomeDeLogin
'lblUser.Caption = Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8)
doData = dtVencto.value
soStatus = cmbStatus.Text
txtLog.Text = ""
CarregaTributo
Eventos ""

End Sub

Private Sub CarregaTributo()

grdTrib.Rows = 1
Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,TRIBUTO.ABREVTRIBUTO "
Sql = Sql & "FROM DEBITOTRIBUTO INNER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO "
Sql = Sql & "WHERE DEBITOTRIBUTO.CODREDUZIDO = " & Val(lblCod.Caption) & " AND ANOEXERCICIO = " & Val(lblAno.Caption) & " AND CODLANCAMENTO = " & Val(lbllanc.Caption) & " AND "
Sql = Sql & "SEQLANCAMENTO = " & Val(lblSeq.Caption) & " AND NUMPARCELA = " & lblParc.Caption & " AND CODCOMPLEMENTO = " & lblCompl.Caption
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdTrib.AddItem Format(!CodTributo, "000") & "-" & !ABREVTRIBUTO & Chr(9) & FormatNumber(!ValorTributo, 2)
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT CODTRIBUTO,ABREVTRIBUTO FROM TRIBUTO ORDER BY ABREVTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTrib.AddItem Format(!CodTributo, "000") & "-" & !ABREVTRIBUTO
       .MoveNext
    Loop
   .Close
End With

grdTrib_Click

End Sub

Private Sub Eventos(sTipo As String)

If sTipo = "" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmbTrib.Enabled = False
   cmbTrib.BackColor = frmParcela.BackColor
   txtValor.Enabled = False
   txtValor.BackColor = frmParcela.BackColor
Else
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   If sTipo = "Novo" Then
      cmbTrib.Enabled = True
      cmbTrib.BackColor = Branco
   End If
   txtValor.Enabled = True
   txtValor.BackColor = Branco
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim nLastSeq As Integer

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM LOGPARCELA WHERE CODREDUZIDO=" & Val(lblCod.Caption) & " AND ANOEXERCICIO=" & Val(lblAno.Caption) & " AND CODLANCAMENTO=" & Val(lbllanc.Caption) & " AND SEQLANCAMENTO=" & Val(lblSeq.Caption)
Sql = Sql & " AND NUMPARCELA=" & Val(lblParc.Caption) & " AND CODCOMPLEMENTO=" & Val(lblCompl.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        nLastSeq = Val(SubNull(!maximo)) + 1
    Else
        nLastSeq = 1
    End If
   .Close
End With


If Trim(txtLog.Text) <> "" Then
    Sql = "INSERT LOGPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,DATALOG,USERID,SEQ,TEXTO) VALUES("
    Sql = Sql & Val(lblCod.Caption) & "," & Val(lblAno.Caption) & "," & Val(lbllanc.Caption) & ","
    Sql = Sql & Val(lblSeq.Caption) & "," & Val(lblParc.Caption) & "," & Val(lblCompl.Caption) & ",'" & Format(Now, sDataFormat & " hh:mm") & "'," & RetornaUsuarioID(lblUser.Caption) & "," & nLastSeq & ",'" & Left$(Mask(txtLog.Text), 5000) & "')"
'    Sql = "INSERT LOGPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,DATALOG,USUARIO,SEQ,TEXTO) VALUES("
'    Sql = Sql & Val(lblCod.Caption) & "," & Val(lblAno.Caption) & "," & Val(lbllanc.Caption) & ","
'    Sql = Sql & Val(lblSeq.Caption) & "," & Val(lblParc.Caption) & "," & Val(lblCompl.Caption) & ",'" & Format(Now, sDataFormat & " hh:mm") & "','" & lblUser.Caption & "'," & nLastSeq & ",'" & Left$(Mask(txtLog.Text), 5000) & "')"
    cn.Execute Sql, rdExecDirect
    
'    SQL = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & VAL(LBLCOD.Caption) & "," & " AND ANOEXERCICIO=" & VAL(LBLANO.Caption) & " AND SEQLANCAMENTO="
'    Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES("
'    Sql = Sql & Val(lblCod.Caption) & "," & Val(lblAno.Caption) & "," & Val(lblLanc.Caption) & ","
'    Sql = Sql & Val(lblSeq.Caption) & "," & Val(lblParc.Caption) & "," & Val(lblCompl.Caption) & ","
'    cn.Execute Sql, rdExecDirect
End If

End Sub

Private Sub grdTrib_Click()
If grdTrib.Row > 0 Then
   cmbTrib.Text = grdTrib.TextMatrix(grdTrib.Row, 0)
   txtValor.Text = grdTrib.TextMatrix(grdTrib.Row, 1)
End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub GravaLog(sLog As String)
If txtLog.Text <> "" Then
    txtLog.Text = txtLog.Text & vbCrLf & sLog
Else
    txtLog.Text = txtLog.Text & sLog
End If
End Sub
