VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmVVDeclarado 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valor venal declarado"
   ClientHeight    =   4170
   ClientLeft      =   5205
   ClientTop       =   3720
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   5505
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1005
      Left            =   60
      TabIndex        =   11
      Top             =   2670
      Width           =   5385
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3870
         MaxLength       =   15
         TabIndex        =   2
         Top             =   210
         Width           =   1305
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   3
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Atrib.......:"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   15
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Processo...:"
         Height          =   195
         Index           =   3
         Left            =   2700
         TabIndex        =   14
         Top             =   255
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venal.........:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   255
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdVV 
      Height          =   2565
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   4524
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   192
      ForeColorSel    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "^Código         |Processo            |>Valor Venal          |^Data                  "
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      ToolTipText     =   "Sair da Tela"
      Top             =   3765
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
      MICON           =   "frmVVDeclarado.frx":0000
      PICN            =   "frmVVDeclarado.frx":001C
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
      TabIndex        =   6
      ToolTipText     =   "Cancelar Edição"
      Top             =   3765
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
      MICON           =   "frmVVDeclarado.frx":008A
      PICN            =   "frmVVDeclarado.frx":00A6
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
      TabIndex        =   7
      ToolTipText     =   "Novo Registro"
      Top             =   3765
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
      MICON           =   "frmVVDeclarado.frx":0200
      PICN            =   "frmVVDeclarado.frx":021C
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
      TabIndex        =   8
      ToolTipText     =   "Editar Registro"
      Top             =   3765
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
      MICON           =   "frmVVDeclarado.frx":0376
      PICN            =   "frmVVDeclarado.frx":0392
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
      TabIndex        =   9
      ToolTipText     =   "Excluir Registro"
      Top             =   3765
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
      MICON           =   "frmVVDeclarado.frx":04EC
      PICN            =   "frmVVDeclarado.frx":0508
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
      Left            =   3270
      TabIndex        =   10
      ToolTipText     =   "Gravar os Dados"
      Top             =   3765
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
      MICON           =   "frmVVDeclarado.frx":05AA
      PICN            =   "frmVVDeclarado.frx":05C6
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
Attribute VB_Name = "frmVVDeclarado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String

Private Sub cmdAlterar_Click()
    Evento = "Alterar"
    Eventos "INCLUIR"
    txtCod.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()
Dim Sql As String, RdoAux As rdoResultset, nSeq As Integer, sHist As String
On Error GoTo Erro
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
        
    If MsgBox("Excluir este registro ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       
        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & Val(txtCod.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!maximo) Then
                nSeq = 1
            Else
                nSeq = !maximo + 1
            End If
           .Close
        End With
       
       
       Sql = "DELETE FROM VVDECLARADO WHERE CODREDUZIDO=" & Val(txtCod.Text)
       cn.Execute Sql, rdExecDirect
        sHist = "Excluido valor venal declarado no valor de " & FormatNumber(txtValor.Text, 2) & ",processo: " & txtNumProc.Text & ", na data de: " & txtData.Text & ", pelo usuário: " & NomeDeLogin
'        Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'        Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "','" & "GTI" & "','" & Format(Now, "mm/dd/yyyy") & "')"
        Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
        Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "'," & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
       
       Limpa
       CarregaLista
       grdVV_RowColChange
    End If
Exit Sub
Erro:
MsgBox Err.Description

End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset, x As Integer, nSeq As Integer, sHist As String

Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Código não cadastrado.", vbExclamation, "Atenção"
        RdoAux.Close
        Exit Sub
    End If
   .Close
End With

If Val(txtValor.Text) = 0 Then
    MsgBox "Digite o Valor.", vbExclamation, "Atenção"
    Exit Sub
End If

If txtNumProc.Text = "" Then
    MsgBox "Digite o processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Evento = "Novo" Then
    For x = 1 To grdVV.Rows - 1
        If Val(txtCod.Text) = Val(grdVV.TextMatrix(x, 0)) Then
            MsgBox "Código já cadastrado.", vbExclamation, "Atenção"
            Exit Sub
        End If
    Next
End If

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With

If Evento = "Novo" Then
    Sql = "INSERT VVDECLARADO(CODREDUZIDO,NUMPROC,VALOR,DATA) VALUES(" & Val(txtCod.Text) & ",'" & txtNumProc.Text & "'," & Virg2Ponto(RemovePonto(txtValor.Text)) & ",'" & Format(txtData.Text, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
    
    sHist = "Inserido valor venal declarado no valor de " & FormatNumber(txtValor.Text, 2) & ",processo: " & txtNumProc.Text & ", na data de: " & txtData.Text & ", pelo usuário: " & NomeDeLogin
'    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'    Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "','" & "GTI" & "','" & Format(Now, "mm/dd/yyyy") & "')"
    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
    Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "'," & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
    
    grdVV.AddItem Format(txtCod.Text, "000000") & Chr(9) & txtNumProc.Text & Chr(9) & FormatNumber(txtValor.Text, 2) & Chr(9) & Format(Now, "dd/mm/yyyy")
Else
    Sql = "UPDATE VVDECLARADO SET NUMPROC='" & txtNumProc.Text & "',VALOR=" & Virg2Ponto(RemovePonto(txtValor.Text)) & ",DATA='" & Format(Now, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
    grdVV.TextMatrix(grdVV.Row, 1) = txtNumProc.Text
    grdVV.TextMatrix(grdVV.Row, 2) = FormatNumber(txtValor.Text, 2)
    grdVV.TextMatrix(grdVV.Row, 3) = Format(Now, "dd/mm/yyyy")
    sHist = "Alteração no valor venal declarado no valor de " & FormatNumber(txtValor.Text, 2) & ",processo: " & txtNumProc.Text & ", na data de: " & txtData.Text & ", pelo usuário: " & NomeDeLogin
'    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'    Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "','" & "GTI" & "','" & Format(Now, "mm/dd/yyyy") & "')"
    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
    Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "'," & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect

End If

Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
    txtData.Text = Format(Now, "dd/mm/yyyy")
    txtCod.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
CarregaLista
Eventos "INICIAR"
End Sub

Private Sub grdVV_RowColChange()
If grdVV.Rows = 1 Then Exit Sub
If grdVV.Row > 0 Then
    txtCod.Text = grdVV.TextMatrix(grdVV.Row, 0)
    txtNumProc.Text = grdVV.TextMatrix(grdVV.Row, 1)
    txtValor.Text = grdVV.TextMatrix(grdVV.Row, 2)
    txtData.Text = grdVV.TextMatrix(grdVV.Row, 3)
End If

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
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
   For Each Ct In frmVVDeclarado
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   grdVV.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmVVDeclarado
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   grdVV.Enabled = False
   
End If

End Sub

Private Sub Limpa()
txtCod.Text = ""
txtValor.Text = ""
txtData.Text = ""
txtNumProc.Text = ""
End Sub

Private Sub txtNumProc_LostFocus()
Dim sValidaProc As String, nNumproc As Long, nAnoproc As Long
If Trim(txtNumProc.Text) = "" Then Exit Sub
On Error Resume Next
nNumproc = Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 1))
nAnoproc = Val(Right$(txtNumProc.Text, 4))


sValidaProc = ValidaProcesso(txtNumProc.Text)
If Left$(sValidaProc, 24) = "Nº do processo inválido." Then
    MsgBox sValidaProc, vbCritical, "Atenção"
    txtNumProc.SetFocus
    Exit Sub
ElseIf Right$(nNumproc, 1) <> RetornaDVProcesso(Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))) Then
    MsgBox "Número de Processo inválido", vbExclamation, "Atenção"
    txtNumProc.SetFocus
    Exit Sub
ElseIf Left$(sValidaProc, 24) = "Processo não Cadastrado." Then
    MsgBox sValidaProc, vbCritical, "Atenção"
    txtNumProc.SetFocus
    Exit Sub
End If

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset
grdVV.Rows = 1
Sql = "SELECT * FROM VVDECLARADO ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        grdVV.AddItem Format(!CODREDUZIDO, "000000") & Chr(9) & !NumProc & Chr(9) & FormatNumber(!Valor, 2) & Chr(9) & Format(!Data, "dd/mm/yyyy")
       .MoveNext
    Loop
   .Close
End With

If grdVV.Rows > 1 Then grdVV_RowColChange

End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub
