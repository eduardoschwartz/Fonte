VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmMovEconomico 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento Econômico"
   ClientHeight    =   4845
   ClientLeft      =   5700
   ClientTop       =   3840
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   6585
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Valor do Imposto"
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   10
      Top             =   3735
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Valor Faturado"
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   4005
      Width           =   1500
   End
   Begin VB.ComboBox cmbAtiv 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3045
      Width           =   5175
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2535
      MaxLength       =   15
      TabIndex        =   7
      Top             =   3885
      Width           =   1320
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5145
      TabIndex        =   5
      ToolTipText     =   "Sair da Tela"
      Top             =   4380
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmMovEconomico.frx":0000
      PICN            =   "frmMovEconomico.frx":001C
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
      Left            =   3870
      TabIndex        =   6
      ToolTipText     =   "Gravar Movimento Econômico"
      Top             =   4380
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmMovEconomico.frx":008A
      PICN            =   "frmMovEconomico.frx":00A6
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
      Height          =   255
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   2
      Top             =   90
      Width           =   1200
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   4785
      MaxLength       =   4
      TabIndex        =   1
      Top             =   90
      Width           =   945
   End
   Begin MSComctlLib.ListView lvLanc 
      Height          =   1860
      Left            =   15
      TabIndex        =   0
      Top             =   480
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   3281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lancamento"
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Seq"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Parc"
         Object.Width           =   1237
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Compl"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Data Vencto."
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de ISS...:"
      Height          =   240
      Index           =   2
      Left            =   195
      TabIndex        =   20
      Top             =   2775
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Atividades.....:"
      Height          =   240
      Index           =   3
      Left            =   195
      TabIndex        =   19
      Top             =   3105
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota.........:"
      Height          =   240
      Index           =   4
      Left            =   195
      TabIndex        =   18
      Top             =   3420
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor..:"
      Height          =   240
      Index           =   5
      Left            =   1965
      TabIndex        =   17
      Top             =   3915
      Width           =   615
   End
   Begin VB.Label lblTipo 
      BackStyle       =   0  'Transparent
      Caption         =   "ISS ESTIMADO"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1335
      TabIndex        =   16
      Top             =   2775
      Width           =   2760
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Razão Social.:"
      Height          =   240
      Index           =   6
      Left            =   195
      TabIndex        =   15
      Top             =   2460
      Width           =   1065
   End
   Begin VB.Label lblRazao 
      BackStyle       =   0  'Transparent
      Caption         =   "TESTE"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1335
      TabIndex        =   14
      Top             =   2460
      Width           =   2760
   End
   Begin VB.Label lblAliq 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1335
      TabIndex        =   13
      Top             =   3435
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total..:"
      Height          =   240
      Index           =   7
      Left            =   4410
      TabIndex        =   12
      Top             =   3915
      Width           =   615
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5025
      TabIndex        =   11
      Top             =   3930
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código da Empresa..:"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano do Exercício..:"
      Height          =   240
      Index           =   1
      Left            =   3345
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "frmMovEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim nValorUFIR As Double

Private Sub cmbAtiv_Click()
Dim x As Integer, Y As Integer
If cmbAtiv.ListIndex = -1 Then Exit Sub



If lvLanc.ListItems.Count = 0 Then
   MsgBox "Não existem lancamentos.", vbCritical, "atenção"
   cmbAtiv.ListIndex = -1
   txtAno.SetFocus
   Exit Sub
End If


For x = 1 To lvLanc.ListItems.Count
   If lvLanc.ListItems(x).Checked = True Then
      Y = x
      nCount = nCount + 1
      nSeq = lvLanc.ListItems(x).SubItems(1)
      nNumParc = lvLanc.ListItems(x).SubItems(2)
      nCompl = lvLanc.ListItems(x).SubItems(3)
      sDataVencto = lvLanc.ListItems(x).SubItems(4)
      If nCount = 2 Then Exit For
   End If
Next
bCompl = False
If nCount = 0 Then
   MsgBox "Selecione um lançamento.", vbExclamation, "Atenção"
   Exit Sub
ElseIf nCount > 1 Then
   MsgBox "Selecione apenas um lançamento.", vbExclamation, "Atenção"
   Exit Sub
End If

lblAliq.Caption = "0,00"
lblAliq.Caption = RetornaAliquotaISS(cmbAtiv.ItemData(cmbAtiv.ListIndex), lvLanc.ListItems(Y).SubItems(4))
'If Val(Left$(lblTipo, 2)) = 3 Then 'EST
'    Sql = "SELECT ALIQUOTA FROM TABELAISS WHERE TIPOISS=12 AND CODIGOATIV=" & cmbAtiv.ItemData(cmbAtiv.ListIndex)
'Else
'    Sql = "SELECT ALIQUOTA FROM TABELAISS WHERE TIPOISS=13 AND CODIGOATIV=" & cmbAtiv.ItemData(cmbAtiv.ListIndex)
'End If
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'If RdoAux.RowCount > 0 Then
'   lblAliq.Caption = FormatNumber(RdoAux!Aliquota, 2)
'End If
'RdoAux.Close

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, nCount As Integer
Dim s As String, nNumParc As Integer, nSeq As Integer, nCompl As Integer, sDataVencto As String
Dim sDataBase As String, nPos As Integer, nPagina As Integer, sTypeBook As String, nLivro As Integer
Dim nSomaTributo As Double, bCompl As Boolean, t As String

nCount = 0
For x = 1 To lvLanc.ListItems.Count
   If lvLanc.ListItems(x).Checked = True Then
      nCount = nCount + 1
      nSeq = lvLanc.ListItems(x).SubItems(1)
      nNumParc = lvLanc.ListItems(x).SubItems(2)
      nCompl = lvLanc.ListItems(x).SubItems(3)
      sDataVencto = lvLanc.ListItems(x).SubItems(4)
      If nCount = 2 Then Exit For
   End If
Next
bCompl = False
If nCount = 0 Then
   MsgBox "Selecione um lançamento.", vbExclamation, "Atenção"
   Exit Sub
ElseIf nCount > 1 Then
   MsgBox "Selecione apenas um lançamento.", vbExclamation, "Atenção"
   Exit Sub
End If

If CDbl(lblTotal.Caption) = 0 Then
   MsgBox "Faltando valor do movimento.", vbExclamation, "Atenção"
   Exit Sub
End If

Ocupado
Sql = "SELECT CODREDUZIDO,DATADEBASE FROM DEBITOPARCELA WHERE "
Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
Sql = Sql & "CODLANCAMENTO=" & Val(Left$(lblTipo, 2)) & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
   sDataBase = Format(!DATADEBASE, "dd/mm/yyyy")
End With

If Val(Left$(lblTipo, 2)) = 3 Then 'EST sempre tem complemento
   Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE "
   Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
   Sql = Sql & "CODLANCAMENTO=3 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
      nCompl = !maximo + 1
     .Close
   End With
End If

If Val(Left$(lblTipo, 2)) = 5 Then 'VAR só tem complemento se for > 0
   nSomaTributo = 0
   Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO WHERE "
   Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
   Sql = Sql & "CODLANCAMENTO=5 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODTRIBUTO<>3"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
      Do Until .EOF
         nSomaTributo = nSomaTributo + !ValorTributo
        .MoveNext
      Loop
     .Close
   End With
   If nSomaTributo > 0 Then
        Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE "
        Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
        Sql = Sql & "CODLANCAMENTO=5 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
           nCompl = !maximo + 1
          .Close
        End With
   Else
        Sql = "SELECT STATUSLANC FROM DEBITOPARCELA WHERE "
        Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
        Sql = Sql & "CODLANCAMENTO=5 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
           If !statuslanc = 4 Or !statuslanc = 2 Or !statuslanc = 14 Then
                Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE "
                Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
                Sql = Sql & "CODLANCAMENTO=5 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                   bCompl = True
                   nCompl = !maximo + 1
                  .Close
                End With
           Else
                bCompl = False
                nCompl = 0
           End If
          .Close
        End With
   End If
End If

s = "Inscrição: " & txtCod.Text & vbCrLf
s = s & "Razão: " & lblRazao.Caption & vbCrLf
s = s & "Exercício: " & txtAno.Text & vbCrLf
s = s & "Lançamento: " & lblTipo.Caption & vbCrLf
s = s & "Sequência: " & nSeq & vbCrLf
s = s & "Parcela: " & nNumParc & vbCrLf
s = s & "Complemento: " & nCompl & vbCrLf
s = s & "Vencimento: " & sDataVencto & vbCrLf
s = s & "Valor: " & lblTotal.Caption
t = "Empresa: " & txtCod.Text & "-" & lblRazao.Caption & " Ano: " & txtAno.Text
t = t & " Lançamento: " & lblTipo.Caption & " Seq: " & nSeq & " Parcela: " & nNumParc
t = t & " Compl: " & nCompl & " Valor: " & lblTotal.Caption
Log Form, Me.Caption, Inclusão, t
If MsgBox(s, vbYesNo + vbQuestion, "Confirmação de Movimento") = vbYes Then
    If Val(Left$(lblTipo, 2)) = 3 Then 'ESTIMADO
        'GRAVA LANCAMENTO NA TABELA DEBITOPARCELA
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
'        Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,USUARIO) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(Left$(lblTipo, 2)) & ","
'        Sql = Sql & nSeq & "," & nNumParc & "," & nCompl & "," & "3" & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "','" & Left$(NomeDeLogin, 25) & "')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
        Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(Left$(lblTipo, 2)) & ","
        Sql = Sql & nSeq & "," & nNumParc & "," & nCompl & "," & "3" & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute Sql, rdExecDirect
        'GRAVA LANCAMENTO NA TABELA DEBITOTRIBUTO
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
        Sql = Sql & "CODTRIBUTO,VALORTRIBUTO) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(Left$(lblTipo, 2)) & ","
        Sql = Sql & nSeq & "," & nNumParc & "," & nCompl & ",'" & IIf(Val(Left$(lblTipo, 2)) = 3, 12, 13) & "','" & Virg2Ponto(lblTotal.Caption) & "')"
        cn.Execute Sql, rdExecDirect
        'DIVIDA ATIVA
        If Val(txtAno.Text) < Year(Now) Then
            Sql = "SELECT NUMERO From LIVRO WHERE CODTIPO = 2 AND ANO = " & Year(Now)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                 nLivro = Val(SubNull(!Numero))
                .Close
            End With
            sTypeBook = "CODREDUZIDO > 100000 AND CODREDUZIDO < 500000"
            cn.QueryTimeout = 0
            Sql = "SELECT MAX(PAGINALIVRO) AS MAXIMO FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Val(txtAno.Text) & " AND " & sTypeBook
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                   nPagina = 1
                Else
                    If !maximo > 0 Then
                        Sql = "SELECT distinct codreduzido,ANOEXERCICIO,CODLANCAMENTO FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Year(Now) & " AND " & sTypeBook & " AND PAGINALIVRO=" & !maximo
                        'Sql = "SELECT COUNT(CODREDUZIDO) AS CONTADOR FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Val(txtAno.text) & " AND " & sTypeBook & " AND PAGINALIVRO=" & !MAXIMO
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            nPos = .RowCount
                            If nPos < 31 Then
                               nPagina = RdoAux!maximo
                            Else
                               nPagina = RdoAux!maximo + 1
                            End If
                           .Close
                        End With
                    Else
                        nPagina = 1
                    End If
                End If
               .Close
            End With
            Sql = "UPDATE DEBITOPARCELA SET NUMEROLIVRO=" & nLivro & " ,PAGINALIVRO=" & nPagina
            Sql = Sql & " ,DATAINSCRICAO='" & Format(Now, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & Val(txtCod.Text)
            Sql = Sql & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND CODLANCAMENTO=" & Val(Left$(lblTipo, 2))
            Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
            cn.Execute Sql, rdExecDirect
        End If
    Else
        
        If nSomaTributo > 0 Or bCompl Then
            'GRAVA LANCAMENTO NA TABELA DEBITOPARCELA
'            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
'            Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,USUARIO) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(Left$(lblTipo, 2)) & ","
'            Sql = Sql & nSeq & "," & nNumParc & "," & nCompl & "," & "3" & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "','" & Left$(NomeDeLogin, 25) & "')"
            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
            Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(Left$(lblTipo, 2)) & ","
            Sql = Sql & nSeq & "," & nNumParc & "," & nCompl & "," & "3" & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
            cn.Execute Sql, rdExecDirect
            'GRAVA LANCAMENTO NA TABELA DEBITOTRIBUTO
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
            Sql = Sql & "CODTRIBUTO,VALORTRIBUTO) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(Left$(lblTipo, 2)) & ","
            Sql = Sql & nSeq & "," & nNumParc & "," & nCompl & ",'" & IIf(Val(Left$(lblTipo, 2)) = 3, 12, 13) & "'," & Virg2Ponto(RemovePonto(lblTotal.Caption)) & ")"
            cn.Execute Sql, rdExecDirect
            'DIVIDA ATIVA
            If Val(txtAno.Text) < Year(Now) Then
                Sql = "SELECT NUMERO From LIVRO WHERE CODTIPO = 2 AND ANO = " & Year(Now)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                     If .RowCount > 0 Then
                         nLivro = !Numero
                     Else
                         nLivro = 0
                     End If
                    .Close
                End With
                sTypeBook = "CODREDUZIDO > 100000 AND CODREDUZIDO < 500000"
                cn.QueryTimeout = 0
                Sql = "SELECT MAX(PAGINALIVRO) AS MAXIMO FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Val(txtAno.Text) & " AND " & sTypeBook
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If IsNull(!maximo) Then
                       nPagina = 1
                    Else
                        If !maximo > 0 Then
                            Sql = "SELECT distinct codreduzido,ANOEXERCICIO,CODLANCAMENTO FROM DEBITOPARCELA WHERE ANOEXERCICIO=" & Year(Now) & " AND " & sTypeBook & " AND PAGINALIVRO=" & !maximo
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                nPos = .RowCount
                                If nPos < 31 Then
                                   nPagina = RdoAux!maximo
                                Else
                                   nPagina = RdoAux!maximo + 1
                                End If
                               .Close
                            End With
                        Else
                            nPagina = 1
                        End If
                    End If
                   .Close
                End With
                Sql = "UPDATE DEBITOPARCELA SET NUMEROLIVRO=" & nLivro & " ,PAGINALIVRO=" & nPagina
                Sql = Sql & " ,DATAINSCRICAO='" & Format(Now, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & Val(txtCod.Text)
                Sql = Sql & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND CODLANCAMENTO=" & Val(Left$(lblTipo, 2))
                Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
                cn.Execute Sql, rdExecDirect
            End If
        Else
            'GRAVA LANCAMENTO NA TABELA DEBITOTRIBUTO
            Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=" & Virg2Ponto(RemovePonto(lblTotal.Caption)) & " WHERE "
            Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
            Sql = Sql & "CODLANCAMENTO=5 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
            Sql = Sql & " AND (CODTRIBUTO=13 OR CODTRIBUTO=502)"
            cn.Execute Sql, rdExecDirect
        End If
    End If
    'GRAVA MOVIMENTO
    Sql = "DELETE FROM MOVIMENTOECONOMICO WHERE "
    Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND "
    Sql = Sql & "CODLANCAMENTO=5 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
    cn.Execute Sql, rdExecDirect
    Sql = "INSERT MOVIMENTOECONOMICO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
    Sql = Sql & "DATAMOV) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(Left$(lblTipo, 2)) & ","
    Sql = Sql & nSeq & "," & nNumParc & "," & nCompl & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
    
    Limpa
    txtCod.SetFocus
    MsgBox "Movimento Econômico efetuado com sucesso.", vbExclamation, "Atenção"
End If
Liberado
End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txtCod.SetFocus
End Sub

Private Sub Form_Load()
Limpa
Centraliza Me
End Sub

Private Sub lvLanc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = True Then
    lblTipo.Caption = lvLanc.ListItems(Item.Index).Text
End If
End Sub

Private Sub Opt_Click(Index As Integer)
txtValor.Text = ""
lblTotal.Caption = "0,00"
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Dim itmX As ListItem, i As Integer
Dim z As Long
z = SendMessage(lvLanc.hwnd, LVM_DELETEALLITEMS, 0, 0)

If KeyAscii = vbKeyReturn Then
    
      'TIPO DE ISS
       Sql = "SELECT CODLANCAMENTO "
       Sql = Sql & "From DEBITOPARCELA Where CODREDUZIDO =" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND (CODLANCAMENTO=3 OR CODLANCAMENTO=5) AND STATUSLANC<>5 "
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
          If .RowCount > 0 Then
               Select Case !CodLancamento
                    Case 2
                       sTipoIss = "F"
                    Case 3
                       sTipoIss = "E" '3
                    Case 5
                       sTipoIss = "V" '5
               End Select
          Else
               sTipoIss = "N"
          End If
         .Close
       End With
       If sTipoIss <> "E" And sTipoIss <> "V" Then
           MsgBox "Apenas empresas com ISS Estimado ou Variável podem ser selecionadas.", vbCritical, "Atenção"
           txtCod.SetFocus
       Else
           If sTipoIss = "E" Then
              lblTipo.Caption = "03 - ISS ESTIMADO "
              Opt(1).Enabled = True
           Else
              lblTipo.Caption = "05 - ISS VARIAVEL "
           End If
       End If
    
    If Val(txtAno.Text) > 1990 And Val(txtAno.Text) < Year(Now) Then
       Sql = "SELECT CODREDUZIDO,ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO, DESCREDUZ,SEQLANCAMENTO, NUMPARCELA, CODCOMPLEMENTO,DATAVENCIMENTO "
       Sql = Sql & "FROM DEBITOPARCELA INNER JOIN LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text)
       'Sql = Sql & " AND (DEBITOPARCELA.CODLANCAMENTO=3 or DEBITOPARCELA.CODLANCAMENTO=5) AND ANOEXERCICIO=" & Val(txtAno.text) & " AND STATUSLANC=3"
       Sql = Sql & " AND (DEBITOPARCELA.CODLANCAMENTO=3 or DEBITOPARCELA.CODLANCAMENTO=5) AND ANOEXERCICIO=" & Val(txtAno.Text) & " AND STATUSLANC<>5"
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
           If .RowCount > 0 Then
                i = 0
                Do Until .EOF
                    i = i + 1
                    Set itmX = lvLanc.ListItems.Add(, "C" & Format(!CodLancamento, "00") & CStr(i), Format(!CodLancamento, "00") & " - " & !descreduz)
                    itmX.SubItems(1) = Format(!SeqLancamento, "00")
                    itmX.SubItems(2) = Format(!NumParcela, "00")
                    itmX.SubItems(3) = Format(!CODCOMPLEMENTO, "00")
                    itmX.SubItems(4) = Format(!DataVencimento, "dd/mm/yyyy")
                   .MoveNext
                Loop
                nValorUFIR = RetornaUFIR(Val(txtAno.Text))
           Else
                MsgBox "Não existem débitos para este ano.", vbCritical, "Atenção"
                txtAno.SetFocus
           End If
          .Close
       End With
    Else
       MsgBox "Ano Inválido.", vbExclamation, "atenção"
       txtAno.SetFocus
    End If
Else
    If Val(txtCod.Text) = 0 Then
        MsgBox "Selecione a Empresa.", vbCritical, "Atenção"
        KeyAscii = 0
        txtCod.SetFocus
    Else
        Tweak txtAno, KeyAscii, IntegerPositive
    End If
End If
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
If Val(txtAno.Text) > 0 Then Limpa
If KeyAscii = vbKeyReturn Then
    txtCod_LostFocus
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If

End Sub

Private Sub CarregaEmpresa()
Dim nCodReduz As Long, sTipoIss As String

nCodReduz = Val(txtCod.Text)

Limpa
Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO "
Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
       If Not IsNull(!dataencerramento) Or !dataencerramento <> CDate("01/01/1900") Then
          MsgBox "Esta empresa foi encerrada em " & Format(!dataencerramento, "dd/mm/yyyy"), vbExclamation, "Atenção"
'          Exit Sub
       End If
      'suspenção
       Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & txtCod.Text
       Sql = Sql & " ORDER BY DATAEVENTO DESC"
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
           If .RowCount > 0 Then
               If !CODTIPOEVENTO = 2 Then
                   MsgBox "Esta empresa esta SUSPENSA", vbExclamation, "Atenção"
                   Exit Sub
               End If
           End If
          .Close
       End With
       
      'TIPO DE ISS
       Sql = "SELECT CODTRIBUTO,CODATIVIDADE,SEQ,QTDEISS,ValorISS "
       Sql = Sql & "From MOBILIARIOATIVIDADEISS Where CODMOBILIARIO =" & nCodReduz
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
          If .RowCount > 0 Then
               Select Case !CodTributo
                    Case 11
                       sTipoIss = "F"
                    Case 12
                       sTipoIss = "E" '3
                    Case 13
                       sTipoIss = "V" '5
               End Select
          Else
               sTipoIss = "N"
          End If
         .Close
       End With
       GoTo PULA
       If sTipoIss <> "E" And sTipoIss <> "V" Then
           MsgBox "Apenas empresas com ISS Estimado ou Variável podem ser selecionadas.", vbCritical, "Atenção"
           txtCod.SetFocus
       Else
           If sTipoIss = "E" Then
              lblTipo.Caption = "03 - ISS ESTIMADO "
              Opt(1).Enabled = True
           Else
              lblTipo.Caption = "05 - ISS VARIAVEL "
           End If
PULA:
           lblRazao.Caption = !RazaoSocial
           'CARREGA ATIVIDADE
           Sql = "SELECT MOBILIARIOATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE FROM MOBILIARIOATIVIDADEISS INNER JOIN "
           Sql = Sql & "ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE "
           Sql = Sql & "Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO = " & nCodReduz
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
           With RdoAux
               Do Until .EOF
                  cmbAtiv.AddItem !descatividade
                  cmbAtiv.ItemData(cmbAtiv.NewIndex) = !codatividade
                 .MoveNext
               Loop
           End With
           txtAno.SetFocus
       End If
    Else
        MsgBox "Empresa não cadastrada.", vbCritical, "Atenção"
        txtCod.SetFocus
    End If
End With

End Sub

Private Sub txtCod_LostFocus()
If Val(txtCod.Text) > 0 Then
    CarregaEmpresa
End If
End Sub

Private Sub Limpa()
Dim z As Long
z = SendMessage(lvLanc.hwnd, LVM_DELETEALLITEMS, 0, 0)

txtAno.Text = ""
lblRazao.Caption = ""
lblTipo.Caption = ""
cmbAtiv.Clear
lblAliq.Caption = "0,00"
lblTotal.Caption = "0,00"

End Sub

Private Sub txtValor_Change()
Dim nValorAliq As Double

If Opt(0).value = True Then
   lblTotal.Caption = txtValor.Text
Else
   nValorAliq = CDbl(lblAliq.Caption)
   If Val(txtValor.Text) > 0 Then
      lblTotal.Caption = FormatNumber(CDbl(txtValor.Text) * nValorAliq, 2)
   Else
      lblTotal.Caption = "0,00"
   End If
End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

Tweak txtValor, KeyAscii, DecimalPositive

If CDbl(lblAliq.Caption) = 0 And Opt(1).value = True Then
    MsgBox "Aliquota inválida.", vbCritical, "Atenção"
    KeyAscii = 0
End If

End Sub
