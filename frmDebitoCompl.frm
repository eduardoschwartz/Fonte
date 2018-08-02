VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDebitoCompl 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complemento de Pagamento"
   ClientHeight    =   4395
   ClientLeft      =   4920
   ClientTop       =   3180
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   5985
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      ToolTipText     =   "Imprimir complementos"
      Top             =   4005
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmDebitoCompl.frx":0000
      PICN            =   "frmDebitoCompl.frx":001C
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
      Left            =   1395
      TabIndex        =   2
      ToolTipText     =   "Excluir os complementos selecionados"
      Top             =   4005
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmDebitoCompl.frx":0176
      PICN            =   "frmDebitoCompl.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   3855
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   6800
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1766
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Ano"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Lanc"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Seq"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Parc"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Compl"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdAutorizar 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Autorizar os complementos selecionados"
      Top             =   4005
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Autorizar"
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
      MICON           =   "frmDebitoCompl.frx":0234
      PICN            =   "frmDebitoCompl.frx":0250
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
Attribute VB_Name = "frmDebitoCompl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAutorizar_Click()
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, X As Integer

If MsgBox("Autorizar os complementos selecionados?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

For X = 1 To lvMain.ListItems.Count
    With lvMain.ListItems(X)
        If .Checked Then
            nCodReduz = .Text
            nAno = .SubItems(1)
            nLanc = .SubItems(2)
            nSeq = .SubItems(3)
            nParc = .SubItems(4)
            nCompl = .SubItems(5)
            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=3 WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
            cn.Execute Sql, rdExecDirect
            
            Sql = "UPDATE COMPLEMENTOPAGTO SET AUTORIZADO='S', USUARIO='" & NomeDeLogin & "',DATAEVENTO='" & Format(Now, "mm/dd/yyyy") & "' "
            Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTOCP=" & nCompl
            cn.Execute Sql, rdExecDirect
            
        End If
    End With
Next

CarregaLista

End Sub

Private Sub cmdExcluir_Click()
If MsgBox("Excluir os complementos selecionados?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

For X = 1 To lvMain.ListItems.Count
    With lvMain.ListItems(X)
        If .Checked Then
            nCodReduz = .Text
            nAno = .SubItems(1)
            nLanc = .SubItems(2)
            nSeq = .SubItems(3)
            nParc = .SubItems(4)
            nCompl = .SubItems(5)
            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
            cn.Execute Sql, rdExecDirect
            
            Sql = "UPDATE COMPLEMENTOPAGTO SET AUTORIZADO='N', USUARIO='" & NomeDeLogin & "',DATAEVENTO='" & Format(Now, "mm/dd/yyyy") & "' "
            Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTOCP=" & nCompl
            cn.Execute Sql, rdExecDirect
            
        End If
    End With
Next

CarregaLista

End Sub

Private Sub cmdPrint_Click()
Dim X As Integer

sNomeArq = sPathBin & "\REPCOML1.TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1

Print #FF1, "*******************************************************"
Print #FF1, "LISTA DE COMPLEMENTOS GERADOS PELO SISTEMA TRIBUTÁRIO"
Print #FF1, "QUE ESTÃO EM ANÁLISE PARA AUTORIZAÇÃO OU CANCELAMENTO"
Print #FF1, "*******************************************************"
Print #FF1, ""
ax = FillSpace("CÓDIGO", 8) & FillSpace(" ANO", 5) & FillLeft("LANC", 5) & FillLeft(" SEQ", 6) & FillLeft("PARC", 6) & FillLeft("COMPL", 6) & FillLeft("VALOR", 11)
Print #FF1, ax
Print #FF1, "****************************************************************************"

For X = 1 To lvMain.ListItems.Count
    With lvMain.ListItems(X)
        ax = FillSpace(.Text, 8) & FillSpace(.SubItems(1), 5) & FillLeft(.SubItems(2), 5) & FillLeft(.SubItems(3), 6) & FillLeft(.SubItems(4), 6) & FillLeft(.SubItems(5), 6) & FillLeft(.SubItems(6), 11)
        Print #FF1, ax
    End With
Next

Print #FF1, ""
Print #FF1, "PMJ - REPCOML1.TXT - " & Format(Now, "dd/mm/yyyy")

Close #FF1
ret = Shell("NOTEPAD2" & " " & sNomeArq, vbNormalFocus)

Liberado
'MsgBox "Relatório disponível em " & sPathBin & "\REPORTMOB1.TXT"

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaLista
End Sub

Private Sub CarregaLista()
Dim itmX As ListItem, z As Long, Sql As String, RdoAux As rdoResultset
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
Sql = Sql & "debitoparcela.CODCOMPLEMENTO , complementopagto.valor FROM debitoparcela INNER JOIN complementopagto ON debitoparcela.codreduzido = complementopagto.codreduzido AND debitoparcela.anoexercicio = complementopagto.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = complementopagto.codlancamento AND debitoparcela.seqlancamento = complementopagto.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = complementopagto.NumParcela And debitoparcela.CODCOMPLEMENTO = complementopagto.codcomplementocp Where (debitoparcela.statuslanc = 25) order by debitoparcela.codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, "C" & Format(!CODREDUZIDO, "000000") & CStr(!AnoExercicio) & Format(!CodLancamento, "000") & Format(!SeqLancamento, "0000") & Format(!NumParcela, "000") & Format(!CODCOMPLEMENTO, "00"), Format(!CODREDUZIDO, "000000"))
        itmX.SubItems(1) = !AnoExercicio
        itmX.SubItems(2) = Format(!CodLancamento, "000")
        itmX.SubItems(3) = Format(!SeqLancamento, "0000")
        itmX.SubItems(4) = Format(!NumParcela, "000")
        itmX.SubItems(5) = Format(!CODCOMPLEMENTO, "00")
        itmX.SubItems(6) = FormatNumber(!Valor, 2)
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvMain.SortKey = ColumnHeader.Position - 1
lvMain.Sorted = True
lvMain.SortOrder = lvwAscending
End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

