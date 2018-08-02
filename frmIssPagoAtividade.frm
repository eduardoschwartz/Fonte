VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmIssPagoAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iss pago por atividade"
   ClientHeight    =   1650
   ClientLeft      =   8640
   ClientTop       =   4140
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   5355
   Begin VB.CheckBox chkAmbulantes 
      Caption         =   "Somente Ambulantes"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   765
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dtDataDe 
      Height          =   285
      Left            =   900
      TabIndex        =   2
      Top             =   180
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      Format          =   100663297
      CurrentDate     =   41232
   End
   Begin MSComCtl2.DTPicker dtDataAte 
      Height          =   285
      Left            =   3780
      TabIndex        =   3
      Top             =   180
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      Format          =   100663297
      CurrentDate     =   41232
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4050
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   1215
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmIssPagoAtividade.frx":0000
      PICN            =   "frmIssPagoAtividade.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAddAll 
      Height          =   285
      Left            =   495
      TabIndex        =   5
      ToolTipText     =   "Marcar todos"
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIssPagoAtividade.frx":008A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRemoveAll 
      Height          =   285
      Left            =   855
      TabIndex        =   6
      ToolTipText     =   "Desmarcar todos"
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIssPagoAtividade.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvCid 
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   3915
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1164
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição da Atividade"
         Object.Width           =   5292
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdImprimir 
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      ToolTipText     =   "Impressão do Protocolo de Entrada e Requerimento"
      Top             =   1215
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmIssPagoAtividade.frx":00C2
      PICN            =   "frmIssPagoAtividade.frx":00DE
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
      Caption         =   "Data Até:"
      Height          =   195
      Index           =   1
      Left            =   2790
      TabIndex        =   1
      Top             =   225
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   780
   End
End
Attribute VB_Name = "frmIssPagoAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddAll_Click()
Dim x As Integer

For x = 1 To lvCid.ListItems.Count
    lvCid.ListItems(x).Checked = True
Next

End Sub

Private Sub cmdImprimir_Click()
Dim RdoAux As rdoResultset, Sql As String, nCodReduz As Long, RdoAux2 As rdoResultset
Dim sDataIni As String, sSimples As String, sSemMov As String, sAtiv As String, x As Integer

If chkAmbulantes.value = vbChecked Then
    Ambulantes
    
Else
    frmReport.ShowReport2 "isspagoperiodo", frmMdi.hwnd, Me.hwnd
End If

End Sub

Private Sub cmdRemoveAll_Click()
Dim x As Integer

For x = 1 To lvCid.ListItems.Count
    lvCid.ListItems(x).Checked = False
Next

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Dim itmX As ListItem

dtDataDe.value = Now
dtDataAte.value = Now
Centraliza Me
'Sql = "select codatividade,descatividade from atividadeiss order by codatividade"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
 '   Do Until .EOF
 '      lstISS.AddItem !codatividade & " - " & !descatividade
 '     .MoveNext
'    Loop
 '  .Close
'End With

        
Sql = "select codatividade,descatividade from atividade order by descatividade"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Set itmX = lvCid.ListItems.Add(, "C" & Format(!codatividade, "00000"), Format(!codatividade, "00000"))
       itmX.SubItems(1) = !descatividade

       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub lvCid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvCid.SortKey = ColumnHeader.Position - 1
lvCid.Sorted = True
lvCid.SortOrder = lvwAscending
End Sub

Private Sub Ambulantes()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, RdoAux2 As rdoResultset, Sql2 As String

Sql = "delete from relpagamentoambulantes where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Ocupado
Sql = "select vwrelatorioatividadetl.codigomob,vwrelatorioatividadetl.razaosocial,vwrelatorioatividadetl.codatividade,vwrelatorioatividadetl.descatividade,vwrelatorioatividadetl.ativextenso from vwrelatorioatividadetl where vwrelatorioatividadetl.codatividade>40000"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        DoEvents
        nCodReduz = !codigomob
        Sql2 = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela, "
        Sql2 = Sql2 & "debitoparcela.codcomplemento, debitoparcela.statuslanc, debitoparcela.datavencimento, SUM(debitotributo.valortributo) AS Total, lancamento.descreduz "
        Sql2 = Sql2 & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
        Sql2 = Sql2 & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
        Sql2 = Sql2 & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
        Sql2 = Sql2 & "lancamento ON debitoparcela.codlancamento = lancamento.codlancamento GROUP BY debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
        Sql2 = Sql2 & "debitoparcela.CODCOMPLEMENTO , debitoparcela.statuslanc, debitoparcela.DataVencimento, lancamento.descreduz HAVING (debitoparcela.codreduzido = " & nCodReduz & ") AND (debitoparcela.numparcela > 0) AND (debitoparcela.statuslanc = 3) AND "
        Sql2 = Sql2 & "(debitoparcela.datavencimento < CONVERT(DATETIME,'" & Format(Now, "mm/dd/yyyy") & "', 102)) "
        Set RdoAux2 = cn.OpenResultset(Sql2, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            Do Until RdoAux2.EOF
                Sql = "insert relpagamentoambulantes(usuario,codigo,razao,ano,lanc,seq,parc,compl,pago,desclanc,valor) values('"
                Sql = Sql & NomeDeLogin & "'," & !codigomob & ",'" & Left(Mask(!RazaoSocial), 50) & "'," & RdoAux2!AnoExercicio & "," & RdoAux2!CodLancamento & ","
                Sql = Sql & RdoAux2!SeqLancamento & "," & RdoAux2!NumParcela & "," & RdoAux2!CODCOMPLEMENTO & "," & 0 & ",'" & RdoAux2!descreduz & "'," & Virg2Ponto(RdoAux2!Total) & ")"
                cn.Execute Sql, rdExecDirect
                RdoAux2.MoveNext
            Loop
        Else
            Sql = "insert relpagamentoambulantes(usuario,codigo,razao,ano,lanc,seq,parc,compl,pago,desclanc,valor) values('"
            Sql = Sql & NomeDeLogin & "'," & !codigomob & ",'" & Left(Mask(!RazaoSocial), 50) & "'," & "0,0,0,0,0,1,'',0)"
            cn.Execute Sql, rdExecDirect
        End If
       .MoveNext
    Loop
   .Close
End With

Liberado

frmReport.ShowReport2 "RELPAGAMENTOS", frmMdi.hwnd, Me.hwnd

Sql = "delete from relpagamentoambulantes where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub
