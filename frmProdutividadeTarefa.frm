VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProdutividadeTarefa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtividade - Cadastro de pontos por processo"
   ClientHeight    =   5925
   ClientLeft      =   6615
   ClientTop       =   2655
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   7395
   Begin Tributacao.jcFrames pnlObs 
      Height          =   3705
      Left            =   405
      Top             =   1080
      Visible         =   0   'False
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   6535
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Observação da Tarefa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtObsTmp 
         Height          =   2715
         Left            =   45
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   450
         Width           =   6495
      End
      Begin prjChameleon.chameleonButton cmdGravarObs 
         Height          =   315
         Left            =   5400
         TabIndex        =   15
         ToolTipText     =   "Gravar os Dados da Observação"
         Top             =   3285
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProdutividadeTarefa.frx":0000
         PICN            =   "frmProdutividadeTarefa.frx":001C
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
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   585
      Locked          =   -1  'True
      MaxLength       =   1000
      TabIndex        =   12
      Top             =   5040
      Width           =   6720
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4155
      Left            =   30
      TabIndex        =   3
      Top             =   810
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7329
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
         Text            =   "Item"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descricao"
         Object.Width           =   8468
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qtde"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Obs"
         Object.Width           =   0
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   345
      Left            =   6150
      TabIndex        =   8
      ToolTipText     =   "Gravar os Dados"
      Top             =   5460
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      MICON           =   "frmProdutividadeTarefa.frx":03C1
      PICN            =   "frmProdutividadeTarefa.frx":03DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdQtde 
      Height          =   345
      Left            =   4830
      TabIndex        =   10
      ToolTipText     =   "Alterar a quantidade do item selecionado"
      Top             =   5460
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Qtde"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmProdutividadeTarefa.frx":0782
      PICN            =   "frmProdutividadeTarefa.frx":079E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmbObs 
      Height          =   345
      Left            =   3510
      TabIndex        =   11
      ToolTipText     =   "Observação do item selecionado"
      Top             =   5460
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Observ."
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmProdutividadeTarefa.frx":087D
      PICN            =   "frmProdutividadeTarefa.frx":0899
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
      Caption         =   "Obs.:"
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   13
      Top             =   5085
      Width           =   390
   End
   Begin VB.Label lblNumProc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1590
      TabIndex        =   9
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label lblPontos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2040
      TabIndex        =   7
      Top             =   5550
      Width           =   465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pontos no Processo..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   90
      TabIndex        =   6
      Top             =   5550
      Width           =   1935
   End
   Begin VB.Label lblNome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1620
      TabIndex        =   5
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Fiscal......:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1485
   End
   Begin VB.Label lblDataTramite 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4860
      TabIndex        =   2
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Trâmite....:"
      Height          =   225
      Index           =   2
      Left            =   3360
      TabIndex        =   1
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo......:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1485
   End
End
Attribute VB_Name = "frmProdutividadeTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExec As Boolean, bAsk As Boolean

Private Sub cmbObs_Click()
Dim n As Variant, sOld As String
If lvMain.SelectedItem.Checked = False Then
    MsgBox "Marque primeiro o ítem que deseja observar.", vbExclamation, "Atenção"
Else
    sOld = Val(lvMain.SelectedItem.SubItems(4))
'    n = InputBox("Digite a observação para o item -> " & lvMain.SelectedItem.SubItems(1), "Nova observação", lvMain.SelectedItem.SubItems(4))
'    If n = "" And sOld <> "" Then
'        If MsgBox("Deseja remover a observação?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'            lvMain.SelectedItem.SubItems(4) = ""
'        End If
'    Else
'        lvMain.SelectedItem.SubItems(4) = Left(n, 1000)
'        txtObs.Text = lvMain.SelectedItem.SubItems(4)
'    End If
    pnlObs.Visible = True
    txtObsTmp.Text = txtObs.Text
    EnableObs (False)
    txtObsTmp.SetFocus

End If

End Sub

Private Sub EnableObs(bEnable As Boolean)
cmdGravar.Enabled = bEnable
cmdQtde.Enabled = bEnable
cmbObs.Enabled = bEnable
lvMain.Enabled = bEnable

End Sub


Private Sub cmdGravar_Click()
If Not Valida() Then Exit Sub
Grava
Unload Me
End Sub

Private Function Valida() As Boolean
'Dim x As Integer, bApuracao As Boolean, bOutros As Boolean, sItem As String

'bApuracao = False
'bOutros = False

'For x = 1 To lvMain.ListItems.Count
'    If lvMain.ListItems(x).Checked Then
'        sItem = lvMain.ListItems(x).Text
'        If sItem = "15" Or sItem = "15.1" Or sItem = "16.1" Then
'            bApuracao = True
'        End If
'        If sItem <> "15" And sItem <> "15.1" And sItem <> "16.1" Then
'            bOutros = True
'        End If
'    End If
'Next
'
'If bApuracao And bOutros And Not ProdIsBossLogin Then
'    MsgBox "Apuração fiscal não pode conter outros ítens.", vbCritical, "Erro"
'    Valida = False
'    Exit Function
'End If

Valida = True
End Function

Private Sub cmdGravarObs_Click()
pnlObs.Visible = False
txtObs.Text = txtObsTmp.Text
lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(4) = txtObs.Text
EnableObs (True)

End Sub

Private Sub cmdQtde_Click()
Dim n As Variant, nOld As Double
If lvMain.SelectedItem.Checked = False Then
    MsgBox "Marque primeiro o ítem que deseja alterar.", vbExclamation, "Atenção"
Else
    nOld = CDbl(lvMain.SelectedItem.SubItems(2))
    n = InputBox("Digite a quantidade para o item -> " & lvMain.SelectedItem.SubItems(1), "Nova quantidade", lvMain.SelectedItem.SubItems(2))
    If Not IsNumeric(n) Then
        MsgBox "Quantidade inválida", vbCritical, "Atenção"
        lvMain.SelectedItem.SubItems(2) = nOld
    Else
        If CDbl(n) = 0 Or CDbl(n) > 10000 Then
            MsgBox "Quantidade inválida", vbCritical, "Atenção"
            lvMain.SelectedItem.SubItems(2) = nOld
        Else
            lvMain.SelectedItem.SubItems(2) = CDbl(n)
        End If
    End If
    Totaliza
End If
End Sub

Private Sub Form_Load()
Centraliza Me
bAsk = False
Limpa
bExec = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim nResult As Long
If cmdGravar.Enabled And bAsk Then
    nResult = MsgBox("Deseja salvar as alterações?", vbQuestion + vbYesNoCancel, "Atenção")
    If nResult = vbYes Then
        Grava
    ElseIf nResult = vbCancel Then
        Cancel = 1
    End If
End If
End Sub

Private Sub lvMain_Click()
txtObs.Text = lvMain.SelectedItem.SubItems(4)
End Sub

Private Sub lvMain_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Not bExec Then Exit Sub
Item.Selected = True
If Item.Text = "13" And Not ProdIsBossLogin() Then
    Item.Checked = Not Item.Checked
End If
bAsk = True
cmdGravar.Enabled = True
Totaliza
End Sub

Public Sub ProdutCarregaLista()
Dim RdoAux As rdoResultset, Sql As String, itmX As ListItem

Sql = "select * from vwprodutividade where dataini<='" & Format(lblDataTramite.Caption, "mm/dd/yyyy") & "' and "
Sql = Sql & " datafim>='" & Format(lblDataTramite.Caption, "mm/dd/yyyy") & "' "
'Sql = Sql & " and item not in ('01','14','15','22')"
Sql = Sql & " and item not in ('01','22')"
Sql = Sql & " order by item"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, , !Item)
        itmX.SubItems(1) = !DESCRICAO
        itmX.SubItems(2) = 1
        itmX.SubItems(3) = FormatNumber(!valor, 2)
       .MoveNext
    Loop
   .Close
End With
cmdGravar.Enabled = False

End Sub

Private Sub Limpa()
Dim itmX As ListItem, z As Long

lblDataTramite.Caption = ""
lblNome.Caption = ""
lblPontos.Caption = "0,00"
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)

End Sub

Private Sub Totaliza()
Dim x As Integer, nTotal As Double

nTotal = 0
For x = 1 To lvMain.ListItems.Count
    If lvMain.ListItems(x).Checked Then
        nTotal = nTotal + (CDbl(lvMain.ListItems(x).SubItems(2)) * CDbl(lvMain.ListItems(x).SubItems(3)))
    End If
Next

lblPontos.Caption = FormatNumber(nTotal, 2)

End Sub

Private Sub Grava()
Dim RdoAux As rdoResultset, Sql As String, sNumProcesso As String, sItem As String
Dim nCodFiscal As Integer, nSeq As Integer, nRow As Integer, nAno As Integer, nNumProc As Long
Dim nQtde As Double, nValor As Double, sObs As String

If MsgBox("Deseja gravar as alterações?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Sql = "select codigo from produtividadefiscal where nome='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nCodFiscal = RdoAux!Codigo
RdoAux.Close

Sql = "SELECT max(seq) as maximo from produtividadetarefa where data='" & Format(Now, "mm/dd/yyyy") & "' "
Sql = Sql & " and fiscal=" & nCodFiscal
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!MAXIMO) Then
    nSeq = 1
Else
    nSeq = RdoAux!MAXIMO + 1
End If
RdoAux.Close

sNumProcesso = lblNumProc.Caption
nAno = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
nNumProc = Val(Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 1))

For nRow = 1 To lvMain.ListItems.Count
    If lvMain.ListItems(nRow).Checked Then
        sItem = lvMain.ListItems(nRow).Text
        nQtde = CDbl(lvMain.ListItems(nRow).SubItems(2))
        nValor = CDbl(lvMain.ListItems(nRow).SubItems(3))
        sObs = lvMain.ListItems(nRow).SubItems(4)
        
        Sql = "insert produtividadetarefa(data,fiscal,seq,item,qtde,valor,ano,numero,processo,obs) values('"
        Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & nCodFiscal & "," & nSeq & ",'" & sItem & "',"
        Sql = Sql & Virg2Ponto(CStr(nQtde)) & "," & Virg2Ponto(CStr(nValor)) & "," & nAno & "," & nNumProc & ",'" & sNumProcesso & "','" & Mask(sObs) & "')"
        cn.Execute Sql, rdExecDirect
    End If
Next
bAsk = False
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtObs.Text = lvMain.SelectedItem.SubItems(4)
End Sub
