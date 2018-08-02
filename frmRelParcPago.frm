VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmRelParcPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamento mensal de parcelamentos"
   ClientHeight    =   3930
   ClientLeft      =   5820
   ClientTop       =   3075
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   5235
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   90
      TabIndex        =   6
      Top             =   3420
      Width           =   5055
      Begin VB.Label lblTotal 
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3330
         TabIndex        =   10
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label lblQtde 
         Caption         =   "0000"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1575
         TabIndex        =   9
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Valor total:"
         Height          =   195
         Index           =   1
         Left            =   2475
         TabIndex        =   8
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Qtde de processos:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmRelParcPago.frx":0000
      Left            =   630
      List            =   "frmRelParcPago.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   1410
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   1155
   End
   Begin prjChameleon.chameleonButton cmdCalc 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3960
      TabIndex        =   2
      ToolTipText     =   "Carregar parcelamentos"
      Top             =   90
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Carregar"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmRelParcPago.frx":0094
      PICN            =   "frmRelParcPago.frx":00B0
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
      Height          =   2820
      Left            =   90
      TabIndex        =   3
      Top             =   585
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   4974
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº Processo"
         Object.Width           =   2717
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contribuinte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mês...:"
      Height          =   225
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   195
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano...:"
      Height          =   225
      Index           =   2
      Left            =   2220
      TabIndex        =   4
      Top             =   195
      Width           =   705
   End
End
Attribute VB_Name = "frmRelParcPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAno_Click()
Limpa
End Sub

Private Sub cmbMes_Click()
Limpa
End Sub

Private Sub cmdCalc_Click()
Dim Sql As String, RdoAux As rdoResultset, z As Long, nValor As Double
nValor = 0
Ocupado
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)
Sql = "SELECT debitopago.codreduzido, debitopago.anoexercicio, debitopago.codlancamento, debitopago.seqlancamento, debitopago.numparcela, debitopago.codcomplemento, "
Sql = Sql & "debitopago.seqpag, debitopago.datapagamento, debitopago.datarecebimento, debitopago.valorpago, debitopago.codbanco, debitopago.codagencia,"
Sql = Sql & "debitopago.restituido, debitopago.numdocumento, debitopago.valorpagoreal, debitopago.intacto, debitopago.valortarifa, debitopago.arquivobanco, debitopago.valordif,"
Sql = Sql & "debitopago.datapagamentocalc , debitoparcela.numprocesso FROM debitopago INNER JOIN debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND debitopago.numparcela = debitoparcela.numparcela AND debitopago.codcomplemento = debitoparcela.codcomplemento "
Sql = Sql & "Where (debitopago.CodLancamento = 20) And (Year(datarecebimento) = " & Val(cmbAno.Text) & ") And (Month(datarecebimento) = " & cmbMes.ItemData(cmbMes.ListIndex) & ") And (restituido Is Null)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblQtde.Caption = .RowCount
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, , SubNull(!numprocesso))
        itmX.SubItems(1) = !CODREDUZIDO
        itmX.SubItems(2) = FormatNumber(!valorpagoreal, 2)
        nValor = nValor + !valorpagoreal
       .MoveNext
    Loop
   .Close
End With

lblTotal.Caption = FormatNumber(nValor)

Liberado

End Sub

Private Sub Limpa()
Dim z As Long
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)
lblQtde.Caption = "0000"
lblTotal.Caption = "0,00"

End Sub

Private Sub Form_Load()
Dim X As Integer
Centraliza Me

For X = 2011 To Year(Now)
    cmbAno.AddItem (CStr(X))
Next

cmbMes.ListIndex = Month(Now) - 1
cmbAno.Text = Year(Now)

End Sub
