VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmPagamentoDuplicidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamento em Duplicidade"
   ClientHeight    =   5205
   ClientLeft      =   6615
   ClientTop       =   3855
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   10080
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   4920
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1335
      TabIndex        =   1
      Top             =   180
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmPagamentoDuplicidade.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   1335
      TabIndex        =   2
      Top             =   525
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmPagamentoDuplicidade.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   2490
      TabIndex        =   5
      ToolTipText     =   "Consultar as empresas"
      Top             =   540
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Pesquisar"
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
      MICON           =   "frmPagamentoDuplicidade.frx":0038
      PICN            =   "frmPagamentoDuplicidade.frx":0054
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
      Left            =   60
      TabIndex        =   6
      Top             =   930
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   6800
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1322
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ano"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lc"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sq"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Parc"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cp"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Dt.Pag"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Valor"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Documento"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Arquivo"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   300
      TabIndex        =   4
      Top             =   570
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início.:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   225
      Width           =   1005
   End
End
Attribute VB_Name = "frmPagamentoDuplicidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdConsultar_Click()
If Not IsDate(mskDataIni.Text) Then
    MsgBox "Data de Inicio inválido", vbExclamation, "atenção"
    Exit Sub
End If

If Not IsDate(mskDataFim.Text) Then
    MsgBox "Data de Fim inválido", vbExclamation, "atenção"
    Exit Sub
End If

If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    MsgBox "Data de Inicio tem que ser maior que data de termino", vbExclamation, "atenção"
    Exit Sub
End If

Ocupado
Pesquisar
Liberado

End Sub

Private Sub PrintExcel()

If lvMain.ListItems.Count = 0 Then Exit Sub

Dim x As Long, y As Long, ax As String, Scr_hdc As Long, z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
PBar.value = 0
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For y = 1 To lvMain.ColumnHeaders.Count
    ax = ax & RemoveSpace(lvMain.ColumnHeaders(y).Text) & " char(255), "
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "[Table1$]", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For x = 1 To lvMain.ListItems.Count
    Rs.AddNew
    nCont = 0
    Rs.Fields(nCont).value = lvMain.ListItems(x).Text
    nCont = nCont + 1
    For y = 2 To lvMain.ColumnHeaders.Count
         
         Rs.Fields(nCont).value = lvMain.ListItems(x).SubItems(y - 1)
         nCont = nCont + 1
    
        
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)


End Sub

Private Sub Form_Load()
Centraliza Me

End Sub

Private Sub Pesquisar()
Dim RdoAux As rdoResultset, Sql As String, itmX As ListItem, RdoAux2 As rdoResultset
Dim nTot As Long, nPos As Long
PBar.value = 0
nPos = 1
Ocupado
Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, seqpag, datapagamento, datarecebimento, valorpago, codbanco, "
Sql = Sql & "CodAgencia , restituido, NumDocumento, valorpagoreal, intacto, ValorTarifa, arquivobanco, valordif, datapagamentocalc, dataintegracao "
Sql = Sql & "from debitopago WHERE (seqpag > 0) AND (datapagamento BETWEEN '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' AND '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "')"
Sql = Sql & " order by codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        CallPb nPos, nTot
        Sql = "SELECT * from debitopago WHERE codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and "
        Sql = Sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                Set itmX = lvMain.ListItems.Add(, , !CODREDUZIDO)
                itmX.SubItems(1) = !AnoExercicio
                itmX.SubItems(2) = !CodLancamento
                itmX.SubItems(3) = !SeqLancamento
                itmX.SubItems(4) = !NumParcela
                itmX.SubItems(5) = !CODCOMPLEMENTO
                itmX.SubItems(6) = Format(!DataPagamento, "dd/mm/yyyy")
                itmX.SubItems(7) = !valorpagoreal
                itmX.SubItems(8) = !NumDocumento
                itmX.SubItems(9) = SubNull(!arquivobanco)
               .MoveNext
            Loop
           .Close
        End With
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
Liberado
If lvMain.ListItems.Count > 0 Then
    PrintExcel
Else
    MsgBox "Nada encontrado.", vbInformation, "Atenção"
End If

End Sub


Private Sub CallPb(nPos As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents
If nTotal = 0 Then Exit Sub
If ((nPos * 100) / nTotal) <= 100 Then
   PBar.value = (nPos * 100) / nTotal
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

