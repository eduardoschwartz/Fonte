VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmNotificaISS 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificação de ISS Eletrônico (parcelas vencidas e não pagas)"
   ClientHeight    =   4920
   ClientLeft      =   4185
   ClientTop       =   2310
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   10545
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   360
      Left            =   8730
      TabIndex        =   1
      ToolTipText     =   "Cancelar Edição"
      Top             =   4470
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Notificar débito"
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
      MICON           =   "frmNotificaISS.frx":0000
      PICN            =   "frmNotificaISS.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvTmp 
      Height          =   4365
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   7699
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1669
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   5733
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Ano"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Lanc"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Seq"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Parc"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Comp"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dt.Vencto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Documento"
         Object.Width           =   2964
      EndProperty
   End
End
Attribute VB_Name = "frmNotificaISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExec_Click()
Dim bAchou As Boolean, x As Integer
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer

nCodReduz = 0
With lvTmp
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked Then
            If nCodReduz = 0 Then
                nCodReduz = Val(.ListItems(x).Text)
            Else
                If nCodReduz <> Val(.ListItems(x).Text) Then
                    MsgBox "Selecione apenas uma empresa.", vbExclamation, "Atenção"
                    Exit Sub
                End If
            End If
        End If
    Next
End With

bAchou = False
With lvTmp
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked Then
            bAchou = True
            'Notifica
            nCodReduz = Val(.ListItems(x).Text)
            nAno = Val(.ListItems(x).SubItems(2))
            nLanc = Val(.ListItems(x).SubItems(3))
            nSeqLanc = Val(.ListItems(x).SubItems(4))
            nParc = Val(.ListItems(x).SubItems(5))
            nCompl = Val(.ListItems(x).SubItems(6))
            
            Sql = "UPDATE DEBITOPARCELA SET NOTIFICADO=1 WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND "
            Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nParc & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & nCompl
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With

If Not bAchou Then
    MsgBox "Selecione ao menos uma parcela a notificar!", vbExclamation, "Atenção"
Else
    frmReport.ShowReport "NOTIFICACAO", frmMdi.hwnd, Me.hwnd
    Le
End If

End Sub

Private Sub Form_Load()
Centraliza Me
Ocupado
Le
Liberado
End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem


lvTmp.ListItems.Clear
Sql = "SELECT * FROM vwNOTIFICACAOISS ORDER BY CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        
        Set itmX = lvTmp.ListItems.Add(, , !CODREDUZIDO)
        itmX.SubItems(1) = SubNull(!NOME)
        itmX.SubItems(2) = !AnoExercicio
        itmX.SubItems(3) = Format(!CodLancamento, "00")
        itmX.SubItems(4) = Format(!SeqLancamento, "000")
        itmX.SubItems(5) = Format(!NumParcela, "00")
        itmX.SubItems(6) = Format(!CODCOMPLEMENTO, "00")
        itmX.SubItems(7) = Format(!DataVencimento, "dd/mm/yyyy")
        itmX.SubItems(8) = !NumDocumento

        .MoveNext
    Loop
   .Close
End With



End Sub

