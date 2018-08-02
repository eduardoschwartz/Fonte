VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmReparcOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correção Reparcelamentos Sistema Antigo"
   ClientHeight    =   4395
   ClientLeft      =   1155
   ClientTop       =   2025
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9000
   Begin MSComctlLib.ListView lvOrigem 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lanc"
         Object.Width           =   2998
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Seq"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pc."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Co."
         Object.Width           =   811
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Vencto."
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Valor"
         Object.Width           =   2470
      EndProperty
   End
   Begin MSComctlLib.ListView lvDestino 
      Height          =   1575
      Left            =   60
      TabIndex        =   3
      Top             =   2280
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lanc"
         Object.Width           =   2998
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Seq"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pc."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Co."
         Object.Width           =   811
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Vencto."
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Valor"
         Object.Width           =   2470
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   315
      Left            =   7560
      TabIndex        =   6
      ToolTipText     =   "Gera Reparcelamento SMAR"
      Top             =   3960
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gerar"
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
      MICON           =   "frmReparcOld.frx":0000
      PICN            =   "frmReparcOld.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Exercicio...:"
      Height          =   225
      Index           =   2
      Left            =   2820
      TabIndex        =   10
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   225
      Left            =   3930
      TabIndex        =   9
      Top             =   3990
      Width           =   465
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Sequencia...:"
      Height          =   225
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Label lblSeq 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   225
      Left            =   2550
      TabIndex        =   7
      Top             =   3990
      Width           =   465
   End
   Begin VB.Label lblParc 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   3990
      Width           =   465
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Parcelas...:"
      Height          =   225
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas de Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   2010
      Width           =   1845
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas de Origem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   1845
   End
End
Attribute VB_Name = "frmReparcOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itmX As ListItem, bAchou As Boolean
Dim x As Integer, nCod As Long, nAno As Integer, nLanc As Integer
Dim nSeq As Integer, nParc As Integer, nCompl As Integer

Private Sub cmdExec_Click()
bAchou = False

For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked = True Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    MsgBox "Selecione os lancamentos de origem ", vbExclamation, "Atenção"
    Exit Sub
End If
bAchou = False
For x = 1 To lvDestino.ListItems.Count
    If lvDestino.ListItems(x).Checked = True Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    MsgBox "Selecione os lancamentos de destino ", vbExclamation, "Atenção"
    Exit Sub
End If


Sql = "INSERT REPARC2TMP (CODREDUZ,NUMSEQ,CODREDUZ2,ANOEXERC,CODLANC,CODSEQ,PARCELAS,DATAVENCTO,NUMprotocolo) VALUES("
Sql = Sql & nCod & "," & Val(lblSeq.Caption) & "," & nCod & "," & Val(lblAno.Caption) & "," & "20" & "," & Val(lblSeq.Caption)
Sql = Sql & "," & Val(lblParc.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','" & "0" & "')"
cn.Execute Sql, rdExecDirect

For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked = True Then
        nAno = lvOrigem.ListItems(x).text
        nLanc = Val(Left$(lvOrigem.ListItems(x).SubItems(1), 3))
        nSeq = lvOrigem.ListItems(x).SubItems(2)
        nParc = lvOrigem.ListItems(x).SubItems(3)
        nComp = lvOrigem.ListItems(x).SubItems(4)
        Sql = "INSERT REPARCTMP (CODREDUZD,ANOEXERCD,CODLANCD,CODSEQD,CODREDUZO,ANOEXERCO,CODLANCO,CODSEQO,NUMPARCO,CODCOMPLO,SSTATUS,CODSIT) VALUES("
        Sql = Sql & nCod & "," & Val(lblAno.Caption) & "," & "20" & "," & Val(lblSeq.Caption) & "," & nCod & "," & nAno & ","
        Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & "24" & "," & "0" & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next
MsgBox "Executado"
Unload Me
End Sub

Private Sub Form_Load()

With frmDebitoImob.grdExtrato
    nCod = frmDebitoImob.txtCod.text
    For x = 1 To .Rows
        nAno = .CellText(x, 1)
        nLanc = Val(Left$(.CellText(x, 2), 3))
        nSeq = .CellText(x, 3)
        nParc = .CellText(x, 4)
        nComp = .CellText(x, 5)
        
        If nLanc <> 20 Then
            Set itmX = lvOrigem.ListItems.Add(, Format(nCod, "000000") & nAno & Format(nLanc, "00") & Format(nSeq, "00") & Format(nParc, "00") & Format(nComp, "00"), nAno)
            itmX.SubItems(1) = .CellText(x, 2)
            itmX.SubItems(2) = .CellText(x, 3)
            itmX.SubItems(3) = .CellText(x, 4)
            itmX.SubItems(4) = .CellText(x, 5)
            itmX.SubItems(5) = .CellText(x, 7)
            itmX.SubItems(6) = .CellText(x, 6)
            itmX.SubItems(7) = .CellText(x, 10)
        Else
            Set itmX = lvDestino.ListItems.Add(, Format(nCod, "000000") & nAno & Format(nLanc, "00") & Format(nSeq, "00") & Format(nParc, "00") & Format(nComp, "00"), nAno)
            itmX.SubItems(1) = .CellText(x, 2)
            itmX.SubItems(2) = .CellText(x, 3)
            itmX.SubItems(3) = .CellText(x, 4)
            itmX.SubItems(4) = .CellText(x, 5)
            itmX.SubItems(5) = .CellText(x, 7)
            itmX.SubItems(6) = .CellText(x, 6)
            itmX.SubItems(7) = .CellText(x, 10)
        End If
    
    Next
End With

End Sub

Private Sub lvDestino_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = True Then
    If Val(lblParc.Caption) = 0 Then
        lblAno.Caption = Item.text
        lblSeq.Caption = Item.SubItems(2)
    End If
    lblParc.Caption = Val(lblParc.Caption) + 1
Else
    lblParc.Caption = Val(lblParc.Caption) - 1
    If Val(lblParc.Caption) = 0 Then
        lblAno.Caption = "0"
        lblSeq.Caption = "0"
    End If
End If

End Sub

