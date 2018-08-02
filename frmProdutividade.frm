VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProdutividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Produtividade"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   7275
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   0
      Top             =   150
      Width           =   1185
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4215
      Left            =   30
      TabIndex        =   4
      Top             =   810
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   7435
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descricao"
         Object.Width           =   8997
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   1411
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   345
      Left            =   6030
      TabIndex        =   9
      ToolTipText     =   "Gravar os Dados"
      Top             =   5130
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
      MICON           =   "frmProdutividade.frx":0000
      PICN            =   "frmProdutividade.frx":001C
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
      Left            =   2820
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmProdutividade.frx":03C1
      PICN            =   "frmProdutividade.frx":03DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   8
      Top             =   5190
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
      TabIndex        =   7
      Top             =   5190
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
      TabIndex        =   6
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Fiscal......:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   5
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
      TabIndex        =   3
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Trâmite....:"
      Height          =   225
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo......:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1485
   End
End
Attribute VB_Name = "frmProdutividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExec As Boolean

Private Sub cmdCancel_Click()
If MsgBox("Deseja cancelar ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
Limpa
txtNumProc.Text = ""
txtNumProc.Locked = False
txtNumProc.BackColor = Branco
txtNumProc.SetFocus
End Sub

Private Sub cmdGravar_Click()
Grava
cmdGravar.Enabled = False
End Sub

Private Sub Form_Load()
Centraliza Me
Limpa
bExec = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim nResult As Long
If cmdGravar.Enabled Then
    nResult = MsgBox("Deseja salvar as alterações?", vbQuestion + vbYesNoCancel, "Atenção")
    If nResult = vbYes Then
        Grava
    ElseIf nResult = vbCancel Then
        Cancel = 1
    End If
End If
End Sub

Private Sub lvMain_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Not bExec Then Exit Sub
cmdGravar.Enabled = True
Totaliza
End Sub

Private Sub txtNumProc_GotFocus()
txtNumProc.SelStart = 0
txtNumProc.SelLength = Len(txtNumProc.Text)
End Sub

Private Sub txtNumProc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    txtNumProc_LostFocus
End If
End Sub

Private Sub txtNumProc_LostFocus()
Dim RdoAux As rdoResultset, Sql As String, z As Long

z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)
lblDataTramite.Caption = ""

If Trim$(txtNumProc.Text) <> "" Then
    If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) > 0 Then
        nNumProc = Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))
        nAnoProc = Val(Right$(txtNumProc.Text, 4))
        Sql = "SELECT DATAENTRADA FROM PROCESSOGTI WHERE NUMERO=" & nNumProc & " AND ANO=" & nAnoProc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Processo de não cadastrado.", vbExclamation, "Atenção"
                txtNumProc.SetFocus
                Exit Sub
            Else
                lblDataTramite.Caption = Format(!dataentrada, "dd/mm/yyyy")
                ProdutCarregaLista
                txtNumProc.Locked = True
                txtNumProc.BackColor = Me.BackColor
                lvMain.SetFocus
            End If
           .Close
        End With
    End If
End If

End Sub

Public Sub ProdutCarregaLista()
Dim RdoAux As rdoResultset, Sql As String, itmX As ListItem

Sql = "select * from vwprodutividade where dataini<='" & Format(lblDataTramite.Caption, "mm/dd/yyyy") & "' and "
Sql = Sql & " datafim>='" & Format(lblDataTramite.Caption, "mm/dd/yyyy") & "' order by item"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, , !Item)
        itmX.SubItems(1) = !DESCRICAO
        itmX.SubItems(2) = !valor

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
lblPontos.Caption = "0"
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)

End Sub

Private Sub Totaliza()
Dim x As Integer, nTotal As Integer

nTotal = 0
For x = 1 To lvMain.ListItems.Count
    If lvMain.ListItems(x).Checked Then
        nTotal = nTotal + Val(lvMain.ListItems(x).SubItems(2))
    End If
Next

lblPontos.Caption = CStr(nTotal)

End Sub

Private Sub Grava()
End Sub
