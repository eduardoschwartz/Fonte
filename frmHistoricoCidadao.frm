VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHistoricoCidadao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico do Cidadão - Eventos Registrados no Sistema"
   ClientHeight    =   5400
   ClientLeft      =   10170
   ClientTop       =   3675
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   9960
   Begin VB.TextBox txtHist 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1095
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4260
      Width           =   9885
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   90
      Width           =   8295
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   3825
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   6747
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
         Text            =   "Data"
         Object.Width           =   2998
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Usuário"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Histórico"
         Object.Width           =   9419
      EndProperty
   End
   Begin VB.Label lblCod 
      Caption         =   "000000"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   960
      TabIndex        =   2
      Top             =   90
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Cidadão..:"
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   90
      Width           =   795
   End
End
Attribute VB_Name = "frmHistoricoCidadao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodCidadao As Long, sOldHist As String

Public Property Let nContribuinte(nCodigoContribuinte As Long)
    nCodCidadao = nCodigoContribuinte
End Property



Private Sub Form_Activate()
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem, z As Long
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)

If nCodCidadao > 0 Then
    lblCod.Caption = nCodCidadao
    Sql = "select nomecidadao from cidadao where codcidadao=" & nCodCidadao
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    
    txtNome.Text = SubNull(RdoAux!nomecidadao)
    RdoAux.Close
    
    Sql = "SELECT distinct historicocidadao.*, usuario.nomecompleto FROM historicocidadao INNER JOIN "
    Sql = Sql & "usuario ON historicocidadao.userid = usuario.Id Where Codigo=" & nCodCidadao & " order by data"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Set itmX = lvMain.ListItems.Add(, , Format(!Data, "dd/mm/yyyy hh:mm:ss"))
            itmX.SubItems(1) = SubNull(!NomeCompleto)
            itmX.SubItems(2) = !obs
           .MoveNext
        Loop
       .Close
    End With
    
    If lvMain.ListItems.Count > 0 Then lvMain_Click
    
End If

nCodCidadao = 0

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub lvMain_Click()
If lvMain.ListItems.Count = 0 Then Exit Sub
txtHist.Text = lvMain.SelectedItem.SubItems(2)
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtHist.Text = Item.SubItems(2)
End Sub
