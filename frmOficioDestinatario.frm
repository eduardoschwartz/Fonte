VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmOficioDestinatario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Destinatário(s) do Ofício"
   ClientHeight    =   2775
   ClientLeft      =   1590
   ClientTop       =   4680
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstDest 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   90
      TabIndex        =   3
      Top             =   810
      Width           =   4875
   End
   Begin VB.TextBox txtBusca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "0"
      Top             =   135
      Width           =   3885
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   300
      Left            =   5040
      TabIndex        =   2
      ToolTipText     =   "Digite uma parte do nome e clique para pesquisar"
      Top             =   135
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmOficioDestinatario.frx":0000
      PICN            =   "frmOficioDestinatario.frx":001C
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
      Left            =   5040
      TabIndex        =   5
      ToolTipText     =   "Remover Destinatário"
      Top             =   1890
      Width           =   450
      _ExtentX        =   794
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
      MICON           =   "frmOficioDestinatario.frx":0176
      PICN            =   "frmOficioDestinatario.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4410
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   2340
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmOficioDestinatario.frx":0234
      PICN            =   "frmOficioDestinatario.frx":0250
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstBusca 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F8F8&
      Height          =   1785
      ItemData        =   "frmOficioDestinatario.frx":02BE
      Left            =   1080
      List            =   "frmOficioDestinatario.frx":02C5
      TabIndex        =   7
      Top             =   135
      Visible         =   0   'False
      Width           =   4350
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista do(s) centro(s) de custo selecionado(s)"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   4
      Top             =   585
      Width           =   5280
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa..:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   180
      Width           =   870
   End
End
Attribute VB_Name = "frmOficioDestinatario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()
If lstDest.ListIndex = -1 Then
    MsgBox "Selecione um ítem para remover.", vbCritical, "Atenção"
Else
    lstDest.RemoveItem (lstDest.ListIndex)
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub lstBusca_DblClick()
If lstBusca.ListIndex > -1 Then
    lstDest.AddItem (lstBusca.Text)
    lstDest.ItemData(lstDest.NewIndex) = lstBusca.ItemData(lstBusca.ListIndex)
End If
txtBusca.Text = ""
txtBusca.SetFocus
lstBusca.Visible = False

End Sub

Private Sub lstBusca_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstBusca.ListIndex > -1 Then
        lstDest.AddItem (lstBusca.Text)
        lstDest.ItemData(lstDest.NewIndex) = lstBusca.ItemData(lstBusca.ListIndex)
    End If
    txtBusca.Text = ""
    txtBusca.SetFocus
    lstBusca.Visible = False
ElseIf KeyAscii = vbKeyEscape Then
   lstBusca.Visible = False
End If

End Sub

Private Sub lstBusca_LostFocus()
lstBusca.Visible = False
End Sub

Private Sub txtBusca_GotFocus()
txtBusca.SelStart = 0
txtBusca.SelLength = Len(txtBusca.Text)
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstBusca.Clear
   If txtBusca.Text <> "" Then
      Sql = "SELECT CODIGO, DESCRICAO From centrocusto Where Ativo = 1 and descricao like '%" & Mask(txtBusca.Text) & "%' ORDER BY DESCRICAO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstBusca.AddItem Trim$(!Descricao)
                lstBusca.ItemData(lstBusca.NewIndex) = !Codigo
               .MoveNext
             Loop
             lstBusca.Visible = True
             lstBusca.ZOrder 0
             lstBusca.ListIndex = 0
             lstBusca.SetFocus
          Else
             MsgBox "Destinatário não encontrado.", vbInformation, "Atenção"
             lstBusca.Visible = False
             txtBusca.SetFocus
          End If
      End With
   End If
Else
   txtBusca.Tag = "0"
End If

End Sub

Private Sub txtBusca_LostFocus()
txtBusca.Text = UCase$(txtBusca.Text)
End Sub

Private Sub OutputUpdate()
'Dim s As String,

frmOficioInfo.txtDest.Text = ""
End Sub
