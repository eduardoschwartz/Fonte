VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEmissao2ViaEspecial 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de 2ª Via Especial"
   ClientHeight    =   5985
   ClientLeft      =   6030
   ClientTop       =   2775
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7665
   Begin VB.TextBox txtNumDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      MaxLength       =   9
      TabIndex        =   6
      Top             =   2700
      Width           =   1275
   End
   Begin VB.ComboBox cmbLanc 
      Height          =   315
      ItemData        =   "frmEmissao2ViaEspecial.frx":0000
      Left            =   1485
      List            =   "frmEmissao2ViaEspecial.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3060
      Width           =   6015
   End
   Begin VB.ComboBox cmbProc 
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3405
      Width           =   1545
   End
   Begin VB.TextBox txtQtde 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   4770
      MaxLength       =   6
      TabIndex        =   10
      Top             =   3435
      Width           =   945
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1695
      MaxLength       =   6
      TabIndex        =   4
      Top             =   150
      Width           =   945
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   2745
      TabIndex        =   5
      ToolTipText     =   "Consulta Imóvel"
      Top             =   120
      Width           =   465
      _ExtentX        =   820
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmEmissao2ViaEspecial.frx":0004
      PICN            =   "frmEmissao2ViaEspecial.frx":0020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   1980
      Left            =   15
      TabIndex        =   3
      Top             =   6330
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   3493
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Código     |^Ano     |^Lanc. |^Seq  |^Parc. |^Compl. |^Vencimento      |>Vl.Lançado  |<Num.Documento      |Stat"
   End
   Begin prjChameleon.chameleonButton cmdSelAll 
      Height          =   345
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Marcar/Desmarcar os lançamentos"
      Top             =   5550
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Marcar Todos"
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
      MICON           =   "frmEmissao2ViaEspecial.frx":017A
      PICN            =   "frmEmissao2ViaEspecial.frx":0196
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
      Height          =   345
      Left            =   6210
      TabIndex        =   0
      ToolTipText     =   "Sair da Tela"
      Top             =   5550
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
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
      MICON           =   "frmEmissao2ViaEspecial.frx":02F0
      PICN            =   "frmEmissao2ViaEspecial.frx":030C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   4785
      TabIndex        =   1
      ToolTipText     =   "Imprimir 1ª Via"
      Top             =   5535
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
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
      MICON           =   "frmEmissao2ViaEspecial.frx":037A
      PICN            =   "frmEmissao2ViaEspecial.frx":0396
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvDeb 
      Height          =   1575
      Left            =   90
      TabIndex        =   11
      Top             =   3810
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15658734
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lanc"
         Object.Width           =   3882
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
         SubItemIndex    =   7
         Text            =   "CodReduzido"
         Object.Width           =   2540
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdAddDoc 
      Height          =   315
      Left            =   2820
      TabIndex        =   7
      ToolTipText     =   "Adicionar número de documento"
      Top             =   2670
      Width           =   390
      _ExtentX        =   688
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmEmissao2ViaEspecial.frx":04F0
      PICN            =   "frmEmissao2ViaEspecial.frx":050C
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
      Caption         =   "(APENAS PARA LANCAM. DE ISS FIXO/TLL)"
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
      Index           =   16
      Left            =   3300
      TabIndex        =   46
      Top             =   2730
      Width           =   3930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Documento...:"
      Height          =   225
      Index           =   15
      Left            =   90
      TabIndex        =   45
      Top             =   2760
      Width           =   1470
   End
   Begin VB.Label lblTipoEnd 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2550
      TabIndex        =   44
      Top             =   5580
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Lançamento.:"
      Height          =   225
      Index           =   13
      Left            =   90
      TabIndex        =   43
      Top             =   3120
      Width           =   1470
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo...:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   42
      Top             =   3465
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Parcelas...:"
      Height          =   225
      Index           =   14
      Left            =   3240
      TabIndex        =   41
      Top             =   3465
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UF...:"
      Height          =   225
      Index           =   12
      Left            =   4290
      TabIndex        =   40
      Top             =   2265
      Width           =   390
   End
   Begin VB.Label lblUF 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4785
      TabIndex        =   39
      Top             =   2250
      Width           =   330
   End
   Begin VB.Label lblCepEntrega 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6030
      TabIndex        =   38
      Top             =   2250
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cep......:"
      Height          =   225
      Index           =   8
      Left            =   5340
      TabIndex        =   37
      Top             =   2265
      Width           =   585
   End
   Begin VB.Label lblCidadeEntrega 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1290
      TabIndex        =   36
      Top             =   2235
      Width           =   2730
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade...........:"
      Height          =   225
      Index           =   7
      Left            =   60
      TabIndex        =   35
      Top             =   2235
      Width           =   1155
   End
   Begin VB.Label lblBairroEntrega 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5010
      TabIndex        =   34
      Top             =   1950
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro...:"
      Height          =   225
      Index           =   5
      Left            =   4305
      TabIndex        =   33
      Top             =   1965
      Width           =   690
   End
   Begin VB.Label lblComplentrega 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1290
      TabIndex        =   32
      Top             =   1950
      Width           =   2730
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento.:"
      Height          =   225
      Index           =   4
      Left            =   60
      TabIndex        =   31
      Top             =   1950
      Width           =   1155
   End
   Begin VB.Label lblNumEntrega 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6765
      TabIndex        =   30
      Top             =   1650
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº...:"
      Height          =   225
      Index           =   3
      Left            =   6345
      TabIndex        =   29
      Top             =   1665
      Width           =   405
   End
   Begin VB.Label lblRuaEntrega 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1290
      TabIndex        =   28
      Top             =   1665
      Width           =   4860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço.......:"
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   27
      Top             =   1665
      Width           =   1155
   End
   Begin VB.Label lblBairro 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4035
      TabIndex        =   26
      Top             =   1050
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro..:"
      Height          =   225
      Index           =   11
      Left            =   3435
      TabIndex        =   25
      Top             =   1050
      Width           =   570
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento.............:"
      Height          =   225
      Index           =   10
      Left            =   30
      TabIndex        =   24
      Top             =   1035
      Width           =   1740
   End
   Begin VB.Label lblCompl 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1695
      TabIndex        =   23
      Top             =   1050
      Width           =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cep.:"
      Height          =   225
      Index           =   9
      Left            =   6045
      TabIndex        =   22
      Top             =   1065
      Width           =   420
   End
   Begin VB.Label lblCep 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6465
      TabIndex        =   21
      Top             =   1050
      Width           =   990
   End
   Begin VB.Label lblNumImovel 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6465
      TabIndex        =   20
      Top             =   765
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº...:"
      Height          =   225
      Index           =   1
      Left            =   6045
      TabIndex        =   19
      Top             =   780
      Width           =   405
   End
   Begin VB.Label lblNumInsc 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4695
      TabIndex        =   18
      Top             =   165
      Width           =   2790
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido/I.M.:"
      Height          =   225
      Index           =   0
      Left            =   45
      TabIndex        =   17
      Top             =   180
      Width           =   1695
   End
   Begin VB.Label lblRua 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1695
      TabIndex        =   16
      Top             =   765
      Width           =   3690
   End
   Begin VB.Label lblProp 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1695
      TabIndex        =   15
      Top             =   480
      Width           =   5790
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço...................:"
      Height          =   225
      Index           =   6
      Left            =   45
      TabIndex        =   14
      Top             =   735
      Width           =   1695
   End
   Begin VB.Label lblRS 
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietário.................:"
      Height          =   225
      Left            =   45
      TabIndex        =   13
      Top             =   465
      Width           =   1695
   End
   Begin VB.Label lblNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Insc.Cadastral:"
      Height          =   225
      Left            =   3555
      TabIndex        =   12
      Top             =   165
      Width           =   1200
   End
End
Attribute VB_Name = "frmEmissao2ViaEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuContrib As cPopupMenu
Attribute m_cMenuContrib.VB_VarHelpID = -1
'TIPOS
Private Type CARNE
    sDesc As String
    sUn As String
    nValor As Double
End Type
Private Type PROFUNDIDADE
    Distrito As Integer
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORPROFUN
    Distrito As Integer
    Codigo As Integer
    Fator As Double
End Type
Private Type GLEBA
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORCATEG
    Uso As Integer
    Tipo As Integer
    Categoria As Integer
    Fator As Double
End Type
Dim RdoAux4 As rdoResultset
'MATRIZES
Dim aCarne() As CARNE
Dim aFatorD() As Double
Dim aFatorD98() As Double
Dim aFatorP() As Double
Dim aFatorP98() As Double
Dim aFatorT() As Double
Dim aFatorT98() As Double
Dim aFatorS() As Double
Dim aFatorS98() As Double
Dim aFatorG() As Double
Dim aFatorG98() As Double
Dim aFatorR() As Double
Dim aFatorR98() As Double
Dim aProf() As PROFUNDIDADE
Dim aFatorF() As FATORPROFUN
Dim aFatorF98() As FATORPROFUN
Dim aFatorC() As FATORCATEG
Dim aFatorC98() As FATORCATEG
Dim aGleba() As GLEBA

Dim RdoAux As rdoResultset, Sql As String
Dim xImovel As clsImovel, bExec As Boolean
Dim z As Long, nTotalParc As Integer
Dim itmX As ListItem, bISS As Boolean, bTaxa As Boolean
Dim nTestada As Double, nAreaTotalTerreno As Double, nAreaConstruida As Double
Dim nVVT As Double, nVVC As Double, nVVI As Double, nAliq As Double, nValorFinal As Double

Private Sub cmbLanc_Click()
Dim nCodReduz As Long
If Not bExec Then Exit Sub
If txtCod.Text = "" Then Exit Sub
nCodReduz = CLng(txtCod.Text)
z = SendMessage(lvDeb.hwnd, LVM_DELETEALLITEMS, 0, 0)
cmbProc.Enabled = False
cmbProc.BackColor = Kde
txtNumDoc.Text = "": bISS = False: bTaxa = False
If cmbLanc.ItemData(cmbLanc.ListIndex) = 20 Then GoTo proximo


Sql = "SELECT DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,"
Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.CODCOMPLEMENTO,DATAAJUIZA, "
Sql = Sql & "DEBITOPARCELA.STATUSLANC, LANCAMENTO.DESCREDUZ,SITUACAOLANCAMENTO.DescSituacao,DATAVENCIMENTO,DATADEBASE "
Sql = Sql & "FROM DEBITOPARCELA INNER JOIN LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
Sql = Sql & "Inner Join SITUACAOLANCAMENTO ON DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO "
Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & "AND DEBITOPARCELA.CODLANCAMENTO=" & cmbLanc.ItemData(cmbLanc.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    nTotalParc = .RowCount
    txtQtde.Text = nTotalParc
   .Close
End With

Sql = "SELECT DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,"
Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.CODCOMPLEMENTO,DATAAJUIZA, "
Sql = Sql & "DEBITOPARCELA.STATUSLANC, LANCAMENTO.DESCREDUZ,SITUACAOLANCAMENTO.DescSituacao,DATAVENCIMENTO,DATADEBASE "
Sql = Sql & "FROM DEBITOPARCELA INNER JOIN LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
Sql = Sql & "Inner Join SITUACAOLANCAMENTO ON DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO "
Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & "AND DEBITOPARCELA.CODLANCAMENTO=" & cmbLanc.ItemData(cmbLanc.ListIndex) & "  AND STATUSLANC<4  ORDER BY DATAVENCIMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    If RdoAux.RowCount > 0 Then
        Do Until .EOF
           Set itmX = lvDeb.ListItems.Add(, Format(!CODREDUZIDO, "000000") & !AnoExercicio & Format(!CodLancamento, "00") & Format(!SeqLancamento, "00") & Format(!NumParcela, "00") & Format(!CODCOMPLEMENTO, "00"), !AnoExercicio)
           itmX.SubItems(1) = Format(!CodLancamento, "00") & "-" & !descreduz
           itmX.SubItems(2) = Format(!SeqLancamento, "00")
           itmX.SubItems(3) = Format(!NumParcela, "00")
           itmX.SubItems(4) = Format(!CODCOMPLEMENTO, "00")
           itmX.SubItems(5) = Format(!DataVencimento, "dd/mm/yyyy")
           itmX.SubItems(6) = !DescSituacao
          .MoveNext
        Loop
    Else
        MsgBox "Não existem Débitos com este Lançamento.", vbExclamation, "Atenção"
    End If
   .Close
End With
Exit Sub
proximo:
If cmbLanc.ItemData(cmbLanc.ListIndex) = 20 Then
    cmbProc.Enabled = True
    cmbProc.BackColor = Branco
    CarregaProc
Else
    cmbProc.Enabled = False
    cmbProc.BackColor = Kde
End If

End Sub

Private Sub cmbProc_Click()
Dim RdoAux2 As rdoResultset
Dim sprotocolo As String
Dim sDescSit As String
z = SendMessage(lvDeb.hwnd, LVM_DELETEALLITEMS, 0, 0)

If Right$(cmbProc.Text, 4) <> "SMAR" Then
    Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & cmbProc.Text & "' AND STATUSLANC<4"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Sql = "SELECT DESCSITUACAO FROM SITUACAOLANCAMENTO WHERE CODSITUACAO=" & !statuslanc
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            sDescSit = RdoAux2!DescSituacao
            RdoAux2.Close
            Set itmX = lvDeb.ListItems.Add(, Format(!CODREDUZIDO, "000000") & !AnoExercicio & Format(!CodLancamento, "00") & Format(!numsequencia, "00") & Format(!NumParcela, "00") & Format(!CODCOMPLEMENTO, "00"), !AnoExercicio)
            itmX.SubItems(1) = Format(!CodLancamento, "00") & "-" & cmbLanc.Text
            itmX.SubItems(2) = Format(!numsequencia, "00")
            itmX.SubItems(3) = Format(!NumParcela, "00")
            itmX.SubItems(4) = Format(!CODCOMPLEMENTO, "00")
            itmX.SubItems(5) = Format(!DataVencimento, "dd/mm/yyyy")
            itmX.SubItems(6) = sDescSit
           .MoveNext
        Loop
    End With
Else
    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & Left$(cmbProc.Text, Len(cmbProc) - 5) & " AND (STATUSLANC=1 or STATUSLANC=2 or STATUSLANC=3 or STATUSLANC=7)  "
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Sql = "SELECT DESCSITUACAO FROM SITUACAOLANCAMENTO WHERE CODSITUACAO=" & !statuslanc
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            sDescSit = RdoAux2!DescSituacao
            RdoAux2.Close
            Set itmX = lvDeb.ListItems.Add(, Format(!CODREDUZIDO, "000000") & !AnoExercicio & Format(!CodLancamento, "00") & Format(!SeqLancamento, "00") & Format(!NumParcela, "00") & Format(!CODCOMPLEMENTO, "00"), !AnoExercicio)
            itmX.SubItems(1) = Format(!CodLancamento, "00") & "-" & cmbLanc.Text
            itmX.SubItems(2) = Format(!SeqLancamento, "00")
            itmX.SubItems(3) = Format(!NumParcela, "00")
            itmX.SubItems(4) = Format(!CODCOMPLEMENTO, "00")
            itmX.SubItems(5) = Format(!DataVencimento, "dd/mm/yyyy")
            itmX.SubItems(6) = sDescSit
          .MoveNext
        Loop
       .Close
    End With
End If

End Sub

Private Sub cmdAddDoc_Click()
Dim sDoc As String, nDV As Single

MsgBox "Esta opção foi desativada, para emitir segunda via de Taxa de Licença/ISS por favor utilize a opção DAM.", vbInformation, "Atenção"
Exit Sub

bISS = True
bTaxa = True
'If Val(txtCod.text) = 0 Then
'    MsgBox "Selecione o contribuinte", vbExclamation, "Atenção"
'    Exit Sub
'End If
If Val(txtNumDoc.Text) = 0 Then
    MsgBox "Digite o número do documento", vbExclamation, "Atenção"
    Exit Sub
End If
sDoc = CStr(Left$(txtNumDoc.Text, Len(txtNumDoc.Text) - 1))
nDV = Val(Right$(txtNumDoc.Text, 1))
If nDV <> RetornaDVNumDoc(CLng(sDoc)) Then
    MsgBox "Digito verificador inválido.", vbExclamation, "Atenção"
    Exit Sub
End If
'z = SendMessage(lvDeb.hwnd, LVM_DELETEALLITEMS, 0, 0)
Sql = "SELECT parceladocumento.codreduzido, parceladocumento.numdocumento, debitoparcela.anoexercicio, debitoparcela.codlancamento, "
Sql = Sql & "debitoparcela.seqlancamento, debitoparcela.numparcela, debitoparcela.codcomplemento, debitoparcela.statuslanc, debitoparcela.datavencimento,"
Sql = Sql & "situacaolancamento.DescSituacao , lancamento.descreduz FROM parceladocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND "
Sql = Sql & "parceladocumento.anoexercicio = debitoparcela.anoexercicio AND parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND "
Sql = Sql & "parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN situacaolancamento ON debitoparcela.statuslanc = situacaolancamento.codsituacao INNER JOIN "
Sql = Sql & "lancamento ON debitoparcela.codlancamento = lancamento.codlancamento  Where parceladocumento.NumDocumento = " & Val(sDoc)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvDeb.ListItems.Add(, Format(!CODREDUZIDO, "000000") & !AnoExercicio & Format(2, "00") & Format(!SeqLancamento, "00") & Format(!NumParcela, "00") & Format(!CODCOMPLEMENTO, "00") & CStr(.AbsolutePosition), !AnoExercicio)
        itmX.SubItems(1) = !CodLancamento & " - " & !descreduz
        itmX.SubItems(2) = Format(!SeqLancamento, "00")
        itmX.SubItems(3) = Format(!NumParcela, "00")
        itmX.SubItems(4) = Format(!CODCOMPLEMENTO, "00")
        itmX.SubItems(5) = Format(!DataVencimento, "dd/mm/yyyy")
        itmX.SubItems(6) = !DescSituacao
        itmX.SubItems(7) = !CODREDUZIDO
        .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdCnsImovel_Click()
lIndex = m_cMenuContrib.ShowPopupMenu(cmdCnsImovel.Left, cmdCnsImovel.Top, cmdCnsImovel.Left, cmdCnsImovel.Top, Me.ScaleWidth - cmdCnsImovel.Left - cmdCnsImovel.Width, cmdCnsImovel.Top + cmdCnsImovel.Height, False)
End Sub

Private Sub cmdPrint_Click()
Dim x As Integer, bAchou As Boolean

MsgBox "Bloqueado"
Exit Sub



If lvDeb.ListItems.Count = 0 Then
    MsgBox "Não existe nada a imprimir", vbCritical, "atenção"
    Exit Sub
End If

bAchou = False
For x = 1 To lvDeb.ListItems.Count
    If lvDeb.ListItems(x).Checked = True Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    MsgBox "Selecione algum lançamento", vbCritical, "atenção"
    Exit Sub
End If

'GravaCarneTmp

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSelAll_Click()

If lvDeb.ListItems.Count = 0 Then
    Exit Sub
End If

If Left$(cmdSelAll.Caption, 1) = "M" Then
    cmdSelAll.Caption = "Desmarcar Todos"
    For x = 1 To lvDeb.ListItems.Count
        lvDeb.ListItems(x).Checked = True
    Next
Else
    cmdSelAll.Caption = "Marcar Todos"
    For x = 1 To lvDeb.ListItems.Count
        lvDeb.ListItems(x).Checked = False
    Next
End If

End Sub

Private Sub Form_Activate()
If Val(CodImovel) > 0 Then
     txtCod.Text = Val(Left$(CodImovel, 7))
     CodImovel = 0
     txtCod_LostFocus
Else
    If Val(CodEmpresa) > 0 Then
         txtCod.Text = Val(Left$(CodEmpresa, 7))
         CodEmpresa = 0
         txtCod_LostFocus
    Else
        If Val(CodCidadao) > 0 Then
             Unload frmCnsCidadao
             If cGetInputState() <> 0 Then DoEvents
             txtCod.Text = Val(CodCidadao)
             CodCidadao = 0
             txtCod_LostFocus
        End If
    End If
End If

End Sub

Private Sub Form_Load()
Centraliza Me
MontaMenu
Set xImovel = New clsImovel
bExec = False
Sql = "SELECT CODLANCAMENTO,DESCFULL FROM LANCAMENTO ORDER BY DESCFULL"
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
    Do Until .EOF
        cmbLanc.AddItem !DESCFULL
        cmbLanc.ItemData(cmbLanc.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With
bExec = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
Set m_cMenuContrib = Nothing
End Sub

Private Sub CarregaImovel(nCodigoImovel As Long)
Dim Sql As String, RdoAux As rdoResultset

Ocupado
With xImovel
    .CarregaImovel nCodigoImovel
    If .CodigoImovel > 0 Then
          lblNumInsc.Caption = .Inscricao
          lblProp.Caption = .NomePropPrincipal
          lblRua.Caption = Trim$(.AbrevTipoLog) & " " & Trim$(.AbrevTitLog) & " " & .NomeLogradouro
          lblNumImovel.Caption = .Li_Num
          lblCep.Caption = RetornaCEP(.CodLogr, .Li_Num)
          lblCompl.Caption = .Li_Compl
          lblBairro.Caption = .DescBairro
          Select Case .Ee_TipoEnd
                Case 0
                    lblTipoEnd.Caption = "(Endereço do Imóvel)"
                    lblRuaEntrega.Caption = lblRua.Caption
                    lblNumEntrega.Caption = lblNumImovel.Caption
                    lblComplentrega.Caption = lblCompl.Caption
                    lblBairroEntrega.Caption = lblBairro.Caption
                    lblCidadeEntrega.Caption = "JABOTICABAL"
                    lblCepEntrega.Caption = lblCep.Caption
                    lblUF.Caption = lblUF.Caption
                Case 1
                    lblTipoEnd.Caption = "(Endereço do Proprietário)"
                    CarregaEndCidadao .CodPropPrincipal
                Case 2
                    lblTipoEnd.Caption = "(Endereço de Entrega Específico)"
                    lblRuaEntrega.Caption = .Ee_NomeLog
                    lblNumEntrega.Caption = .Ee_NumImovel
                    lblComplentrega.Caption = .Ee_Complemento
                    Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & .Ee_Uf & "' AND CODCIDADE=" & .Ee_Cidade & " AND CODBAIRRO=" & .Ee_Bairro
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            lblBairroEntrega.Caption = !DescBairro
                        End If
                       .Close
                    End With
                    lblCidadeEntrega.Caption = .Ee_Cidade
                    lblCepEntrega.Caption = .Ee_Cep
                    lblUF.Caption = .Ee_Uf
          End Select
    End If
End With

fim:
Liberado

End Sub

Private Sub Limpa()

lblNum.Caption = ""
lblProp.Caption = ""
lblRua.Caption = ""
lblNumImovel.Caption = ""
lblCompl.Caption = ""
lblBairro.Caption = ""
lblCep.Caption = ""
lblRuaEntrega.Caption = ""
lblNumEntrega.Caption = ""
lblComplentrega.Caption = ""
lblBairroEntrega.Caption = ""
lblCidadeEntrega.Caption = ""
lblCepEntrega.Caption = ""
lblUF.Caption = ""
lblNumInsc.Caption = ""
lblTipoEnd.Caption = ""
cmbProc.Clear
cmbProc.Enabled = False
cmbProc.BackColor = Kde
bExec = False: cmbLanc.ListIndex = -1: bExec = True
End Sub

Private Sub CarregaEndCidadao(nCodigo As Long)

Sql = "SELECT CIDADAO.CODCIDADAO,vwLOGRADOUROCEP.ABREVTIPOLOG,vwLOGRADOUROCEP.ABREVTITLOG,"
Sql = Sql & "vwLOGRADOUROCEP.NOMELOGRADOURO,vwLOGRADOUROCEP.CEP, CIDADAO.NUMIMOVEL,"
Sql = Sql & "CIDADAO.COMPLEMENTO, CIDADAO.CODBAIRRO,CIDADAO.CODCIDADE, CIDADAO.SIGLAUF,"
Sql = Sql & "Cidade.DESCCIDADE , BAIRRO.DescBairro FROM CIDADAO INNER JOIN vwLOGRADOUROCEP ON "
Sql = Sql & "CIDADAO.CODLOGRADOURO = vwLOGRADOUROCEP.CODLOGRADOURO Inner Join BAIRRO ON CIDADAO.SIGLAUF = BAIRRO.SIGLAUF AND "
Sql = Sql & "CIDADAO.CODCIDADE = BAIRRO.CODCIDADE AND CIDADAO.CODBAIRRO = BAIRRO.CODBAIRRO INNER JOIN "
Sql = Sql & "CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = Cidade.CODCIDADE "
Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & nCodigo
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    lblRuaEntrega.Caption = SubNull(!NomeLogradouro)
    lblNumEntrega.Caption = SubNull(!NUMIMOVEL)
    lblComplentrega.Caption = SubNull(!Complemento)
    lblBairroEntrega.Caption = SubNull(!DescBairro)
    lblCidadeEntrega.Caption = SubNull(!desccidade)
    lblCepEntrega.Caption = SubNull(!Cep)
    lblUF.Caption = SubNull(!Cep)
End With

End Sub

Private Sub m_cMenuContrib_Click(ItemNumber As Long)
Select Case m_cMenuContrib.ItemKey(ItemNumber)
    Case "mnuMob"
        sFormMob = "2VIAE"
        frmCnsMob.show
        frmCnsMob.ZOrder 0
    Case "mnuImob"
        sForm = "2VIAE"
        frmCnsImovel.show
        frmCnsImovel.ZOrder 0
    Case "mnuOutros"
        Set frm = frmCnsCidadao
        frm.sForm = "2VIAE"
        frm.show
        frm.ZOrder 0
End Select

End Sub

Private Sub txtCod_GotFocus()

txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    txtCod_LostFocus
    Exit Sub
End If

Tweak txtCod, KeyAscii, IntegerPositive

End Sub

Private Sub txtCod_LostFocus()
Dim nCodImovel As Long, sNomeCidade As String
z = SendMessage(lvDeb.hwnd, LVM_DELETEALLITEMS, 0, 0)
If Val(txtCod.Text) = 0 Then Exit Sub
nCodImovel = Val(txtCod.Text)
Limpa
Sql = "SELECT CODREDUZIDO,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & txtCod.Text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If !Inativo = 1 Then
           MsgBox "Este imóvel encontra-se inativo.", vbExclamation, "Atenção"
           Exit Sub
        End If
        lblNum.Caption = "Insc.Cadastral"
        lblRS.Caption = "Proprietário"
        CarregaImovel nCodImovel
    Else
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO,NOMELOGR,CODCIDADE,DESCCIDADE "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & txtCod.Text
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               If Not IsNull(!dataencerramento) Or !dataencerramento <> CDate("01/01/1900") Then
                  MsgBox "Esta empresa foi encerrada em " & Format(!dataencerramento, "dd/mm/yyyy"), vbExclamation, "Atenção"
'                  Exit Sub
               End If
              'suspenção
               Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & txtCod.Text
               Sql = Sql & " ORDER BY DATAEVENTO DESC"
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                       If !CODTIPOEVENTO = 2 Then
                           MsgBox "Esta empresa esta SUSPENSA", vbExclamation, "Atenção"
                       End If
                   End If
                  .Close
               End With
               
               lblNum.Caption = "Insc.Estadual"
               lblNumInsc.Caption = SubNull(!INSCESTADUAL)
               lblRS.Caption = "Raz.Social"
               lblProp.Caption = !RazaoSocial
               If !CodCidade = 413 Then
                  lblRua.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
               Else
                  sNomeCidade = SubNull(!desccidade)
                  lblRua.Caption = SubNull(!NomeLogr)
               End If
               lblNumImovel.Caption = Val(SubNull(!Numero))
               lblCep.Caption = IIf(IsNull(!Cep), "", Left$(!Cep, 5) & "-" & Right$(!Cep, 3))
               lblCompl.Caption = SubNull(!Complemento)
               Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO=" & !CodBairro
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                        lblBairro.Caption = !DescBairro
                   Else
                        lblBairro.Caption = ""
                   End If
                  .Close
               End With
               Sql = "SELECT NOMELOGRADOURO,NUMIMOVEL,COMPLEMENTO,UF,CIDADE.DESCCIDADE AS DESCCIDADE1,"
               Sql = Sql & "BAIRRO.DESCBAIRRO AS DESCBAIRRO1,CEP,MOBILIARIOENDENTREGA.DESCBAIRRO,"
               Sql = Sql & "MOBILIARIOENDENTREGA.DESCCIDADE FROM CIDADE INNER JOIN BAIRRO ON "
               Sql = Sql & "CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE RIGHT OUTER Join "
               Sql = Sql & "MOBILIARIOENDENTREGA ON BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND "
               Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & Val(txtCod.Text)
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                    If .RowCount > 0 Then
                        lblTipoEnd.Caption = "(Endereço de Entrega Específico)"
                        lblRuaEntrega.Caption = SubNull(!NomeLogradouro)
                        lblNumEntrega.Caption = SubNull(!NUMIMOVEL)
                        lblComplentrega.Caption = SubNull(!Complemento)
                        lblBairroEntrega.Caption = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        lblCidadeEntrega.Caption = IIf(IsNull(!desccidade), SubNull(!DESCCIDADE1), SubNull(!desccidade))
                        lblCepEntrega.Caption = SubNull(!Cep)
                        lblUF.Caption = SubNull(!UF)
                    Else
                        lblTipoEnd.Caption = "(Endereço da Empresa)"
                        lblRuaEntrega.Caption = lblRua.Caption
                        lblNumEntrega.Caption = lblNumImovel.Caption
                        lblComplentrega.Caption = lblCompl.Caption
                        lblBairroEntrega.Caption = lblBairro.Caption
                        lblCidadeEntrega.Caption = sNomeCidade
                        lblCepEntrega.Caption = lblCep.Caption
                        lblUF.Caption = "SP"
                    End If
                   .Close
               End With
            Else
               Sql = "SELECT CIDADAO.CODCIDADAO,CIDADAO.NOMECIDADAO,CIDADAO.CPF, CIDADAO.CNPJ, CIDADAO.CODLOGRADOURO,vwLOGRADOURO.ABREVTIPOLOG,"
               Sql = Sql & "vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,CIDADAO.NUMIMOVEL, CIDADAO.COMPLEMENTO,CIDADAO.CODBAIRRO, BAIRRO.DESCBAIRRO,"
               Sql = Sql & "CIDADAO.CODCIDADE, CIDADE.DESCCIDADE,CIDADAO.SIGLAUF, UF.DESCUF, CIDADAO.CEP,CIDADAO.NOMELOGRADOURO AS RUA2 "
               Sql = Sql & "FROM vwLOGRADOURO RIGHT OUTER JOIN CIDADAO ON vwLOGRADOURO.CODLOGRADOURO = CIDADAO.CODLOGRADOURO "
               Sql = Sql & "LEFT OUTER JOIN CIDADE INNER JOIN BAIRRO ON CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE INNER JOIN "
               Sql = Sql & "UF ON CIDADE.SIGLAUF = UF.SIGLAUF ON CIDADAO.SIGLAUF = BAIRRO.SIGLAUF AND CIDADAO.CODCIDADE = BAIRRO.CODCIDADE AND CIDADAO.CODBAIRRO = BAIRRO.CODBAIRRO WHERE CODCIDADAO=" & Val(txtCod.Text)
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux
                   If .RowCount > 0 Then
                       lblProp.Caption = !nomecidadao
                       If !CodLogradouro > 0 Then
                          lblRua.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                       Else
                          lblRua.Caption = SubNull(!RUA2)
                       End If
                       lblNumImovel.Caption = Val(SubNull(!NUMIMOVEL))
                       If IsNull(!Cep) Then
                           lblCep = "00000-000"
                       Else
                          lblCep.Caption = IIf(IsNull(!Cep), "", Left$(!Cep, 5) & "-" & Right$(!Cep, 3))
                       End If
                       lblCompl.Caption = SubNull(!Complemento)
                       lblBairro.Caption = SubNull(!DescBairro)
                       lblNumImovel.Caption = Val(SubNull(!NUMIMOVEL))
                       lblCompl.Caption = SubNull(!Complemento)
                       'lblBairro.Caption = IIf(IsNull(!DescBairro), SubNull(!NOMEBairro), !DescBairro)
                   
                       lblRuaEntrega.Caption = lblRua.Caption
                       lblNumEntrega.Caption = lblNumImovel.Caption
                       lblComplentrega.Caption = lblCompl.Caption
                       lblBairroEntrega.Caption = lblBairro.Caption
                       lblCidadeEntrega.Caption = SubNull(!desccidade)
                       lblCepEntrega.Caption = lblCep.Caption
                       lblUF.Caption = SubNull(!SiglaUF)
                   Else
                       MsgBox "Código não cadastrado.", vbCritical, "Atenção"
                   End If
                  .Close
               End With
            End If
           .Close
        End With
    End If
End With
End Sub

Private Sub CarregaProc()
    
cmbProc.Clear
Sql = "SELECT DISTINCT NUMPROCESSO FROM ORIGEMREPARC WHERE "
Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      Do Until .EOF
            cmbProc.AddItem !NUMPROCESSO
           .MoveNext
      Loop
     .Close
End With

Sql = "SELECT DISTINCT(CODSEQD) From REPARCTMP Where CODREDUZD =" & Val(txtCod.Text) & " Or CODREDUZO = " & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
'            MsgBox "Não existem reparcelamentos.", vbExclamation, "atenção"
'            Exit Sub
    Else
        Do Until .EOF
           cmbProc.AddItem CStr(!CODSEQD) & "/SMAR"
          .MoveNext
        Loop
        cmbProc.ListIndex = 0
    End If
   .Close
End With

    
If cmbProc.ListCount = 0 Then
    MsgBox "Não existem reparcelamentos.", vbExclamation, "atenção"
    cmbProc.Enabled = False
    cmbProc.BackColor = Kde
    Exit Sub
End If
End Sub

Private Sub CalculoIndividual(nCodReduz As Long)
On Error GoTo Erro
Dim nSomaTestada As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer
Dim nAliquotaPredial As Double
Dim nAliquotaTerritorial As Double
Dim bTemPredial As Boolean
Dim bFracaoIdeal As Boolean
Dim nAreaTerreno As Double
Dim nAreaPrincipal As Double
Dim nCodAgrupamento As Integer
Dim nValorAgrupamento As Double
Dim nValorAgrupamento98 As Double
Dim nNumTestadas As Integer
Dim nTestadaPrincipal As Double
Dim nCodGleba As Integer
Dim nFatorGleba As Double
Dim nFatorGleba98 As Double
Dim nCodProfundidade As Integer
Dim nValorProfundidade As Double
Dim nFatorProfundidade As Double
Dim nFatorProfundidade98 As Double
Dim nCodSituacao As Integer
Dim nFatorSituacao As Double
Dim nFatorSituacao98 As Double
Dim nCodPedologia As Integer
Dim nFatorPedologia As Double
Dim nFatorPedologia98 As Double
Dim nCodTopografia As Integer
Dim nFatorTopografia As Double
Dim nFatorTopografia98 As Double
Dim nFatorDistrito As Double
Dim nFatorDistrito98 As Double
Dim nValorFatores As Double
Dim nValorFatores98 As Double
Dim nFatorCategoria As Double
Dim nFatorCategoria98 As Double
Dim nValorVenalTerritorial As Double
Dim nValorVenalTerritorial98 As Double
Dim nValorVenalPredial As Double
Dim nValorVenalPredial98 As Double
Dim nCodTributo As Integer
Dim nValorVenalImovel As Double
Dim nValorVenalImovel98 As Double
Dim nValorIptu As Double, nValorITU As Double
Dim nValorIPTU98 As Double, nValorITU98 As Double
Dim nTaxaLimpeza As Double, nTaxaConservacao As Double

nAliquotaPredial = 1.5
nAliquotaTerritorial = 3
nAnoCalculo = Year(Now)
'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = !Li_CodBairro
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nTestadaPrincipal = 1
            nTestada1 = 1
        Else
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                nTestadaPrincipal = nSomaTestada / nNumTestadas
            End If
        End If
       .Close
    End With
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    'BUSCA ÁREA PRINCIPAL
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P' AND YEAR(DATAAPROVA) < " & Year(Now)
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & " AND YEAR(DATAAPROVA) < " & Year(Now)
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 0 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) And RdoAux2!QTDEPAV < 2 And nAreaTerreno < 600 Then
                        bIsento = True
                        MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                        Limpa
                    End If
                End If
               .Close
            End With
        Else
            bIsento = False
        End If
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
        If bTemPredial Then
             nUso = !USOCONSTR
             nTipo = !TIPOCONSTR
             nCat = !CATCONSTR
             Select Case !USOCONSTR
                  Case 0
                     nTaxaLimpeza = 3.78
                  Case 1, 2, 3, 4, 5
                     nTaxaLimpeza = 10.57
                  Case Else
                     nTaxaLimpeza = 3.01
             End Select
        Else
             nTaxaLimpeza = 3.01
        End If
        nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
       '--CÁLCULO DA TAXA DE CONSERVAÇÃO
        If RdoAux!PAVIMENTO = 1 Then
           nTaxaConservacao = 1.35 * nTestadaPrincipal
        Else
           nTaxaConservacao = 0
        End If
        If nCodBairro = 81 Then
           nTaxaLimpeza = 1
           nTaxaConservacao = 1
        End If
       .Close
    End With
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
       nValorAgrupamento98 = aFatorR98(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
       nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
    End If
    
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
    For x = 1 To UBound(aGleba)
        If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
             Exit For
        ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
             Exit For
        End If
    Next
    nCodGleba = aGleba(x).Codigo
    'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
    nFatorGleba = aFatorG(nCodGleba)
    'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
    nFatorGleba98 = aFatorG98(nCodGleba)
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= aProf(x).Min And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
        nFatorProfundidade98 = 0
        For x = 1 To UBound(aFatorF98)
            If aFatorF98(x).Distrito = !Distrito And aFatorF98(x).Codigo = nCodProfundidade Then
               nFatorProfundidade98 = aFatorF98(x).Fator
               Exit For
            End If
        Next
     Else
        nFatorProfundidade = 1
        nFatorProfundidade98 = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    'FATOR SITUACAO 98
    nFatorSituacao98 = aFatorS98(nCodSituacao)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    'FATOR PEDOLOGIA 98
    nFatorPedologia98 = aFatorP98(nCodPedologia)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    'FATOR TOPOGRAFIA 98
    nFatorTopografia98 = aFatorT98(nCodTopografia)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
    nValorFatores98 = nFatorTopografia98 * nFatorSituacao98 * nFatorPedologia98 * nFatorProfundidade98 * nFatorGleba98
    'CÁLCULO VALOR VENAL TERRITORIAL
    nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores
    nValorVenalTerritorial98 = nAreaTerreno * nValorAgrupamento98 * nValorFatores98
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        'FATOR DISTRITO 98
        nFatorDistrito98 = aFatorD98(!Distrito)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        nValorVenalPredial = 0
        nValorVenalPredial98 = 0
        nFatorCategoria = 0
        For x = 1 To UBound(aFatorC)
            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
               nFatorCategoria = aFatorC(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
        
       'FATOR CATEGORIA 98
        nFatorCategoria98 = 0
        For x = 1 To UBound(aFatorC98)
            If aFatorC98(x).Uso = nUso And aFatorC98(x).Tipo = nTipo And aFatorC98(x).Categoria = nCat Then
               nFatorCategoria98 = aFatorC98(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial98 = nValorVenalPredial98 + (nAreaPrincipal * nFatorCategoria98)
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        nValorVenalPredial98 = nValorVenalPredial98 * nFatorDistrito98
    Else
        nFatorDistrito = 0
        nFatorDistrito98 = 0
        nFatorCategoria = 0
        nFatorCategoria98 = 0
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nCodTributo = 1
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
        nValorVenalImovel98 = nValorVenalTerritorial98 + nValorVenalPredial98
        'nValorIPTU = nValorVenalImovel * (nAliquotaPredial / 100) * 1.062 'reajuste 2004-2005
        nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)  'reajuste 2004-2005
        nValorIPTU98 = nValorVenalImovel98 * (nAliquotaPredial / 100)
        nValorIPTU98 = nValorIPTU98 + nTaxaConservacao + nTaxaLimpeza
        nValorIPTU98 = nValorIPTU98 * 1.7947
    Else
        nCodTributo = 2
        nValorVenalImovel = nValorVenalTerritorial
        nValorVenalImovel98 = nValorVenalTerritorial98
        'nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100) * 1.062 'reajuste 2004-2005
        nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)  'reajuste 2004-2005
        nValorITU98 = nValorVenalImovel98 * (nAliquotaTerritorial / 100)
        nValorITU98 = nValorITU98 + nTaxaConservacao + nTaxaLimpeza
        nValorITU98 = nValorITU98 * 1.7947
    End If
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    If bTemPredial Then
'        If nValorIPTU98 > nValorIPTU Then
           nValorFinal = nValorIptu
'        Else
'           nValorFinal = nValorIPTU98
'        End If
    Else
 '       If nValorITU98 > nValorITU Then
           nValorFinal = nValorITU
 '       Else
 '          nValorFinal = nValorITU98
 '       End If
    End If
End With
nValorParc = nValorFinal
nTestada = nTestadaPrincipal
nAreaTotalTerreno = nAreaTerreno
nAreaConstruida = nAreaPrincipal
If bTemPredial Then
'    If nValorIPTU98 < nValorIPTU Then
'        nVVT = nValorVenalTerritorial98
'        nVVC = nValorVenalPredial98
'        nVVI = nValorVenalImovel98
'    Else
        nVVT = nValorVenalTerritorial
        nVVC = nValorVenalPredial
        nVVI = nValorVenalImovel
'    End If
Else
 '   If nValorITU98 < nValorITU Then
 '       nVVT = nValorVenalTerritorial98
 '       nVVC = nValorVenalPredial98
 '       nVVI = nValorVenalImovel98
 '   Else
        nVVT = nValorVenalTerritorial
        nVVC = nValorVenalPredial
        nVVI = nValorVenalImovel
 '   End If
End If

Exit Sub
Erro:
Resume Next

End Sub

Private Sub LoadMatrix()

ReDim aFatorD(3)
ReDim aFatorD98(3)
ReDim aFatorP(6)
ReDim aFatorP98(6)
ReDim aFatorT(6)
ReDim aFatorT98(6)
ReDim aFatorS(6)
ReDim aFatorS98(6)
ReDim aFatorG(23)
ReDim aFatorG98(23)
ReDim aFatorR(8)
ReDim aFatorR98(8)

On Error Resume Next
RdoAux4.Close
On Error GoTo 0

nAnoCalculo = Year(Now)
Sql = "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & nAnoCalculo & " ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA= 1998 ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG= 1998 ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO= 1998 ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA= 1998 ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO= 1998 ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR=" & nAnoCalculo & "  AND  CODMOEDA=1; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR= 1998  AND  CODMOEDA=1"
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
     Do Until .EOF
        aFatorP(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorP98(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT98(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS98(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG98(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD98(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR98(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .Close
End With



ReDim aProf(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN FROM PROFUNDIDADE ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
     Do Until .EOF
        ReDim Preserve aProf(UBound(aProf) + 1)
        aProf(UBound(aProf)).Distrito = !CODDISTRITO
        aProf(UBound(aProf)).Codigo = !CODPROFUN
        aProf(UBound(aProf)).Min = !MINPROFUN
        aProf(UBound(aProf)).Max = !MAXPROFUN
       .MoveNext
     Loop
    .Close
End With


ReDim aFatorF(0)
ReDim aFatorF98(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN; " & _
      "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN= 1998 ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
     Do Until .EOF
        ReDim Preserve aFatorF(UBound(aFatorF) + 1)
        aFatorF(UBound(aFatorF)).Distrito = !CODDISTRITO
        aFatorF(UBound(aFatorF)).Codigo = !CODPROFUN
        aFatorF(UBound(aFatorF)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        ReDim Preserve aFatorF98(UBound(aFatorF98) + 1)
        aFatorF98(UBound(aFatorF98)).Distrito = !CODDISTRITO
        aFatorF98(UBound(aFatorF98)).Codigo = !CODPROFUN
        aFatorF98(UBound(aFatorF98)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .Close
End With

ReDim aGleba(0)
Sql = "SELECT CODGLEBA,MINGLEBA,MAXGLEBA FROM GLEBA ORDER BY CODGLEBA "
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
     Do Until .EOF
        ReDim Preserve aGleba(UBound(aGleba) + 1)
        aGleba(UBound(aGleba)).Codigo = !CODGLEBA
        aGleba(UBound(aGleba)).Min = !MINGLEBA
        aGleba(UBound(aGleba)).Max = !MAXGLEBA
       .MoveNext
     Loop
    .Close
End With

ReDim aFatorC(0)
ReDim aFatorC98(0)
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1; " & _
      "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=1998 AND CODMOEDA=1 "
Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux4
     Do Until .EOF
        ReDim Preserve aFatorC(UBound(aFatorC) + 1)
        aFatorC(UBound(aFatorC)).Uso = !CODUSO
        aFatorC(UBound(aFatorC)).Tipo = !CodTipo
        aFatorC(UBound(aFatorC)).Categoria = !CODCATEG
        aFatorC(UBound(aFatorC)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        ReDim Preserve aFatorC98(UBound(aFatorC98) + 1)
        aFatorC98(UBound(aFatorC98)).Uso = !CODUSO
        aFatorC98(UBound(aFatorC98)).Tipo = !CodTipo
        aFatorC98(UBound(aFatorC98)).Categoria = !CODCATEG
        aFatorC98(UBound(aFatorC98)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
Tweak txtNumDoc, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtde_KeyPress(KeyAscii As Integer)
Tweak txtQtde, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtde_LostFocus()
nTotalParc = Val(txtQtde.Text)
End Sub

Private Sub MontaMenu()

   Set m_cMenuContrib = New cPopupMenu
   With m_cMenuContrib
      .hwndOwner = Me.hwnd
      .GradientHighlight = True
      
      i = .AddItem("Mobiliário", "", 1, , , , , "mnuMob")
      .OwnerDraw(i) = True
      i = .AddItem("Imobiliário", "", 1, , , , , "mnuImob")
      .OwnerDraw(i) = True
      i = .AddItem("Outros", "", 1, , , , , "mnuOutros")
      .OwnerDraw(i) = True
   End With
   
End Sub

